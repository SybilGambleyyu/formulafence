from __future__ import annotations

import json
import struct
from dataclasses import replace
from zipfile import ZIP_DEFLATED, ZipFile

import pytest

from formulafence.cli import main
from formulafence.diff import analyze_downstream_impact, compare_snapshots
from formulafence.formulas import has_broken_reference, inspect_formula
from formulafence.lint import lint_snapshot
from formulafence.models import FormulaFenceError, WorkbookLoadError
from formulafence.output import profile_to_markdown
from formulafence.workbook import load_snapshot, profile_snapshot
from formulafence.xlsb import (
    XlsbFormulaError,
    XlsbFormulaUnsupportedError,
    XlsbLimitError,
    XlsbParseError,
    XlsbReaderLimits,
    decode_xlsb_formula,
    parse_xlsb_workbook_parts,
    xlsb_records_from_bytes,
)


def _variable(value: int, *, max_bytes: int) -> bytes:
    if value < 0:
        raise ValueError("XLSB variable integer must not be negative")
    pieces: list[int] = []
    while True:
        piece = value & 0x7F
        value >>= 7
        if value:
            piece |= 0x80
        pieces.append(piece)
        if not value:
            break
    if len(pieces) > max_bytes:
        raise ValueError("XLSB variable integer does not fit the requested header")
    return bytes(pieces)


def _record(record_type: int, payload: bytes) -> bytes:
    return _variable(record_type, max_bytes=2) + _variable(len(payload), max_bytes=4) + payload


def _reference(
    row: int,
    column: int,
    *,
    row_relative: bool = True,
    column_relative: bool = True,
) -> bytes:
    flags = column
    if column_relative:
        flags |= 0x4000
    if row_relative:
        flags |= 0x8000
    return struct.pack("<IH", row, flags)


def _area(
    first_row: int,
    last_row: int,
    first_column: int,
    last_column: int,
    *,
    first_row_relative: bool = True,
    last_row_relative: bool = True,
    first_column_relative: bool = True,
    last_column_relative: bool = True,
) -> bytes:
    first_flags = first_column
    last_flags = last_column
    if first_column_relative:
        first_flags |= 0x4000
    if last_column_relative:
        last_flags |= 0x4000
    if first_row_relative:
        first_flags |= 0x8000
    if last_row_relative:
        last_flags |= 0x8000
    return struct.pack("<IIHH", first_row, last_row, first_flags, last_flags)


def _wide(value: str) -> bytes:
    encoded = value.encode("utf-16le")
    return struct.pack("<I", len(encoded) // 2) + encoded


def _cell_header(column: int) -> bytes:
    return struct.pack("<II", column, 0)


def _formula_num_cell(column: int, cached_value: float, tokens: bytes) -> bytes:
    return (
        _cell_header(column)
        + struct.pack("<d", cached_value)
        + b"\x00\x00"
        + struct.pack("<I", len(tokens))
        + tokens
    )


def _bundle_sheet(state: int, tab_id: int, relationship_id: str, title: str) -> bytes:
    return struct.pack("<II", state, tab_id) + _wide(relationship_id) + _wide(title)


def _defined_name(name: str, tokens: bytes, *, scope: int | None = None) -> bytes:
    scope_value = 0xFFFF_FFFF if scope is None else scope
    return (
        struct.pack("<IBI", 0, 0, scope_value)
        + _wide(name)
        + struct.pack("<I", len(tokens))
        + tokens
    )


def _relationships(*targets: str) -> bytes:
    if not targets:
        targets = ("worksheets/sheet1.bin",)
    relationship_entries = b"".join(
        (
            b'<Relationship Id="rId'
            + str(index).encode("ascii")
            + (
                b'" Type="http://schemas.openxmlformats.org/officeDocument/2006/'
                b'relationships/worksheet" '
            )
            + b'Target="'
            + target.encode("ascii")
            + b'"/>'
        )
        for index, target in enumerate(targets, start=1)
    )
    return (
        b'<?xml version="1.0" encoding="UTF-8"?>'
        b'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        + relationship_entries
        + b"</Relationships>"
    )


def _core_parts(
    *,
    sheet_records: bytes | None = None,
    workbook_records: bytes | None = None,
    relationship_target: str = "worksheets/sheet1.bin",
    shared_strings: bytes | None = None,
) -> dict[str, bytes]:
    if workbook_records is None:
        workbook_records = (
            _record(0x0099, struct.pack("<I", 1))
            + _record(0x008F, b"")
            + _record(0x009C, _bundle_sheet(0, 1, "rId1", "Model"))
            + _record(0x0090, b"")
        )
    if sheet_records is None:
        sheet_records = _record(0x0091, b"") + _record(0x0092, b"")
    parts = {
        "xl/workbook.bin": workbook_records,
        "xl/_rels/workbook.bin.rels": _relationships(relationship_target),
        "xl/worksheets/sheet1.bin": sheet_records,
    }
    if shared_strings is not None:
        parts["xl/sharedStrings.bin"] = shared_strings
    return parts


def _write_xlsb(path, parts: dict[str, bytes]) -> None:
    with ZipFile(path, "w", compression=ZIP_DEFLATED) as archive:
        for name, payload in parts.items():
            archive.writestr(name, payload)


def test_xlsb_record_iterator_decodes_one_and_two_byte_types() -> None:
    payload = _record(0x0008, b"formula") + _record(0x009C, b"sheet")

    records = xlsb_records_from_bytes(payload)

    assert [(record.record_type, record.payload) for record in records] == [
        (0x0008, b"formula"),
        (0x009C, b"sheet"),
    ]
    assert [record.offset for record in records] == [0, 9]


@pytest.mark.parametrize(
    "payload, message",
    [
        (b"\x80", "Truncated XLSB record header."),
        (b"\x80\x80", "XLSB record type uses more than two bytes."),
        (
            b"\x08\x80\x80\x80\x80",
            "XLSB record length uses more than four bytes.",
        ),
        (b"\x08\x02x", "Truncated XLSB record payload."),
    ],
)
def test_xlsb_record_iterator_rejects_malformed_framing(
    payload: bytes, message: str
) -> None:
    with pytest.raises(XlsbParseError, match=message):
        xlsb_records_from_bytes(payload)


def test_xlsb_record_iterator_enforces_local_limits() -> None:
    with pytest.raises(XlsbLimitError, match="max_record_bytes"):
        xlsb_records_from_bytes(_record(8, b"12"), max_record_bytes=1)
    with pytest.raises(XlsbLimitError, match="max_records"):
        xlsb_records_from_bytes(_record(8, b"") + _record(9, b""), max_records=1)


def test_xlsb_formula_decoder_reconstructs_references_and_operators() -> None:
    tokens = b"\x44" + _reference(12, 2) + b"\x1e\x02\x00\x05"

    formula = decode_xlsb_formula(tokens)

    assert formula == "=(C13*2)"
    inspection = inspect_formula(formula, origin=("Model", "B13"))
    assert inspection.references[0].sheet is None
    assert inspection.references[0].min_row == 13
    assert inspection.references[0].min_column == 3


def test_xlsb_formula_decoder_preserves_absolute_and_area_flags() -> None:
    tokens = b"\x45" + _area(
        1,
        3,
        2,
        4,
        first_row_relative=False,
        last_column_relative=False,
    )

    assert decode_xlsb_formula(tokens) == "=C$2:$E4"


def test_xlsb_formula_decoder_supports_fixed_and_variable_functions() -> None:
    fixed = b"\x21\x0A\x00\x21\x10\x00"
    variable = (
        b"\x44"
        + _reference(0, 0)
        + b"\x44"
        + _reference(0, 1)
        + b"\x42\x02\x04\x00"
    )

    assert decode_xlsb_formula(fixed) == "=COS(NA())"
    assert decode_xlsb_formula(variable) == "=SUM(A1,B1)"
    assert decode_xlsb_formula(
        b"\x17"
        + struct.pack("<H", 1)
        + "a".encode("utf-16le")
        + b"\x17"
        + struct.pack("<H", 1)
        + "b".encode("utf-16le")
        + b"\x42\x02\x50\x01"
    ) == '=CONCATENATE("a","b")'


def test_xlsb_formula_decoder_supports_verified_udf_and_spacing_tokens() -> None:
    tokens = (
        b"\x43"
        + struct.pack("<I", 1)
        + b"\x17"
        + struct.pack("<H", 1)
        + "A".encode("utf-16le")
        + b"\x19\x40\x00\x01"
        + b"\x17"
        + struct.pack("<H", 1)
        + "b".encode("utf-16le")
        + b"\x42\x03\xFF\x00"
    )

    assert decode_xlsb_formula(tokens, defined_names=("_xlfn.CONCAT",)) == (
        '=_xlfn.CONCAT("A","b")'
    )

    with pytest.raises(XlsbFormulaUnsupportedError, match="attribute token"):
        decode_xlsb_formula(b"\x19\x10\x00\x00")


def test_xlsb_formula_decoder_handles_strings_names_and_internal_3d_references() -> None:
    string = 'a"b'.encode("utf-16le")
    string_tokens = b"\x17" + struct.pack("<H", 3) + string
    name_tokens = b"\x43" + struct.pack("<I", 1)
    three_d_tokens = b"\x5A" + struct.pack("<H", 0) + _reference(4, 1)

    assert decode_xlsb_formula(string_tokens) == '="a""b"'
    assert decode_xlsb_formula(name_tokens, defined_names=("Scenario",)) == "=Scenario"
    assert (
        decode_xlsb_formula(three_d_tokens, external_sheets=("Inputs & Assumptions",))
        == "='Inputs & Assumptions'!B5"
    )


def test_xlsb_formula_decoder_handles_broken_reference_operands() -> None:
    formula = decode_xlsb_formula(b"\x4A" + b"\x00" * 6)

    assert formula == "=#REF!"
    assert has_broken_reference(formula)


@pytest.mark.parametrize(
    "tokens, expected",
    [
        (b"\x01" + b"\x00" * 4, XlsbFormulaUnsupportedError),
        (b"\x20" + b"\x00" * 14, XlsbFormulaUnsupportedError),
        (b"\x39" + b"\x00" * 6, XlsbFormulaUnsupportedError),
        (b"\x7F", XlsbFormulaUnsupportedError),
        (b"\x03", XlsbFormulaError),
    ],
)
def test_xlsb_formula_decoder_fails_closed_for_ambiguous_tokens(
    tokens: bytes, expected: type[Exception]
) -> None:
    with pytest.raises(expected):
        decode_xlsb_formula(tokens)


def test_xlsb_formula_decoder_rejects_invalid_indexes_and_literals() -> None:
    with pytest.raises(XlsbFormulaError, match="defined-name index"):
        decode_xlsb_formula(b"\x43" + struct.pack("<I", 0))
    with pytest.raises(XlsbFormulaError, match="external-sheet index"):
        decode_xlsb_formula(b"\x5A" + struct.pack("<H", 0) + _reference(0, 0))
    with pytest.raises(XlsbFormulaError, match="Boolean"):
        decode_xlsb_formula(b"\x1D\x02")
    with pytest.raises(XlsbFormulaError, match="unknown Excel error"):
        decode_xlsb_formula(b"\x1C\xFF")
    with pytest.raises(XlsbFormulaError, match="non-finite"):
        decode_xlsb_formula(b"\x1F" + struct.pack("<d", float("nan")))


def test_xlsb_formula_decoder_enforces_formula_specific_limits() -> None:
    tokens = b"\x1E\x01\x00\x1E\x02\x00\x03"
    with pytest.raises(XlsbLimitError, match="max_formula_tokens"):
        decode_xlsb_formula(tokens, max_tokens=2)
    with pytest.raises(XlsbLimitError, match="max_formula_stack_items"):
        decode_xlsb_formula(tokens, max_stack_items=1)
    with pytest.raises(XlsbLimitError, match="max_formula_chars"):
        decode_xlsb_formula(tokens, max_formula_chars=3)
    with pytest.raises(XlsbLimitError, match="max_function_args"):
        decode_xlsb_formula(
            b"\x1E\x01\x00\x1E\x02\x00\x42\x02\x04\x00",
            max_function_args=1,
        )


def test_xlsb_formula_decoder_renders_deep_nesting_without_recursive_expansion() -> None:
    literal = b"\x1E\x01\x00"
    tokens = literal + (literal + b"\x03") * 20_000

    formula = decode_xlsb_formula(tokens)

    assert formula.startswith("=(")
    assert formula.endswith(")")
    assert formula.count("+") == 20_000


def test_xlsb_core_reader_parses_cells_names_and_formula_cache() -> None:
    name_tokens = b"\x44" + _reference(0, 0)
    formula_tokens = b"\x43" + struct.pack("<I", 1)
    arithmetic_tokens = b"\x44" + _reference(0, 0) + b"\x1E\x02\x00\x05"
    workbook_records = (
        _record(0x0099, struct.pack("<I", 1))
        + _record(0x008F, b"")
        + _record(0x009C, _bundle_sheet(0, 1, "rId1", "Model"))
        + _record(0x0090, b"")
        + _record(0x0027, _defined_name("Scenario", name_tokens))
    )
    shared_strings = _record(0x009F, struct.pack("<II", 1, 1)) + _record(
        0x0013, b"\x00" + _wide("Input label")
    )
    sheet_records = (
        _record(0x0091, b"")
        + _record(0x0000, struct.pack("<I", 0))
        + _record(0x0007, _cell_header(0) + struct.pack("<I", 0))
        + _record(0x0002, _cell_header(1) + struct.pack("<i", (25 << 2) | 2))
        + _record(0x0009, _formula_num_cell(2, 25, formula_tokens))
        + _record(0x0009, _formula_num_cell(3, 50, arithmetic_tokens))
        + _record(0x0092, b"")
    )

    core = parse_xlsb_workbook_parts(
        _core_parts(
            sheet_records=sheet_records,
            workbook_records=workbook_records,
            shared_strings=shared_strings,
        )
    )

    assert core.date_1904
    assert core.formula_text_coverage_complete
    assert core.sheets[0].title == "Model"
    assert core.sheets[0].tab_id == 1
    assert core.defined_names[0].formula == "=A1"
    assert core.cells[("Model", 0, 0)].value == "Input label"
    assert core.cells[("Model", 0, 1)].value == 25
    assert core.cells[("Model", 0, 2)].formula == "=Scenario"
    assert core.cells[("Model", 0, 2)].cached_value == 25
    assert core.cells[("Model", 0, 3)].formula == "=(A1*2)"
    assert core.cells[("Model", 0, 3)].cached_value == 50


def test_xlsb_core_reader_marks_unsupported_formula_text_as_incomplete_coverage() -> None:
    sheet_records = (
        _record(0x0091, b"")
        + _record(0x0000, struct.pack("<I", 0))
        + _record(
            0x0009,
            _formula_num_cell(0, 1, b"\x01" + b"\x00" * 4),
        )
        + _record(0x0092, b"")
    )

    core = parse_xlsb_workbook_parts(_core_parts(sheet_records=sheet_records))

    assert not core.formula_text_coverage_complete
    assert core.unsupported_formula_cells == frozenset({("Model", 0, 0)})
    assert core.cells[("Model", 0, 0)].cell_type == "formula"
    assert core.cells[("Model", 0, 0)].formula is None
    assert core.cells[("Model", 0, 0)].value_type == "unsupported_formula"


def test_xlsb_core_reader_rejects_malformed_formula_tokens() -> None:
    sheet_records = (
        _record(0x0091, b"")
        + _record(0x0000, struct.pack("<I", 0))
        + _record(0x0009, _formula_num_cell(0, 1, b"\x03"))
        + _record(0x0092, b"")
    )

    with pytest.raises(XlsbFormulaError, match="stack underflow"):
        parse_xlsb_workbook_parts(_core_parts(sheet_records=sheet_records))


def test_xlsb_core_reader_rejects_unsafe_relationship_targets() -> None:
    with pytest.raises(XlsbParseError, match="unsafe target"):
        parse_xlsb_workbook_parts(_core_parts(relationship_target="../outside.bin"))


def test_xlsb_core_reader_bounds_relationship_xml_surface() -> None:
    parts = _core_parts()
    parts["xl/_rels/workbook.bin.rels"] = _relationships().replace(
        b' Target="worksheets/sheet1.bin"',
        b' Unexpected="x" Target="worksheets/sheet1.bin"',
    )
    with pytest.raises(XlsbParseError, match="unsupported attribute"):
        parse_xlsb_workbook_parts(parts)

    parts = _core_parts()
    parts["xl/_rels/workbook.bin.rels"] = _relationships().replace(
        b'/>', b' TargetMode="remote"/>', 1
    )
    with pytest.raises(XlsbParseError, match="invalid target mode"):
        parse_xlsb_workbook_parts(parts)

    limits = replace(XlsbReaderLimits(), max_relationship_part_bytes=1)
    with pytest.raises(XlsbLimitError, match="max_relationship_part_bytes"):
        parse_xlsb_workbook_parts(_core_parts(), limits=limits)


def test_xlsb_profile_rejects_an_oversized_relationship_part_before_reading_it(
    tmp_path,
) -> None:
    workbook_path = tmp_path / "oversized-relationships.xlsb"
    parts = _core_parts()
    # Keep the generated member below the outer ZIP compression-ratio limit so
    # this test reaches the narrower XLSB relationship-part gate.
    parts["xl/_rels/workbook.bin.rels"] = bytes(range(256)) * 4_097
    _write_xlsb(workbook_path, parts)

    with pytest.raises(WorkbookLoadError, match="relationship part exceeds"):
        load_snapshot(workbook_path, inspection_scope="profile")


def test_xlsb_core_reader_rejects_ambiguous_tab_and_defined_name_identifiers() -> None:
    duplicate_tab_catalog = (
        _record(0x0099, struct.pack("<I", 0))
        + _record(0x008F, b"")
        + _record(0x009C, _bundle_sheet(0, 0, "rId1", "Model"))
        + _record(0x009C, _bundle_sheet(0, 0, "rId2", "Inputs"))
        + _record(0x0090, b"")
    )
    duplicate_tab_parts = _core_parts(workbook_records=duplicate_tab_catalog)
    duplicate_tab_parts["xl/worksheets/sheet2.bin"] = _record(0x0091, b"") + _record(
        0x0092, b""
    )
    duplicate_tab_parts["xl/_rels/workbook.bin.rels"] = _relationships(
        "worksheets/sheet1.bin", "worksheets/sheet2.bin"
    )
    with pytest.raises(XlsbParseError, match="duplicate sheet tab identifiers"):
        parse_xlsb_workbook_parts(duplicate_tab_parts)

    duplicate_names_catalog = (
        _record(0x0099, struct.pack("<I", 0))
        + _record(0x008F, b"")
        + _record(0x009C, _bundle_sheet(0, 1, "rId1", "Model"))
        + _record(0x0090, b"")
        + _record(0x0027, _defined_name("Scenario", b""))
        + _record(0x0027, _defined_name("scenario", b""))
    )
    with pytest.raises(XlsbParseError, match="duplicate defined names"):
        parse_xlsb_workbook_parts(_core_parts(workbook_records=duplicate_names_catalog))


def test_xlsb_core_reader_requires_a_well_delimited_sheet_catalog() -> None:
    malformed_catalog = (
        _record(0x0099, struct.pack("<I", 0))
        + _record(0x009C, _bundle_sheet(0, 1, "rId1", "Model"))
        + _record(0x0090, b"")
    )

    with pytest.raises(XlsbParseError, match="outside BrtBundleShs"):
        parse_xlsb_workbook_parts(_core_parts(workbook_records=malformed_catalog))


def test_xlsb_core_reader_rejects_invalid_shared_string_and_cell_state() -> None:
    invalid_shared_string_cell = (
        _record(0x0091, b"")
        + _record(0x0000, struct.pack("<I", 0))
        + _record(0x0007, _cell_header(0) + struct.pack("<I", 0))
        + _record(0x0092, b"")
    )
    with pytest.raises(XlsbParseError, match="shared-string index"):
        parse_xlsb_workbook_parts(_core_parts(sheet_records=invalid_shared_string_cell))

    cell_before_row = _record(0x0091, b"") + _record(
        0x0002, _cell_header(0) + struct.pack("<i", (1 << 2) | 2)
    )
    with pytest.raises(XlsbParseError, match="outside a row"):
        parse_xlsb_workbook_parts(_core_parts(sheet_records=cell_before_row))

    mismatched_shared_strings = _record(0x009F, struct.pack("<II", 2, 2)) + _record(
        0x0013, b"\x00" + _wide("only one")
    )
    with pytest.raises(XlsbParseError, match="count does not match"):
        parse_xlsb_workbook_parts(
            _core_parts(shared_strings=mismatched_shared_strings)
        )


def test_xlsb_core_reader_rejects_nested_relationship_content_before_tree_growth() -> None:
    parts = _core_parts()
    parts["xl/_rels/workbook.bin.rels"] = (
        b'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        b"<opaque><Relationship Id=\"rId1\" Type=\"ignored\" Target=\"ignored\"/></opaque>"
        b"</Relationships>"
    )

    with pytest.raises(XlsbParseError, match="unknown element"):
        parse_xlsb_workbook_parts(parts)


def test_xlsb_core_reader_accepts_one_case_variant_but_rejects_a_collision() -> None:
    shared_strings = _record(0x009F, struct.pack("<II", 1, 1)) + _record(
        0x0013, b"\x00" + _wide("case-normalized")
    )
    sheet_records = (
        _record(0x0091, b"")
        + _record(0x0000, struct.pack("<I", 0))
        + _record(0x0007, _cell_header(0) + struct.pack("<I", 0))
        + _record(0x0092, b"")
    )
    parts = _core_parts(sheet_records=sheet_records, shared_strings=shared_strings)
    parts["xl/SharedStrings.bin"] = parts.pop("xl/sharedStrings.bin")

    core = parse_xlsb_workbook_parts(parts)

    assert core.cells[("Model", 0, 0)].value == "case-normalized"

    parts["XL/WORKBOOK.BIN"] = parts["xl/workbook.bin"]
    with pytest.raises(XlsbParseError, match="case-colliding"):
        parse_xlsb_workbook_parts(parts)


def test_xlsb_core_reader_enforces_semantic_limits() -> None:
    value_sheet_records = (
        _record(0x0091, b"")
        + _record(0x0000, struct.pack("<I", 0))
        + _record(0x0002, _cell_header(0) + struct.pack("<i", (1 << 2) | 2))
        + _record(0x0092, b"")
    )
    default_limits = XlsbReaderLimits()
    limits = replace(default_limits, max_worksheet_cells=0)
    with pytest.raises(ValueError, match="max_worksheet_cells"):
        parse_xlsb_workbook_parts(
            _core_parts(sheet_records=value_sheet_records), limits=limits
        )

    limits = replace(
        default_limits,
        max_relationship_part_bytes=default_limits.max_binary_part_bytes + 1,
    )
    with pytest.raises(ValueError, match="max_relationship_part_bytes"):
        parse_xlsb_workbook_parts(_core_parts(), limits=limits)

    limits = replace(default_limits, max_worksheet_cells=1)
    core = parse_xlsb_workbook_parts(
        _core_parts(sheet_records=value_sheet_records), limits=limits
    )
    assert core.cells[("Model", 0, 0)].value == 1

    two_sheet_workbook_records = (
        _record(0x0099, struct.pack("<I", 0))
        + _record(0x008F, b"")
        + _record(0x009C, _bundle_sheet(0, 1, "rId1", "Model"))
        + _record(0x009C, _bundle_sheet(0, 2, "rId2", "Inputs"))
        + _record(0x0090, b"")
    )
    two_sheet_parts = _core_parts(
        sheet_records=value_sheet_records,
        workbook_records=two_sheet_workbook_records,
    )
    two_sheet_parts["xl/worksheets/sheet2.bin"] = value_sheet_records
    two_sheet_parts["xl/_rels/workbook.bin.rels"] = _relationships(
        "worksheets/sheet1.bin", "worksheets/sheet2.bin"
    )
    limits = replace(
        default_limits,
        max_worksheet_cells=1,
        max_total_worksheet_cells=1,
    )
    with pytest.raises(XlsbLimitError, match="max_total_worksheet_cells"):
        parse_xlsb_workbook_parts(two_sheet_parts, limits=limits)

    formula_sheet_records = (
        _record(0x0091, b"")
        + _record(0x0000, struct.pack("<I", 0))
        + _record(0x0009, _formula_num_cell(0, 1, b"\x1E\x01\x00"))
        + _record(0x0009, _formula_num_cell(1, 1, b"\x1E\x01\x00"))
        + _record(0x0092, b"")
    )
    limits = replace(
        default_limits,
        max_worksheet_cells=2,
        max_total_worksheet_cells=2,
        max_formula_characters=3,
        max_total_formula_characters=3,
    )
    with pytest.raises(XlsbLimitError, match="max_total_formula_characters"):
        parse_xlsb_workbook_parts(
            _core_parts(sheet_records=formula_sheet_records), limits=limits
        )


def test_xlsb_profile_workflow_is_explicit_and_does_not_publish_cell_contents(
    tmp_path, capsys
) -> None:
    formula_tokens = b"\x44" + _reference(0, 0) + b"\x1E\x02\x00\x05"
    shared_strings = _record(0x009F, struct.pack("<II", 1, 1)) + _record(
        0x0013, b"\x00" + _wide("PRIVATE-XLSB-CELL-VALUE")
    )
    sheet_records = (
        _record(0x0091, b"")
        + _record(0x0000, struct.pack("<I", 0))
        + _record(0x0007, _cell_header(0) + struct.pack("<I", 0))
        + _record(0x0009, _formula_num_cell(1, 2, formula_tokens))
        + _record(0x0092, b"")
    )
    workbook_records = (
        _record(0x0099, struct.pack("<I", 0))
        + _record(0x008F, b"")
        + _record(0x009C, _bundle_sheet(0, 1, "rId1", "Model"))
        + _record(0x0090, b"")
        + _record(
            0x0027,
            _defined_name("PRIVATE-XLSB-DEFINED-NAME", b"\x44" + _reference(0, 0)),
        )
    )
    workbook_path = tmp_path / "model.xlsb"
    parts = _core_parts(
        sheet_records=sheet_records,
        workbook_records=workbook_records,
        shared_strings=shared_strings,
    )
    parts["xl/SharedStrings.bin"] = parts.pop("xl/sharedStrings.bin")
    _write_xlsb(workbook_path, parts)

    snapshot = load_snapshot(workbook_path, inspection_scope="profile")
    profile = profile_snapshot(snapshot)
    rendered = json.dumps(profile, sort_keys=True)

    assert snapshot.file_type == "xlsb"
    assert snapshot.inspection_scope == "xlsb_core_profile"
    assert snapshot.formula_text_coverage_complete
    assert snapshot.summary()["formula_cells"] == 1
    assert snapshot.summary()["defined_names"] == 1
    assert all(cell.value is None and cell.formula is None for cell in snapshot.cells.values())
    assert all("PRIVATE" not in name for name in snapshot.defined_names)
    with pytest.raises(FormulaFenceError, match="Formula lint requires a full workbook"):
        lint_snapshot(snapshot)
    with pytest.raises(FormulaFenceError, match="Workbook comparison requires a full workbook"):
        compare_snapshots(snapshot, snapshot)
    with pytest.raises(
        FormulaFenceError, match="Downstream impact analysis requires a full workbook"
    ):
        analyze_downstream_impact(("Model", "A1"), snapshot)
    assert profile["coverage"] == {
        "scope": "xlsb_core_profile",
        "supported_workflows": ["profile"],
        "formula_text_coverage_complete": True,
        "limitations": list(snapshot.parser_warnings),
    }
    assert "PRIVATE-XLSB-CELL-VALUE" not in rendered
    assert "PRIVATE-XLSB-DEFINED-NAME" not in rendered
    assert "A1*2" not in rendered
    markdown = profile_to_markdown(profile)
    assert "## Inspection scope" in markdown
    assert "Counts outside the stated scope are unassessed" in markdown
    assert "Formula-text token coverage:** complete" in markdown

    with pytest.raises(WorkbookLoadError, match="only by the profile workflow"):
        load_snapshot(workbook_path)

    output = tmp_path / "profile.json"
    assert (
        main(
            [
                "profile",
                str(workbook_path),
                "--format",
                "json",
                "--output",
                str(output),
            ]
        )
        == 0
    )
    cli_profile = json.loads(output.read_text(encoding="utf-8"))
    assert cli_profile["coverage"]["scope"] == "xlsb_core_profile"
    assert "PRIVATE-XLSB-CELL-VALUE" not in output.read_text(encoding="utf-8")

    assert main(["lint", str(workbook_path)]) == 2
    assert "only by the profile workflow" in capsys.readouterr().err
