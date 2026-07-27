from __future__ import annotations

import binascii
import hashlib
import stat
import struct
import zipfile
from copy import deepcopy
from pathlib import Path
from xml.etree import ElementTree
from zipfile import ZIP_DEFLATED, ZipFile, ZipInfo

import pytest

import formulafence.workbook as workbook_module
from formulafence.cli import main
from formulafence.models import WorkbookLoadError
from formulafence.workbook import load_snapshot

from .helpers import (
    make_external_data_refresh_model,
    make_external_link_package_model,
    make_model,
    make_strict_worksheet_print_layout_model,
)


def _append_member(
    path: Path,
    name: str,
    payload: bytes,
    *,
    compression: int = ZIP_DEFLATED,
) -> None:
    with ZipFile(path, "a", compression=compression) as archive:
        archive.writestr(name, payload)


def _replace_member(path: Path, name: str, payload: bytes) -> None:
    """Replace one test package part without creating a duplicate ZIP entry."""
    staging = path.with_suffix(".replacement.xlsx")
    with ZipFile(path) as archive:
        contents = {
            member.filename: archive.read(member)
            for member in archive.infolist()
        }
    contents[name] = payload
    with ZipFile(staging, "w", compression=ZIP_DEFLATED) as archive:
        for member_name, member_payload in contents.items():
            archive.writestr(member_name, member_payload)
    staging.replace(path)


def _append_workbook_sheet_declarations(
    path: Path,
    count: int,
    *,
    namespace_prefix: str | None = None,
) -> None:
    """Repeat one declared sheet to exercise pre-reader sheet amplification."""
    member_name = "xl/workbook.xml"
    with ZipFile(path) as archive:
        workbook = archive.read(member_name)
    start = workbook.index(b"<sheet ")
    end = workbook.index(b"/>", start) + 2
    declaration = workbook[start:end]
    if namespace_prefix is not None:
        namespace = b"urn:formulafence:archive-safety"
        workbook = workbook.replace(
            b"<workbook ",
            b"<workbook xmlns:" + namespace_prefix.encode() + b'=\"' + namespace + b'\" ',
            1,
        )
        declaration = declaration.replace(
            b"<sheet ",
            b"<" + namespace_prefix.encode() + b":sheet ",
            1,
        )
    _replace_member(
        path,
        member_name,
        workbook.replace(b"</sheets>", declaration * count + b"</sheets>"),
    )


def _append_workbook_defined_name_declarations(
    path: Path,
    count: int,
    *,
    namespace_prefix: str | None = None,
) -> None:
    """Append unique declared names without relying on an object-model writer."""
    member_name = "xl/workbook.xml"
    with ZipFile(path) as archive:
        workbook = archive.read(member_name)
    tag_name = "definedName"
    if namespace_prefix is not None:
        namespace = b"urn:formulafence:archive-safety"
        workbook = workbook.replace(
            b"<workbook ",
            b"<workbook xmlns:" + namespace_prefix.encode() + b'=\"' + namespace + b'\" ',
            1,
        )
        tag_name = f"{namespace_prefix}:definedName"
    declarations = b"".join(
        (
            f'<{tag_name} name="FormulaFenceAuditDefinedName{index:06d}">'
            f"Inputs!$B$2</{tag_name}>"
        ).encode()
        for index in range(count)
    )
    _replace_member(
        path,
        member_name,
        workbook.replace(
            b"</definedNames>",
            declarations + b"</definedNames>",
        ),
    )


def _append_workbook_relationship_catalog_declarations(
    path: Path,
    catalog_name: str,
    count: int,
    *,
    alternate_namespace: str | None = None,
) -> None:
    """Repeat one relationship-backed workbook declaration in a test package."""
    member_name = "xl/workbook.xml"
    with ZipFile(path) as archive:
        workbook = archive.read(member_name)
    root = ElementTree.fromstring(workbook)
    container = next(
        (
            child
            for child in root
            if child.tag.rsplit("}", maxsplit=1)[-1] == catalog_name
        ),
        None,
    )
    if container is None:
        raise ValueError(f"workbook fixture has no {catalog_name!r} catalog")
    declaration = next(iter(container), None)
    if declaration is None:
        raise ValueError(f"workbook fixture has an empty {catalog_name!r} catalog")
    if alternate_namespace is not None:
        declaration = deepcopy(declaration)
        declaration.tag = (
            f"{{{alternate_namespace}}}" + declaration.tag.rsplit("}", maxsplit=1)[-1]
        )
    for _ in range(count):
        container.append(deepcopy(declaration))
    _replace_member(
        path,
        member_name,
        ElementTree.tostring(root, encoding="utf-8", xml_declaration=True),
    )


def _append_workbook_package_catalog_declarations(
    path: Path,
    catalog_name: str,
    child_name: str,
    count: int,
    *,
    attributes: dict[str, str] | None = None,
    alternate_namespace: str | None = None,
) -> None:
    """Add repeated workbook-package catalog entries without using a reader."""
    member_name = "xl/workbook.xml"
    with ZipFile(path) as archive:
        workbook = archive.read(member_name)
    root = ElementTree.fromstring(workbook)
    namespace = root.tag.partition("}")[0].removeprefix("{")
    qualified_catalog_name = (
        f"{{{namespace}}}{catalog_name}" if namespace else catalog_name
    )
    qualified_child_name = f"{{{namespace}}}{child_name}" if namespace else child_name
    container = next(
        (
            child
            for child in root
            if child.tag.rsplit("}", maxsplit=1)[-1] == catalog_name
        ),
        None,
    )
    if container is None:
        container = ElementTree.Element(qualified_catalog_name)
        root.insert(0, container)
    declaration = ElementTree.Element(qualified_child_name)
    for name, value in (attributes or {}).items():
        declaration.set(name, value)
    if alternate_namespace is not None:
        declaration.tag = f"{{{alternate_namespace}}}{child_name}"
    for _ in range(count):
        container.append(deepcopy(declaration))
    _replace_member(
        path,
        member_name,
        ElementTree.tostring(root, encoding="utf-8", xml_declaration=True),
    )


def _append_custom_sheet_view_declarations(
    path: Path,
    count: int,
    *,
    alternate_namespace: str | None = None,
) -> None:
    """Repeat direct sheet Custom View entries for the preflight boundary."""
    member_name = "xl/worksheets/sheet1.xml"
    with ZipFile(path) as archive:
        worksheet = archive.read(member_name)
    root = ElementTree.fromstring(worksheet)
    namespace = root.tag.partition("}")[0].removeprefix("{")
    qualified_container_name = (
        f"{{{namespace}}}customSheetViews" if namespace else "customSheetViews"
    )
    qualified_child_name = (
        f"{{{namespace}}}customSheetView" if namespace else "customSheetView"
    )
    container = next(
        (
            child
            for child in root
            if child.tag.rsplit("}", maxsplit=1)[-1] == "customSheetViews"
        ),
        None,
    )
    if container is None:
        container = ElementTree.Element(qualified_container_name)
        root.insert(0, container)
    declaration = ElementTree.Element(qualified_child_name)
    declaration.set("guid", "{11111111-1111-1111-1111-111111111111}")
    if alternate_namespace is not None:
        declaration.tag = f"{{{alternate_namespace}}}customSheetView"
    for _ in range(count):
        container.append(deepcopy(declaration))
    _replace_member(
        path,
        member_name,
        ElementTree.tostring(root, encoding="utf-8", xml_declaration=True),
    )


def _append_merged_cell_declarations(
    path: Path,
    references: tuple[str, ...],
    *,
    alternate_namespace: str | None = None,
) -> None:
    """Add direct worksheet merge declarations without invoking a workbook reader."""
    member_name = "xl/worksheets/sheet1.xml"
    with ZipFile(path) as archive:
        worksheet = archive.read(member_name)
    root = ElementTree.fromstring(worksheet)
    namespace = root.tag.partition("}")[0].removeprefix("{")
    qualified_container_name = (
        f"{{{namespace}}}mergeCells" if namespace else "mergeCells"
    )
    qualified_child_name = f"{{{namespace}}}mergeCell" if namespace else "mergeCell"
    container = next(
        (
            child
            for child in root
            if child.tag.rsplit("}", maxsplit=1)[-1] == "mergeCells"
        ),
        None,
    )
    if container is None:
        container = ElementTree.Element(qualified_container_name)
        root.insert(0, container)
    for reference in references:
        declaration = ElementTree.Element(qualified_child_name, {"ref": reference})
        if alternate_namespace is not None:
            declaration.tag = f"{{{alternate_namespace}}}mergeCell"
        container.append(declaration)
    _replace_member(
        path,
        member_name,
        ElementTree.tostring(root, encoding="utf-8", xml_declaration=True),
    )


def _append_row_dimension_declarations(
    path: Path,
    count: int,
    *,
    member_name: str = "xl/worksheets/sheet1.xml",
    attributes: dict[str, str] | None = None,
    alternate_namespace: str | None = None,
) -> None:
    """Append empty rows that can make the workbook reader retain dimensions."""
    with ZipFile(path) as archive:
        worksheet = archive.read(member_name)
    root = ElementTree.fromstring(worksheet)
    namespace = root.tag.partition("}")[0].removeprefix("{")
    qualified_row_name = f"{{{namespace}}}row" if namespace else "row"
    if alternate_namespace is not None:
        qualified_row_name = f"{{{alternate_namespace}}}row"
    sheet_data = next(
        (
            child
            for child in root
            if child.tag.rsplit("}", maxsplit=1)[-1] == "sheetData"
        ),
        None,
    )
    if sheet_data is None:
        raise ValueError("worksheet fixture has no sheetData element")
    row_attributes = (
        {"ht": "15", "customHeight": "1"}
        if attributes is None
        else attributes
    )
    for index in range(count):
        declaration_attributes = {"r": str(10_000 + index), **row_attributes}
        ElementTree.SubElement(sheet_data, qualified_row_name, declaration_attributes)
    _replace_member(
        path,
        member_name,
        ElementTree.tostring(root, encoding="utf-8", xml_declaration=True),
    )


def _append_column_dimension_declarations(
    path: Path,
    count: int,
    *,
    member_name: str = "xl/worksheets/sheet1.xml",
    attributes: dict[str, str] | None = None,
    alternate_namespace: str | None = None,
) -> None:
    """Append reader-visible column declarations without an object-model writer."""
    with ZipFile(path) as archive:
        worksheet = archive.read(member_name)
    root = ElementTree.fromstring(worksheet)
    namespace = root.tag.partition("}")[0].removeprefix("{")
    qualified_container_name = f"{{{namespace}}}cols" if namespace else "cols"
    qualified_column_name = f"{{{namespace}}}col" if namespace else "col"
    if alternate_namespace is not None:
        qualified_column_name = f"{{{alternate_namespace}}}col"
    columns = next(
        (child for child in root if child.tag == qualified_container_name),
        None,
    )
    if columns is None:
        columns = ElementTree.Element(qualified_container_name)
        sheet_data_index = next(
            index
            for index, child in enumerate(root)
            if child.tag.rsplit("}", maxsplit=1)[-1] == "sheetData"
        )
        root.insert(sheet_data_index, columns)
    column_attributes = {"min": "1", "max": "1"}
    if attributes is None:
        column_attributes.update({"width": "10", "customWidth": "1"})
    else:
        column_attributes.update(attributes)
    for _ in range(count):
        ElementTree.SubElement(columns, qualified_column_name, column_attributes)
    _replace_member(
        path,
        member_name,
        ElementTree.tostring(root, encoding="utf-8", xml_declaration=True),
    )


def _append_column_dimension_containers(
    path: Path,
    count: int,
    *,
    member_name: str = "xl/worksheets/sheet1.xml",
    alternate_namespace: str | None = None,
) -> None:
    """Append direct column containers that raw dimension scanners must retain."""
    with ZipFile(path) as archive:
        worksheet = archive.read(member_name)
    root = ElementTree.fromstring(worksheet)
    namespace = root.tag.partition("}")[0].removeprefix("{")
    qualified_container_name = f"{{{namespace}}}cols" if namespace else "cols"
    if alternate_namespace is not None:
        qualified_container_name = f"{{{alternate_namespace}}}cols"
    sheet_data_index = next(
        index
        for index, child in enumerate(root)
        if child.tag.rsplit("}", maxsplit=1)[-1] == "sheetData"
    )
    for _ in range(count):
        root.insert(sheet_data_index, ElementTree.Element(qualified_container_name))
    _replace_member(
        path,
        member_name,
        ElementTree.tostring(root, encoding="utf-8", xml_declaration=True),
    )


def _append_page_break_declarations(
    path: Path,
    count: int,
    *,
    axis: str = "row",
    member_name: str = "xl/worksheets/sheet1.xml",
    child_name: str = "brk",
    alternate_namespace: str | None = None,
) -> None:
    """Append direct print-break records without invoking a workbook reader."""
    container_name = {"row": "rowBreaks", "column": "colBreaks"}.get(axis)
    if container_name is None:
        raise ValueError(f"Unsupported page-break axis {axis!r}")
    with ZipFile(path) as archive:
        worksheet = archive.read(member_name)
    root = ElementTree.fromstring(worksheet)
    namespace = root.tag.partition("}")[0].removeprefix("{")
    qualified_container_name = (
        f"{{{namespace}}}{container_name}" if namespace else container_name
    )
    qualified_child_name = f"{{{namespace}}}{child_name}" if namespace else child_name
    if alternate_namespace is not None:
        qualified_child_name = f"{{{alternate_namespace}}}{child_name}"
    container = next(
        (child for child in root if child.tag == qualified_container_name),
        None,
    )
    if container is None:
        container = ElementTree.Element(qualified_container_name)
        _insert_worksheet_child_after_sheet_data(root, container)
    for _ in range(count):
        ElementTree.SubElement(
            container,
            qualified_child_name,
            {"id": "10", "min": "0", "max": "16383", "man": "1"},
        )
    _replace_member(
        path,
        member_name,
        ElementTree.tostring(root, encoding="utf-8", xml_declaration=True),
    )


def _append_page_break_containers(
    path: Path,
    count: int,
    *,
    axis: str = "row",
    member_name: str = "xl/worksheets/sheet1.xml",
    alternate_namespace: str | None = None,
) -> None:
    """Append direct break containers that raw print scanners must retain."""
    container_name = {"row": "rowBreaks", "column": "colBreaks"}.get(axis)
    if container_name is None:
        raise ValueError(f"Unsupported page-break axis {axis!r}")
    with ZipFile(path) as archive:
        worksheet = archive.read(member_name)
    root = ElementTree.fromstring(worksheet)
    namespace = root.tag.partition("}")[0].removeprefix("{")
    qualified_container_name = (
        f"{{{namespace}}}{container_name}" if namespace else container_name
    )
    if alternate_namespace is not None:
        qualified_container_name = f"{{{alternate_namespace}}}{container_name}"
    for _ in range(count):
        _insert_worksheet_child_after_sheet_data(
            root,
            ElementTree.Element(qualified_container_name),
        )
    _replace_member(
        path,
        member_name,
        ElementTree.tostring(root, encoding="utf-8", xml_declaration=True),
    )


def _insert_worksheet_child_after_sheet_data(
    root: ElementTree.Element,
    child: ElementTree.Element,
) -> None:
    """Insert a worksheet child in a reader-compatible location for a fixture."""
    sheet_data_index = next(
        (
            index
            for index, current in enumerate(root)
            if current.tag.rsplit("}", maxsplit=1)[-1] == "sheetData"
        ),
        None,
    )
    root.insert(0 if sheet_data_index is None else sheet_data_index + 1, child)


def _append_data_validation_declarations(
    path: Path,
    references: tuple[str, ...],
    *,
    alternate_namespace: str | None = None,
    formula: str | None = None,
) -> None:
    """Add direct validation declarations without invoking a workbook reader."""
    member_name = "xl/worksheets/sheet1.xml"
    with ZipFile(path) as archive:
        worksheet = archive.read(member_name)
    root = ElementTree.fromstring(worksheet)
    namespace = root.tag.partition("}")[0].removeprefix("{")
    qualified_container_name = (
        f"{{{namespace}}}dataValidations" if namespace else "dataValidations"
    )
    qualified_child_name = (
        f"{{{namespace}}}dataValidation" if namespace else "dataValidation"
    )
    qualified_formula_name = f"{{{namespace}}}formula1" if namespace else "formula1"
    container = next(
        (
            child
            for child in root
            if child.tag.rsplit("}", maxsplit=1)[-1] == "dataValidations"
        ),
        None,
    )
    if container is None:
        container = ElementTree.Element(qualified_container_name)
        _insert_worksheet_child_after_sheet_data(root, container)
    for reference in references:
        declaration = ElementTree.Element(
            qualified_child_name,
            {"type": "whole", "sqref": reference},
        )
        if alternate_namespace is not None:
            declaration.tag = f"{{{alternate_namespace}}}dataValidation"
        if formula is not None:
            ElementTree.SubElement(declaration, qualified_formula_name).text = formula
        container.append(declaration)
    _replace_member(
        path,
        member_name,
        ElementTree.tostring(root, encoding="utf-8", xml_declaration=True),
    )


def _append_conditional_formatting_declarations(
    path: Path,
    references: tuple[str, ...],
    *,
    rule_namespaces: tuple[str | None, ...] = (None,),
    formula: str | None = None,
) -> None:
    """Add direct conditional-formatting declarations without a reader."""
    member_name = "xl/worksheets/sheet1.xml"
    with ZipFile(path) as archive:
        worksheet = archive.read(member_name)
    root = ElementTree.fromstring(worksheet)
    namespace = root.tag.partition("}")[0].removeprefix("{")
    qualified_container_name = (
        f"{{{namespace}}}conditionalFormatting" if namespace else "conditionalFormatting"
    )
    qualified_rule_name = f"{{{namespace}}}cfRule" if namespace else "cfRule"
    qualified_formula_name = f"{{{namespace}}}formula" if namespace else "formula"
    priority = 1
    for reference in references:
        container = ElementTree.Element(qualified_container_name, {"sqref": reference})
        for rule_namespace in rule_namespaces:
            rule_name = (
                qualified_rule_name
                if rule_namespace is None
                else f"{{{rule_namespace}}}cfRule"
            )
            rule = ElementTree.SubElement(
                container,
                rule_name,
                {"type": "expression", "priority": str(priority)},
            )
            priority += 1
            if formula is not None:
                ElementTree.SubElement(rule, qualified_formula_name).text = formula
        _insert_worksheet_child_after_sheet_data(root, container)
    _replace_member(
        path,
        member_name,
        ElementTree.tostring(root, encoding="utf-8", xml_declaration=True),
    )


def _append_scenario_manager_containers(
    path: Path,
    references: tuple[str, ...],
    *,
    scenario_count: int = 1,
    input_cell_count: int = 1,
    alternate_namespace: str | None = None,
) -> None:
    """Add Scenario Manager declarations without invoking a workbook reader."""
    member_name = "xl/worksheets/sheet1.xml"
    with ZipFile(path) as archive:
        worksheet = archive.read(member_name)
    root = ElementTree.fromstring(worksheet)
    namespace = root.tag.partition("}")[0].removeprefix("{")
    qualified_container_name = f"{{{namespace}}}scenarios" if namespace else "scenarios"
    qualified_scenario_name = f"{{{namespace}}}scenario" if namespace else "scenario"
    qualified_input_name = f"{{{namespace}}}inputCells" if namespace else "inputCells"
    scenario_name = (
        qualified_scenario_name
        if alternate_namespace is None
        else f"{{{alternate_namespace}}}scenario"
    )
    input_name = (
        qualified_input_name
        if alternate_namespace is None
        else f"{{{alternate_namespace}}}inputCells"
    )
    for container_index, reference in enumerate(references):
        container = ElementTree.Element(qualified_container_name, {"sqref": reference})
        for scenario_index in range(scenario_count):
            scenario = ElementTree.SubElement(
                container,
                scenario_name,
                {
                    "name": f"FormulaFence audit {container_index}-{scenario_index}",
                    "locked": "0",
                    "hidden": "0",
                    "count": str(input_cell_count),
                },
            )
            for _ in range(input_cell_count):
                ElementTree.SubElement(scenario, input_name, {"r": "A1", "val": "1"})
        _insert_worksheet_child_after_sheet_data(root, container)
    _replace_member(
        path,
        member_name,
        ElementTree.tostring(root, encoding="utf-8", xml_declaration=True),
    )


def _append_stylesheet_catalog_declarations(
    path: Path,
    parent_names: tuple[str, ...],
    child_name: str,
    count: int,
    *,
    alternate_parent_namespace: str | None = None,
) -> int:
    """Append reader-visible style records and return their direct-child count."""
    member_name = "xl/styles.xml"
    with ZipFile(path) as archive:
        styles = archive.read(member_name)
    root = ElementTree.fromstring(styles)
    namespace = root.tag.partition("}")[0].removeprefix("{")

    def qualified(name: str) -> str:
        return f"{{{namespace}}}{name}" if namespace else name

    parent = root
    for parent_name in parent_names:
        child = next(
            (
                current
                for current in parent
                if current.tag.rsplit("}", maxsplit=1)[-1] == parent_name
            ),
            None,
        )
        if child is None:
            child = ElementTree.Element(qualified(parent_name))
            parent.append(child)
        parent = child
    if alternate_parent_namespace is not None:
        parent.tag = f"{{{alternate_parent_namespace}}}{parent_names[-1]}"

    for index in range(count):
        declaration = ElementTree.Element(qualified(child_name))
        if child_name == "numFmt":
            declaration.attrib.update(
                {"numFmtId": str(164 + index), "formatCode": "0.000"}
            )
        elif child_name == "font":
            ElementTree.SubElement(declaration, qualified("name"), {"val": "Arial"})
        elif child_name == "fill":
            ElementTree.SubElement(
                declaration,
                qualified("patternFill"),
                {"patternType": "none"},
            )
        elif child_name == "xf":
            declaration.attrib.update(
                {"numFmtId": "0", "fontId": "0", "fillId": "0", "borderId": "0"}
            )
        elif child_name == "cellStyle":
            declaration.attrib.update(
                {"name": f"FormulaFence audit {index}", "xfId": "0"}
            )
        elif child_name in {"rgbColor", "color"}:
            declaration.set("rgb", "FF000000")
        elif child_name == "tableStyle":
            declaration.set("name", f"FormulaFence audit {index}")
        elif child_name == "tableStyleElement":
            declaration.set("type", "wholeTable")
        elif child_name == "ext":
            declaration.set("uri", f"urn:formulafence:archive-safety:{index}")
        elif child_name == "patternFill":
            declaration.set("patternType", "none")
        elif child_name == "stop":
            declaration.set("position", str(index))
        parent.append(declaration)

    _replace_member(
        path,
        member_name,
        ElementTree.tostring(root, encoding="utf-8", xml_declaration=True),
    )
    return len(parent)


def _last_central_directory_offset(contents: bytes | bytearray) -> int:
    offset = contents.rfind(b"PK\x01\x02")
    assert offset >= 0
    return offset


def _reject_before_workbook_readers(monkeypatch: pytest.MonkeyPatch, path: Path) -> str:
    def unexpected_reader(*args, **kwargs):
        raise AssertionError("a workbook reader ran before archive preflight rejected input")

    monkeypatch.setattr(workbook_module, "_workbook_tab_order_metadata", unexpected_reader)
    with pytest.raises(WorkbookLoadError, match="safety preflight") as error:
        load_snapshot(path)
    return str(error.value)


_WORKBOOK_AUXILIARY_CATALOG_CASES = (
    (
        "bookViews",
        "workbookView",
        {},
        "_OOXML_READER_MAX_WORKBOOK_BOOK_VIEW_COUNT",
        "workbook book-view declarations",
    ),
    (
        "customWorkbookViews",
        "customWorkbookView",
        {
            "name": "FormulaFence audit",
            "guid": "{11111111-1111-1111-1111-111111111111}",
            "windowWidth": "100",
            "windowHeight": "100",
            "activeSheetId": "0",
        },
        "_OOXML_READER_MAX_WORKBOOK_CUSTOM_VIEW_COUNT",
        "workbook custom-view declarations",
    ),
    (
        "functionGroups",
        "functionGroup",
        {"name": "FormulaFence audit"},
        "_OOXML_READER_MAX_WORKBOOK_FUNCTION_GROUP_COUNT",
        "workbook function-group declarations",
    ),
    (
        "smartTagTypes",
        "smartTagType",
        {
            "namespaceUri": "urn:formulafence:archive-safety",
            "name": "audit",
            "url": "https://example.invalid/formulafence",
        },
        "_OOXML_READER_MAX_WORKBOOK_SMART_TAG_TYPE_COUNT",
        "workbook smart-tag type declarations",
    ),
    (
        "webPublishObjects",
        "webPublishObject",
        {"id": "1", "divId": "audit", "destinationFile": "audit.html"},
        "_OOXML_READER_MAX_WORKBOOK_WEB_PUBLISH_OBJECT_COUNT",
        "workbook web-publish-object declarations",
    ),
)


_STYLESHEET_CATALOG_CASES = (
    (
        ("numFmts",),
        "numFmt",
        "_OOXML_READER_MAX_NUMBER_FORMAT_COUNT",
        "number-format records",
    ),
    (("fonts",), "font", "_OOXML_READER_MAX_FONT_COUNT", "font records"),
    (("fills",), "fill", "_OOXML_READER_MAX_FILL_COUNT", "fill records"),
    (
        ("fills", "fill"),
        "patternFill",
        "_OOXML_READER_MAX_FILL_CHILD_COUNT",
        "fill child records",
    ),
    (
        ("fills", "fill", "gradientFill"),
        "stop",
        "_OOXML_READER_MAX_GRADIENT_FILL_STOP_COUNT",
        "gradient-fill stops",
    ),
    (("borders",), "border", "_OOXML_READER_MAX_BORDER_COUNT", "border records"),
    (
        ("cellStyleXfs",),
        "xf",
        "_OOXML_READER_MAX_BASE_CELL_STYLE_COUNT",
        "base cell styles",
    ),
    (
        ("cellXfs",),
        "xf",
        "_OOXML_READER_MAX_CELL_STYLE_COUNT",
        "cell styles",
    ),
    (
        ("cellStyles",),
        "cellStyle",
        "_OOXML_READER_MAX_NAMED_CELL_STYLE_COUNT",
        "named cell styles",
    ),
    (
        ("dxfs",),
        "dxf",
        "_OOXML_READER_MAX_DIFFERENTIAL_STYLE_COUNT",
        "differential styles",
    ),
    (
        ("colors", "indexedColors"),
        "rgbColor",
        "_OOXML_READER_MAX_STYLE_COLOR_COUNT",
        "stylesheet colour records",
    ),
    (
        ("colors", "mruColors"),
        "color",
        "_OOXML_READER_MAX_STYLE_COLOR_COUNT",
        "stylesheet colour records",
    ),
    (
        ("tableStyles",),
        "tableStyle",
        "_OOXML_READER_MAX_TABLE_STYLE_COUNT",
        "table styles",
    ),
    (
        ("tableStyles", "tableStyle"),
        "tableStyleElement",
        "_OOXML_READER_MAX_TABLE_STYLE_ELEMENT_COUNT",
        "table-style elements",
    ),
    (
        ("extLst",),
        "ext",
        "_OOXML_READER_MAX_STYLE_EXTENSION_COUNT",
        "stylesheet extension records",
    ),
)


def test_archive_preflight_accepts_an_ordinary_workbook(tmp_path: Path) -> None:
    workbook = make_model(tmp_path / "ordinary.xlsx")

    snapshot = load_snapshot(workbook)

    assert snapshot.file_type == "xlsx"
    assert set(snapshot.sheets) == {"Inputs", "Model", "Dashboard", "Control"}


def test_archive_preflight_accepts_a_valid_zip64_workbook(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.setattr(zipfile, "ZIP64_LIMIT", 1)
    workbook = make_model(tmp_path / "zip64.xlsx")

    assert bytes((80, 75, 6, 6)) in workbook.read_bytes()
    assert load_snapshot(workbook).file_type == "xlsx"


def test_archive_preflight_rejects_source_size_before_any_zip_reader(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    workbook = make_model(tmp_path / "oversized.xlsx")
    monkeypatch.setattr(workbook_module, "_OOXML_ARCHIVE_MAX_SOURCE_BYTES", 1)

    def unexpected_zip_reader(*args, **kwargs):
        raise AssertionError("ZipFile ran before the source-size safety gate")

    monkeypatch.setattr(workbook_module, "ZipFile", unexpected_zip_reader)

    _reject_before_workbook_readers(monkeypatch, workbook)


def test_archive_preflight_rejects_excessive_entry_count_before_zip_reader(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    workbook = make_model(tmp_path / "too-many-members.xlsx")
    with ZipFile(workbook) as archive:
        member_count = len(archive.infolist())
    monkeypatch.setattr(workbook_module, "_OOXML_ARCHIVE_MAX_ENTRY_COUNT", member_count - 1)

    def unexpected_zip_reader(*args, **kwargs):
        raise AssertionError("ZipFile ran before the central-directory safety gate")

    monkeypatch.setattr(workbook_module, "ZipFile", unexpected_zip_reader)

    _reject_before_workbook_readers(monkeypatch, workbook)


def test_archive_preflight_rejects_an_oversized_member_before_readers(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    workbook = make_model(tmp_path / "oversized-member.xlsx")
    with ZipFile(workbook) as archive:
        largest_member = max(member.file_size for member in archive.infolist())
    _append_member(workbook, "xl/media/filler.bin", b"x" * (largest_member + 1))
    monkeypatch.setattr(
        workbook_module,
        "_OOXML_ARCHIVE_MAX_MEMBER_UNCOMPRESSED_BYTES",
        largest_member,
    )

    _reject_before_workbook_readers(monkeypatch, workbook)


def test_archive_preflight_rejects_excessive_aggregate_size_before_readers(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    workbook = make_model(tmp_path / "oversized-total.xlsx")
    with ZipFile(workbook) as archive:
        original_total = sum(member.file_size for member in archive.infolist())
    _append_member(workbook, "xl/media/filler.bin", b"x")
    monkeypatch.setattr(
        workbook_module,
        "_OOXML_ARCHIVE_MAX_TOTAL_UNCOMPRESSED_BYTES",
        original_total,
    )

    _reject_before_workbook_readers(monkeypatch, workbook)


def test_archive_preflight_rejects_a_compression_bomb_before_readers(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    workbook = make_model(tmp_path / "compression-bomb.xlsx")
    with ZipFile(workbook) as archive:
        ordinary_ratio = max(
            member.file_size // max(member.compress_size, 1)
            for member in archive.infolist()
        )
    _append_member(workbook, "xl/media/compressed.bin", b"x" * 200_000)
    with ZipFile(workbook) as archive:
        compressed_member = archive.getinfo("xl/media/compressed.bin")
    assert compressed_member.file_size > compressed_member.compress_size * ordinary_ratio
    monkeypatch.setattr(
        workbook_module,
        "_OOXML_ARCHIVE_MAX_COMPRESSION_RATIO",
        ordinary_ratio,
    )

    _reject_before_workbook_readers(monkeypatch, workbook)


def test_semantic_reader_preflight_rejects_an_oversized_xml_part_before_scanners(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    workbook = make_model(tmp_path / "oversized-xml.xlsx")
    monkeypatch.setattr(workbook_module, "_OOXML_READER_MAX_XML_PART_BYTES", 1)

    message = _reject_before_workbook_readers(monkeypatch, workbook)

    assert "semantic reader" in message
    assert "XML part" in message


def test_semantic_reader_preflight_rejects_aggregate_xml_before_scanners(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    workbook = make_model(tmp_path / "oversized-xml-total.xlsx")
    monkeypatch.setattr(workbook_module, "_OOXML_READER_MAX_TOTAL_XML_BYTES", 1)

    message = _reject_before_workbook_readers(monkeypatch, workbook)

    assert "aggregate XML" in message


def test_semantic_reader_preflight_rejects_excessive_cells_before_scanners(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    workbook = make_model(tmp_path / "too-many-cells.xlsx")
    monkeypatch.setattr(workbook_module, "_OOXML_READER_MAX_WORKSHEET_CELL_COUNT", 1)

    message = _reject_before_workbook_readers(monkeypatch, workbook)

    assert "populated worksheet cells" in message


def test_semantic_reader_preflight_rejects_excessive_row_dimensions_before_scanners(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    workbook = make_model(tmp_path / "too-many-row-dimensions.xlsx")
    _append_row_dimension_declarations(workbook, 2)
    monkeypatch.setattr(workbook_module, "_OOXML_READER_MAX_ROW_DIMENSION_COUNT", 1)

    message = _reject_before_workbook_readers(monkeypatch, workbook)

    assert "row-dimension declarations" in message


def test_semantic_reader_preflight_rejects_default_row_dimension_limit_before_scanners(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    workbook = make_model(tmp_path / "default-row-dimension-limit.xlsx")
    _append_row_dimension_declarations(
        workbook,
        workbook_module._OOXML_READER_MAX_ROW_DIMENSION_COUNT + 1,
    )

    message = _reject_before_workbook_readers(monkeypatch, workbook)

    assert "row-dimension declarations" in message


def test_semantic_reader_preflight_counts_row_dimensions_across_worksheets(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    workbook = make_model(tmp_path / "aggregate-row-dimensions.xlsx")
    _append_row_dimension_declarations(workbook, 1)
    _append_row_dimension_declarations(workbook, 1, member_name="xl/worksheets/sheet2.xml")
    monkeypatch.setattr(workbook_module, "_OOXML_READER_MAX_ROW_DIMENSION_COUNT", 1)

    message = _reject_before_workbook_readers(monkeypatch, workbook)

    assert "row-dimension declarations" in message


def test_semantic_reader_preflight_accepts_row_dimensions_at_the_limit(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    workbook = make_model(tmp_path / "row-dimensions-at-limit.xlsx")
    _append_row_dimension_declarations(workbook, 2)
    monkeypatch.setattr(workbook_module, "_OOXML_READER_MAX_ROW_DIMENSION_COUNT", 2)

    snapshot = load_snapshot(workbook)

    assert snapshot.file_type == "xlsx"


@pytest.mark.parametrize(
    "attributes",
    (
        {"spans": "1:1"},
        {"{urn:formulafence:archive-safety}audit": "1"},
    ),
)
def test_semantic_reader_preflight_ignores_row_attributes_that_do_not_create_dimensions(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
    attributes: dict[str, str],
) -> None:
    workbook = make_model(tmp_path / "non-dimension-row-attributes.xlsx")
    _append_row_dimension_declarations(workbook, 1, attributes=attributes)
    monkeypatch.setattr(workbook_module, "_OOXML_READER_MAX_ROW_DIMENSION_COUNT", 0)

    snapshot = load_snapshot(workbook)

    assert snapshot.file_type == "xlsx"


def test_semantic_reader_preflight_counts_unknown_unqualified_row_attributes(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    workbook = make_model(tmp_path / "unknown-row-dimension-attribute.xlsx")
    _append_row_dimension_declarations(
        workbook,
        1,
        attributes={"formulafenceAudit": "1"},
    )
    monkeypatch.setattr(workbook_module, "_OOXML_READER_MAX_ROW_DIMENSION_COUNT", 0)

    message = _reject_before_workbook_readers(monkeypatch, workbook)

    assert "row-dimension declarations" in message


def test_semantic_reader_preflight_ignores_foreign_namespace_rows(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    workbook = make_model(tmp_path / "foreign-row-dimensions.xlsx")
    _append_row_dimension_declarations(
        workbook,
        1,
        alternate_namespace="urn:formulafence:archive-safety",
    )
    monkeypatch.setattr(workbook_module, "_OOXML_READER_MAX_ROW_DIMENSION_COUNT", 0)

    snapshot = load_snapshot(workbook)

    assert snapshot.file_type == "xlsx"


def test_semantic_reader_preflight_rejects_excessive_column_dimensions_before_scanners(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    workbook = make_model(tmp_path / "too-many-column-dimensions.xlsx")
    _append_column_dimension_declarations(workbook, 2)
    monkeypatch.setattr(workbook_module, "_OOXML_READER_MAX_COLUMN_DIMENSION_COUNT", 1)

    message = _reject_before_workbook_readers(monkeypatch, workbook)

    assert "column-dimension declarations" in message


def test_semantic_reader_preflight_rejects_default_column_dimension_limit_before_scanners(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    workbook = make_model(tmp_path / "default-column-dimension-limit.xlsx")
    _append_column_dimension_declarations(
        workbook,
        workbook_module._OOXML_READER_MAX_COLUMN_DIMENSION_COUNT + 1,
    )

    message = _reject_before_workbook_readers(monkeypatch, workbook)

    assert "column-dimension declarations" in message


def test_semantic_reader_preflight_counts_column_dimensions_across_worksheets(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    workbook = make_model(tmp_path / "aggregate-column-dimensions.xlsx")
    _append_column_dimension_declarations(workbook, 1)
    _append_column_dimension_declarations(
        workbook,
        1,
        member_name="xl/worksheets/sheet2.xml",
    )
    monkeypatch.setattr(workbook_module, "_OOXML_READER_MAX_COLUMN_DIMENSION_COUNT", 1)

    message = _reject_before_workbook_readers(monkeypatch, workbook)

    assert "column-dimension declarations" in message


def test_semantic_reader_preflight_accepts_column_dimensions_at_the_limit(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    workbook = make_model(tmp_path / "column-dimensions-at-limit.xlsx")
    _append_column_dimension_declarations(workbook, 2)
    monkeypatch.setattr(workbook_module, "_OOXML_READER_MAX_COLUMN_DIMENSION_COUNT", 2)

    snapshot = load_snapshot(workbook)

    assert snapshot.file_type == "xlsx"


def test_semantic_reader_preflight_counts_unknown_unqualified_column_attributes(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    workbook = make_model(tmp_path / "unknown-column-dimension-attribute.xlsx")
    _append_column_dimension_declarations(
        workbook,
        1,
        attributes={"formulafenceAudit": "1"},
    )
    monkeypatch.setattr(workbook_module, "_OOXML_READER_MAX_COLUMN_DIMENSION_COUNT", 0)

    message = _reject_before_workbook_readers(monkeypatch, workbook)

    assert "column-dimension declarations" in message


def test_semantic_reader_preflight_ignores_foreign_namespace_columns(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    workbook = make_model(tmp_path / "foreign-column-dimensions.xlsx")
    _append_column_dimension_declarations(
        workbook,
        1,
        alternate_namespace="urn:formulafence:archive-safety",
    )
    monkeypatch.setattr(workbook_module, "_OOXML_READER_MAX_COLUMN_DIMENSION_COUNT", 0)

    snapshot = load_snapshot(workbook)

    assert snapshot.file_type == "xlsx"


def test_semantic_reader_preflight_rejects_excessive_column_dimension_containers_before_scanners(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    workbook = make_model(tmp_path / "too-many-column-dimension-containers.xlsx")
    _append_column_dimension_containers(workbook, 2)
    monkeypatch.setattr(
        workbook_module,
        "_OOXML_READER_MAX_COLUMN_DIMENSION_CONTAINER_COUNT",
        1,
    )

    message = _reject_before_workbook_readers(monkeypatch, workbook)

    assert "column-dimension containers" in message


def test_semantic_reader_preflight_rejects_default_column_dimension_container_limit_before_scanners(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    workbook = make_model(tmp_path / "default-column-dimension-container-limit.xlsx")
    _append_column_dimension_containers(
        workbook,
        workbook_module._OOXML_READER_MAX_COLUMN_DIMENSION_CONTAINER_COUNT + 1,
    )

    message = _reject_before_workbook_readers(monkeypatch, workbook)

    assert "column-dimension containers" in message


def test_semantic_reader_preflight_counts_column_dimension_containers_across_worksheets(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    workbook = make_model(tmp_path / "aggregate-column-dimension-containers.xlsx")
    _append_column_dimension_containers(workbook, 1)
    _append_column_dimension_containers(
        workbook,
        1,
        member_name="xl/worksheets/sheet2.xml",
    )
    monkeypatch.setattr(
        workbook_module,
        "_OOXML_READER_MAX_COLUMN_DIMENSION_CONTAINER_COUNT",
        1,
    )

    message = _reject_before_workbook_readers(monkeypatch, workbook)

    assert "column-dimension containers" in message


def test_semantic_reader_preflight_accepts_column_dimension_containers_at_the_limit(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    workbook = make_model(tmp_path / "column-dimension-containers-at-limit.xlsx")
    _append_column_dimension_containers(workbook, 2)
    monkeypatch.setattr(
        workbook_module,
        "_OOXML_READER_MAX_COLUMN_DIMENSION_CONTAINER_COUNT",
        2,
    )

    snapshot = load_snapshot(workbook)

    assert snapshot.file_type == "xlsx"


def test_semantic_reader_preflight_ignores_foreign_namespace_column_containers(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    workbook = make_model(tmp_path / "foreign-column-dimension-containers.xlsx")
    _append_column_dimension_containers(
        workbook,
        1,
        alternate_namespace="urn:formulafence:archive-safety",
    )
    monkeypatch.setattr(
        workbook_module,
        "_OOXML_READER_MAX_COLUMN_DIMENSION_CONTAINER_COUNT",
        0,
    )

    snapshot = load_snapshot(workbook)

    assert snapshot.file_type == "xlsx"


def test_semantic_reader_preflight_rejects_excessive_page_break_declarations_before_scanners(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    workbook = make_model(tmp_path / "too-many-page-break-declarations.xlsx")
    _append_page_break_declarations(workbook, 2)
    monkeypatch.setattr(
        workbook_module,
        "_OOXML_READER_MAX_PAGE_BREAK_DECLARATION_COUNT",
        1,
    )

    message = _reject_before_workbook_readers(monkeypatch, workbook)

    assert "page-break declarations" in message


def test_semantic_reader_preflight_rejects_default_page_break_declaration_limit_before_scanners(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    workbook = make_model(tmp_path / "default-page-break-declaration-limit.xlsx")
    _append_page_break_declarations(
        workbook,
        workbook_module._OOXML_READER_MAX_PAGE_BREAK_DECLARATION_COUNT + 1,
    )

    message = _reject_before_workbook_readers(monkeypatch, workbook)

    assert "page-break declarations" in message


def test_semantic_reader_preflight_counts_page_break_declarations_across_worksheets_and_axes(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    workbook = make_model(tmp_path / "aggregate-page-break-declarations.xlsx")
    _append_page_break_declarations(workbook, 1)
    _append_page_break_declarations(
        workbook,
        1,
        axis="column",
        member_name="xl/worksheets/sheet2.xml",
    )
    monkeypatch.setattr(
        workbook_module,
        "_OOXML_READER_MAX_PAGE_BREAK_DECLARATION_COUNT",
        1,
    )

    message = _reject_before_workbook_readers(monkeypatch, workbook)

    assert "page-break declarations" in message


def test_semantic_reader_preflight_accepts_one_complete_published_page_break_allowance(
    tmp_path: Path,
) -> None:
    workbook = make_model(tmp_path / "page-breaks-at-published-limit.xlsx")
    _append_page_break_declarations(workbook, 1_026)
    _append_page_break_declarations(workbook, 1_026, axis="column")

    snapshot = load_snapshot(workbook)

    assert snapshot.file_type == "xlsx"


@pytest.mark.parametrize(
    ("child_name", "alternate_namespace"),
    (
        ("brk", "urn:formulafence:archive-safety"),
        ("formulafenceAudit", None),
    ),
)
def test_semantic_reader_preflight_counts_all_direct_page_break_children(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
    child_name: str,
    alternate_namespace: str | None,
) -> None:
    workbook = make_model(tmp_path / "nonstandard-page-break-child.xlsx")
    _append_page_break_declarations(
        workbook,
        1,
        child_name=child_name,
        alternate_namespace=alternate_namespace,
    )
    monkeypatch.setattr(
        workbook_module,
        "_OOXML_READER_MAX_PAGE_BREAK_DECLARATION_COUNT",
        0,
    )

    message = _reject_before_workbook_readers(monkeypatch, workbook)

    assert "page-break declarations" in message


def test_semantic_reader_preflight_counts_strict_page_break_declarations_before_scanners(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    workbook = make_strict_worksheet_print_layout_model(
        tmp_path / "strict-page-break-declarations.xlsx"
    )
    monkeypatch.setattr(
        workbook_module,
        "_OOXML_READER_MAX_PAGE_BREAK_DECLARATION_COUNT",
        1,
    )

    message = _reject_before_workbook_readers(monkeypatch, workbook)

    assert "page-break declarations" in message


def test_semantic_reader_preflight_rejects_excessive_page_break_containers_before_scanners(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    workbook = make_model(tmp_path / "too-many-page-break-containers.xlsx")
    _append_page_break_containers(workbook, 2)
    monkeypatch.setattr(
        workbook_module,
        "_OOXML_READER_MAX_PAGE_BREAK_CONTAINER_COUNT",
        1,
    )

    message = _reject_before_workbook_readers(monkeypatch, workbook)

    assert "page-break containers" in message


def test_semantic_reader_preflight_rejects_default_page_break_container_limit_before_scanners(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    workbook = make_model(tmp_path / "default-page-break-container-limit.xlsx")
    _append_page_break_containers(
        workbook,
        workbook_module._OOXML_READER_MAX_PAGE_BREAK_CONTAINER_COUNT + 1,
    )

    message = _reject_before_workbook_readers(monkeypatch, workbook)

    assert "page-break containers" in message


def test_semantic_reader_preflight_counts_page_break_containers_across_worksheets_and_axes(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    workbook = make_model(tmp_path / "aggregate-page-break-containers.xlsx")
    _append_page_break_containers(workbook, 1)
    _append_page_break_containers(
        workbook,
        1,
        axis="column",
        member_name="xl/worksheets/sheet2.xml",
    )
    monkeypatch.setattr(
        workbook_module,
        "_OOXML_READER_MAX_PAGE_BREAK_CONTAINER_COUNT",
        1,
    )

    message = _reject_before_workbook_readers(monkeypatch, workbook)

    assert "page-break containers" in message


def test_semantic_reader_preflight_accepts_page_break_containers_at_the_limit(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    workbook = make_model(tmp_path / "page-break-containers-at-limit.xlsx")
    _append_page_break_containers(workbook, 2)
    monkeypatch.setattr(
        workbook_module,
        "_OOXML_READER_MAX_PAGE_BREAK_CONTAINER_COUNT",
        2,
    )

    snapshot = load_snapshot(workbook)

    assert snapshot.file_type == "xlsx"


def test_semantic_reader_preflight_ignores_foreign_namespace_page_break_containers(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    workbook = make_model(tmp_path / "foreign-page-break-containers.xlsx")
    _append_page_break_containers(
        workbook,
        1,
        alternate_namespace="urn:formulafence:archive-safety",
    )
    monkeypatch.setattr(
        workbook_module,
        "_OOXML_READER_MAX_PAGE_BREAK_CONTAINER_COUNT",
        0,
    )

    snapshot = load_snapshot(workbook)

    assert snapshot.file_type == "xlsx"


def test_semantic_reader_preflight_rejects_excessive_content_type_declarations_before_scanners(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    workbook = make_model(tmp_path / "too-many-content-types.xlsx")
    monkeypatch.setattr(
        workbook_module,
        "_OOXML_READER_MAX_CONTENT_TYPE_DECLARATION_COUNT",
        0,
    )

    message = _reject_before_workbook_readers(monkeypatch, workbook)

    assert "content-type declarations" in message


def test_semantic_reader_preflight_rejects_excessive_workbook_relationships_before_scanners(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    workbook = make_model(tmp_path / "too-many-workbook-relationships.xlsx")
    monkeypatch.setattr(
        workbook_module,
        "_OOXML_READER_MAX_WORKBOOK_RELATIONSHIP_COUNT",
        0,
    )

    message = _reject_before_workbook_readers(monkeypatch, workbook)

    assert "workbook relationships" in message


def test_semantic_reader_preflight_rejects_excessive_workbook_sheet_declarations_before_scanners(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    workbook = make_model(tmp_path / "too-many-workbook-sheets.xlsx")
    monkeypatch.setattr(
        workbook_module,
        "_OOXML_READER_MAX_WORKBOOK_SHEET_COUNT",
        0,
    )

    message = _reject_before_workbook_readers(monkeypatch, workbook)

    assert "workbook sheet declarations" in message


def test_semantic_reader_preflight_counts_alternate_namespace_workbook_sheet_declarations(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    workbook = make_model(tmp_path / "alternate-namespace-workbook-sheets.xlsx")
    _append_workbook_sheet_declarations(
        workbook,
        1,
        namespace_prefix="audit",
    )
    monkeypatch.setattr(
        workbook_module,
        "_OOXML_READER_MAX_WORKBOOK_SHEET_COUNT",
        4,
    )

    message = _reject_before_workbook_readers(monkeypatch, workbook)

    assert "workbook sheet declarations" in message


def test_semantic_reader_preflight_rejects_default_sheet_declaration_limit_before_scanners(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    workbook = make_model(tmp_path / "repeated-workbook-sheets.xlsx")
    _append_workbook_sheet_declarations(
        workbook,
        workbook_module._OOXML_READER_MAX_WORKBOOK_SHEET_COUNT,
    )

    message = _reject_before_workbook_readers(monkeypatch, workbook)

    assert "workbook sheet declarations" in message


def test_semantic_reader_preflight_rejects_excessive_workbook_defined_names_before_scanners(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    workbook = make_model(tmp_path / "too-many-workbook-defined-names.xlsx")
    monkeypatch.setattr(
        workbook_module,
        "_OOXML_READER_MAX_WORKBOOK_DEFINED_NAME_COUNT",
        0,
    )

    message = _reject_before_workbook_readers(monkeypatch, workbook)

    assert "workbook defined-name declarations" in message


def test_semantic_reader_preflight_counts_alternate_namespace_workbook_defined_names(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    workbook = make_model(tmp_path / "alternate-namespace-workbook-defined-names.xlsx")
    _append_workbook_defined_name_declarations(
        workbook,
        1,
        namespace_prefix="audit",
    )
    monkeypatch.setattr(
        workbook_module,
        "_OOXML_READER_MAX_WORKBOOK_DEFINED_NAME_COUNT",
        1,
    )

    message = _reject_before_workbook_readers(monkeypatch, workbook)

    assert "workbook defined-name declarations" in message


def test_semantic_reader_preflight_rejects_default_defined_name_declaration_limit_before_scanners(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    workbook = make_model(tmp_path / "repeated-workbook-defined-names.xlsx")
    _append_workbook_defined_name_declarations(
        workbook,
        workbook_module._OOXML_READER_MAX_WORKBOOK_DEFINED_NAME_COUNT,
    )

    message = _reject_before_workbook_readers(monkeypatch, workbook)

    assert "workbook defined-name declarations" in message


def test_semantic_reader_preflight_rejects_excessive_workbook_external_references_before_scanners(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    workbook = make_external_link_package_model(
        tmp_path / "too-many-workbook-external-references.xlsx"
    )
    monkeypatch.setattr(
        workbook_module,
        "_OOXML_READER_MAX_WORKBOOK_EXTERNAL_REFERENCE_COUNT",
        0,
    )

    message = _reject_before_workbook_readers(monkeypatch, workbook)

    assert "workbook external-reference declarations" in message


def test_semantic_reader_preflight_counts_alternate_namespace_workbook_external_references(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    workbook = make_external_link_package_model(
        tmp_path / "alternate-namespace-workbook-external-references.xlsx"
    )
    _append_workbook_relationship_catalog_declarations(
        workbook,
        "externalReferences",
        1,
        alternate_namespace="urn:formulafence:archive-safety",
    )
    monkeypatch.setattr(
        workbook_module,
        "_OOXML_READER_MAX_WORKBOOK_EXTERNAL_REFERENCE_COUNT",
        3,
    )

    message = _reject_before_workbook_readers(monkeypatch, workbook)

    assert "workbook external-reference declarations" in message


def test_semantic_reader_preflight_rejects_default_external_reference_limit_before_scanners(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    workbook = make_external_link_package_model(
        tmp_path / "repeated-workbook-external-references.xlsx"
    )
    _append_workbook_relationship_catalog_declarations(
        workbook,
        "externalReferences",
        workbook_module._OOXML_READER_MAX_WORKBOOK_EXTERNAL_REFERENCE_COUNT,
    )

    message = _reject_before_workbook_readers(monkeypatch, workbook)

    assert "workbook external-reference declarations" in message


def test_semantic_reader_preflight_rejects_excessive_workbook_pivot_caches_before_scanners(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    workbook = make_external_data_refresh_model(
        tmp_path / "too-many-workbook-pivot-caches.xlsx"
    )
    monkeypatch.setattr(
        workbook_module,
        "_OOXML_READER_MAX_WORKBOOK_PIVOT_CACHE_COUNT",
        0,
    )

    message = _reject_before_workbook_readers(monkeypatch, workbook)

    assert "workbook pivot-cache declarations" in message


def test_semantic_reader_preflight_counts_alternate_namespace_workbook_pivot_caches(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    workbook = make_external_data_refresh_model(
        tmp_path / "alternate-namespace-workbook-pivot-caches.xlsx"
    )
    _append_workbook_relationship_catalog_declarations(
        workbook,
        "pivotCaches",
        1,
        alternate_namespace="urn:formulafence:archive-safety",
    )
    monkeypatch.setattr(
        workbook_module,
        "_OOXML_READER_MAX_WORKBOOK_PIVOT_CACHE_COUNT",
        1,
    )

    message = _reject_before_workbook_readers(monkeypatch, workbook)

    assert "workbook pivot-cache declarations" in message


def test_semantic_reader_preflight_rejects_default_pivot_cache_declaration_limit_before_scanners(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    workbook = make_external_data_refresh_model(
        tmp_path / "repeated-workbook-pivot-caches.xlsx"
    )
    _append_workbook_relationship_catalog_declarations(
        workbook,
        "pivotCaches",
        workbook_module._OOXML_READER_MAX_WORKBOOK_PIVOT_CACHE_COUNT,
    )

    message = _reject_before_workbook_readers(monkeypatch, workbook)

    assert "workbook pivot-cache declarations" in message


@pytest.mark.parametrize(
    "catalog_name, child_name, attributes, limit_name, expected_message",
    _WORKBOOK_AUXILIARY_CATALOG_CASES,
)
def test_semantic_reader_preflight_rejects_excessive_auxiliary_workbook_catalogs_before_scanners(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
    catalog_name: str,
    child_name: str,
    attributes: dict[str, str],
    limit_name: str,
    expected_message: str,
) -> None:
    workbook = make_model(tmp_path / f"too-many-{catalog_name}.xlsx")
    _append_workbook_package_catalog_declarations(
        workbook,
        catalog_name,
        child_name,
        1,
        attributes=attributes,
    )
    monkeypatch.setattr(workbook_module, limit_name, 0)

    message = _reject_before_workbook_readers(monkeypatch, workbook)

    assert expected_message in message


@pytest.mark.parametrize(
    "catalog_name, child_name, attributes, limit_name, expected_message",
    _WORKBOOK_AUXILIARY_CATALOG_CASES,
)
def test_semantic_reader_preflight_counts_alternate_namespace_auxiliary_workbook_catalogs(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
    catalog_name: str,
    child_name: str,
    attributes: dict[str, str],
    limit_name: str,
    expected_message: str,
) -> None:
    workbook = make_model(tmp_path / f"alternate-namespace-{catalog_name}.xlsx")
    if catalog_name != "bookViews":
        _append_workbook_package_catalog_declarations(
            workbook,
            catalog_name,
            child_name,
            1,
            attributes=attributes,
        )
    _append_workbook_package_catalog_declarations(
        workbook,
        catalog_name,
        child_name,
        1,
        attributes=attributes,
        alternate_namespace="urn:formulafence:archive-safety",
    )
    monkeypatch.setattr(workbook_module, limit_name, 1)

    message = _reject_before_workbook_readers(monkeypatch, workbook)

    assert expected_message in message


@pytest.mark.parametrize(
    "catalog_name, child_name, attributes, limit_name, expected_message",
    _WORKBOOK_AUXILIARY_CATALOG_CASES,
)
def test_semantic_reader_preflight_rejects_default_auxiliary_workbook_catalogs_before_scanners(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
    catalog_name: str,
    child_name: str,
    attributes: dict[str, str],
    limit_name: str,
    expected_message: str,
) -> None:
    workbook = make_model(tmp_path / f"repeated-{catalog_name}.xlsx")
    _append_workbook_package_catalog_declarations(
        workbook,
        catalog_name,
        child_name,
        getattr(workbook_module, limit_name) + 1,
        attributes=attributes,
    )

    message = _reject_before_workbook_readers(monkeypatch, workbook)

    assert expected_message in message


def test_semantic_reader_preflight_rejects_excessive_custom_sheet_views_before_scanners(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    workbook = make_model(tmp_path / "too-many-custom-sheet-views.xlsx")
    _append_custom_sheet_view_declarations(workbook, 1)
    monkeypatch.setattr(workbook_module, "_OOXML_READER_MAX_CUSTOM_SHEET_VIEW_COUNT", 0)

    message = _reject_before_workbook_readers(monkeypatch, workbook)

    assert "custom sheet-view declarations" in message


def test_semantic_reader_preflight_counts_alternate_namespace_custom_sheet_views(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    workbook = make_model(tmp_path / "alternate-namespace-custom-sheet-views.xlsx")
    _append_custom_sheet_view_declarations(workbook, 1)
    _append_custom_sheet_view_declarations(
        workbook,
        1,
        alternate_namespace="urn:formulafence:archive-safety",
    )
    monkeypatch.setattr(workbook_module, "_OOXML_READER_MAX_CUSTOM_SHEET_VIEW_COUNT", 1)

    message = _reject_before_workbook_readers(monkeypatch, workbook)

    assert "custom sheet-view declarations" in message


def test_semantic_reader_preflight_rejects_default_custom_sheet_view_limit_before_scanners(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    workbook = make_model(tmp_path / "repeated-custom-sheet-views.xlsx")
    _append_custom_sheet_view_declarations(
        workbook,
        workbook_module._OOXML_READER_MAX_CUSTOM_SHEET_VIEW_COUNT + 1,
    )

    message = _reject_before_workbook_readers(monkeypatch, workbook)

    assert "custom sheet-view declarations" in message


def test_semantic_reader_preflight_rejects_excessive_merged_cell_declarations_before_scanners(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    workbook = make_model(tmp_path / "too-many-merged-cell-declarations.xlsx")
    _append_merged_cell_declarations(workbook, ("A1:B1",))
    monkeypatch.setattr(workbook_module, "_OOXML_READER_MAX_MERGED_CELL_RANGE_COUNT", 0)

    message = _reject_before_workbook_readers(monkeypatch, workbook)

    assert "merged-cell declarations" in message


def test_semantic_reader_preflight_counts_alternate_namespace_merged_cell_declarations(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    workbook = make_model(tmp_path / "alternate-namespace-merged-cell-declarations.xlsx")
    _append_merged_cell_declarations(workbook, ("A1:B1",))
    _append_merged_cell_declarations(
        workbook,
        ("C1:D1",),
        alternate_namespace="urn:formulafence:archive-safety",
    )
    monkeypatch.setattr(workbook_module, "_OOXML_READER_MAX_MERGED_CELL_RANGE_COUNT", 1)

    message = _reject_before_workbook_readers(monkeypatch, workbook)

    assert "merged-cell declarations" in message


def test_semantic_reader_preflight_rejects_default_merged_cell_declaration_limit_before_scanners(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    workbook = make_model(tmp_path / "repeated-merged-cell-declarations.xlsx")
    _append_merged_cell_declarations(
        workbook,
        ("A1",) * (workbook_module._OOXML_READER_MAX_MERGED_CELL_RANGE_COUNT + 1),
    )

    message = _reject_before_workbook_readers(monkeypatch, workbook)

    assert "merged-cell declarations" in message


def test_semantic_reader_preflight_rejects_an_oversized_merged_cell_range_before_scanners(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    workbook = make_model(tmp_path / "oversized-merged-cell-range.xlsx")
    _append_merged_cell_declarations(workbook, ("A1:K10",))
    monkeypatch.setattr(workbook_module, "_OOXML_READER_MAX_MERGED_CELL_COUNT", 100)

    message = _reject_before_workbook_readers(monkeypatch, workbook)

    assert "a merged-cell range" in message


def test_semantic_reader_preflight_rejects_excessive_merged_cell_area_before_scanners(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    workbook = make_model(tmp_path / "excessive-merged-cell-area.xlsx")
    _append_merged_cell_declarations(workbook, ("A1:J10", "K1:T10"))
    monkeypatch.setattr(workbook_module, "_OOXML_READER_MAX_MERGED_CELL_COUNT", 150)

    message = _reject_before_workbook_readers(monkeypatch, workbook)

    assert "merged-cell area" in message


def test_semantic_reader_preflight_rejects_a_full_worksheet_merge_before_scanners(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    workbook = make_model(tmp_path / "full-worksheet-merge.xlsx")
    _append_merged_cell_declarations(workbook, ("A1:XFD1048576",))

    message = _reject_before_workbook_readers(monkeypatch, workbook)

    assert "a merged-cell range" in message


def test_semantic_reader_preflight_rejects_an_oversized_merged_cell_reference_before_scanners(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    workbook = make_model(tmp_path / "oversized-merged-cell-reference.xlsx")
    _append_merged_cell_declarations(workbook, ("A1:B1",))
    monkeypatch.setattr(
        workbook_module,
        "_OOXML_READER_MAX_MERGED_CELL_REFERENCE_CHARACTERS",
        4,
    )

    message = _reject_before_workbook_readers(monkeypatch, workbook)

    assert "a merged-cell reference" in message


@pytest.mark.parametrize("reference", ("!", "A1:A0"))
def test_semantic_reader_preflight_rejects_an_unmeasurable_merged_cell_reference_before_scanners(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
    reference: str,
) -> None:
    workbook = make_model(tmp_path / "unmeasurable-merged-cell-reference.xlsx")
    _append_merged_cell_declarations(workbook, (reference,))

    message = _reject_before_workbook_readers(monkeypatch, workbook)

    assert "could not be measured safely" in message


def test_semantic_reader_preflight_accepts_merged_cell_area_at_the_configured_limit(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    workbook = make_model(tmp_path / "merged-cell-area-at-limit.xlsx")
    _append_merged_cell_declarations(workbook, ("Inputs!A1:J10",))
    monkeypatch.setattr(workbook_module, "_OOXML_READER_MAX_MERGED_CELL_COUNT", 100)

    snapshot = load_snapshot(workbook)

    assert snapshot.sheets


def test_semantic_reader_preflight_rejects_excessive_data_validation_declarations_before_scanners(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    workbook = make_model(tmp_path / "too-many-data-validations.xlsx")
    _append_data_validation_declarations(workbook, ("A1",))
    monkeypatch.setattr(workbook_module, "_OOXML_READER_MAX_DATA_VALIDATION_COUNT", 0)

    message = _reject_before_workbook_readers(monkeypatch, workbook)

    assert "data-validation declarations" in message


def test_semantic_reader_preflight_counts_alternate_namespace_data_validation_declarations(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    workbook = make_model(tmp_path / "alternate-namespace-data-validations.xlsx")
    _append_data_validation_declarations(workbook, ("A1",))
    _append_data_validation_declarations(
        workbook,
        ("B1",),
        alternate_namespace="urn:formulafence:archive-safety",
    )
    monkeypatch.setattr(workbook_module, "_OOXML_READER_MAX_DATA_VALIDATION_COUNT", 1)

    message = _reject_before_workbook_readers(monkeypatch, workbook)

    assert "data-validation declarations" in message


def test_semantic_reader_preflight_rejects_default_data_validation_limit_before_scanners(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    workbook = make_model(tmp_path / "repeated-data-validations.xlsx")
    _append_data_validation_declarations(
        workbook,
        ("A1",) * (workbook_module._OOXML_READER_MAX_DATA_VALIDATION_COUNT + 1),
    )

    message = _reject_before_workbook_readers(monkeypatch, workbook)

    assert "data-validation declarations" in message


def test_semantic_reader_preflight_rejects_default_data_validation_reference_limit_before_scanners(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    workbook = make_model(tmp_path / "repeated-data-validation-reference-ranges.xlsx")
    _append_data_validation_declarations(
        workbook,
        (
            " ".join(
                ("A1",)
                * (workbook_module._OOXML_READER_MAX_MULTI_RANGE_REFERENCE_COUNT + 1)
            ),
        ),
    )

    message = _reject_before_workbook_readers(monkeypatch, workbook)

    assert "a data-validation target reference" in message


def test_preflight_rejects_data_validation_reference_range_limit(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    workbook = make_model(tmp_path / "data-validation-reference-ranges.xlsx")
    _append_data_validation_declarations(workbook, ("A1 B1",))
    monkeypatch.setattr(workbook_module, "_OOXML_READER_MAX_MULTI_RANGE_REFERENCE_COUNT", 1)

    message = _reject_before_workbook_readers(monkeypatch, workbook)

    assert "a data-validation target reference" in message


def test_preflight_rejects_aggregate_data_validation_reference_ranges(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    workbook = make_model(tmp_path / "aggregate-data-validation-reference-ranges.xlsx")
    _append_data_validation_declarations(workbook, ("A1", "B1"))
    monkeypatch.setattr(
        workbook_module,
        "_OOXML_READER_MAX_DATA_VALIDATION_TARGET_RANGE_COUNT",
        1,
    )

    message = _reject_before_workbook_readers(monkeypatch, workbook)

    assert "data-validation target ranges" in message


def test_semantic_reader_preflight_rejects_an_oversized_data_validation_reference_before_scanners(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    workbook = make_model(tmp_path / "oversized-data-validation-reference.xlsx")
    _append_data_validation_declarations(workbook, ("A1:B1",))
    monkeypatch.setattr(
        workbook_module,
        "_OOXML_READER_MAX_MULTI_RANGE_REFERENCE_CHARACTERS",
        4,
    )

    message = _reject_before_workbook_readers(monkeypatch, workbook)

    assert "a data-validation target reference" in message


def test_semantic_reader_preflight_rejects_an_oversized_data_validation_formula_before_scanners(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    workbook = make_model(tmp_path / "oversized-data-validation-formula.xlsx")
    _append_data_validation_declarations(workbook, ("A1",), formula="TRUE")
    monkeypatch.setattr(workbook_module, "_OOXML_READER_MAX_FORMULA_CHARACTERS", 3)

    message = _reject_before_workbook_readers(monkeypatch, workbook)

    assert "formula text" in message


def test_semantic_reader_preflight_accepts_data_validation_reference_ranges_at_configured_limits(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    workbook = make_model(tmp_path / "data-validation-reference-ranges-at-limit.xlsx")
    _append_data_validation_declarations(workbook, ("A1 B1", "C1 D1"))
    monkeypatch.setattr(workbook_module, "_OOXML_READER_MAX_MULTI_RANGE_REFERENCE_COUNT", 2)
    monkeypatch.setattr(
        workbook_module,
        "_OOXML_READER_MAX_DATA_VALIDATION_TARGET_RANGE_COUNT",
        4,
    )

    snapshot = load_snapshot(workbook)

    assert snapshot.summary()["data_validation_target_ranges"] == 4


def test_preflight_rejects_conditional_formatting_declaration_limit(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    workbook = make_model(tmp_path / "too-many-conditional-formatting-declarations.xlsx")
    _append_conditional_formatting_declarations(
        workbook,
        ("A1",),
        rule_namespaces=(),
    )
    monkeypatch.setattr(
        workbook_module,
        "_OOXML_READER_MAX_CONDITIONAL_FORMATTING_COUNT",
        0,
    )

    message = _reject_before_workbook_readers(monkeypatch, workbook)

    assert "conditional-formatting declarations" in message


def test_preflight_rejects_default_conditional_formatting_declaration_limit(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    workbook = make_model(tmp_path / "repeated-conditional-formatting-declarations.xlsx")
    _append_conditional_formatting_declarations(
        workbook,
        ("A1",) * (workbook_module._OOXML_READER_MAX_CONDITIONAL_FORMATTING_COUNT + 1),
        rule_namespaces=(),
    )

    message = _reject_before_workbook_readers(monkeypatch, workbook)

    assert "conditional-formatting declarations" in message


def test_semantic_reader_preflight_counts_alternate_namespace_conditional_formatting_rules(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    workbook = make_model(tmp_path / "alternate-namespace-conditional-formatting-rules.xlsx")
    _append_conditional_formatting_declarations(
        workbook,
        ("A1",),
        rule_namespaces=(None, "urn:formulafence:archive-safety"),
    )
    monkeypatch.setattr(
        workbook_module,
        "_OOXML_READER_MAX_CONDITIONAL_FORMATTING_RULE_COUNT",
        1,
    )

    message = _reject_before_workbook_readers(monkeypatch, workbook)

    assert "conditional-formatting rules" in message


def test_preflight_rejects_default_conditional_formatting_rule_limit(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    workbook = make_model(tmp_path / "repeated-conditional-formatting-rules.xlsx")
    _append_conditional_formatting_declarations(
        workbook,
        ("A1",),
        rule_namespaces=(None,)
        * (workbook_module._OOXML_READER_MAX_CONDITIONAL_FORMATTING_RULE_COUNT + 1),
    )

    message = _reject_before_workbook_readers(monkeypatch, workbook)

    assert "conditional-formatting rules" in message


def test_preflight_rejects_conditional_formatting_reference_range_limit(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    workbook = make_model(tmp_path / "conditional-formatting-reference-ranges.xlsx")
    _append_conditional_formatting_declarations(workbook, ("A1 B1",))
    monkeypatch.setattr(workbook_module, "_OOXML_READER_MAX_MULTI_RANGE_REFERENCE_COUNT", 1)

    message = _reject_before_workbook_readers(monkeypatch, workbook)

    assert "a conditional-formatting target reference" in message


def test_preflight_rejects_aggregate_conditional_formatting_reference_ranges(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    workbook = make_model(tmp_path / "aggregate-conditional-formatting-reference-ranges.xlsx")
    _append_conditional_formatting_declarations(workbook, ("A1", "B1"))
    monkeypatch.setattr(
        workbook_module,
        "_OOXML_READER_MAX_CONDITIONAL_FORMATTING_TARGET_RANGE_COUNT",
        1,
    )

    message = _reject_before_workbook_readers(monkeypatch, workbook)

    assert "conditional-formatting target ranges" in message


def test_preflight_rejects_oversized_conditional_formatting_reference(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    workbook = make_model(tmp_path / "oversized-conditional-formatting-reference.xlsx")
    _append_conditional_formatting_declarations(workbook, ("A1:B1",))
    monkeypatch.setattr(
        workbook_module,
        "_OOXML_READER_MAX_MULTI_RANGE_REFERENCE_CHARACTERS",
        4,
    )

    message = _reject_before_workbook_readers(monkeypatch, workbook)

    assert "a conditional-formatting target reference" in message


def test_preflight_rejects_oversized_conditional_formatting_formula(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    workbook = make_model(tmp_path / "oversized-conditional-formatting-formula.xlsx")
    _append_conditional_formatting_declarations(workbook, ("A1",), formula="TRUE")
    monkeypatch.setattr(workbook_module, "_OOXML_READER_MAX_FORMULA_CHARACTERS", 3)

    message = _reject_before_workbook_readers(monkeypatch, workbook)

    assert "formula text" in message


def test_preflight_accepts_conditional_formatting_reference_ranges_at_limit(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    workbook = make_model(tmp_path / "conditional-formatting-reference-ranges-at-limit.xlsx")
    _append_conditional_formatting_declarations(workbook, ("A1 B1", "C1 D1"))
    monkeypatch.setattr(workbook_module, "_OOXML_READER_MAX_MULTI_RANGE_REFERENCE_COUNT", 2)
    monkeypatch.setattr(
        workbook_module,
        "_OOXML_READER_MAX_CONDITIONAL_FORMATTING_TARGET_RANGE_COUNT",
        4,
    )

    snapshot = load_snapshot(workbook)

    assert len(snapshot.conditional_formatting) == 2


def test_semantic_reader_preflight_rejects_excessive_scenario_manager_containers_before_scanners(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    workbook = make_model(tmp_path / "too-many-scenario-manager-containers.xlsx")
    _append_scenario_manager_containers(workbook, ("A1",), scenario_count=0)
    monkeypatch.setattr(
        workbook_module,
        "_OOXML_READER_MAX_SCENARIO_CONTAINER_COUNT",
        0,
    )

    message = _reject_before_workbook_readers(monkeypatch, workbook)

    assert "Scenario Manager containers" in message


def test_semantic_reader_preflight_rejects_default_scenario_manager_container_limit_before_scanners(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    workbook = make_model(tmp_path / "repeated-scenario-manager-containers.xlsx")
    _append_scenario_manager_containers(
        workbook,
        ("A1",) * (workbook_module._OOXML_READER_MAX_SCENARIO_CONTAINER_COUNT + 1),
        scenario_count=0,
    )

    message = _reject_before_workbook_readers(monkeypatch, workbook)

    assert "Scenario Manager containers" in message


def test_semantic_reader_preflight_counts_alternate_namespace_scenario_manager_declarations(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    workbook = make_model(tmp_path / "alternate-namespace-scenario-manager-declarations.xlsx")
    _append_scenario_manager_containers(workbook, ("A1",))
    _append_scenario_manager_containers(
        workbook,
        ("B1",),
        alternate_namespace="urn:formulafence:archive-safety",
    )
    monkeypatch.setattr(workbook_module, "_OOXML_READER_MAX_SCENARIO_COUNT", 1)

    message = _reject_before_workbook_readers(monkeypatch, workbook)

    assert "Scenario Manager declarations" in message


def test_semantic_reader_preflight_counts_alternate_namespace_scenario_manager_input_cells(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    workbook = make_model(tmp_path / "alternate-namespace-scenario-manager-input-cells.xlsx")
    _append_scenario_manager_containers(workbook, ("A1",))
    _append_scenario_manager_containers(
        workbook,
        ("B1",),
        alternate_namespace="urn:formulafence:archive-safety",
    )
    monkeypatch.setattr(workbook_module, "_OOXML_READER_MAX_SCENARIO_INPUT_CELL_COUNT", 1)

    message = _reject_before_workbook_readers(monkeypatch, workbook)

    assert "Scenario Manager input cells" in message


def test_semantic_reader_preflight_rejects_default_scenario_manager_limit_before_scanners(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    workbook = make_model(tmp_path / "repeated-scenario-manager-declarations.xlsx")
    _append_scenario_manager_containers(
        workbook,
        ("A1",),
        scenario_count=workbook_module._OOXML_READER_MAX_SCENARIO_COUNT + 1,
        input_cell_count=0,
    )

    message = _reject_before_workbook_readers(monkeypatch, workbook)

    assert "Scenario Manager declarations" in message


def test_preflight_rejects_default_scenario_manager_input_cell_limit(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    workbook = make_model(tmp_path / "repeated-scenario-manager-input-cells.xlsx")
    _append_scenario_manager_containers(
        workbook,
        ("A1",),
        input_cell_count=workbook_module._OOXML_READER_MAX_SCENARIO_INPUT_CELL_COUNT + 1,
    )

    message = _reject_before_workbook_readers(monkeypatch, workbook)

    assert "Scenario Manager input cells" in message


def test_preflight_rejects_scenario_manager_reference_range_limit(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    workbook = make_model(tmp_path / "scenario-manager-reference-ranges.xlsx")
    _append_scenario_manager_containers(workbook, ("A1 B1",), scenario_count=0)
    monkeypatch.setattr(workbook_module, "_OOXML_READER_MAX_MULTI_RANGE_REFERENCE_COUNT", 1)

    message = _reject_before_workbook_readers(monkeypatch, workbook)

    assert "a Scenario Manager target reference" in message


def test_preflight_rejects_aggregate_scenario_manager_reference_ranges(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    workbook = make_model(tmp_path / "aggregate-scenario-manager-reference-ranges.xlsx")
    _append_scenario_manager_containers(workbook, ("A1", "B1"), scenario_count=0)
    monkeypatch.setattr(
        workbook_module,
        "_OOXML_READER_MAX_SCENARIO_TARGET_RANGE_COUNT",
        1,
    )

    message = _reject_before_workbook_readers(monkeypatch, workbook)

    assert "Scenario Manager target ranges" in message


def test_semantic_reader_preflight_rejects_an_oversized_scenario_manager_reference_before_scanners(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    workbook = make_model(tmp_path / "oversized-scenario-manager-reference.xlsx")
    _append_scenario_manager_containers(workbook, ("A1:B1",), scenario_count=0)
    monkeypatch.setattr(
        workbook_module,
        "_OOXML_READER_MAX_MULTI_RANGE_REFERENCE_CHARACTERS",
        4,
    )

    message = _reject_before_workbook_readers(monkeypatch, workbook)

    assert "a Scenario Manager target reference" in message


def test_semantic_reader_preflight_accepts_scenario_manager_reference_ranges_at_configured_limits(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    workbook = make_model(tmp_path / "scenario-manager-reference-ranges-at-limit.xlsx")
    _append_scenario_manager_containers(workbook, ("A1 B1",))
    monkeypatch.setattr(workbook_module, "_OOXML_READER_MAX_MULTI_RANGE_REFERENCE_COUNT", 2)
    monkeypatch.setattr(
        workbook_module,
        "_OOXML_READER_MAX_SCENARIO_TARGET_RANGE_COUNT",
        2,
    )

    snapshot = load_snapshot(workbook)

    assert snapshot.scenario_manager.summary_reference_count == 2


def test_semantic_reader_preflight_rejects_excessive_xml_nesting_before_scanners(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    workbook = make_model(tmp_path / "deep-reader-xml.xlsx")
    monkeypatch.setattr(workbook_module, "_OOXML_READER_MAX_XML_NESTING_DEPTH", 1)

    message = _reject_before_workbook_readers(monkeypatch, workbook)

    assert "XML nesting" in message


def test_semantic_reader_preflight_rejects_excessive_xml_elements_before_scanners(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    workbook = make_model(tmp_path / "many-reader-elements.xlsx")
    monkeypatch.setattr(workbook_module, "_OOXML_READER_MAX_XML_ELEMENT_COUNT", 1)

    message = _reject_before_workbook_readers(monkeypatch, workbook)

    assert "XML element count" in message


def test_semantic_reader_preflight_rejects_excessive_cell_styles_before_scanners(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    workbook = make_model(tmp_path / "too-many-styles.xlsx")
    monkeypatch.setattr(workbook_module, "_OOXML_READER_MAX_CELL_STYLE_COUNT", 0)

    message = _reject_before_workbook_readers(monkeypatch, workbook)

    assert "cell styles" in message


def test_semantic_reader_preflight_rejects_excessive_stylesheet_containers_before_scanners(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    workbook = make_model(tmp_path / "too-many-stylesheet-containers.xlsx")
    monkeypatch.setattr(workbook_module, "_OOXML_READER_MAX_STYLESHEET_CONTAINER_COUNT", 0)

    message = _reject_before_workbook_readers(monkeypatch, workbook)

    assert "stylesheet containers" in message


@pytest.mark.parametrize(
    ("parent_names", "child_name", "limit_name", "message_fragment"),
    _STYLESHEET_CATALOG_CASES,
)
def test_semantic_reader_preflight_rejects_excessive_stylesheet_catalogs_before_scanners(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
    parent_names: tuple[str, ...],
    child_name: str,
    limit_name: str,
    message_fragment: str,
) -> None:
    workbook = make_model(tmp_path / f"too-many-stylesheet-{child_name}.xlsx")
    record_count = _append_stylesheet_catalog_declarations(
        workbook,
        parent_names,
        child_name,
        1,
    )
    monkeypatch.setattr(workbook_module, limit_name, record_count - 1)

    message = _reject_before_workbook_readers(monkeypatch, workbook)

    assert message_fragment in message


def test_semantic_reader_preflight_rejects_default_font_catalog_over_limit_before_scanners(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    workbook = make_model(tmp_path / "too-many-font-records.xlsx")
    _append_stylesheet_catalog_declarations(
        workbook,
        ("fonts",),
        "font",
        workbook_module._OOXML_READER_MAX_FONT_COUNT,
    )

    message = _reject_before_workbook_readers(monkeypatch, workbook)

    assert "font records" in message


def test_semantic_reader_preflight_accepts_font_catalog_at_configured_limit(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    workbook = make_model(tmp_path / "font-records-at-limit.xlsx")
    record_count = _append_stylesheet_catalog_declarations(
        workbook,
        ("fonts",),
        "font",
        1,
    )
    monkeypatch.setattr(workbook_module, "_OOXML_READER_MAX_FONT_COUNT", record_count)

    snapshot = load_snapshot(workbook)

    assert snapshot.file_type == "xlsx"


def test_semantic_reader_preflight_accepts_font_catalog_at_default_limit(
    tmp_path: Path,
) -> None:
    workbook = make_model(tmp_path / "font-records-at-default-limit.xlsx")
    existing_count = _append_stylesheet_catalog_declarations(
        workbook,
        ("fonts",),
        "font",
        0,
    )
    record_count = _append_stylesheet_catalog_declarations(
        workbook,
        ("fonts",),
        "font",
        workbook_module._OOXML_READER_MAX_FONT_COUNT - existing_count,
    )

    assert record_count == workbook_module._OOXML_READER_MAX_FONT_COUNT
    assert load_snapshot(workbook).file_type == "xlsx"


def test_semantic_reader_preflight_counts_alternate_namespace_cell_styles_before_scanners(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    workbook = make_model(tmp_path / "alternate-cell-styles.xlsx")
    _append_stylesheet_catalog_declarations(
        workbook,
        ("cellXfs",),
        "xf",
        0,
        alternate_parent_namespace="urn:formulafence:alternate-style",
    )
    monkeypatch.setattr(workbook_module, "_OOXML_READER_MAX_CELL_STYLE_COUNT", 0)

    message = _reject_before_workbook_readers(monkeypatch, workbook)

    assert "cell styles" in message


def test_semantic_reader_preflight_counts_unknown_font_children_before_scanners(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    workbook = make_model(tmp_path / "unexpected-font-record.xlsx")
    record_count = _append_stylesheet_catalog_declarations(
        workbook,
        ("fonts",),
        "unexpectedFontRecord",
        1,
    )
    monkeypatch.setattr(workbook_module, "_OOXML_READER_MAX_FONT_COUNT", record_count - 1)

    message = _reject_before_workbook_readers(monkeypatch, workbook)

    assert "font records" in message


def test_semantic_reader_preflight_rejects_excessive_cell_text_before_scanners(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    workbook = make_model(tmp_path / "too-much-cell-text.xlsx")
    monkeypatch.setattr(workbook_module, "_OOXML_READER_MAX_CELL_TEXT_CHARACTERS", 1)

    message = _reject_before_workbook_readers(monkeypatch, workbook)

    assert "cell text" in message


def test_semantic_reader_preflight_rejects_excessive_formula_text_before_scanners(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    workbook = make_model(tmp_path / "too-much-formula-text.xlsx")
    monkeypatch.setattr(workbook_module, "_OOXML_READER_MAX_FORMULA_CHARACTERS", 1)

    message = _reject_before_workbook_readers(monkeypatch, workbook)

    assert "formula text" in message


def test_semantic_reader_preflight_rejects_cell_text_beyond_excel_limit(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    workbook = make_model(tmp_path / "cell-text-limit.xlsx")
    member_name = "xl/worksheets/sheet1.xml"
    with ZipFile(workbook) as archive:
        worksheet = archive.read(member_name)
    _replace_member(
        workbook,
        member_name,
        worksheet.replace(
            b"<t>Revenue</t>",
            b"<t>" + b"x" * 32_768 + b"</t>",
        ),
    )

    message = _reject_before_workbook_readers(monkeypatch, workbook)

    assert "cell text" in message


def test_semantic_reader_preflight_rejects_formula_text_beyond_excel_limit(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    workbook = make_model(tmp_path / "formula-text-limit.xlsx")
    member_name = "xl/worksheets/sheet2.xml"
    with ZipFile(workbook) as archive:
        worksheet = archive.read(member_name)
    _replace_member(
        workbook,
        member_name,
        worksheet.replace(
            b"<f>Inputs!B2*2</f>",
            b"<f>=" + b"1" * 8_192 + b"</f>",
        ),
    )

    message = _reject_before_workbook_readers(monkeypatch, workbook)

    assert "formula text" in message


def test_semantic_reader_preflight_rejects_excessive_shared_strings_before_scanners(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    workbook = make_model(tmp_path / "too-many-shared-strings.xlsx")
    _append_member(
        workbook,
        "xl/sharedStrings.xml",
        (
            b'<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
            b"<si><t>one</t></si><si><t>two</t></si></sst>"
        ),
    )
    monkeypatch.setattr(workbook_module, "_OOXML_READER_MAX_SHARED_STRING_COUNT", 1)

    message = _reject_before_workbook_readers(monkeypatch, workbook)

    assert "shared-string table entries" in message


def test_semantic_reader_preflight_follows_relationship_selected_shared_strings(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    workbook = make_model(tmp_path / "relationship-selected-shared-strings.xlsx")
    relationship_member = "xl/_rels/workbook.xml.rels"
    shared_member = "xl/private/strings.xml"
    with ZipFile(workbook) as archive:
        relationships = archive.read(relationship_member)
    _append_member(
        workbook,
        shared_member,
        (
            b'<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
            b"<si><t>one</t></si><si><t>two</t></si></sst>"
        ),
    )
    _replace_member(
        workbook,
        relationship_member,
        relationships.replace(
            b"</Relationships>",
            (
                b'<Relationship Id="rIdSharedStrings" '
                b'Type="http://schemas.openxmlformats.org/officeDocument/2006/'
                b'relationships/sharedStrings" Target="private/strings.xml" />'
                b"</Relationships>"
            ),
        ),
    )
    monkeypatch.setattr(workbook_module, "_OOXML_READER_MAX_SHARED_STRING_COUNT", 1)

    message = _reject_before_workbook_readers(monkeypatch, workbook)

    assert "shared-string table entries" in message


def test_semantic_reader_preflight_follows_manifest_selected_shared_strings(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    workbook = make_model(tmp_path / "manifest-selected-shared-strings.xlsx")
    content_types_member = "[Content_Types].xml"
    shared_member = "xl/private/strings.xml"
    with ZipFile(workbook) as archive:
        content_types = archive.read(content_types_member)
    _append_member(
        workbook,
        shared_member,
        (
            b'<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
            b"<si><t>one</t></si><si><t>two</t></si></sst>"
        ),
    )
    _replace_member(
        workbook,
        content_types_member,
        content_types.replace(
            b"</Types>",
            (
                b'<Override PartName="/xl/private/strings.xml" '
                b'ContentType="application/vnd.openxmlformats-officedocument.'
                b'spreadsheetml.sharedStrings+xml" />'
                b"</Types>"
            ),
        ),
    )
    monkeypatch.setattr(workbook_module, "_OOXML_READER_MAX_SHARED_STRING_COUNT", 1)

    message = _reject_before_workbook_readers(monkeypatch, workbook)

    assert "shared-string table entries" in message


def test_semantic_reader_preflight_counts_relationship_selected_worksheet_parts(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    workbook = make_model(tmp_path / "relationship-selected-cells.xlsx")
    source_member = "xl/worksheets/sheet1.xml"
    selected_member = "xl/private/sheet-one.xml"
    relationship_member = "xl/_rels/workbook.xml.rels"
    with ZipFile(workbook) as archive:
        worksheet_payload = archive.read(source_member)
        relationships = archive.read(relationship_member)
    assert b"worksheets/sheet1.xml" in relationships
    _append_member(workbook, selected_member, worksheet_payload)
    _replace_member(
        workbook,
        relationship_member,
        relationships.replace(b"worksheets/sheet1.xml", b"private/sheet-one.xml"),
    )
    monkeypatch.setattr(workbook_module, "_OOXML_READER_MAX_WORKSHEET_CELL_COUNT", 1)

    message = _reject_before_workbook_readers(monkeypatch, workbook)

    assert "populated worksheet cells" in message


def test_semantic_reader_preflight_rejects_xml_entity_declarations_before_scanners(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    workbook = make_model(tmp_path / "xml-entity.xlsx")
    member_name = "xl/worksheets/sheet1.xml"
    with ZipFile(workbook) as archive:
        original = archive.read(member_name)
    _replace_member(
        workbook,
        member_name,
        b"<!DOCTYPE worksheet [<!ENTITY forbidden 'private'>]>\n" + original,
    )

    message = _reject_before_workbook_readers(monkeypatch, workbook)

    assert "workbook sheet metadata could not be scanned safely" in message


def test_macro_hash_streams_the_payload_without_zipfile_read(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    workbook = make_model(tmp_path / "macro-stream.xlsx")
    payload = b"private macro payload" * 4096
    _append_member(workbook, "xl/vbaProject.bin", payload)

    original_read = ZipFile.read

    def reject_macro_read(self: ZipFile, name, *args, **kwargs):
        if name == "xl/vbaProject.bin":
            raise AssertionError("macro hashing loaded the complete payload")
        return original_read(self, name, *args, **kwargs)

    monkeypatch.setattr(ZipFile, "read", reject_macro_read)

    assert workbook_module._vba_hash(workbook) == hashlib.sha256(payload).hexdigest()


@pytest.mark.parametrize(
    ("member_name", "expected_detail"),
    [
        ("../not-a-workbook-part.xml", "unsafe or non-canonical"),
        ("XL/WORKBOOK.XML", "ambiguous"),
    ],
)
def test_archive_preflight_rejects_ambiguous_or_unsafe_member_paths(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
    member_name: str,
    expected_detail: str,
) -> None:
    workbook = make_model(tmp_path / "unsafe-member-path.xlsx")
    _append_member(workbook, member_name, b"untrusted")

    message = _reject_before_workbook_readers(monkeypatch, workbook)

    assert expected_detail in message
    assert member_name not in message


def test_archive_preflight_rejects_duplicate_zip_members_before_readers(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    workbook = make_model(tmp_path / "duplicate-member.xlsx")
    with ZipFile(workbook) as archive:
        workbook_payload = archive.read("xl/workbook.xml")
    with pytest.warns(UserWarning, match="Duplicate name"):
        _append_member(workbook, "xl/workbook.xml", workbook_payload)

    message = _reject_before_workbook_readers(monkeypatch, workbook)

    assert "ambiguous" in message
    assert "workbook.xml" not in message


def test_archive_preflight_rejects_encrypted_central_member_before_readers(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    workbook = make_model(tmp_path / "encrypted-member.xlsx")
    contents = bytearray(workbook.read_bytes())
    central_directory_offset = _last_central_directory_offset(contents)
    flag_offset = central_directory_offset + 8
    flag_bits = int.from_bytes(contents[flag_offset : flag_offset + 2], "little")
    contents[flag_offset : flag_offset + 2] = (flag_bits | 0x1).to_bytes(2, "little")
    workbook.write_bytes(contents)

    assert "encrypted" in _reject_before_workbook_readers(monkeypatch, workbook)


def test_archive_preflight_rejects_unicode_path_aliases_before_readers(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    workbook = make_model(tmp_path / "unicode-path-alias.xlsx")
    original_name = "xl/media/original.bin"
    unicode_alias = "xl/media/alias.bin"
    entry = ZipInfo(original_name)
    entry.extra = struct.pack(
        "<HHBI",
        0x7075,
        5 + len(unicode_alias.encode("utf-8")),
        1,
        binascii.crc32(original_name.encode("ascii")),
    ) + unicode_alias.encode("utf-8")
    with ZipFile(workbook, "a") as archive:
        archive.writestr(entry, b"untrusted")

    assert "Unicode-path aliases" in _reject_before_workbook_readers(
        monkeypatch,
        workbook,
    )


def test_archive_preflight_rejects_malformed_zip64_member_metadata_before_readers(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    workbook = make_model(tmp_path / "malformed-zip64.xlsx")
    contents = bytearray(workbook.read_bytes())
    central_directory_offset = _last_central_directory_offset(contents)
    uncompressed_size_offset = central_directory_offset + 24
    contents[uncompressed_size_offset : uncompressed_size_offset + 4] = b"\xff" * 4
    workbook.write_bytes(contents)

    assert "ZIP64" in _reject_before_workbook_readers(monkeypatch, workbook)


def test_archive_preflight_rejects_local_header_mismatch_before_readers(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    workbook = make_model(tmp_path / "local-header-mismatch.xlsx")
    contents = bytearray(workbook.read_bytes())
    local_flag_offset = 6
    flag_bits = int.from_bytes(contents[local_flag_offset : local_flag_offset + 2], "little")
    contents[local_flag_offset : local_flag_offset + 2] = (flag_bits | 0x8).to_bytes(
        2,
        "little",
    )
    workbook.write_bytes(contents)

    assert "local member metadata is inconsistent" in _reject_before_workbook_readers(
        monkeypatch,
        workbook,
    )


def test_archive_preflight_rejects_symbolic_link_members_before_readers(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    workbook = make_model(tmp_path / "symbolic-link.xlsx")
    link = ZipInfo("xl/media/link")
    link.create_system = 3
    link.external_attr = (stat.S_IFLNK | 0o777) << 16
    with ZipFile(workbook, "a") as archive:
        archive.writestr(link, b"not a real part")

    assert "symbolic-link" in _reject_before_workbook_readers(monkeypatch, workbook)


def test_cli_surfaces_archive_preflight_as_an_input_error(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
    capsys: pytest.CaptureFixture[str],
) -> None:
    workbook = make_model(tmp_path / "cli-limit.xlsx")
    monkeypatch.setattr(workbook_module, "_OOXML_ARCHIVE_MAX_ENTRY_COUNT", 1)

    assert main(["profile", str(workbook)]) == 2

    assert "safety preflight" in capsys.readouterr().err


def test_cli_surfaces_malformed_openpyxl_cell_metadata_as_an_input_error(
    tmp_path: Path,
    capsys: pytest.CaptureFixture[str],
) -> None:
    workbook = make_model(tmp_path / "empty-style.xlsx")
    member_name = "xl/worksheets/sheet1.xml"
    with ZipFile(workbook) as archive:
        worksheet = archive.read(member_name)
    _replace_member(
        workbook,
        member_name,
        worksheet.replace(b"<c ", b'<c s="" ', 1),
    )

    assert main(["profile", str(workbook)]) == 2

    assert "Could not read workbook" in capsys.readouterr().err


def test_load_snapshot_surfaces_openpyxl_index_errors_as_input_errors(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    workbook = make_model(tmp_path / "reader-index-error.xlsx")

    def malformed_reader(*args, **kwargs):
        raise IndexError("malformed workbook metadata")

    monkeypatch.setattr(workbook_module, "load_workbook", malformed_reader)

    with pytest.raises(WorkbookLoadError, match="Could not read workbook"):
        load_snapshot(workbook)
