"""Bounded primitives for reading formula-bearing XLSB binary records.

This module deliberately starts below workbook semantics.  XLSB stores cell
formulas as BIFF12 reverse-polish token streams inside a ZIP package, so a
safe workbook adapter needs strict record framing and a decoder that declines
unknown constructs instead of inventing an Excel formula string.  The public
helpers here do no calculation, do not follow relationships, and never expose
raw payloads in exceptions.

The supported token subset is intentionally explicit.  A caller must treat an
``XlsbFormulaUnsupportedError`` as an incomplete-coverage condition, not as a
formula that is safe to silently omit.
"""

from __future__ import annotations

import io
import math
import re
import struct
from collections.abc import Iterator, Mapping, Sequence
from dataclasses import dataclass
from typing import Any, BinaryIO

from defusedxml import ElementTree
from defusedxml.common import DefusedXmlException

from formulafence.models import WorkbookLoadError

# The XLSB record header encodes its type in at most two bytes and its payload
# size in at most four 7-bit groups.  The workbook-level reader will set
# narrower limits when a component has a tighter documented boundary.
DEFAULT_MAX_XLSB_RECORD_BYTES = 16 * 1024 * 1024
DEFAULT_MAX_XLSB_RECORDS = 1_000_000
DEFAULT_MAX_XLSB_FORMULA_BYTES = 1 * 1024 * 1024
DEFAULT_MAX_XLSB_FORMULA_TOKENS = 65_536
DEFAULT_MAX_XLSB_FORMULA_STACK_ITEMS = 8_192
DEFAULT_MAX_XLSB_FORMULA_CHARS = 1 * 1024 * 1024
DEFAULT_MAX_XLSB_FUNCTION_ARGS = 255
DEFAULT_MAX_XLSB_BINARY_PART_BYTES = 64 * 1024 * 1024
DEFAULT_MAX_XLSB_TOTAL_BINARY_BYTES = 256 * 1024 * 1024
DEFAULT_MAX_XLSB_RELATIONSHIP_PART_BYTES = 1 * 1024 * 1024
DEFAULT_MAX_XLSB_SHEETS = 512
DEFAULT_MAX_XLSB_DEFINED_NAMES = 100_000
DEFAULT_MAX_XLSB_WORKSHEET_CELLS = 500_000
DEFAULT_MAX_XLSB_TOTAL_WORKSHEET_CELLS = 500_000
DEFAULT_MAX_XLSB_SHARED_STRINGS = 500_000
DEFAULT_MAX_XLSB_CELL_TEXT_CHARACTERS = 32_767
DEFAULT_MAX_XLSB_RELATIONSHIPS = 4_096
DEFAULT_MAX_XLSB_TOTAL_FORMULA_CHARACTERS = 64 * 1024 * 1024

_MAX_EXCEL_ROW_INDEX = 1_048_575
_MAX_EXCEL_COLUMN_INDEX = 16_383


class XlsbParseError(WorkbookLoadError):
    """An XLSB binary record was malformed or exceeded a local limit."""


class XlsbLimitError(XlsbParseError):
    """An XLSB component exceeded an explicit retained-work limit."""


class XlsbFormulaError(XlsbParseError):
    """A formula token stream could not be converted without ambiguity."""


class XlsbFormulaUnsupportedError(XlsbFormulaError):
    """A valid-looking token construct is outside the verified decoder subset."""


@dataclass(frozen=True)
class XlsbReaderLimits:
    """Resource ceilings for the core XLSB binary reader.

    Package ZIP validation remains a separate concern.  These limits protect
    the binary record and semantic objects retained after a package member has
    been selected for inspection.
    """

    max_binary_part_bytes: int = DEFAULT_MAX_XLSB_BINARY_PART_BYTES
    max_total_binary_bytes: int = DEFAULT_MAX_XLSB_TOTAL_BINARY_BYTES
    max_relationship_part_bytes: int = DEFAULT_MAX_XLSB_RELATIONSHIP_PART_BYTES
    max_record_bytes: int = DEFAULT_MAX_XLSB_RECORD_BYTES
    max_records_per_part: int = DEFAULT_MAX_XLSB_RECORDS
    max_sheets: int = DEFAULT_MAX_XLSB_SHEETS
    max_defined_names: int = DEFAULT_MAX_XLSB_DEFINED_NAMES
    max_worksheet_cells: int = DEFAULT_MAX_XLSB_WORKSHEET_CELLS
    max_total_worksheet_cells: int = DEFAULT_MAX_XLSB_TOTAL_WORKSHEET_CELLS
    max_shared_strings: int = DEFAULT_MAX_XLSB_SHARED_STRINGS
    max_cell_text_characters: int = DEFAULT_MAX_XLSB_CELL_TEXT_CHARACTERS
    max_relationships: int = DEFAULT_MAX_XLSB_RELATIONSHIPS
    max_formula_bytes: int = DEFAULT_MAX_XLSB_FORMULA_BYTES
    max_formula_tokens: int = DEFAULT_MAX_XLSB_FORMULA_TOKENS
    max_formula_stack_items: int = DEFAULT_MAX_XLSB_FORMULA_STACK_ITEMS
    max_formula_characters: int = DEFAULT_MAX_XLSB_FORMULA_CHARS
    max_total_formula_characters: int = DEFAULT_MAX_XLSB_TOTAL_FORMULA_CHARACTERS
    max_function_args: int = DEFAULT_MAX_XLSB_FUNCTION_ARGS

    def validate(self) -> None:
        positive_values = {
            "max_binary_part_bytes": self.max_binary_part_bytes,
            "max_total_binary_bytes": self.max_total_binary_bytes,
            "max_relationship_part_bytes": self.max_relationship_part_bytes,
            "max_record_bytes": self.max_record_bytes,
            "max_records_per_part": self.max_records_per_part,
            "max_sheets": self.max_sheets,
            "max_defined_names": self.max_defined_names,
            "max_worksheet_cells": self.max_worksheet_cells,
            "max_total_worksheet_cells": self.max_total_worksheet_cells,
            "max_shared_strings": self.max_shared_strings,
            "max_cell_text_characters": self.max_cell_text_characters,
            "max_relationships": self.max_relationships,
            "max_formula_bytes": self.max_formula_bytes,
            "max_formula_tokens": self.max_formula_tokens,
            "max_formula_stack_items": self.max_formula_stack_items,
            "max_formula_characters": self.max_formula_characters,
            "max_total_formula_characters": self.max_total_formula_characters,
        }
        for name, value in positive_values.items():
            if value < 1:
                raise ValueError(f"{name} must be at least 1.")
        if not 0 <= self.max_function_args <= 255:
            raise ValueError("max_function_args must be between 0 and 255.")
        if self.max_record_bytes > self.max_binary_part_bytes:
            raise ValueError("max_record_bytes must not exceed max_binary_part_bytes.")
        if self.max_binary_part_bytes > self.max_total_binary_bytes:
            raise ValueError(
                "max_binary_part_bytes must not exceed max_total_binary_bytes."
            )
        if self.max_relationship_part_bytes > self.max_binary_part_bytes:
            raise ValueError(
                "max_relationship_part_bytes must not exceed max_binary_part_bytes."
            )
        if self.max_worksheet_cells > self.max_total_worksheet_cells:
            raise ValueError(
                "max_worksheet_cells must not exceed max_total_worksheet_cells."
            )
        if self.max_formula_characters > self.max_total_formula_characters:
            raise ValueError(
                "max_formula_characters must not exceed max_total_formula_characters."
            )


DEFAULT_XLSB_READER_LIMITS = XlsbReaderLimits()


@dataclass(frozen=True)
class XlsbRecord:
    """One fully framed XLSB record.

    ``offset`` is byte-relative to the beginning of the binary part.  It is
    useful for internal diagnostics but callers must not put it in user-facing
    formula reports alongside sensitive workbook material.
    """

    record_type: int
    payload: bytes
    offset: int


class _Reader:
    """A small bounds-checked little-endian reader over one private payload."""

    def __init__(self, payload: bytes, *, context: str) -> None:
        self._payload = payload
        self._context = context
        self._offset = 0

    @property
    def remaining(self) -> int:
        return len(self._payload) - self._offset

    def take(self, count: int) -> bytes:
        if count < 0 or count > self.remaining:
            raise XlsbParseError(f"Truncated {self._context} record.")
        start = self._offset
        self._offset += count
        return self._payload[start : start + count]

    def u8(self) -> int:
        return self.take(1)[0]

    def u16(self) -> int:
        return struct.unpack("<H", self.take(2))[0]

    def u32(self) -> int:
        return struct.unpack("<I", self.take(4))[0]

    def f64(self) -> float:
        return struct.unpack("<d", self.take(8))[0]

    def utf16(self, characters: int, *, max_characters: int) -> str:
        if characters > max_characters:
            raise XlsbLimitError(f"{self._context} string exceeds its safety limit.")
        try:
            return self.take(characters * 2).decode("utf-16le", errors="strict")
        except UnicodeDecodeError as error:
            raise XlsbParseError(f"Invalid UTF-16 data in {self._context} record.") from error


def _read_header_byte(stream: BinaryIO, *, required: bool, context: str) -> int | None:
    value = stream.read(1)
    if value == b"":
        if required:
            raise XlsbParseError(f"Truncated {context} record header.")
        return None
    return value[0]


def _read_record_type(stream: BinaryIO) -> tuple[int, int] | None:
    first = _read_header_byte(stream, required=False, context="XLSB")
    if first is None:
        return None
    if first & 0x80 == 0:
        return first, 1
    second = _read_header_byte(stream, required=True, context="XLSB")
    assert second is not None  # Help static type checking after ``required=True``.
    if second & 0x80:
        raise XlsbParseError("XLSB record type uses more than two bytes.")
    return (first & 0x7F) | ((second & 0x7F) << 7), 2


def _read_record_length(stream: BinaryIO) -> tuple[int, int]:
    length = 0
    for index in range(4):
        current = _read_header_byte(stream, required=True, context="XLSB")
        assert current is not None
        length |= (current & 0x7F) << (index * 7)
        if current & 0x80 == 0:
            return length, index + 1
    raise XlsbParseError("XLSB record length uses more than four bytes.")


def _read_exact(stream: BinaryIO, count: int, *, context: str) -> bytes:
    chunks: list[bytes] = []
    remaining = count
    while remaining:
        chunk = stream.read(remaining)
        if not chunk:
            raise XlsbParseError(f"Truncated {context} record payload.")
        chunks.append(chunk)
        remaining -= len(chunk)
    return b"".join(chunks)


def iter_xlsb_records(
    stream: BinaryIO,
    *,
    max_record_bytes: int = DEFAULT_MAX_XLSB_RECORD_BYTES,
    max_records: int = DEFAULT_MAX_XLSB_RECORDS,
) -> Iterator[XlsbRecord]:
    """Yield strict BIFF12 records from one already-bounded binary part.

    This function intentionally takes a binary stream rather than a path or a
    ``ZipFile``.  The caller remains responsible for ZIP path validation and
    compressed/uncompressed archive limits; this layer prevents a single
    record's framing from bypassing those package-level controls.
    """
    if max_record_bytes < 0:
        raise ValueError("max_record_bytes must not be negative.")
    if max_records < 1:
        raise ValueError("max_records must be at least 1.")

    record_count = 0
    offset = 0
    while True:
        record_offset = offset
        record_header = _read_record_type(stream)
        if record_header is None:
            return
        record_type, type_size = record_header
        record_length, length_size = _read_record_length(stream)
        if record_length > max_record_bytes:
            raise XlsbLimitError("XLSB record payload exceeds max_record_bytes.")
        record_count += 1
        if record_count > max_records:
            raise XlsbLimitError("XLSB part exceeds max_records.")
        payload = _read_exact(stream, record_length, context="XLSB")
        offset = record_offset + type_size + length_size + record_length
        yield XlsbRecord(record_type=record_type, payload=payload, offset=record_offset)


def _column_name(index: int) -> str:
    if not 0 <= index <= _MAX_EXCEL_COLUMN_INDEX:
        raise XlsbFormulaError("Formula reference has an invalid column index.")
    value = index + 1
    letters: list[str] = []
    while value:
        value, remainder = divmod(value - 1, 26)
        letters.append(chr(ord("A") + remainder))
    return "".join(reversed(letters))


def _quoted_sheet_name(name: str) -> str:
    if not name or name.startswith("#"):
        raise XlsbFormulaUnsupportedError(
            "Formula uses an unsupported external-sheet reference."
        )
    return "'" + name.replace("'", "''") + "'"


def _error_literal(code: int) -> str:
    values = {
        0x00: "#NULL!",
        0x07: "#DIV/0!",
        0x0F: "#VALUE!",
        0x17: "#REF!",
        0x1D: "#NAME?",
        0x24: "#NUM!",
        0x2A: "#N/A",
        0x2B: "#GETTING_DATA",
    }
    try:
        return values[code]
    except KeyError as error:
        raise XlsbFormulaError("Formula contains an unknown Excel error code.") from error


def _number_literal(value: float) -> str:
    if not math.isfinite(value):
        raise XlsbFormulaError("Formula contains a non-finite numeric literal.")
    # Excel's general stored binary-number spelling is round-trippable with
    # 17 significant digits.  It is formula text only; no number is evaluated.
    return format(value, ".17g")


def _cell_reference(reader: _Reader) -> str:
    row = reader.u32()
    column_flags = reader.u16()
    column = column_flags & 0x3FFF
    if row > _MAX_EXCEL_ROW_INDEX:
        raise XlsbFormulaError("Formula reference has an invalid row index.")
    column_text = _column_name(column)
    if column_flags & 0x4000 == 0:
        column_text = "$" + column_text
    row_text = str(row + 1)
    if column_flags & 0x8000 == 0:
        row_text = "$" + row_text
    return column_text + row_text


def _area_reference(reader: _Reader) -> str:
    first_row = reader.u32()
    last_row = reader.u32()
    first_column_flags = reader.u16()
    last_column_flags = reader.u16()

    def format_endpoint(row: int, column_flags: int) -> str:
        if row > _MAX_EXCEL_ROW_INDEX:
            raise XlsbFormulaError("Formula reference has an invalid row index.")
        column_text = _column_name(column_flags & 0x3FFF)
        if column_flags & 0x4000 == 0:
            column_text = "$" + column_text
        row_text = str(row + 1)
        if column_flags & 0x8000 == 0:
            row_text = "$" + row_text
        return column_text + row_text

    return format_endpoint(first_row, first_column_flags) + ":" + format_endpoint(
        last_row, last_column_flags
    )


@dataclass(frozen=True)
class _FunctionSpec:
    name: str
    fixed_arity: int | None


# The standard IDs below are the conservative core used by the first adapter.
# ``None`` means the function is valid only through PtgFuncVar, whose encoded
# argument count is checked independently.  Unknown IDs remain fail-closed.
_FUNCTION_SPECS: Mapping[int, _FunctionSpec] = {
    0: _FunctionSpec("COUNT", None),
    1: _FunctionSpec("IF", 3),
    2: _FunctionSpec("ISNA", 1),
    3: _FunctionSpec("ISERROR", 1),
    4: _FunctionSpec("SUM", None),
    5: _FunctionSpec("AVERAGE", None),
    6: _FunctionSpec("MIN", None),
    7: _FunctionSpec("MAX", None),
    8: _FunctionSpec("ROW", 1),
    9: _FunctionSpec("COLUMN", 1),
    10: _FunctionSpec("NA", 0),
    11: _FunctionSpec("NPV", None),
    12: _FunctionSpec("STDEV", None),
    13: _FunctionSpec("DOLLAR", 2),
    14: _FunctionSpec("FIXED", 3),
    15: _FunctionSpec("SIN", 1),
    16: _FunctionSpec("COS", 1),
    17: _FunctionSpec("TAN", 1),
    18: _FunctionSpec("ATAN", 1),
    19: _FunctionSpec("PI", 0),
    20: _FunctionSpec("SQRT", 1),
    21: _FunctionSpec("EXP", 1),
    22: _FunctionSpec("LN", 1),
    23: _FunctionSpec("LOG10", 1),
    24: _FunctionSpec("ABS", 1),
    25: _FunctionSpec("INT", 1),
    26: _FunctionSpec("SIGN", 1),
    27: _FunctionSpec("ROUND", 2),
    28: _FunctionSpec("LOOKUP", 3),
    29: _FunctionSpec("INDEX", 4),
    30: _FunctionSpec("REPT", 2),
    31: _FunctionSpec("MID", 3),
    32: _FunctionSpec("LEN", 1),
    33: _FunctionSpec("VALUE", 1),
    34: _FunctionSpec("TRUE", 0),
    35: _FunctionSpec("FALSE", 0),
    36: _FunctionSpec("AND", None),
    37: _FunctionSpec("OR", None),
    38: _FunctionSpec("NOT", 1),
    39: _FunctionSpec("MOD", 2),
    48: _FunctionSpec("TEXT", 2),
    56: _FunctionSpec("PV", 5),
    57: _FunctionSpec("FV", 5),
    58: _FunctionSpec("NPER", 5),
    59: _FunctionSpec("PMT", 5),
    60: _FunctionSpec("RATE", 6),
    61: _FunctionSpec("MIRR", 3),
    62: _FunctionSpec("IRR", 2),
    63: _FunctionSpec("RAND", 0),
    64: _FunctionSpec("MATCH", 3),
    65: _FunctionSpec("DATE", 3),
    66: _FunctionSpec("TIME", 3),
    67: _FunctionSpec("DAY", 1),
    68: _FunctionSpec("MONTH", 1),
    69: _FunctionSpec("YEAR", 1),
    70: _FunctionSpec("WEEKDAY", 2),
    71: _FunctionSpec("HOUR", 1),
    72: _FunctionSpec("MINUTE", 1),
    73: _FunctionSpec("SECOND", 1),
    74: _FunctionSpec("NOW", 0),
    75: _FunctionSpec("AREAS", 1),
    76: _FunctionSpec("ROWS", 1),
    77: _FunctionSpec("COLUMNS", 1),
    78: _FunctionSpec("OFFSET", 5),
    82: _FunctionSpec("SEARCH", 3),
    83: _FunctionSpec("TRANSPOSE", 1),
    97: _FunctionSpec("ATAN2", 2),
    98: _FunctionSpec("ASIN", 1),
    99: _FunctionSpec("ACOS", 1),
    100: _FunctionSpec("CHOOSE", None),
    101: _FunctionSpec("HLOOKUP", 4),
    102: _FunctionSpec("VLOOKUP", 4),
    109: _FunctionSpec("LOG", 2),
    111: _FunctionSpec("CHAR", 1),
    112: _FunctionSpec("LOWER", 1),
    113: _FunctionSpec("UPPER", 1),
    114: _FunctionSpec("PROPER", 1),
    115: _FunctionSpec("LEFT", 2),
    116: _FunctionSpec("RIGHT", 2),
    117: _FunctionSpec("EXACT", 2),
    118: _FunctionSpec("TRIM", 1),
    119: _FunctionSpec("REPLACE", 4),
    120: _FunctionSpec("SUBSTITUTE", 4),
    121: _FunctionSpec("CODE", 1),
    124: _FunctionSpec("FIND", 3),
    125: _FunctionSpec("CELL", 2),
    126: _FunctionSpec("ISERR", 1),
    127: _FunctionSpec("ISTEXT", 1),
    128: _FunctionSpec("ISNUMBER", 1),
    129: _FunctionSpec("ISBLANK", 1),
    130: _FunctionSpec("T", 1),
    131: _FunctionSpec("N", 1),
    140: _FunctionSpec("DATEVALUE", 1),
    141: _FunctionSpec("TIMEVALUE", 1),
    142: _FunctionSpec("SLN", 3),
    143: _FunctionSpec("SYD", 4),
    144: _FunctionSpec("DDB", 5),
    148: _FunctionSpec("INDIRECT", 2),
    163: _FunctionSpec("MDETERM", 1),
    164: _FunctionSpec("MINVERSE", 1),
    165: _FunctionSpec("MMULT", 2),
    169: _FunctionSpec("COUNTA", None),
    198: _FunctionSpec("ISLOGICAL", 1),
    199: _FunctionSpec("DCOUNTA", 3),
    183: _FunctionSpec("PRODUCT", None),
    184: _FunctionSpec("FACT", 1),
    227: _FunctionSpec("MEDIAN", None),
    228: _FunctionSpec("SUMPRODUCT", None),
    261: _FunctionSpec("ERROR.TYPE", 1),
    269: _FunctionSpec("AVEDEV", 1),
    321: _FunctionSpec("SUMSQ", 1),
    325: _FunctionSpec("LARGE", 2),
    326: _FunctionSpec("SMALL", 2),
    336: _FunctionSpec("CONCATENATE", None),
    337: _FunctionSpec("POWER", 2),
    344: _FunctionSpec("SUBTOTAL", 2),
    345: _FunctionSpec("SUMIF", 3),
    346: _FunctionSpec("COUNTIF", 2),
    358: _FunctionSpec("GETPIVOTDATA", None),
    359: _FunctionSpec("HYPERLINK", 2),
    379: _FunctionSpec("RTD", None),
    480: _FunctionSpec("IFERROR", 2),
    481: _FunctionSpec("COUNTIFS", None),
    482: _FunctionSpec("SUMIFS", None),
    483: _FunctionSpec("AVERAGEIF", 3),
}


@dataclass(frozen=True)
class _Expression:
    """One lazily rendered formula expression.

    Repeatedly joining an RPN expression into a string makes a deeply nested
    but small token stream quadratic in decoded characters.  Retain a compact
    immutable fragment tree instead, accounting for its eventual text size at
    every push and rendering it iteratively only after the stack resolves.
    """

    fragments: tuple[str | _Expression, ...]
    character_count: int
    literal_text: str | None = None


def _literal_expression(text: str) -> _Expression:
    return _Expression((text,), len(text), literal_text=text)


def _joined_expression(*fragments: str | _Expression) -> _Expression:
    return _Expression(
        fragments=fragments,
        character_count=sum(
            len(fragment)
            if isinstance(fragment, str)
            else fragment.character_count
            for fragment in fragments
        ),
    )


def _render_expression(expression: _Expression) -> str:
    """Flatten a bounded expression tree without recursion."""
    pieces: list[str] = []
    pending: list[str | _Expression] = list(reversed(expression.fragments))
    while pending:
        fragment = pending.pop()
        if isinstance(fragment, str):
            pieces.append(fragment)
        else:
            pending.extend(reversed(fragment.fragments))
    return "".join(pieces)


def _push_expression(
    stack: list[_Expression],
    expression: _Expression,
    *,
    max_stack_items: int,
    max_formula_chars: int,
) -> None:
    if expression.character_count > max_formula_chars:
        raise XlsbLimitError("Decoded XLSB formula exceeds max_formula_chars.")
    if len(stack) >= max_stack_items:
        raise XlsbLimitError("XLSB formula exceeds max_formula_stack_items.")
    stack.append(expression)


def _pop_expression(stack: list[_Expression]) -> _Expression:
    if not stack:
        raise XlsbFormulaError("XLSB formula token stack underflow.")
    return stack.pop()


def _function_spec(identifier: int) -> _FunctionSpec:
    try:
        return _FUNCTION_SPECS[identifier]
    except KeyError as error:
        raise XlsbFormulaUnsupportedError(
            "Formula uses an unsupported built-in function ID."
        ) from error


def _call_expression(
    stack: list[_Expression],
    specification: _FunctionSpec,
    argument_count: int,
    *,
    max_stack_items: int,
    max_formula_chars: int,
) -> None:
    if argument_count > len(stack):
        raise XlsbFormulaError("XLSB formula function has too few arguments.")
    arguments = stack[-argument_count:] if argument_count else []
    if argument_count:
        del stack[-argument_count:]
    fragments: list[str | _Expression] = [specification.name, "("]
    for index, argument in enumerate(arguments):
        if index:
            fragments.append(",")
        fragments.append(argument)
    fragments.append(")")
    _push_expression(
        stack,
        _joined_expression(*fragments),
        max_stack_items=max_stack_items,
        max_formula_chars=max_formula_chars,
    )


_UDF_NAME_PATTERN = re.compile(r"(?:_xlfn\.)?[A-Za-z_][A-Za-z0-9_.]*\Z")


def _user_defined_function_expression(
    stack: list[_Expression],
    argument_count: int,
    *,
    max_stack_items: int,
    max_formula_chars: int,
) -> None:
    """Render the documented UDF PtgFuncVar form without inventing a name."""
    if argument_count < 1 or argument_count > len(stack):
        raise XlsbFormulaError("XLSB UDF formula has an invalid argument count.")
    expressions = stack[-argument_count:]
    del stack[-argument_count:]
    function_name = expressions[0].literal_text
    if function_name is None or not _UDF_NAME_PATTERN.fullmatch(function_name):
        raise XlsbFormulaUnsupportedError(
            "Formula uses an unsupported user-defined function name."
        )
    fragments: list[str | _Expression] = [function_name, "("]
    for index, expression in enumerate(expressions[1:]):
        if index:
            fragments.append(",")
        fragments.append(expression)
    fragments.append(")")
    _push_expression(
        stack,
        _joined_expression(*fragments),
        max_stack_items=max_stack_items,
        max_formula_chars=max_formula_chars,
    )


def decode_xlsb_formula(
    token_bytes: bytes,
    *,
    external_sheets: Sequence[str] = (),
    defined_names: Sequence[str] = (),
    max_formula_bytes: int = DEFAULT_MAX_XLSB_FORMULA_BYTES,
    max_tokens: int = DEFAULT_MAX_XLSB_FORMULA_TOKENS,
    max_stack_items: int = DEFAULT_MAX_XLSB_FORMULA_STACK_ITEMS,
    max_formula_chars: int = DEFAULT_MAX_XLSB_FORMULA_CHARS,
    max_function_args: int = DEFAULT_MAX_XLSB_FUNCTION_ARGS,
) -> str:
    """Return a conservative Excel formula text from a BIFF12 RPN stream.

    Formula strings are reconstructed only for token forms represented by this
    decoder.  They are never evaluated.  Unknown, array/shared, memory, and
    external-name token forms raise ``XlsbFormulaUnsupportedError`` so callers
    can surface a coverage gap instead of producing a deceptively plausible
    formula.
    """
    if len(token_bytes) > max_formula_bytes:
        raise XlsbLimitError("XLSB formula exceeds max_formula_bytes.")
    if max_tokens < 1:
        raise ValueError("max_tokens must be at least 1.")
    if max_stack_items < 1:
        raise ValueError("max_stack_items must be at least 1.")
    if max_formula_chars < 1:
        raise ValueError("max_formula_chars must be at least 1.")
    if not 0 <= max_function_args <= 255:
        raise ValueError("max_function_args must be between 0 and 255.")

    reader = _Reader(token_bytes, context="XLSB formula")
    stack: list[_Expression] = []
    token_count = 0

    while reader.remaining:
        token_count += 1
        if token_count > max_tokens:
            raise XlsbLimitError("XLSB formula exceeds max_formula_tokens.")
        token = reader.u8()

        if token in (
            0x03,
            0x04,
            0x05,
            0x06,
            0x07,
            0x08,
            0x09,
            0x0A,
            0x0B,
            0x0C,
            0x0D,
            0x0E,
            0x0F,
            0x10,
            0x11,
        ):
            right = _pop_expression(stack)
            left = _pop_expression(stack)
            operator = {
                0x03: "+",
                0x04: "-",
                0x05: "*",
                0x06: "/",
                0x07: "^",
                0x08: "&",
                0x09: "<",
                0x0A: "<=",
                0x0B: "=",
                0x0C: ">=",
                0x0D: ">",
                0x0E: "<>",
                0x0F: " ",
                0x10: ",",
                0x11: ":",
            }[token]
            _push_expression(
                stack,
                _joined_expression("(", left, operator, right, ")"),
                max_stack_items=max_stack_items,
                max_formula_chars=max_formula_chars,
            )
        elif token in (0x12, 0x13):
            expression = _pop_expression(stack)
            _push_expression(
                stack,
                _joined_expression("(", "+" if token == 0x12 else "-", expression, ")"),
                max_stack_items=max_stack_items,
                max_formula_chars=max_formula_chars,
            )
        elif token == 0x14:
            expression = _pop_expression(stack)
            _push_expression(
                stack,
                _joined_expression("(", expression, "%)"),
                max_stack_items=max_stack_items,
                max_formula_chars=max_formula_chars,
            )
        elif token == 0x15:
            expression = _pop_expression(stack)
            _push_expression(
                stack,
                _joined_expression("(", expression, ")"),
                max_stack_items=max_stack_items,
                max_formula_chars=max_formula_chars,
            )
        elif token == 0x16:
            _push_expression(
                stack,
                _literal_expression(""),
                max_stack_items=max_stack_items,
                max_formula_chars=max_formula_chars,
            )
        elif token == 0x17:
            value = reader.utf16(reader.u16(), max_characters=max_formula_chars)
            _push_expression(
                stack,
                _literal_expression('"' + value.replace('"', '""') + '"'),
                max_stack_items=max_stack_items,
                max_formula_chars=max_formula_chars,
            )
        elif token == 0x19:
            attribute_flags = reader.u8()
            reader.u16()
            # PtgAttrSpace is a display-only spacing token.  Its whitespace
            # count does not affect formula evaluation, so the conservative
            # textual form omits it. Other attribute flags alter control flow
            # or token interpretation and remain outside this decoder.
            if attribute_flags != 0x40:
                raise XlsbFormulaUnsupportedError(
                    "Formula uses an XLSB attribute token outside the verified decoder subset."
                )
        elif token == 0x1C:
            _push_expression(
                stack,
                _literal_expression(_error_literal(reader.u8())),
                max_stack_items=max_stack_items,
                max_formula_chars=max_formula_chars,
            )
        elif token == 0x1D:
            boolean = reader.u8()
            if boolean not in (0, 1):
                raise XlsbFormulaError("Formula contains an invalid Boolean literal.")
            _push_expression(
                stack,
                _literal_expression("TRUE" if boolean else "FALSE"),
                max_stack_items=max_stack_items,
                max_formula_chars=max_formula_chars,
            )
        elif token == 0x1E:
            _push_expression(
                stack,
                _literal_expression(str(reader.u16())),
                max_stack_items=max_stack_items,
                max_formula_chars=max_formula_chars,
            )
        elif token == 0x1F:
            _push_expression(
                stack,
                _literal_expression(_number_literal(reader.f64())),
                max_stack_items=max_stack_items,
                max_formula_chars=max_formula_chars,
            )
        elif token in (0x21, 0x41, 0x61):
            specification = _function_spec(reader.u16())
            if specification.fixed_arity is None:
                raise XlsbFormulaUnsupportedError(
                    "Formula encodes a variable-arity function as PtgFunc."
                )
            _call_expression(
                stack,
                specification,
                specification.fixed_arity,
                max_stack_items=max_stack_items,
                max_formula_chars=max_formula_chars,
            )
        elif token in (0x22, 0x42, 0x62):
            argument_count = reader.u8() & 0x7F
            function_word = reader.u16()
            if function_word & 0x8000:
                raise XlsbFormulaUnsupportedError(
                    "Formula uses an unsupported command-equivalent function."
                )
            if argument_count > max_function_args:
                raise XlsbLimitError("Formula function exceeds max_function_args.")
            if function_word == 0x00FF:
                _user_defined_function_expression(
                    stack,
                    argument_count,
                    max_stack_items=max_stack_items,
                    max_formula_chars=max_formula_chars,
                )
            else:
                specification = _function_spec(function_word)
                _call_expression(
                    stack,
                    specification,
                    argument_count,
                    max_stack_items=max_stack_items,
                    max_formula_chars=max_formula_chars,
                )
        elif token in (0x23, 0x43, 0x63):
            name_index = reader.u32()
            if name_index == 0 or name_index > len(defined_names):
                raise XlsbFormulaError("Formula has an invalid defined-name index.")
            name = defined_names[name_index - 1]
            if not name:
                raise XlsbFormulaError("Formula refers to an empty defined name.")
            _push_expression(
                stack,
                _literal_expression(name),
                max_stack_items=max_stack_items,
                max_formula_chars=max_formula_chars,
            )
        elif token in (0x24, 0x44, 0x64):
            _push_expression(
                stack,
                _literal_expression(_cell_reference(reader)),
                max_stack_items=max_stack_items,
                max_formula_chars=max_formula_chars,
            )
        elif token in (0x25, 0x45, 0x65):
            _push_expression(
                stack,
                _literal_expression(_area_reference(reader)),
                max_stack_items=max_stack_items,
                max_formula_chars=max_formula_chars,
            )
        elif token in (0x2A, 0x4A, 0x6A):
            reader.take(6)
            _push_expression(
                stack,
                _literal_expression("#REF!"),
                max_stack_items=max_stack_items,
                max_formula_chars=max_formula_chars,
            )
        elif token in (0x2B, 0x4B, 0x6B):
            reader.take(12)
            _push_expression(
                stack,
                _literal_expression("#REF!"),
                max_stack_items=max_stack_items,
                max_formula_chars=max_formula_chars,
            )
        elif token in (0x3A, 0x5A, 0x7A):
            sheet_index = reader.u16()
            if sheet_index >= len(external_sheets):
                raise XlsbFormulaError("Formula has an invalid external-sheet index.")
            reference = _cell_reference(reader)
            _push_expression(
                stack,
                _literal_expression(
                    _quoted_sheet_name(external_sheets[sheet_index]) + "!" + reference
                ),
                max_stack_items=max_stack_items,
                max_formula_chars=max_formula_chars,
            )
        elif token in (0x3B, 0x5B, 0x7B):
            sheet_index = reader.u16()
            if sheet_index >= len(external_sheets):
                raise XlsbFormulaError("Formula has an invalid external-sheet index.")
            reference = _area_reference(reader)
            _push_expression(
                stack,
                _literal_expression(
                    _quoted_sheet_name(external_sheets[sheet_index]) + "!" + reference
                ),
                max_stack_items=max_stack_items,
                max_formula_chars=max_formula_chars,
            )
        elif token in (0x3C, 0x5C, 0x7C):
            sheet_index = reader.u16()
            if sheet_index >= len(external_sheets):
                raise XlsbFormulaError("Formula has an invalid external-sheet index.")
            reader.take(6)
            _push_expression(
                stack,
                _literal_expression(_quoted_sheet_name(external_sheets[sheet_index]) + "!#REF!"),
                max_stack_items=max_stack_items,
                max_formula_chars=max_formula_chars,
            )
        elif token in (0x3D, 0x5D, 0x7D):
            sheet_index = reader.u16()
            if sheet_index >= len(external_sheets):
                raise XlsbFormulaError("Formula has an invalid external-sheet index.")
            reader.take(12)
            _push_expression(
                stack,
                _literal_expression(_quoted_sheet_name(external_sheets[sheet_index]) + "!#REF!"),
                max_stack_items=max_stack_items,
                max_formula_chars=max_formula_chars,
            )
        elif token in (
            0x01,
            0x20,
            0x40,
            0x60,
            0x26,
            0x27,
            0x28,
            0x29,
            0x2C,
            0x2D,
            0x39,
            0x59,
            0x79,
        ):
            raise XlsbFormulaUnsupportedError(
                "Formula uses an XLSB token outside the verified decoder subset."
            )
        else:
            raise XlsbFormulaUnsupportedError("Formula uses an unknown XLSB token.")

    if len(stack) != 1:
        raise XlsbFormulaError("XLSB formula token stack did not resolve to one expression.")
    text = "=" + _render_expression(stack[0])
    if len(text) > max_formula_chars:
        raise XlsbLimitError("Decoded XLSB formula exceeds max_formula_chars.")
    return text


def xlsb_records_from_bytes(
    payload: bytes,
    *,
    max_record_bytes: int = DEFAULT_MAX_XLSB_RECORD_BYTES,
    max_records: int = DEFAULT_MAX_XLSB_RECORDS,
) -> tuple[XlsbRecord, ...]:
    """Convenience wrapper for controlled fixtures and small private parts."""
    return tuple(
        iter_xlsb_records(
            io.BytesIO(payload),
            max_record_bytes=max_record_bytes,
            max_records=max_records,
        )
    )


@dataclass(frozen=True)
class XlsbCoreSheet:
    """One workbook tab retained by the XLSB core reader."""

    title: str
    state: str
    tab_id: int
    part_name: str | None
    kind: str


@dataclass(frozen=True)
class XlsbCoreCell:
    """A non-empty XLSB worksheet cell without formatting semantics."""

    sheet: str
    row: int
    column: int
    cell_type: str
    value: Any
    value_type: str
    formula: str | None = None
    cached_value: Any = None

    @property
    def location(self) -> tuple[str, int, int]:
        """Return an internal, zero-based sheet/row/column key."""
        return self.sheet, self.row, self.column


@dataclass(frozen=True)
class XlsbCoreDefinedName:
    """A defined name whose formula text was decoded only when verified."""

    name: str
    scope_sheet: str | None
    formula: str | None


@dataclass(frozen=True)
class XlsbCoreWorkbook:
    """Formula-bearing XLSB state before FormulaFence's full snapshot adapter.

    ``formula_text_coverage_complete`` says only that every parsed formula and
    defined name used the verified token subset.  It does *not* claim array,
    rich-data, control, formatting, relationship, or calculation coverage.
    """

    sheets: tuple[XlsbCoreSheet, ...]
    cells: Mapping[tuple[str, int, int], XlsbCoreCell]
    defined_names: tuple[XlsbCoreDefinedName, ...]
    formula_text_coverage_complete: bool
    unsupported_formula_cells: frozenset[tuple[str, int, int]]
    unsupported_defined_name_indexes: frozenset[int]
    unsupported_sheet_types: bool
    date_1904: bool


@dataclass(frozen=True)
class _WorkbookRelationship:
    target: str | None
    relationship_type: str


@dataclass(frozen=True)
class _RawDefinedName:
    name: str
    scope_index: int | None
    token_bytes: bytes


@dataclass
class _FormulaTextBudget:
    """Bound decoded formula strings across one retained XLSB core."""

    max_characters: int
    consumed_characters: int = 0

    def consume(self, formula: str) -> None:
        formula_characters = len(formula)
        if formula_characters > self.max_characters - self.consumed_characters:
            raise XlsbLimitError(
                "XLSB decoded formula text exceeds max_total_formula_characters."
            )
        self.consumed_characters += formula_characters


_PACKAGE_RELATIONSHIPS_NAMESPACE = (
    "http://schemas.openxmlformats.org/package/2006/relationships"
)
_STRICT_PACKAGE_RELATIONSHIPS_NAMESPACE = (
    "http://purl.oclc.org/ooxml/package/relationships"
)
_WORKSHEET_RELATIONSHIP_SUFFIX = "/worksheet"
_CHARTSHEET_RELATIONSHIP_SUFFIX = "/chartsheet"
_DIALOGSHEET_RELATIONSHIP_SUFFIX = "/dialogsheet"
_MACROSHEET_RELATIONSHIP_SUFFIX = "/macrosheet"
_INTL_MACROSHEET_RELATIONSHIP_SUFFIX = "/intlmacrosheet"

_BRT_ROW_HDR = 0x0000
_BRT_CELL_BLANK = 0x0001
_BRT_CELL_RK = 0x0002
_BRT_CELL_ERROR = 0x0003
_BRT_CELL_BOOL = 0x0004
_BRT_CELL_REAL = 0x0005
_BRT_CELL_STRING = 0x0006
_BRT_CELL_SHARED_STRING = 0x0007
_BRT_FMLA_STRING = 0x0008
_BRT_FMLA_NUM = 0x0009
_BRT_FMLA_BOOL = 0x000A
_BRT_FMLA_ERROR = 0x000B
_BRT_SHARED_STRING_ITEM = 0x0013
_BRT_NAME = 0x0027
_BRT_BEGIN_SST = 0x009F
_BRT_BEGIN_BUNDLE_SHEETS = 0x008F
_BRT_END_BUNDLE_SHEETS = 0x0090
_BRT_BEGIN_SHEET_DATA = 0x0091
_BRT_END_SHEET_DATA = 0x0092
_BRT_WORKBOOK_PROPERTIES = 0x0099
_BRT_BUNDLE_SHEET = 0x009C
_BRT_EXTERN_SHEET = 0x016A


def _canonical_xlsb_parts(parts: Mapping[str, bytes]) -> dict[str, bytes]:
    """Return a case-normalized, unambiguous package-part lookup.

    OPC part names are case-sensitive in the abstract package model, while
    workbooks in the wild are commonly opened through case-insensitive ZIP
    readers.  The common FormulaFence archive preflight already rejects
    case-colliding members.  Keep the same invariant for direct callers of
    this lower-level parser, then normalize the lookup so a relationship can
    safely resolve an otherwise equivalent historical part spelling.
    """
    canonical_parts: dict[str, bytes] = {}
    for name, payload in parts.items():
        if not isinstance(name, str) or not name:
            raise XlsbParseError("XLSB package has an invalid part name.")
        canonical_name = name.casefold()
        if canonical_name in canonical_parts:
            raise XlsbParseError("XLSB package has case-colliding part names.")
        canonical_parts[canonical_name] = payload
    return canonical_parts


def _part_payload(
    parts: Mapping[str, bytes],
    name: str,
    *,
    limits: XlsbReaderLimits,
    required: bool,
    retained_bytes: list[int],
) -> bytes | None:
    payload = parts.get(name.casefold())
    if payload is None:
        if required:
            raise XlsbParseError(f"XLSB package is missing required part {name!r}.")
        return None
    if not isinstance(payload, bytes):
        raise XlsbParseError("XLSB package part is not binary data.")
    if len(payload) > limits.max_binary_part_bytes:
        raise XlsbLimitError("XLSB binary part exceeds max_binary_part_bytes.")
    retained_bytes[0] += len(payload)
    if retained_bytes[0] > limits.max_total_binary_bytes:
        raise XlsbLimitError("XLSB binary parts exceed max_total_binary_bytes.")
    return payload


def _xml_local_name(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


def _workbook_relationship_target(target: str) -> str:
    if (
        not target
        or len(target) > 1_024
        or target.startswith(("/", "\\"))
        or "\\" in target
        or "%" in target
    ):
        raise XlsbParseError("XLSB workbook relationship has an unsafe target.")
    pieces = target.split("/")
    if any(piece in {"", ".", ".."} for piece in pieces):
        raise XlsbParseError("XLSB workbook relationship has an unsafe target.")
    return "xl/" + target


def _workbook_relationships(
    payload: bytes,
    *,
    limits: XlsbReaderLimits,
) -> dict[str, _WorkbookRelationship]:
    if len(payload) > limits.max_relationship_part_bytes:
        raise XlsbLimitError(
            "XLSB relationship part exceeds max_relationship_part_bytes."
        )
    relationships: dict[str, _WorkbookRelationship] = {}
    depth = 0
    saw_root = False
    try:
        for event, element in ElementTree.iterparse(
            io.BytesIO(payload), events=("start", "end")
        ):
            if event == "start":
                depth += 1
                if depth == 1:
                    if _xml_local_name(element.tag) != "Relationships":
                        raise XlsbParseError(
                            "XLSB workbook relationships have an invalid root."
                        )
                    namespace = element.tag.partition("}")[0].removeprefix("{")
                    if namespace not in {
                        _PACKAGE_RELATIONSHIPS_NAMESPACE,
                        _STRICT_PACKAGE_RELATIONSHIPS_NAMESPACE,
                    }:
                        raise XlsbParseError(
                            "XLSB workbook relationships use an unsupported namespace."
                        )
                    if element.attrib:
                        raise XlsbParseError(
                            "XLSB workbook relationships have unsupported root attributes."
                        )
                    saw_root = True
                    continue
                if depth != 2 or _xml_local_name(element.tag) != "Relationship":
                    raise XlsbParseError(
                        "XLSB workbook relationships contain an unknown element."
                    )
                if len(relationships) >= limits.max_relationships:
                    raise XlsbLimitError("XLSB relationships exceed max_relationships.")
                if set(element.attrib).difference({"Id", "Type", "Target", "TargetMode"}):
                    raise XlsbParseError(
                        "XLSB workbook relationship has an unsupported attribute."
                    )
                identifier = element.get("Id")
                relationship_type = element.get("Type")
                target = element.get("Target")
                if (
                    not isinstance(identifier, str)
                    or not identifier
                    or len(identifier) > 1_024
                    or not isinstance(relationship_type, str)
                    or not relationship_type
                    or len(relationship_type) > 2_048
                    or not isinstance(target, str)
                    or not target
                    or len(target) > 1_024
                ):
                    raise XlsbParseError("XLSB workbook relationship is malformed.")
                if identifier in relationships:
                    raise XlsbParseError("XLSB workbook relationships have duplicate IDs.")
                target_mode = element.get("TargetMode")
                if target_mode is not None and (
                    len(target_mode) > 32
                    or target_mode.casefold() not in {"internal", "external"}
                ):
                    raise XlsbParseError(
                        "XLSB workbook relationship has an invalid target mode."
                    )
                if target_mode is not None and target_mode.casefold() == "external":
                    relationships[identifier] = _WorkbookRelationship(
                        target=None,
                        relationship_type=relationship_type,
                    )
                else:
                    relationships[identifier] = _WorkbookRelationship(
                        target=_workbook_relationship_target(target),
                        relationship_type=relationship_type,
                    )
            else:
                if depth == 2:
                    element.clear()
                depth -= 1
    except (DefusedXmlException, ElementTree.ParseError, UnicodeError, ValueError) as error:
        raise XlsbParseError("XLSB workbook relationships are not valid XML.") from error
    if not saw_root or depth != 0:
        raise XlsbParseError("XLSB workbook relationships are incomplete.")
    return relationships


def _wide_string(reader: _Reader, *, max_characters: int) -> str:
    return reader.utf16(reader.u32(), max_characters=max_characters)


def _nullable_wide_string(reader: _Reader, *, max_characters: int) -> str | None:
    count = reader.u32()
    if count == 0xFFFF_FFFF:
        return None
    return reader.utf16(count, max_characters=max_characters)


def _require_consumed(reader: _Reader) -> None:
    if reader.remaining:
        raise XlsbParseError("XLSB record has trailing data outside its defined structure.")


def _sheet_state(value: int) -> str:
    states = {0: "visible", 1: "hidden", 2: "veryHidden"}
    try:
        return states[value]
    except KeyError as error:
        raise XlsbParseError("XLSB workbook has an invalid sheet state.") from error


def _sheet_kind(relationship_type: str) -> str:
    normalized = relationship_type.casefold()
    if normalized.endswith(_WORKSHEET_RELATIONSHIP_SUFFIX):
        return "worksheet"
    if normalized.endswith(
        (
            _CHARTSHEET_RELATIONSHIP_SUFFIX,
            _DIALOGSHEET_RELATIONSHIP_SUFFIX,
            _MACROSHEET_RELATIONSHIP_SUFFIX,
            _INTL_MACROSHEET_RELATIONSHIP_SUFFIX,
        )
    ):
        return "non_grid"
    raise XlsbParseError("XLSB workbook sheet relationship has an unsupported type.")


def _parse_bundle_sheet(
    payload: bytes,
    relationships: Mapping[str, _WorkbookRelationship],
    *,
    limits: XlsbReaderLimits,
) -> XlsbCoreSheet:
    reader = _Reader(payload, context="XLSB bundle sheet")
    state = _sheet_state(reader.u32())
    tab_id = reader.u32()
    relationship_id = _nullable_wide_string(
        reader, max_characters=limits.max_cell_text_characters
    )
    title = _wide_string(reader, max_characters=limits.max_cell_text_characters)
    _require_consumed(reader)
    if not title:
        raise XlsbParseError("XLSB workbook has an empty sheet name.")
    if relationship_id is None:
        raise XlsbParseError("XLSB workbook sheet has no relationship ID.")
    relationship = relationships.get(relationship_id)
    if relationship is None or relationship.target is None:
        raise XlsbParseError("XLSB workbook sheet has no internal target part.")
    return XlsbCoreSheet(
        title=title,
        state=state,
        tab_id=tab_id,
        part_name=relationship.target,
        kind=_sheet_kind(relationship.relationship_type),
    )


def _parse_extern_sheet(
    payload: bytes,
    sheets: Sequence[XlsbCoreSheet],
    *,
    limits: XlsbReaderLimits,
) -> tuple[str, ...]:
    reader = _Reader(payload, context="XLSB external-sheet")
    count = reader.u32()
    if count > limits.max_sheets:
        raise XlsbLimitError("XLSB external-sheet catalog exceeds max_sheets.")
    if reader.remaining != count * 12:
        raise XlsbParseError("XLSB external-sheet catalog has an invalid length.")

    values: list[str] = []
    for _ in range(count):
        reader.u32()  # iSupBook; external books are not text-decoded by this core.
        first_sheet = struct.unpack("<i", reader.take(4))[0]
        last_sheet = struct.unpack("<i", reader.take(4))[0]
        if first_sheet != last_sheet or not 0 <= first_sheet < len(sheets):
            # This private marker reaches only the formula decoder, which
            # immediately turns it into an explicit coverage gap.
            values.append("#UnsupportedXti")
        else:
            values.append(sheets[first_sheet].title)
    return tuple(values)


def _parse_defined_name(payload: bytes, *, limits: XlsbReaderLimits) -> _RawDefinedName:
    reader = _Reader(payload, context="XLSB defined-name")
    reader.u32()  # name flags
    reader.u8()  # macro shortcut key
    scope_raw = reader.u32()
    name = _wide_string(reader, max_characters=limits.max_cell_text_characters)
    formula_length = reader.u32()
    if formula_length > limits.max_formula_bytes:
        raise XlsbLimitError("XLSB defined-name formula exceeds max_formula_bytes.")
    token_bytes = reader.take(formula_length)
    # A BrtName can carry comments and macro-only descriptive nullable strings
    # after its NameParsedFormula.  They have no formula semantics and are not
    # retained by this core reader.
    if not name:
        raise XlsbParseError("XLSB workbook has an empty defined name.")
    return _RawDefinedName(
        name=name,
        scope_index=None if scope_raw == 0xFFFF_FFFF else scope_raw,
        token_bytes=token_bytes,
    )


def _parse_xlsb_workbook_part(
    payload: bytes,
    relationships: Mapping[str, _WorkbookRelationship],
    *,
    limits: XlsbReaderLimits,
) -> tuple[tuple[XlsbCoreSheet, ...], tuple[_RawDefinedName, ...], tuple[str, ...], bool]:
    sheets: list[XlsbCoreSheet] = []
    sheet_titles: set[str] = set()
    sheet_tab_ids: set[int] = set()
    raw_defined_names: list[_RawDefinedName] = []
    external_sheets: tuple[str, ...] = ()
    date_1904 = False
    bundle_started = False
    bundle_ended = False

    for record in iter_xlsb_records(
        io.BytesIO(payload),
        max_record_bytes=limits.max_record_bytes,
        max_records=limits.max_records_per_part,
    ):
        if record.record_type == _BRT_WORKBOOK_PROPERTIES:
            reader = _Reader(record.payload, context="XLSB workbook properties")
            date_1904 = bool(reader.u32() & 0x1)
        elif record.record_type == _BRT_BEGIN_BUNDLE_SHEETS:
            if bundle_started or bundle_ended or record.payload:
                raise XlsbParseError("XLSB workbook has an invalid BrtBeginBundleShs record.")
            bundle_started = True
        elif record.record_type == _BRT_BUNDLE_SHEET:
            if not bundle_started or bundle_ended:
                raise XlsbParseError("XLSB workbook has a sheet outside BrtBundleShs.")
            if len(sheets) >= limits.max_sheets:
                raise XlsbLimitError("XLSB workbook exceeds max_sheets.")
            sheet = _parse_bundle_sheet(record.payload, relationships, limits=limits)
            title_key = sheet.title.casefold()
            if title_key in sheet_titles:
                raise XlsbParseError("XLSB workbook has duplicate sheet names.")
            if sheet.tab_id in sheet_tab_ids:
                raise XlsbParseError("XLSB workbook has duplicate sheet tab identifiers.")
            sheet_titles.add(title_key)
            sheet_tab_ids.add(sheet.tab_id)
            sheets.append(sheet)
        elif record.record_type == _BRT_END_BUNDLE_SHEETS:
            if not bundle_started or bundle_ended or record.payload:
                raise XlsbParseError("XLSB workbook has an invalid BrtEndBundleShs record.")
            bundle_ended = True
        elif record.record_type == _BRT_EXTERN_SHEET:
            if not bundle_ended:
                raise XlsbParseError("XLSB external-sheet catalog precedes BrtEndBundleShs.")
            external_sheets = _parse_extern_sheet(record.payload, sheets, limits=limits)
        elif record.record_type == _BRT_NAME:
            if not bundle_ended:
                raise XlsbParseError("XLSB defined name precedes BrtEndBundleShs.")
            if len(raw_defined_names) >= limits.max_defined_names:
                raise XlsbLimitError("XLSB workbook exceeds max_defined_names.")
            raw_defined_names.append(_parse_defined_name(record.payload, limits=limits))

    if not sheets or not bundle_started or not bundle_ended:
        raise XlsbParseError("XLSB workbook is missing its complete sheet catalog.")
    defined_name_keys: set[tuple[int | None, str]] = set()
    for definition in raw_defined_names:
        if definition.scope_index is not None and definition.scope_index >= len(sheets):
            raise XlsbParseError("XLSB defined name has an invalid sheet scope.")
        name_key = (definition.scope_index, definition.name.casefold())
        if name_key in defined_name_keys:
            raise XlsbParseError("XLSB workbook has duplicate defined names in one scope.")
        defined_name_keys.add(name_key)
    return tuple(sheets), tuple(raw_defined_names), external_sheets, date_1904


def _parse_shared_strings(
    payload: bytes,
    *,
    limits: XlsbReaderLimits,
) -> tuple[str, ...]:
    strings: list[str] = []
    saw_begin = False
    declared_count: int | None = None
    for record in iter_xlsb_records(
        io.BytesIO(payload),
        max_record_bytes=limits.max_record_bytes,
        max_records=limits.max_records_per_part,
    ):
        if record.record_type == _BRT_BEGIN_SST:
            if saw_begin:
                raise XlsbParseError("XLSB shared-string table has duplicate BrtBeginSst.")
            reader = _Reader(record.payload, context="XLSB shared-string table")
            reader.u32()  # total string count
            declared_count = reader.u32()
            _require_consumed(reader)
            if declared_count > limits.max_shared_strings:
                raise XlsbLimitError("XLSB shared-string table exceeds max_shared_strings.")
            saw_begin = True
        elif record.record_type == _BRT_SHARED_STRING_ITEM:
            if not saw_begin:
                raise XlsbParseError("XLSB shared-string item precedes BrtBeginSst.")
            if len(strings) >= limits.max_shared_strings:
                raise XlsbLimitError("XLSB shared-string table exceeds max_shared_strings.")
            reader = _Reader(record.payload, context="XLSB shared-string item")
            reader.u8()  # rich-text/phonetic flags
            strings.append(
                _wide_string(reader, max_characters=limits.max_cell_text_characters)
            )
            # Rich text and phonetic data are intentionally outside this core.
            # Their presence changes presentation, not this raw cell's string.
    if not saw_begin:
        raise XlsbParseError("XLSB shared-string table is missing BrtBeginSst.")
    assert declared_count is not None
    if len(strings) != declared_count:
        raise XlsbParseError("XLSB shared-string table count does not match its items.")
    return tuple(strings)


def _cell_header(reader: _Reader) -> int:
    column = reader.u32()
    if column > _MAX_EXCEL_COLUMN_INDEX:
        raise XlsbParseError("XLSB cell has an invalid column index.")
    reader.take(4)  # XF index and cell flags
    return column


def _rk_value(reader: _Reader) -> int | float:
    raw_bytes = reader.take(4)
    raw_unsigned = struct.unpack("<I", raw_bytes)[0]
    divide_by_100 = bool(raw_unsigned & 0x01)
    is_integer = bool(raw_unsigned & 0x02)
    if is_integer:
        value: int | float = struct.unpack("<i", raw_bytes)[0] >> 2
    else:
        value = struct.unpack("<d", struct.pack("<Q", (raw_unsigned & ~0x03) << 32))[0]
        if not math.isfinite(value):
            raise XlsbParseError("XLSB RK cell has a non-finite numeric value.")
    if divide_by_100:
        return value / 100
    return value


def _formula_token_bytes(
    record_type: int,
    payload: bytes,
    *,
    limits: XlsbReaderLimits,
) -> tuple[int, Any, bytes]:
    """Return column, cached result, and Rgce bytes from one BrtFmla record."""
    reader = _Reader(payload, context="XLSB formula cell")
    column = _cell_header(reader)
    if record_type == _BRT_FMLA_STRING:
        cached_value: Any = _wide_string(
            reader, max_characters=limits.max_cell_text_characters
        )
    elif record_type == _BRT_FMLA_NUM:
        cached_value = reader.f64()
        if not math.isfinite(cached_value):
            raise XlsbParseError("XLSB formula cache has a non-finite numeric value.")
    elif record_type == _BRT_FMLA_BOOL:
        cached_value = reader.u8()
        if cached_value not in (0, 1):
            raise XlsbParseError("XLSB formula cache has an invalid Boolean value.")
        cached_value = bool(cached_value)
    elif record_type == _BRT_FMLA_ERROR:
        cached_value = _error_literal(reader.u8())
    else:  # pragma: no cover - internal caller uses only BrtFmla record types
        raise AssertionError("not an XLSB formula-cell record")
    reader.take(2)  # formula-cell flags
    formula_length = reader.u32()
    if formula_length > limits.max_formula_bytes:
        raise XlsbLimitError("XLSB formula exceeds max_formula_bytes.")
    return column, cached_value, reader.take(formula_length)


def _parse_xlsb_worksheet(
    sheet: XlsbCoreSheet,
    payload: bytes,
    *,
    shared_strings: Sequence[str],
    external_sheets: Sequence[str],
    defined_name_labels: Sequence[str],
    limits: XlsbReaderLimits,
    formula_text_budget: _FormulaTextBudget,
) -> tuple[dict[tuple[str, int, int], XlsbCoreCell], set[tuple[str, int, int]]]:
    """Read supported cell records from one ordinary worksheet binary part."""
    cells: dict[tuple[str, int, int], XlsbCoreCell] = {}
    unsupported_formula_cells: set[tuple[str, int, int]] = set()
    current_row: int | None = None
    saw_sheet_data = False
    ended_sheet_data = False

    for record in iter_xlsb_records(
        io.BytesIO(payload),
        max_record_bytes=limits.max_record_bytes,
        max_records=limits.max_records_per_part,
    ):
        record_type = record.record_type
        if record_type == _BRT_BEGIN_SHEET_DATA:
            if saw_sheet_data or ended_sheet_data:
                raise XlsbParseError("XLSB worksheet has duplicate BrtBeginSheetData.")
            saw_sheet_data = True
            continue
        if record_type == _BRT_END_SHEET_DATA:
            if not saw_sheet_data or ended_sheet_data:
                raise XlsbParseError("XLSB worksheet has an invalid BrtEndSheetData.")
            ended_sheet_data = True
            continue
        if record_type == _BRT_ROW_HDR:
            if not saw_sheet_data or ended_sheet_data:
                raise XlsbParseError("XLSB worksheet row appears outside sheet data.")
            reader = _Reader(record.payload, context="XLSB row header")
            current_row = reader.u32()
            if current_row > _MAX_EXCEL_ROW_INDEX:
                raise XlsbParseError("XLSB worksheet has an invalid row index.")
            continue
        if record_type not in {
            _BRT_CELL_BLANK,
            _BRT_CELL_RK,
            _BRT_CELL_ERROR,
            _BRT_CELL_BOOL,
            _BRT_CELL_REAL,
            _BRT_CELL_STRING,
            _BRT_CELL_SHARED_STRING,
            _BRT_FMLA_STRING,
            _BRT_FMLA_NUM,
            _BRT_FMLA_BOOL,
            _BRT_FMLA_ERROR,
        }:
            continue
        if not saw_sheet_data or ended_sheet_data or current_row is None:
            raise XlsbParseError("XLSB worksheet cell appears outside a row.")

        if record_type == _BRT_CELL_BLANK:
            reader = _Reader(record.payload, context="XLSB blank cell")
            _cell_header(reader)
            _require_consumed(reader)
            continue

        if record_type in {
            _BRT_FMLA_STRING,
            _BRT_FMLA_NUM,
            _BRT_FMLA_BOOL,
            _BRT_FMLA_ERROR,
        }:
            column, cached_value, token_bytes = _formula_token_bytes(
                record_type, record.payload, limits=limits
            )
            location = (sheet.title, current_row, column)
            if location in cells:
                raise XlsbParseError("XLSB worksheet has duplicate cell records.")
            try:
                formula = decode_xlsb_formula(
                    token_bytes,
                    external_sheets=external_sheets,
                    defined_names=defined_name_labels,
                    max_formula_bytes=limits.max_formula_bytes,
                    max_tokens=limits.max_formula_tokens,
                    max_stack_items=limits.max_formula_stack_items,
                    max_formula_chars=limits.max_formula_characters,
                    max_function_args=limits.max_function_args,
                )
            except XlsbFormulaUnsupportedError:
                formula = None
                unsupported_formula_cells.add(location)
            if formula is not None:
                formula_text_budget.consume(formula)
            cells[location] = XlsbCoreCell(
                sheet=sheet.title,
                row=current_row,
                column=column,
                cell_type="formula",
                value=formula,
                value_type="formula" if formula is not None else "unsupported_formula",
                formula=formula,
                cached_value=cached_value,
            )
        else:
            reader = _Reader(record.payload, context="XLSB cell")
            column = _cell_header(reader)
            if record_type == _BRT_CELL_RK:
                value = _rk_value(reader)
                value_type = "number"
                cell_type = "value"
            elif record_type == _BRT_CELL_ERROR:
                value = _error_literal(reader.u8())
                value_type = "error"
                cell_type = "error"
            elif record_type == _BRT_CELL_BOOL:
                boolean = reader.u8()
                if boolean not in (0, 1):
                    raise XlsbParseError("XLSB cell has an invalid Boolean value.")
                value = bool(boolean)
                value_type = "bool"
                cell_type = "value"
            elif record_type == _BRT_CELL_REAL:
                value = reader.f64()
                if not math.isfinite(value):
                    raise XlsbParseError("XLSB cell has a non-finite numeric value.")
                value_type = "number"
                cell_type = "value"
            elif record_type == _BRT_CELL_STRING:
                value = _wide_string(
                    reader, max_characters=limits.max_cell_text_characters
                )
                value_type = "str"
                cell_type = "value"
            else:
                shared_index = reader.u32()
                if shared_index >= len(shared_strings):
                    raise XlsbParseError("XLSB cell has an invalid shared-string index.")
                value = shared_strings[shared_index]
                value_type = "str"
                cell_type = "value"
            _require_consumed(reader)
            location = (sheet.title, current_row, column)
            if location in cells:
                raise XlsbParseError("XLSB worksheet has duplicate cell records.")
            cells[location] = XlsbCoreCell(
                sheet=sheet.title,
                row=current_row,
                column=column,
                cell_type=cell_type,
                value=value,
                value_type=value_type,
            )

        if len(cells) > limits.max_worksheet_cells:
            raise XlsbLimitError("XLSB worksheet exceeds max_worksheet_cells.")

    if not saw_sheet_data or not ended_sheet_data:
        raise XlsbParseError("XLSB worksheet is missing complete sheet data.")
    return cells, unsupported_formula_cells


def parse_xlsb_workbook_parts(
    parts: Mapping[str, bytes],
    *,
    limits: XlsbReaderLimits = DEFAULT_XLSB_READER_LIMITS,
) -> XlsbCoreWorkbook:
    """Parse the formula-bearing core of an already-safe XLSB package.

    ``parts`` must have unambiguous package names; matching is case-normalized
    only after rejecting case-colliding keys. This helper does not open a ZIP
    file and therefore does not replace FormulaFence's ZIP central-directory,
    member-name, compression, or symlink preflight. Its role is narrower: it
    bounds BIFF12 records and returns only raw sheet, cell, formula-text, and
    defined-name state that the caller can prove was understood.
    """
    limits.validate()
    canonical_parts = _canonical_xlsb_parts(parts)
    retained_bytes = [0]
    workbook_part = _part_payload(
        canonical_parts,
        "xl/workbook.bin",
        limits=limits,
        required=True,
        retained_bytes=retained_bytes,
    )
    relationship_part = _part_payload(
        canonical_parts,
        "xl/_rels/workbook.bin.rels",
        limits=limits,
        required=True,
        retained_bytes=retained_bytes,
    )
    assert workbook_part is not None and relationship_part is not None
    relationships = _workbook_relationships(relationship_part, limits=limits)
    sheets, raw_defined_names, external_sheets, date_1904 = _parse_xlsb_workbook_part(
        workbook_part,
        relationships,
        limits=limits,
    )

    shared_string_part = _part_payload(
        canonical_parts,
        "xl/sharedStrings.bin",
        limits=limits,
        required=False,
        retained_bytes=retained_bytes,
    )
    shared_strings = (
        _parse_shared_strings(shared_string_part, limits=limits)
        if shared_string_part is not None
        else ()
    )

    defined_name_labels = tuple(definition.name for definition in raw_defined_names)
    defined_names: list[XlsbCoreDefinedName] = []
    unsupported_defined_name_indexes: set[int] = set()
    formula_text_budget = _FormulaTextBudget(limits.max_total_formula_characters)
    for index, definition in enumerate(raw_defined_names):
        scope_sheet = (
            sheets[definition.scope_index].title
            if definition.scope_index is not None
            else None
        )
        if not definition.token_bytes:
            # Workbook metadata can declare a future-function name with no
            # expression payload. It contributes no formula text to recover,
            # but its label remains available to a PtgName caller.
            formula = None
        else:
            try:
                formula = decode_xlsb_formula(
                    definition.token_bytes,
                    external_sheets=external_sheets,
                    defined_names=defined_name_labels,
                    max_formula_bytes=limits.max_formula_bytes,
                    max_tokens=limits.max_formula_tokens,
                    max_stack_items=limits.max_formula_stack_items,
                    max_formula_chars=limits.max_formula_characters,
                    max_function_args=limits.max_function_args,
                )
            except XlsbFormulaUnsupportedError:
                formula = None
                unsupported_defined_name_indexes.add(index)
            if formula is not None:
                formula_text_budget.consume(formula)
        defined_names.append(
            XlsbCoreDefinedName(
                name=definition.name,
                scope_sheet=scope_sheet,
                formula=formula,
            )
        )

    cells: dict[tuple[str, int, int], XlsbCoreCell] = {}
    unsupported_formula_cells: set[tuple[str, int, int]] = set()
    unsupported_sheet_types = False
    seen_worksheet_parts: set[str] = set()
    for sheet in sheets:
        if sheet.kind != "worksheet":
            unsupported_sheet_types = True
            continue
        assert sheet.part_name is not None  # Set by _parse_bundle_sheet.
        if sheet.part_name in seen_worksheet_parts:
            raise XlsbParseError("XLSB workbook maps multiple sheets to one worksheet part.")
        seen_worksheet_parts.add(sheet.part_name)
        worksheet_part = _part_payload(
            canonical_parts,
            sheet.part_name,
            limits=limits,
            required=True,
            retained_bytes=retained_bytes,
        )
        assert worksheet_part is not None
        sheet_cells, sheet_unsupported_formula_cells = _parse_xlsb_worksheet(
            sheet,
            worksheet_part,
            shared_strings=shared_strings,
            external_sheets=external_sheets,
            defined_name_labels=defined_name_labels,
            limits=limits,
            formula_text_budget=formula_text_budget,
        )
        cells.update(sheet_cells)
        if len(cells) > limits.max_total_worksheet_cells:
            raise XlsbLimitError("XLSB workbook exceeds max_total_worksheet_cells.")
        unsupported_formula_cells.update(sheet_unsupported_formula_cells)

    return XlsbCoreWorkbook(
        sheets=sheets,
        cells=cells,
        defined_names=tuple(defined_names),
        formula_text_coverage_complete=(
            not unsupported_formula_cells
            and not unsupported_defined_name_indexes
            and not unsupported_sheet_types
        ),
        unsupported_formula_cells=frozenset(unsupported_formula_cells),
        unsupported_defined_name_indexes=frozenset(unsupported_defined_name_indexes),
        unsupported_sheet_types=unsupported_sheet_types,
        date_1904=date_1904,
    )
