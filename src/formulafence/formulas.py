"""Small, safe helpers for inspecting—not evaluating—Excel formulas."""

from __future__ import annotations

import re
from dataclasses import dataclass

from openpyxl.formula import Tokenizer
from openpyxl.utils.cell import (
    column_index_from_string,
    coordinate_to_tuple,
    range_boundaries,
)

MAX_EXCEL_ROW = 1_048_576
MAX_EXCEL_COLUMN = 16_384

_CELL_REFERENCE = re.compile(
    r"(?<![A-Z0-9_])(?P<column_absolute>\$?)(?P<column>[A-Z]{1,3})"
    r"(?P<row_absolute>\$?)(?P<row>[1-9][0-9]*)(?![A-Z0-9_])",
    re.IGNORECASE,
)


@dataclass(frozen=True)
class ParsedReference:
    """An explicit A1-style reference found in a formula token."""

    sheet: str | None
    min_column: int | None
    min_row: int | None
    max_column: int | None
    max_row: int | None
    raw: str
    is_external: bool = False

    @property
    def is_range(self) -> bool:
        return (
            self.min_column != self.max_column
            or self.min_row != self.max_row
        )


def _last_unquoted_bang(value: str) -> int:
    """Find the separator in `'A sheet'!A1` while respecting Excel quote escapes."""
    quoted = False
    last = -1
    index = 0
    while index < len(value):
        character = value[index]
        if character == "'":
            if quoted and index + 1 < len(value) and value[index + 1] == "'":
                index += 2
                continue
            quoted = not quoted
        elif character == "!" and not quoted:
            last = index
        index += 1
    return last


def _split_sheet_reference(value: str) -> tuple[str | None, str]:
    separator = _last_unquoted_bang(value)
    if separator < 0:
        return None, value
    return value[:separator], value[separator + 1 :]


def _normalise_sheet_name(value: str | None) -> tuple[str | None, bool]:
    """Return a local sheet title, or mark explicit external workbook syntax."""
    if value is None:
        return None, False
    value = value.strip()
    if "[" in value and "]" in value:
        return None, True
    if len(value) >= 2 and value[0] == "'" and value[-1] == "'":
        value = value[1:-1].replace("''", "'")
    return value, False


def _relative_cell_reference(match: re.Match[str], origin: str) -> str:
    origin_row, origin_column = coordinate_to_tuple(origin)
    column = column_index_from_string(match.group("column").upper())
    row = int(match.group("row"))
    if match.group("row_absolute"):
        normalised_row = f"R{row}"
    else:
        normalised_row = f"R[{row - origin_row}]"
    if match.group("column_absolute"):
        normalised_column = f"C{column}"
    else:
        normalised_column = f"C[{column - origin_column}]"
    return f"{normalised_row}{normalised_column}"


def _normalise_range_token(value: str, origin: str) -> str:
    sheet, address = _split_sheet_reference(value)
    if sheet is not None:
        # Sheet names are case-insensitive in Excel. Preserve quote semantics only
        # where they are needed to make a fingerprint readable.
        clean_sheet, external = _normalise_sheet_name(sheet)
        if external:
            sheet_prefix = "[EXTERNAL]!"
        else:
            sheet_prefix = f"{(clean_sheet or '').casefold()}!"
    else:
        sheet_prefix = ""
    normalised_address = _CELL_REFERENCE.sub(
        lambda match: _relative_cell_reference(match, origin), address.upper()
    )
    return f"{sheet_prefix}{normalised_address}"


def formula_fingerprint(formula: str, origin: str) -> str:
    """Make copied formulas comparable without losing absolute-reference intent.

    FormulaFence does not attempt to parse every Excel function. It relies on
    openpyxl's tokenizer and only rewrites A1 references in `OPERAND/RANGE`
    tokens, leaving literals and function arguments untouched.
    """
    try:
        tokens = Tokenizer(formula).items
    except Exception:  # pragma: no cover - defensive against malformed formulas
        return formula.strip()

    parts: list[str] = []
    for token in tokens:
        if token.type == "WSPACE":
            continue
        if token.type == "OPERAND" and token.subtype == "RANGE":
            parts.append(_normalise_range_token(token.value, origin))
        else:
            parts.append(token.value)
    return "".join(parts)


def parse_reference_token(value: str) -> ParsedReference | None:
    """Parse a token such as `'Inputs'!$B$2:$B$20` without resolving names."""
    sheet, address = _split_sheet_reference(value)
    normalised_sheet, external = _normalise_sheet_name(sheet)
    if external:
        return ParsedReference(None, None, None, None, None, value, is_external=True)

    # Table references and named ranges arrive as RANGE tokens too, but are not
    # statically resolvable without a full Excel evaluator. Leave them out rather
    # than inventing dependencies.
    try:
        min_column, min_row, max_column, max_row = range_boundaries(address)
    except ValueError:
        return None

    min_column = min_column or 1
    max_column = max_column or MAX_EXCEL_COLUMN
    min_row = min_row or 1
    max_row = max_row or MAX_EXCEL_ROW
    return ParsedReference(
        normalised_sheet,
        min_column,
        min_row,
        max_column,
        max_row,
        value,
    )


def extract_references(formula: str) -> list[ParsedReference]:
    """Return explicit A1-style references from a formula.

    Dynamic references constructed with `INDIRECT`, named ranges, structured
    table references, and add-in functions are deliberately not guessed.
    """
    try:
        tokens = Tokenizer(formula).items
    except Exception:
        return []
    references: list[ParsedReference] = []
    for token in tokens:
        if token.type == "OPERAND" and token.subtype == "RANGE":
            reference = parse_reference_token(token.value)
            if reference is not None:
                references.append(reference)
    return references


def has_broken_reference(formula: str) -> bool:
    return "#REF!" in formula.upper()
