"""Small, safe helpers for inspecting—not evaluating—Excel formulas."""

from __future__ import annotations

import re
from collections.abc import Mapping, Sequence
from dataclasses import dataclass

from openpyxl.formula import Tokenizer
from openpyxl.utils.cell import (
    column_index_from_string,
    coordinate_to_tuple,
    range_boundaries,
)

MAX_EXCEL_ROW = 1_048_576
MAX_EXCEL_COLUMN = 16_384
_DYNAMIC_REFERENCE_FUNCTIONS = {"INDIRECT", "OFFSET"}

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


@dataclass(frozen=True)
class FormulaInspection:
    """Static-reference coverage collected from one formula without evaluation."""

    references: tuple[ParsedReference, ...]
    unresolved_range_tokens: tuple[str, ...]
    dynamic_reference_functions: tuple[str, ...]


@dataclass(frozen=True)
class StructuredTable:
    """Static metadata needed to resolve conservative Excel-table references."""

    name: str
    sheet: str
    ref: str
    columns: tuple[str, ...]
    header_row_count: int
    totals_row_count: int


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
    """Parse an A1 token such as `'Inputs'!$B$2:$B$20` without resolving names."""
    sheet, address = _split_sheet_reference(value)
    normalised_sheet, external = _normalise_sheet_name(sheet)
    if external:
        return ParsedReference(None, None, None, None, None, value, is_external=True)

    # Table references and named ranges arrive as RANGE tokens too. Their
    # resolution is deliberately handled by inspect_formula, where a workbook
    # supplied name map is available.
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


def reference_lookup_key(value: str) -> str:
    """Normalize a name-like formula token for a case-insensitive lookup.

    This also canonicalizes a sheet-qualified local name such as
    `'Debt Schedule'!TaxRate` so it can be matched against workbook metadata.
    External tokens are deliberately returned unchanged apart from case because
    they are handled as explicit external references before name resolution.
    """
    token = value.strip()
    sheet, address = _split_sheet_reference(token)
    if sheet is None:
        return token.casefold()
    normalized_sheet, external = _normalise_sheet_name(sheet)
    if normalized_sheet is None or external:
        return token.casefold()
    return f"{normalized_sheet.casefold()}!{address.strip().casefold()}"


def _structured_reference_parts(
    value: str,
) -> tuple[str, tuple[str, ...], tuple[str, ...]] | None:
    """Parse a conservative subset of the bracket syntax after a table name."""
    token = value.strip()
    opening = token.find("[")
    if opening < 0:
        return (token, (), ()) if token else None
    if opening == 0:
        return None  # Unqualified table references need formula-table context.
    table_name = token[:opening].strip()
    if not table_name:
        return None
    selector_parts = _structured_selector_parts(token[opening:])
    if selector_parts is None:
        return None
    groups, separators = selector_parts
    return table_name, groups, separators


def _structured_selector_parts(
    value: str,
) -> tuple[tuple[str, ...], tuple[str, ...]] | None:
    """Parse one bracketed structured-reference selector conservatively.

    The nested-group reader is intentionally small, but it is needed for the
    normal ``[@[Column Name]]`` spelling: its inner column-name brackets are
    not a second selector. Bracket-escaped column names remain outside this
    subset and are left as explicit coverage gaps.
    """
    selector = value.strip()
    if not (selector.startswith("[") and selector.endswith("]")):
        return None
    inner = selector[1:-1].strip()
    if not inner:
        return None
    if not inner.startswith("["):
        return (inner,), ()

    groups: list[str] = []
    separators: list[str] = []
    position = 0
    while position < len(inner):
        while position < len(inner) and inner[position].isspace():
            position += 1
        if position >= len(inner) or inner[position] != "[":
            return None
        closing = _matching_closing_bracket(inner, position)
        if closing < 0:
            return None
        group = inner[position + 1 : closing].strip()
        if not group:
            return None
        groups.append(group)
        position = closing + 1
        while position < len(inner) and inner[position].isspace():
            position += 1
        if position == len(inner):
            break
        if inner[position] not in {",", ":"}:
            return None
        separators.append(inner[position])
        position += 1
    return tuple(groups), tuple(separators)


def _matching_closing_bracket(value: str, opening: int) -> int:
    """Return the matching close bracket for ordinary nested selector groups."""
    depth = 0
    for position in range(opening, len(value)):
        character = value[position]
        if character == "[":
            depth += 1
        elif character == "]":
            depth -= 1
            if depth == 0:
                return position
            if depth < 0:
                return -1
    return -1


def _unqualified_structured_reference_parts(
    value: str,
) -> tuple[tuple[str, ...], tuple[str, ...]] | None:
    """Parse a selector whose table is supplied by a formula-cell context."""
    return _structured_selector_parts(value)


def _unescape_structured_column_name(value: str) -> str:
    """Undo Excel's bracket/header escape prefix for the supported subset."""
    result = value.strip()
    for character in "[#'@":
        result = result.replace(f"'{character}", character)
    return result


def _structured_table_regions(
    table: StructuredTable, items: set[str]
) -> tuple[tuple[int, int], ...] | None:
    try:
        _, min_row, _, max_row = range_boundaries(table.ref)
    except ValueError:  # pragma: no cover - invalid table refs come from malformed OOXML
        return None
    height = max_row - min_row + 1
    header_rows = min(max(table.header_row_count, 0), height)
    totals_rows = min(max(table.totals_row_count, 0), height - header_rows)
    data_start = min_row + header_rows
    data_end = max_row - totals_rows

    if "#all" in items:
        if len(items) != 1:
            return None
        return ((min_row, max_row),)
    selected = items or {"#data"}
    regions: list[tuple[int, int]] = []
    if "#headers" in selected and header_rows:
        regions.append((min_row, data_start - 1))
    if "#data" in selected and data_start <= data_end:
        regions.append((data_start, data_end))
    if "#totals" in selected and totals_rows:
        regions.append((data_end + 1, max_row))
    return tuple(regions)


def _table_data_row_for_origin(
    table: StructuredTable, origin: tuple[str, str] | None
) -> int | None:
    """Return an origin row only when it lies in this table's data body."""
    if origin is None or table.sheet.casefold() != origin[0].casefold():
        return None
    try:
        row, _ = coordinate_to_tuple(origin[1])
        _, min_row, _, max_row = range_boundaries(table.ref)
    except ValueError:  # pragma: no cover - malformed OOXML cannot supply a usable context
        return None
    height = max_row - min_row + 1
    header_rows = min(max(table.header_row_count, 0), height)
    totals_rows = min(max(table.totals_row_count, 0), height - header_rows)
    data_start = min_row + header_rows
    data_end = max_row - totals_rows
    return row if data_start <= row <= data_end else None


def _origin_table(
    tables: Mapping[str, StructuredTable], origin: tuple[str, str] | None
) -> StructuredTable | None:
    """Find the sole table whose data cells contain the formula origin."""
    if origin is None:
        return None
    try:
        _, origin_column = coordinate_to_tuple(origin[1])
    except ValueError:  # pragma: no cover - generated cell coordinates are valid
        return None
    candidates: list[StructuredTable] = []
    seen: set[StructuredTable] = set()
    for table in tables.values():
        if table in seen:
            continue
        seen.add(table)
        if _table_data_row_for_origin(table, origin) is None:
            continue
        try:
            min_column, _, max_column, _ = range_boundaries(table.ref)
        except ValueError:  # pragma: no cover - malformed OOXML cannot supply a usable context
            continue
        if min_column <= origin_column <= max_column:
            candidates.append(table)
    return candidates[0] if len(candidates) == 1 else None


def _current_row_column_name(value: str) -> str | None:
    """Read the column part of one ``@Column`` selector group."""
    column = value.strip()[1:].strip()
    if column.startswith("[") and column.endswith("]"):
        column = column[1:-1].strip()
    return column or None


def _current_row_column_spans(
    table: StructuredTable,
    groups: tuple[str, ...],
    separators: tuple[str, ...],
    *,
    implicit: bool,
) -> tuple[tuple[int, int], ...] | None:
    """Resolve the narrow current-row column subset without evaluating Excel.

    Explicit forms are ``Table[@Column]`` and
    ``Table[[#This Row],[Column]]``. Inside a table data row, Excel also uses
    an unqualified ``[Column]`` as a current-row reference; that form is
    admitted only when ``implicit`` is true and the origin determines one table.
    """
    if not groups or len(separators) != len(groups) - 1:
        return None

    column_lookup = {
        _unescape_structured_column_name(column).casefold(): index
        for index, column in enumerate(table.columns)
    }
    column_positions: list[tuple[int, int]] = []
    current_markers: list[int] = []
    for position, group in enumerate(groups):
        raw_group = group.strip()
        normalized = raw_group.casefold()
        if normalized in {"#this row", "@"}:
            current_markers.append(position)
            continue
        if raw_group.startswith("@"):
            column_name = _current_row_column_name(raw_group)
            if column_name is None:
                return None
            current_markers.append(position)
        else:
            if normalized.startswith("#"):
                return None
            column_name = raw_group
        column_index = column_lookup.get(
            _unescape_structured_column_name(column_name).casefold()
        )
        if column_index is None:
            return None
        column_positions.append((position, column_index))

    if implicit:
        if current_markers or not column_positions:
            return None
    else:
        if len(current_markers) != 1:
            return None
        marker = current_markers[0]
        if marker > 0:
            return None
        if marker < len(separators) and separators[marker] != ",":
            return None
        if marker in {position for position, _ in column_positions} and len(groups) != 1:
            return None

    if not column_positions:
        # ``Table[#This Row]`` / ``Table[@]`` is the whole table row.
        return ((0, len(table.columns) - 1),) if table.columns else None

    column_separators = [
        separator
        for position, separator in enumerate(separators)
        if position not in current_markers
    ]
    if any(separator not in {",", ":"} for separator in column_separators):
        return None
    if ":" in column_separators:
        if len(column_positions) != 2 or column_separators != [":"]:
            return None
        indexes = [index for _, index in column_positions]
        return ((min(indexes), max(indexes)),)
    return tuple((index, index) for _, index in column_positions)


def _current_row_references(
    table: StructuredTable,
    row: int,
    column_spans: tuple[tuple[int, int], ...],
    raw: str,
) -> tuple[ParsedReference, ...] | None:
    try:
        min_column, _, _, _ = range_boundaries(table.ref)
    except ValueError:  # pragma: no cover - invalid table refs come from malformed OOXML
        return None
    return tuple(
        ParsedReference(
            sheet=table.sheet,
            min_column=min_column + start_column,
            min_row=row,
            max_column=min_column + end_column,
            max_row=row,
            raw=raw,
        )
        for start_column, end_column in column_spans
    )


def _resolve_current_row_structured_reference(
    value: str,
    tables: Mapping[str, StructuredTable],
    origin: tuple[str, str] | None,
) -> tuple[ParsedReference, ...] | None:
    """Resolve context-bound row selectors only for a table data-row formula."""
    parsed = _structured_reference_parts(value)
    if parsed is not None:
        table_name, groups, separators = parsed
        table = tables.get(table_name.casefold())
        if table is None:
            return None
        row = _table_data_row_for_origin(table, origin)
        if row is None:
            return None
        spans = _current_row_column_spans(table, groups, separators, implicit=False)
        if spans is None:
            return None
        return _current_row_references(table, row, spans, value)

    selector_parts = _unqualified_structured_reference_parts(value)
    table = _origin_table(tables, origin)
    if selector_parts is None or table is None:
        return None
    row = _table_data_row_for_origin(table, origin)
    if row is None:  # defensive: _origin_table establishes this already
        return None
    groups, separators = selector_parts
    explicit_current_row = any(
        group.strip().casefold() in {"#this row", "@"}
        or group.strip().startswith("@")
        for group in groups
    )
    spans = _current_row_column_spans(
        table, groups, separators, implicit=not explicit_current_row
    )
    if spans is None:
        return None
    return _current_row_references(table, row, spans, value)


def resolve_structured_reference(
    value: str,
    tables: Mapping[str, StructuredTable],
    origin: tuple[str, str] | None = None,
) -> tuple[ParsedReference, ...] | None:
    """Resolve static, fully qualified Excel-table references without evaluation.

    Supported forms include a table name, a single column, contiguous column
    ranges, and the ``#All``, ``#Data``, ``#Headers``, and ``#Totals`` item
    specifiers. When the formula-cell origin is in a table data row, it also
    resolves conservative ``@``/``#This Row`` forms and unqualified table
    columns. Complex bracket escaping remains an explicit coverage gap.
    """
    current_row_reference = _resolve_current_row_structured_reference(value, tables, origin)
    if current_row_reference is not None:
        return current_row_reference
    parsed = _structured_reference_parts(value)
    if parsed is None:
        return None
    table_name, groups, separators = parsed
    table = tables.get(table_name.casefold())
    if table is None:
        return None
    item_tokens = {"#all", "#data", "#headers", "#totals", "#this row"}
    items: set[str] = set()
    columns: list[int] = []
    column_lookup = {
        _unescape_structured_column_name(column).casefold(): index
        for index, column in enumerate(table.columns)
    }
    for group in groups:
        normalized = group.strip().casefold()
        if normalized in item_tokens:
            if normalized == "#this row":
                return None
            items.add(normalized)
            continue
        if normalized.startswith("#") or normalized.startswith("@"):
            return None
        column_index = column_lookup.get(_unescape_structured_column_name(group).casefold())
        if column_index is None:
            return None
        columns.append(column_index)

    try:
        min_column, _, max_column, _ = range_boundaries(table.ref)
    except ValueError:  # pragma: no cover - invalid table refs come from malformed OOXML
        return None
    table_width = max_column - min_column + 1
    if len(columns) > 2 and ":" in separators:
        return None
    if ":" in separators:
        if len(columns) != 2:
            return None
        column_spans = [(min(columns), max(columns))]
    elif columns:
        column_spans = [(column, column) for column in columns]
    else:
        column_spans = [(0, table_width - 1)]
    regions = _structured_table_regions(table, items)
    if regions is None:
        return None
    references = [
        ParsedReference(
            sheet=table.sheet,
            min_column=min_column + start_column,
            min_row=min_row,
            max_column=min_column + end_column,
            max_row=max_row,
            raw=value,
        )
        for start_column, end_column in column_spans
        for min_row, max_row in regions
    ]
    return tuple(dict.fromkeys(references))


def inspect_formula(
    formula: str,
    named_references: Mapping[str, Sequence[ParsedReference]] | None = None,
    structured_tables: Mapping[str, StructuredTable] | None = None,
    origin: tuple[str, str] | None = None,
) -> FormulaInspection:
    """Inspect static reference coverage while resolving known named ranges.

    A caller provides a case-folded name-to-range map assembled from the
    workbook. Supported fully qualified table references are resolved from table
    metadata, while context-bound row references require the formula origin.
    Other non-A1 tokens are returned explicitly instead of being silently
    omitted from the graph.
    """
    try:
        tokens = Tokenizer(formula).items
    except Exception:
        return FormulaInspection((), (), ())
    resolved_names = named_references or {}
    resolved_tables = structured_tables or {}
    references: list[ParsedReference] = []
    unresolved_range_tokens: list[str] = []
    dynamic_reference_functions: list[str] = []
    for token in tokens:
        if token.type == "OPERAND" and token.subtype == "RANGE":
            reference = parse_reference_token(token.value)
            if reference is not None:
                references.append(reference)
                continue
            named_range = resolved_names.get(reference_lookup_key(token.value))
            if named_range:
                references.extend(named_range)
                continue
            table_reference = resolve_structured_reference(
                token.value, resolved_tables, origin
            )
            if table_reference is not None:
                references.extend(table_reference)
                continue
            unresolved_range_tokens.append(token.value)
        elif token.type == "FUNC" and token.subtype == "OPEN":
            function_name = token.value.rstrip("(").strip().upper()
            if function_name in _DYNAMIC_REFERENCE_FUNCTIONS:
                dynamic_reference_functions.append(function_name)
    return FormulaInspection(
        references=tuple(references),
        unresolved_range_tokens=tuple(dict.fromkeys(unresolved_range_tokens)),
        dynamic_reference_functions=tuple(dict.fromkeys(dynamic_reference_functions)),
    )


def extract_references(
    formula: str,
    named_references: Mapping[str, Sequence[ParsedReference]] | None = None,
    structured_tables: Mapping[str, StructuredTable] | None = None,
    origin: tuple[str, str] | None = None,
) -> list[ParsedReference]:
    """Return A1-style, supplied named-range, and static table references."""
    return list(
        inspect_formula(formula, named_references, structured_tables, origin).references
    )


def has_broken_reference(formula: str) -> bool:
    return "#REF!" in formula.upper()
