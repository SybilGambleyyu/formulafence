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
# These functions can cause a workbook to navigate a reviewer, request content,
# or bind to a host-side data provider.  HYPERLINK also supports in-workbook
# destinations, but its link location may be dynamically computed, so retain
# every call as a reviewable action surface rather than evaluating arguments.
#
# ``STOCKHISTORY`` retrieves provider-backed market history.  The documented
# Cube family can bind a stored formula to a workbook connection, an OLAP server,
# or an offline cube.  Keep all seven Cube functions together: even helpers
# such as ``CUBESETCOUNT`` carry a value whose semantics depend on a cube-set
# chain, and splitting that family would create a misleading partial boundary.
_EXTERNAL_ACTION_CUBE_FUNCTIONS = frozenset(
    {
        "CUBEKPIMEMBER",
        "CUBEMEMBER",
        "CUBEMEMBERPROPERTY",
        "CUBERANKEDMEMBER",
        "CUBESET",
        "CUBESETCOUNT",
        "CUBEVALUE",
    }
)
_EXTERNAL_ACTION_FUNCTIONS = frozenset(
    {
        "HYPERLINK",
        "WEBSERVICE",
        "IMAGE",
        "RTD",
        "STOCKHISTORY",
        *_EXTERNAL_ACTION_CUBE_FUNCTIONS,
    }
)
# ``PY`` is intentionally separate from the remote-content/action ledger. Excel
# stores its executable code in a workbook-level Python part, so callers need a
# dedicated code boundary that can fingerprint both the formula and that part.
_PYTHON_FUNCTIONS = {"PY"}
# ``REGISTER.ID`` is the worksheet-capable member of Excel's legacy
# DLL/code-resource registration family.  Microsoft documents that it
# registers the resource when needed, then returns its registration ID.  Keep
# it separate from generic external-action calls: the module, procedure, type
# string, and arguments are sensitive implementation material, and the
# surrounding XLM-only ``CALL`` / ``REGISTER`` forms are covered by the raw
# macro-sheet boundary instead.
_WORKSHEET_CODE_RESOURCE_REGISTRATION_FUNCTIONS = {"REGISTER.ID"}
# ``REGISTER`` is an XLM macro function rather than a normal worksheet
# function. Microsoft documents that its registration modes can be called from
# a defined-name definition, so FormulaFence inventories it only while it is
# inspecting formula-defined names and named LAMBDAs. Direct worksheet calls
# remain outside this narrow boundary; raw XLM macro sheets have their own
# package-level scanner.
_FORMULA_DEFINED_XLM_REGISTRATION_FUNCTIONS = {"REGISTER"}
# `EVALUATE` is an XLM function that parses a supplied text expression at
# calculation time. Keep it in its own stored-definition boundary rather than
# treating the text as a statically visible formula. Direct worksheet calls
# remain outside this narrow boundary; the caller enables inspection only for
# formula-defined names and named LAMBDAs.
_FORMULA_DEFINED_XLM_EVALUATION_FUNCTIONS = {"EVALUATE"}
# `GET.CELL` is an XLM information function that can observe cell contents,
# formatting, dimensions, protection, and other state outside ordinary
# value/formula dependency semantics. Keep it in a dedicated stored-definition
# boundary. Direct worksheet calls remain outside this narrow boundary; the
# caller enables inspection only for formula-defined names and named LAMBDAs.
_FORMULA_DEFINED_XLM_GET_CELL_FUNCTIONS = {"GET.CELL"}
# These legacy XLM information calls can make a formula-defined name depend on
# workbook, application/workspace, or document state. Keep this intentionally
# small, explicit native set separate from the cell-information boundary.
# Direct worksheet calls remain outside the stored-definition boundary.
_FORMULA_DEFINED_XLM_ENVIRONMENT_INFORMATION_FUNCTIONS = {
    "GET.DOCUMENT",
    "GET.WORKBOOK",
    "GET.WORKSPACE",
}
# ``CELL`` and ``INFO`` can observe the current file, client, folder,
# selection, or other environment state rather than just visible cell
# precedents. ``SHEET`` and ``SHEETS`` can likewise observe the workbook tab
# catalog: sheet numbers, and an omitted SHEETS reference, depend on that
# catalog rather than normal formula precedents. Unlike the intentionally
# narrow XLM boundary above, inspect these native calls wherever a formula can
# be stored. The private marker carries one additional static fact through a
# named-formula chain: ``CELL`` was called without its optional reference.
# Microsoft documents that Excel can then use the currently selected cell at
# calculation time; FormulaFence inventories that surface without evaluating
# the call or simulating a selection.
_FORMULA_ENVIRONMENT_INFORMATION_FUNCTIONS = {"CELL", "INFO", "SHEET", "SHEETS"}
_FORMULA_ENVIRONMENT_INFORMATION_IMPLICIT_CELL_REFERENCE_MARKER = (
    "FORMULAFENCE_CELL_IMPLICIT_REFERENCE_MARKER"
)
_FORMULA_ENVIRONMENT_INFORMATION_IMPLICIT_SHEETS_REFERENCE_MARKER = (
    "FORMULAFENCE_SHEETS_IMPLICIT_REFERENCE_MARKER"
)
# Excel's native function catalog includes a small but important set of dotted
# names.  A namespace separator is also how Office Add-in custom functions are
# displayed (for example, ``CONTOSO.ADD``), so keep these known native names out
# of the custom-function candidate ledger.  This is deliberately a local,
# stable list instead of depending on a version-specific third-party catalog.
_EXCEL_DOTTED_FUNCTIONS = {
    "BETA.DIST",
    "BETA.INV",
    "BINOM.DIST",
    "BINOM.DIST.RANGE",
    "BINOM.INV",
    "CEILING.MATH",
    "CEILING.PRECISE",
    "CHISQ.DIST",
    "CHISQ.DIST.RT",
    "CHISQ.INV",
    "CHISQ.INV.RT",
    "CHISQ.TEST",
    "CONFIDENCE.NORM",
    "CONFIDENCE.T",
    "COVARIANCE.P",
    "COVARIANCE.S",
    "ECMA.CEILING",
    "ERF.PRECISE",
    "ERFC.PRECISE",
    "ERROR.TYPE",
    "EXPON.DIST",
    "F.DIST",
    "F.DIST.RT",
    "F.INV",
    "F.INV.RT",
    "F.TEST",
    "FLOOR.MATH",
    "FLOOR.PRECISE",
    "FORECAST.ETS",
    "FORECAST.ETS.CONFINT",
    "FORECAST.ETS.SEASONALITY",
    "FORECAST.ETS.STAT",
    "FORECAST.LINEAR",
    "GAMMA.DIST",
    "GAMMA.INV",
    "GAMMALN.PRECISE",
    "GET.CELL",
    "GET.DOCUMENT",
    "GET.WORKBOOK",
    "GET.WORKSPACE",
    "HYPGEOM.DIST",
    "ISO.CEILING",
    "LOGNORM.DIST",
    "LOGNORM.INV",
    "MODE.MULT",
    "MODE.SNGL",
    "NEGBINOM.DIST",
    "NETWORKDAYS.INTL",
    "NORM.DIST",
    "NORM.INV",
    "NORM.S.DIST",
    "NORM.S.INV",
    "PERCENTILE.EXC",
    "PERCENTILE.INC",
    "PERCENTRANK.EXC",
    "PERCENTRANK.INC",
    "POISSON.DIST",
    "QUARTILE.EXC",
    "QUARTILE.INC",
    "RANK.AVG",
    "RANK.EQ",
    "REGISTER.ID",
    "SKEW.P",
    "STDEV.P",
    "STDEV.S",
    "T.DIST",
    "T.DIST.2T",
    "T.DIST.RT",
    "T.INV",
    "T.INV.2T",
    "T.TEST",
    "VAR.P",
    "VAR.S",
    "WEIBULL.DIST",
    "WORKDAY.INTL",
    "Z.TEST",
}

_CELL_REFERENCE = re.compile(
    r"(?<![A-Z0-9_])(?P<column_absolute>\$?)(?P<column>[A-Z]{1,3})"
    r"(?P<row_absolute>\$?)(?P<row>[1-9][0-9]*)(?![A-Z0-9_])",
    re.IGNORECASE,
)
_LOCAL_IDENTIFIER = re.compile(r"[A-Z_][A-Z0-9_]*", re.IGNORECASE)
_A1_LOCAL_IDENTIFIER_CONFLICT = re.compile(r"[A-Z]{1,3}[1-9][0-9]*", re.IGNORECASE)
_R1C1_LOCAL_IDENTIFIER_CONFLICT = re.compile(
    r"R[1-9][0-9]*C[1-9][0-9]*", re.IGNORECASE
)
_SERIALIZED_LOCAL_PREFIXES = ("_xlpm.", "_xlop.")
_GROUP_TOKEN_TYPES = {"FUNC", "PAREN", "ARRAY"}
_WHITESPACE_TOKEN_TYPES = {"WSPACE", "WHITE-SPACE"}
_SPILL_REFERENCE_FUNCTION = "ANCHORARRAY"
_IMPLICIT_INTERSECTION_FUNCTION = "SINGLE"
# This private tokenizer-only wrapper preserves the semantic distinction of a
# literal ``@A1`` while avoiding an ambiguity with a user-authored SINGLE().
_LITERAL_IMPLICIT_INTERSECTION_FUNCTION = "_FORMULAFENCE_IMPLICIT_INTERSECTION"
# ``#`` is not understood by openpyxl's tokenizer, even though it is Excel's
# display syntax for a spilled-array reference. Keep the accepted grammar
# intentionally narrow: one internal A1 anchor, optionally sheet-qualified.
# External, 3-D, range, named, malformed, and implicit-intersection variants
# remain tokenizer coverage limits instead of being rewritten into a guess.
_LITERAL_SPILL_REFERENCE = re.compile(
    r"(?<![A-Z0-9_.'\"!:@$\[\]])"
    r"(?P<reference>(?:(?:'(?:[^']|'')*'|[A-Z0-9_.]+)!)?"
    r"\$?[A-Z]{1,3}\$?[1-9][0-9]*)"
    r"#(?=$|[ \t\r\n,;)}+\-*/^&=<>%])",
    re.IGNORECASE,
)
# Like ``#``, Excel's display-only implicit-intersection operator is outside
# openpyxl's reference grammar. Only rewrite a direct, internal A1 reference or
# A1 range. Table current-row syntax, named expressions, external references,
# 3-D spans, spill combinations, and malformed forms remain untouched so they
# can never be mistaken for a static dependency.
_LITERAL_IMPLICIT_INTERSECTION_REFERENCE = re.compile(
    r"(?<![A-Z0-9_.'\"!:@$\[\]])"
    r"@(?P<reference>(?:(?:'(?:[^']|'')*'|[A-Z0-9_.]+)!)?"
    r"\$?[A-Z]{1,3}\$?[1-9][0-9]*(?::\$?[A-Z]{1,3}\$?[1-9][0-9]*)?)"
    r"(?=$|[ \t\r\n,;)}+\-*/^&=<>%])",
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
    """Static formula coverage collected from one formula without evaluation."""

    references: tuple[ParsedReference, ...]
    unresolved_range_tokens: tuple[str, ...]
    dynamic_reference_functions: tuple[str, ...]
    external_action_functions: tuple[str, ...] = ()
    python_functions: tuple[str, ...] = ()
    office_custom_function_candidates: tuple[str, ...] = ()
    worksheet_code_resource_registration_functions: tuple[str, ...] = ()
    formula_defined_xlm_registration_functions: tuple[str, ...] = ()
    formula_defined_xlm_evaluation_functions: tuple[str, ...] = ()
    formula_defined_xlm_get_cell_functions: tuple[str, ...] = ()
    formula_defined_xlm_environment_information_functions: tuple[str, ...] = ()
    formula_environment_information_functions: tuple[str, ...] = ()
    formula_environment_information_implicit_cell_reference_count: int = 0
    formula_environment_information_implicit_sheets_reference_count: int = 0
    three_d_reference_tokens: tuple[str, ...] = ()
    tokenization_failed: bool = False
    spill_reference_tokens: tuple[str, ...] = ()
    implicit_intersection_tokens: tuple[str, ...] = ()

    @property
    def formula_environment_information_function_count(self) -> int:
        """Return native information calls without exposing their arguments."""
        return len(self.formula_environment_information_functions)

    @property
    def formula_environment_information_signal_values(self) -> tuple[str, ...]:
        """Return private propagation values for named-formula resolution.

        The marker is deliberately not part of
        ``formula_environment_information_functions`` so normal consumers see
        only native function names. They remain necessary internally to retain
        statically visible omitted-reference distinctions through a named formula
        or named LAMBDA invocation chain.
        """
        return (
            self.formula_environment_information_functions
            + (
                _FORMULA_ENVIRONMENT_INFORMATION_IMPLICIT_CELL_REFERENCE_MARKER,
            )
            * self.formula_environment_information_implicit_cell_reference_count
            + (
                _FORMULA_ENVIRONMENT_INFORMATION_IMPLICIT_SHEETS_REFERENCE_MARKER,
            )
            * self.formula_environment_information_implicit_sheets_reference_count
        )


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
    tokens, _, _ = _tokenize_formula(formula, preserve_literal_spill_operator=True)
    if tokens is None:  # pragma: no cover - defensive against malformed formulas
        return formula.strip()

    parts: list[str] = []
    for token in tokens:
        if _is_whitespace(token):
            continue
        if token.type == "OPERAND" and token.subtype == "RANGE":
            parts.append(_normalise_range_token(token.value, origin))
        else:
            parts.append(_fingerprint_token_value(token))
    return "".join(parts)


def parse_reference_token(value: str) -> ParsedReference | None:
    """Parse an A1 token such as `'Inputs'!$B$2:$B$20` without resolving names."""
    sheet, address = _split_sheet_reference(value)
    normalised_sheet, external = _normalise_sheet_name(sheet)
    if external:
        return ParsedReference(None, None, None, None, None, value, is_external=True)
    if normalised_sheet is not None and ":" in normalised_sheet:
        # A 3-D reference needs workbook tab order to resolve safely. Returning
        # None here prevents an invented ``Sheet1:Sheet3`` pseudo-sheet edge.
        return None

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


def _mask_double_quoted_strings(value: str) -> str:
    """Blank formula strings while preserving offsets for token-safe matching."""
    masked = list(value)
    in_string = False
    position = 0
    while position < len(value):
        character = value[position]
        if not in_string:
            if character == '"':
                masked[position] = " "
                in_string = True
            position += 1
            continue

        masked[position] = " "
        if character == '"':
            if position + 1 < len(value) and value[position + 1] == '"':
                masked[position + 1] = " "
                position += 2
                continue
            in_string = False
        position += 1
    return "".join(masked)


def _rewrite_literal_spill_references(
    formula: str, *, preserve_operator: bool = False
) -> tuple[str, tuple[str, ...]]:
    """Remove only static ``A1#`` operators so openpyxl can inspect the anchor.

    The graph edge is deliberately to the formula's anchor cell, not to an
    invented fixed spill extent. The caller retains the original spill token in
    ``FormulaInspection`` so profiles and policy can make the remaining dynamic
    shape and blocker behavior visible.
    """
    masked_formula = _mask_double_quoted_strings(formula)
    parts: list[str] = []
    literal_spill_tokens: list[str] = []
    cursor = 0
    for match in _LITERAL_SPILL_REFERENCE.finditer(masked_formula):
        reference = formula[match.start("reference") : match.end("reference")]
        if parse_reference_token(reference) is None:
            continue
        parts.append(formula[cursor : match.start()])
        if preserve_operator:
            # Keep a spill reference distinct from an ordinary direct cell in
            # formula fingerprints while still giving the underlying tokenizer
            # a grammar it understands. This is equivalent to the OOXML-style
            # ANCHORARRAY representation documented by XlsxWriter.
            parts.append(f"{_SPILL_REFERENCE_FUNCTION}({reference})")
        else:
            parts.append(reference)
        cursor = match.end()
        literal_spill_tokens.append(formula[match.start() : match.end()])
    if not literal_spill_tokens:
        return formula, ()
    parts.append(formula[cursor:])
    return "".join(parts), tuple(dict.fromkeys(literal_spill_tokens))


def _rewrite_literal_implicit_intersection_references(
    formula: str,
) -> tuple[str, tuple[str, ...]]:
    """Make direct ``@A1`` references tokenizable without guessing their scope.

    Excel normally serializes persisted implicit intersection as ``SINGLE()``,
    but the display syntax can still appear in authoring tools. The private
    wrapper lets the rest of the inspection retain the operator as a distinct
    semantic feature and, when an origin is available, select a direct static
    range precisely.
    """
    masked_formula = _mask_double_quoted_strings(formula)
    parts: list[str] = []
    literal_tokens: list[str] = []
    cursor = 0
    for match in _LITERAL_IMPLICIT_INTERSECTION_REFERENCE.finditer(masked_formula):
        reference = formula[match.start("reference") : match.end("reference")]
        if parse_reference_token(reference) is None:
            continue
        parts.append(formula[cursor : match.start()])
        parts.append(f"{_LITERAL_IMPLICIT_INTERSECTION_FUNCTION}({reference})")
        cursor = match.end()
        literal_tokens.append(formula[match.start() : match.end()])
    if not literal_tokens:
        return formula, ()
    parts.append(formula[cursor:])
    return "".join(parts), tuple(dict.fromkeys(literal_tokens))


def _tokenize_formula(
    formula: str, *, preserve_literal_spill_operator: bool = False
) -> tuple[tuple[object, ...] | None, tuple[str, ...], tuple[str, ...]]:
    """Tokenize after narrow spill and implicit-intersection compatibility passes."""
    tokenizer_formula, literal_implicit_intersection_tokens = (
        _rewrite_literal_implicit_intersection_references(formula)
    )
    tokenizer_formula, literal_spill_tokens = _rewrite_literal_spill_references(
        tokenizer_formula, preserve_operator=preserve_literal_spill_operator
    )
    try:
        return (
            tuple(Tokenizer(tokenizer_formula).items),
            literal_spill_tokens,
            literal_implicit_intersection_tokens,
        )
    except Exception:
        return None, literal_spill_tokens, literal_implicit_intersection_tokens


def resolve_3d_reference(
    value: str, sheet_order: Sequence[str] | None
) -> tuple[ParsedReference, ...] | None:
    """Expand a static Excel 3-D A1 reference across its tab-order endpoints.

    Excel treats ``Sales:Marketing!B3`` as the same cell on every worksheet
    from ``Sales`` through ``Marketing`` in workbook tab order. There is no
    safe single-sheet approximation, so unknown, external, malformed, or
    endpoint-missing references deliberately return ``None`` for coverage
    reporting instead of creating a fictitious sheet dependency.
    """
    if not sheet_order:
        return None
    sheet, address = _split_sheet_reference(value)
    normalised_sheet, external = _normalise_sheet_name(sheet)
    if external or normalised_sheet is None:
        return None
    first_sheet, separator, last_sheet = normalised_sheet.partition(":")
    if not separator or not first_sheet or not last_sheet or ":" in last_sheet:
        return None

    sheet_positions = {title.casefold(): position for position, title in enumerate(sheet_order)}
    first_position = sheet_positions.get(first_sheet.casefold())
    last_position = sheet_positions.get(last_sheet.casefold())
    if first_position is None or last_position is None or first_position > last_position:
        return None
    try:
        min_column, min_row, max_column, max_row = range_boundaries(address)
    except ValueError:
        return None

    min_column = min_column or 1
    max_column = max_column or MAX_EXCEL_COLUMN
    min_row = min_row or 1
    max_row = max_row or MAX_EXCEL_ROW
    return tuple(
        ParsedReference(
            sheet=sheet_order[position],
            min_column=min_column,
            min_row=min_row,
            max_column=max_column,
            max_row=max_row,
            raw=value,
        )
        for position in range(first_position, last_position + 1)
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


def _is_group_open(token: object) -> bool:
    """Return whether an openpyxl token starts a nested expression group."""
    return (
        getattr(token, "type", None) in _GROUP_TOKEN_TYPES
        and getattr(token, "subtype", None) == "OPEN"
    )


def _is_group_close(token: object) -> bool:
    """Return whether an openpyxl token ends a nested expression group."""
    return (
        getattr(token, "type", None) in _GROUP_TOKEN_TYPES
        and getattr(token, "subtype", None) == "CLOSE"
    )


def _is_whitespace(token: object) -> bool:
    """Handle the whitespace labels emitted by supported openpyxl versions."""
    return getattr(token, "type", None) in _WHITESPACE_TOKEN_TYPES


def _matching_group_close(tokens: Sequence[object], opening: int, end: int) -> int | None:
    """Find a matching formula-token group close without parsing Excel values."""
    depth = 0
    for position in range(opening, end):
        if _is_group_open(tokens[position]):
            depth += 1
        elif _is_group_close(tokens[position]):
            depth -= 1
            if depth == 0:
                return position
    return None


def _function_argument_spans(
    tokens: Sequence[object], start: int, end: int
) -> tuple[tuple[int, int], ...]:
    """Split one function body into top-level comma-separated token spans."""
    spans: list[tuple[int, int]] = []
    argument_start = start
    depth = 0
    for position in range(start, end):
        token = tokens[position]
        if _is_group_open(token):
            depth += 1
        elif _is_group_close(token):
            depth -= 1
        elif (
            depth == 0
            and getattr(token, "type", None) == "SEP"
            and getattr(token, "subtype", None) == "ARG"
        ):
            spans.append((argument_start, position))
            argument_start = position + 1
    spans.append((argument_start, end))
    return tuple(spans)


def _local_scope_key(value: str) -> str:
    """Normalize a local name while hiding OOXML parameter-marker prefixes."""
    token = value.strip()
    normalized_token = token.casefold()
    for prefix in _SERIALIZED_LOCAL_PREFIXES:
        if normalized_token.startswith(prefix):
            identifier = token[len(prefix) :]
            if _LOCAL_IDENTIFIER.fullmatch(identifier):
                return identifier.casefold()
    return reference_lookup_key(token)


def _simple_local_identifier(
    tokens: Sequence[object],
    start: int,
    end: int,
    *,
    allow_serialized_local_prefix: bool = False,
) -> tuple[int, str] | None:
    """Return one conservative LET/LAMBDA identifier declaration token.

    Excel-compatible writers may serialize LAMBDA parameters or local values
    with ``_xlpm.`` or ``_xlop.`` prefixes. They are accepted only at a
    LET/LAMBDA declaration so an ordinary dotted defined name cannot be
    mistaken for a local variable.
    """
    meaningful = [
        position
        for position in range(start, end)
        if not _is_whitespace(tokens[position])
    ]
    if len(meaningful) != 1:
        return None
    position = meaningful[0]
    token = tokens[position]
    if not (
        getattr(token, "type", None) == "OPERAND"
        and getattr(token, "subtype", None) == "RANGE"
    ):
        return None
    value = str(getattr(token, "value", "")).strip()
    identifier = value
    if allow_serialized_local_prefix:
        normalized_value = value.casefold()
        for prefix in _SERIALIZED_LOCAL_PREFIXES:
            if normalized_value.startswith(prefix):
                identifier = value[len(prefix) :]
                break
    if not _LOCAL_IDENTIFIER.fullmatch(identifier):
        return None
    if (
        identifier.casefold() in {"r", "c"}
        or _A1_LOCAL_IDENTIFIER_CONFLICT.fullmatch(identifier)
        or _R1C1_LOCAL_IDENTIFIER_CONFLICT.fullmatch(identifier)
    ):
        return None
    return position, _local_scope_key(value)


def _function_name(token: object) -> str:
    """Normalize function names, including Excel's OOXML namespace prefixes."""
    value = str(getattr(token, "value", "")).rstrip("(").strip().upper()
    return value.rsplit(".", 1)[-1].lstrip("@")


def _function_lookup_key(token: object) -> str:
    """Normalize a callable defined-name token without losing sheet qualification."""
    return reference_lookup_key(str(getattr(token, "value", "")).rstrip("(").strip())


def _office_custom_function_candidate(token: object) -> str | None:
    """Return a normalized namespaced custom-function candidate, if safe.

    Office Add-in custom functions are displayed as a namespace followed by a
    function name, but their manifest and JavaScript runtime are not embedded
    in a normal workbook.  This intentionally recognizes only the documented
    namespaced spelling, excludes OOXML compatibility prefixes and native
    dotted functions, and leaves every other UDF shape out of this boundary.
    """
    raw_name = str(getattr(token, "value", "")).rstrip("(").strip()
    candidate = raw_name.removeprefix("@")
    upper_candidate = candidate.upper()
    if (
        not candidate
        or upper_candidate.startswith(("_XLFN.", "_XLWS."))
        or candidate.startswith("_")
        or "." not in candidate
        or any(character in candidate for character in "![]'")
        or any(character.isspace() for character in candidate)
        or upper_candidate in _EXCEL_DOTTED_FUNCTIONS
    ):
        return None
    namespace, function_name = candidate.split(".", 1)
    if (
        not namespace
        or not function_name
        or not namespace[0].isalpha()
        or not function_name[0].isalpha()
        or any(not segment for segment in candidate.split("."))
        or any(
            not (character.isalnum() or character in "_.")
            for character in candidate
        )
    ):
        return None
    return upper_candidate


def _worksheet_code_resource_registration_function(token: object) -> str | None:
    """Return a documented worksheet code-resource registration call, if present.

    ``_function_name`` deliberately drops dotted namespace prefixes for broad
    native-function handling, so this narrow classifier works from the raw
    callable spelling instead.  A defined name with the same spelling is
    resolved before this classifier runs; FormulaFence will not assert that a
    user-defined callable is Excel's legacy registration primitive.
    """
    raw_name = str(getattr(token, "value", "")).rstrip("(").strip()
    normalized = raw_name.removeprefix("@").upper()
    if normalized in _WORKSHEET_CODE_RESOURCE_REGISTRATION_FUNCTIONS:
        return normalized
    return None


def _formula_defined_xlm_registration_function(token: object) -> str | None:
    """Return a defined-name-only XLM ``REGISTER`` call, if present.

    This raw-spelling classifier intentionally does not run for ordinary
    worksheet formula inspection. A caller can enable it only while examining
    a stored formula-defined name or named LAMBDA, and known defined callables
    are resolved before it runs so a user-defined ``REGISTER`` name is never
    asserted to be Excel's XLM primitive.
    """
    raw_name = str(getattr(token, "value", "")).rstrip("(").strip()
    normalized = raw_name.removeprefix("@").upper()
    if normalized in _FORMULA_DEFINED_XLM_REGISTRATION_FUNCTIONS:
        return normalized
    return None


def _formula_defined_xlm_evaluation_function(token: object) -> str | None:
    """Return a defined-name-only XLM `EVALUATE` call, if present.

    The raw-spelling classifier is deliberately opt-in and only used while
    examining a stored formula-defined name or named LAMBDA. Known defined
    callables are resolved before it runs, so a workbook-defined `EVALUATE`
    name is not asserted to be Excel's legacy expression-evaluation primitive.
    """
    raw_name = str(getattr(token, "value", "")).rstrip("(").strip()
    normalized = raw_name.removeprefix("@").upper()
    if normalized in _FORMULA_DEFINED_XLM_EVALUATION_FUNCTIONS:
        return normalized
    return None


def _formula_defined_xlm_get_cell_function(token: object) -> str | None:
    """Return a defined-name-only XLM `GET.CELL` call, if present.

    The raw-spelling classifier is deliberately opt-in and only used while
    examining a stored formula-defined name or named LAMBDA. Known defined
    callables are resolved before it runs, so a workbook-defined `GET.CELL`
    name is not asserted to be Excel's legacy information primitive.
    """
    raw_name = str(getattr(token, "value", "")).rstrip("(").strip()
    normalized = raw_name.removeprefix("@").upper()
    if normalized in _FORMULA_DEFINED_XLM_GET_CELL_FUNCTIONS:
        return normalized
    return None


def _formula_defined_xlm_environment_information_function(
    token: object,
) -> str | None:
    """Return an opt-in stored-definition XLM environment-information call.

    Known defined callables are resolved before this raw-spelling classifier
    runs, so a workbook-defined name cannot be asserted to be one of Excel's
    legacy native information primitives.
    """
    raw_name = str(getattr(token, "value", "")).rstrip("(").strip()
    normalized = raw_name.removeprefix("@").upper()
    if normalized in _FORMULA_DEFINED_XLM_ENVIRONMENT_INFORMATION_FUNCTIONS:
        return normalized
    return None


def _formula_environment_information_function(
    tokens: Sequence[object], position: int
) -> tuple[str, bool, bool] | None:
    """Return a native information call and statically visible omitted arguments.

    This only reads token structure.  It intentionally does not inspect the
    information type, calculate a formula, resolve a dynamic argument, or
    infer the active cell. A malformed CELL or SHEETS call still remains an
    inventory item, but cannot safely be asserted to have the documented
    omitted-reference behavior.
    """
    token = tokens[position]
    raw_name = str(getattr(token, "value", "")).rstrip("(").strip()
    normalized = raw_name.removeprefix("@").upper()
    if normalized not in _FORMULA_ENVIRONMENT_INFORMATION_FUNCTIONS:
        return None
    if normalized not in {"CELL", "SHEETS"}:
        return normalized, False, False

    closing = _matching_group_close(tokens, position, len(tokens))
    if closing is None:
        return normalized, False, False
    arguments = _function_argument_spans(tokens, position + 1, closing)
    has_single_nonempty_argument = len(arguments) == 1 and any(
        not _is_whitespace(tokens[index])
        for index in range(arguments[0][0], arguments[0][1])
    )
    has_nonempty_argument = any(
        not _is_whitespace(tokens[index])
        for start, end in arguments
        for index in range(start, end)
    )
    if normalized == "CELL":
        return normalized, has_single_nonempty_argument, False
    # Only the documented no-argument form is a statically reliable workbook
    # count.  Keep malformed multi-argument calls in the private inventory, but
    # do not treat e.g. ``SHEETS(,)`` as an omitted reference.
    return normalized, False, len(arguments) == 1 and not has_nonempty_argument


def _fingerprint_token_value(token: object) -> str:
    """Normalize OOXML spellings of the two dynamic-array compatibility functions."""
    if (
        getattr(token, "type", None) == "FUNC"
        and getattr(token, "subtype", None) == "OPEN"
    ):
        function_name = _function_name(token)
        if function_name == _SPILL_REFERENCE_FUNCTION:
            return f"{_SPILL_REFERENCE_FUNCTION}("
        if function_name in {
            _IMPLICIT_INTERSECTION_FUNCTION,
            _LITERAL_IMPLICIT_INTERSECTION_FUNCTION,
        }:
            return f"{_IMPLICIT_INTERSECTION_FUNCTION}("
    return str(getattr(token, "value", ""))


def lambda_parameter_count(formula: str) -> int | None:
    """Return a valid top-level LAMBDA definition's parameter count, if present.

    A workbook-defined LAMBDA is callable only when its full definition is one
    ``LAMBDA(...)`` expression. This intentionally rejects formulas that merely
    contain a lambda, malformed parameter lists, and trailing invocations so a
    caller cannot mistake an arbitrary defined formula for a custom function.
    """
    tokens, _, _ = _tokenize_formula(formula)
    if tokens is None:
        return None
    meaningful = [
        position
        for position, token in enumerate(tokens)
        if not _is_whitespace(token)
    ]
    if not meaningful:
        return None
    opening = meaningful[0]
    token = tokens[opening]
    if not (
        getattr(token, "type", None) == "FUNC"
        and getattr(token, "subtype", None) == "OPEN"
        and _function_name(token) == "LAMBDA"
    ):
        return None
    closing = _matching_group_close(tokens, opening, len(tokens))
    if closing is None or any(position > closing for position in meaningful):
        return None
    arguments = _function_argument_spans(tokens, opening + 1, closing)
    if not arguments:
        return None
    calculation_start, calculation_end = arguments[-1]
    if not any(
        not _is_whitespace(tokens[position])
        for position in range(calculation_start, calculation_end)
    ):
        return None
    parameters: set[str] = set()
    for start, end in arguments[:-1]:
        declaration = _simple_local_identifier(
            tokens, start, end, allow_serialized_local_prefix=True
        )
        if declaration is None or declaration[1] in parameters:
            return None
        parameters.add(declaration[1])
    return len(parameters)


def _local_variable_token_indexes(tokens: Sequence[object]) -> set[int]:
    """Find LET/LAMBDA declarations and lexical uses masquerading as ranges.

    openpyxl's tokenizer correctly preserves the argument structure but labels
    Excel's local names as ``OPERAND/RANGE``. This walk is deliberately narrow:
    malformed declarations fall back to ordinary inspection, where their tokens
    remain visible as unresolved rather than being suppressed.
    """
    local_indexes: set[int] = set()

    def visit(start: int, end: int, scope: frozenset[str]) -> None:
        position = start
        while position < end:
            token = tokens[position]
            if _is_group_open(token):
                closing = _matching_group_close(tokens, position, end)
                if closing is None:
                    position += 1
                    continue
                body_start = position + 1
                if getattr(token, "type", None) == "FUNC":
                    arguments = _function_argument_spans(tokens, body_start, closing)
                    name = _function_name(token)
                    local_callable = _local_scope_key(
                        str(getattr(token, "value", "")).rstrip("(")
                    ) in scope
                    if local_callable:
                        local_indexes.add(position)
                    if name == "LET" and not local_callable:
                        visit_let(arguments, scope)
                    elif name == "LAMBDA" and not local_callable:
                        visit_lambda(arguments, scope)
                    else:
                        visit(body_start, closing, scope)
                else:
                    visit(body_start, closing, scope)
                position = closing + 1
                continue
            if (
                getattr(token, "type", None) == "OPERAND"
                and getattr(token, "subtype", None) == "RANGE"
                and _local_scope_key(str(getattr(token, "value", ""))) in scope
            ):
                local_indexes.add(position)
            position += 1

    def visit_let(arguments: Sequence[tuple[int, int]], scope: frozenset[str]) -> None:
        if len(arguments) < 3 or len(arguments) % 2 == 0:
            for start, end in arguments:
                visit(start, end, scope)
            return
        declarations: list[tuple[int, str]] = []
        for start, end in arguments[:-1:2]:
            declaration = _simple_local_identifier(
                tokens, start, end, allow_serialized_local_prefix=True
            )
            if declaration is None:
                for fallback_start, fallback_end in arguments:
                    visit(fallback_start, fallback_end, scope)
                return
            declarations.append(declaration)

        active_scope = set(scope)
        for pair_index, (declaration_index, identifier) in enumerate(declarations):
            local_indexes.add(declaration_index)
            value_start, value_end = arguments[pair_index * 2 + 1]
            visit(value_start, value_end, frozenset(active_scope))
            active_scope.add(identifier)
        final_start, final_end = arguments[-1]
        visit(final_start, final_end, frozenset(active_scope))

    def visit_lambda(arguments: Sequence[tuple[int, int]], scope: frozenset[str]) -> None:
        if len(arguments) < 2:
            for start, end in arguments:
                visit(start, end, scope)
            return
        declarations: list[tuple[int, str]] = []
        for start, end in arguments[:-1]:
            declaration = _simple_local_identifier(
                tokens, start, end, allow_serialized_local_prefix=True
            )
            if declaration is None:
                for fallback_start, fallback_end in arguments:
                    visit(fallback_start, fallback_end, scope)
                return
            declarations.append(declaration)
        local_scope = set(scope)
        for declaration_index, identifier in declarations:
            local_indexes.add(declaration_index)
            local_scope.add(identifier)
        body_start, body_end = arguments[-1]
        visit(body_start, body_end, frozenset(local_scope))

    try:
        visit(0, len(tokens), frozenset())
    except Exception:  # pragma: no cover - conservative fallback for unknown token shapes
        return set()
    return local_indexes


def _implicit_intersection_selection(
    reference: ParsedReference, origin: tuple[str, str]
) -> ParsedReference | None:
    """Return the one static cell selected by a direct implicit intersection.

    Microsoft documents implicit intersection as selecting the input on the
    formula's row or column. A one-dimensional range can therefore be resolved
    exactly. For a two-dimensional range, selection is only certain when the
    formula cell lies within it. All other forms deliberately fall back to the
    normal conservative range edge instead of inventing a selected cell.
    """
    if reference.is_external or None in {
        reference.min_column,
        reference.min_row,
        reference.max_column,
        reference.max_row,
    }:
        return None
    try:
        origin_row, origin_column = coordinate_to_tuple(origin[1])
    except ValueError:
        return None

    min_column = reference.min_column
    min_row = reference.min_row
    max_column = reference.max_column
    max_row = reference.max_row
    if min_column == max_column and min_row == max_row:
        selected_column, selected_row = min_column, min_row
    elif min_column == max_column and min_row <= origin_row <= max_row:
        selected_column, selected_row = min_column, origin_row
    elif min_row == max_row and min_column <= origin_column <= max_column:
        selected_column, selected_row = origin_column, min_row
    elif (
        min_column <= origin_column <= max_column
        and min_row <= origin_row <= max_row
    ):
        selected_column, selected_row = origin_column, origin_row
    else:
        return None
    return ParsedReference(
        reference.sheet,
        selected_column,
        selected_row,
        selected_column,
        selected_row,
        raw=reference.raw,
        is_external=reference.is_external,
    )


def _implicit_intersection_reference_replacements(
    tokens: Sequence[object], origin: tuple[str, str] | None
) -> dict[int, ParsedReference]:
    """Map direct SINGLE() operands to precise static dependency cells."""
    if origin is None:
        return {}
    replacements: dict[int, ParsedReference] = {}
    for position, token in enumerate(tokens):
        if not (
            getattr(token, "type", None) == "FUNC"
            and getattr(token, "subtype", None) == "OPEN"
            and _function_name(token)
            in {
                _IMPLICIT_INTERSECTION_FUNCTION,
                _LITERAL_IMPLICIT_INTERSECTION_FUNCTION,
            }
        ):
            continue
        closing = _matching_group_close(tokens, position, len(tokens))
        if closing is None:
            continue
        arguments = _function_argument_spans(tokens, position + 1, closing)
        if len(arguments) != 1:
            continue
        start, end = arguments[0]
        meaningful = [
            index for index in range(start, end) if not _is_whitespace(tokens[index])
        ]
        if len(meaningful) != 1:
            continue
        reference_position = meaningful[0]
        reference_token = tokens[reference_position]
        if not (
            getattr(reference_token, "type", None) == "OPERAND"
            and getattr(reference_token, "subtype", None) == "RANGE"
        ):
            continue
        reference = parse_reference_token(str(getattr(reference_token, "value", "")))
        if reference is None:
            continue
        selected = _implicit_intersection_selection(reference, origin)
        if selected is not None:
            replacements[reference_position] = selected
    return replacements


def inspect_formula(
    formula: str,
    named_references: Mapping[str, Sequence[ParsedReference]] | None = None,
    structured_tables: Mapping[str, StructuredTable] | None = None,
    origin: tuple[str, str] | None = None,
    sheet_order: Sequence[str] | None = None,
    named_function_references: (
        Mapping[str, Sequence[ParsedReference] | None] | None
    ) = None,
    named_custom_function_candidates: Mapping[str, Sequence[str]] | None = None,
    named_function_custom_function_candidates: (
        Mapping[str, Sequence[str]] | None
    ) = None,
    named_formula_external_action_functions: (
        Mapping[str, Sequence[str]] | None
    ) = None,
    named_function_formula_external_action_functions: (
        Mapping[str, Sequence[str]] | None
    ) = None,
    named_worksheet_code_resource_registration_functions: (
        Mapping[str, Sequence[str]] | None
    ) = None,
    named_function_worksheet_code_resource_registration_functions: (
        Mapping[str, Sequence[str]] | None
    ) = None,
    named_formula_defined_xlm_registration_functions: (
        Mapping[str, Sequence[str]] | None
    ) = None,
    named_function_formula_defined_xlm_registration_functions: (
        Mapping[str, Sequence[str]] | None
    ) = None,
    named_formula_defined_xlm_evaluation_functions: (
        Mapping[str, Sequence[str]] | None
    ) = None,
    named_function_formula_defined_xlm_evaluation_functions: (
        Mapping[str, Sequence[str]] | None
    ) = None,
    named_formula_defined_xlm_get_cell_functions: (
        Mapping[str, Sequence[str]] | None
    ) = None,
    named_function_formula_defined_xlm_get_cell_functions: (
        Mapping[str, Sequence[str]] | None
    ) = None,
    named_formula_defined_xlm_environment_information_functions: (
        Mapping[str, Sequence[str]] | None
    ) = None,
    named_function_formula_defined_xlm_environment_information_functions: (
        Mapping[str, Sequence[str]] | None
    ) = None,
    named_formula_environment_information_functions: (
        Mapping[str, Sequence[str]] | None
    ) = None,
    named_function_formula_environment_information_functions: (
        Mapping[str, Sequence[str]] | None
    ) = None,
    *,
    inspect_formula_defined_xlm_registrations: bool = False,
    inspect_formula_defined_xlm_evaluations: bool = False,
    inspect_formula_defined_xlm_get_cell_calls: bool = False,
    inspect_formula_defined_xlm_environment_information_calls: bool = False,
) -> FormulaInspection:
    """Inspect static reference coverage while resolving known named ranges.

    A caller provides case-folded name and named-LAMBDA maps assembled from the
    workbook. Supported fully qualified table references are resolved from table
    metadata, context-bound row references require the formula origin, and
    3-D references require workbook tab order. Other non-A1 tokens are returned
    explicitly instead of being silently omitted from the graph. A ``None``
    named-function value records a known LAMBDA whose definition is not safe to
    expand, so its call remains a visible coverage gap.
    """
    (
        tokens,
        literal_spill_tokens,
        literal_implicit_intersection_tokens,
    ) = _tokenize_formula(formula)
    if tokens is None:
        return FormulaInspection(
            (),
            (),
            (),
            tokenization_failed=True,
            spill_reference_tokens=literal_spill_tokens,
            implicit_intersection_tokens=literal_implicit_intersection_tokens,
        )
    resolved_names = named_references or {}
    resolved_named_functions = named_function_references or {}
    resolved_named_custom_functions = named_custom_function_candidates or {}
    resolved_named_function_custom_functions = (
        named_function_custom_function_candidates or {}
    )
    resolved_named_formula_external_actions = (
        named_formula_external_action_functions or {}
    )
    resolved_named_function_formula_external_actions = (
        named_function_formula_external_action_functions or {}
    )
    resolved_named_worksheet_code_resource_registrations = (
        named_worksheet_code_resource_registration_functions or {}
    )
    resolved_named_function_worksheet_code_resource_registrations = (
        named_function_worksheet_code_resource_registration_functions or {}
    )
    resolved_named_formula_defined_xlm_registrations = (
        named_formula_defined_xlm_registration_functions or {}
    )
    resolved_named_function_formula_defined_xlm_registrations = (
        named_function_formula_defined_xlm_registration_functions or {}
    )
    resolved_named_formula_defined_xlm_evaluations = (
        named_formula_defined_xlm_evaluation_functions or {}
    )
    resolved_named_function_formula_defined_xlm_evaluations = (
        named_function_formula_defined_xlm_evaluation_functions or {}
    )
    resolved_named_formula_defined_xlm_get_cell_calls = (
        named_formula_defined_xlm_get_cell_functions or {}
    )
    resolved_named_function_formula_defined_xlm_get_cell_calls = (
        named_function_formula_defined_xlm_get_cell_functions or {}
    )
    resolved_named_formula_defined_xlm_environment_information_calls = (
        named_formula_defined_xlm_environment_information_functions or {}
    )
    resolved_named_function_formula_defined_xlm_environment_information_calls = (
        named_function_formula_defined_xlm_environment_information_functions or {}
    )
    resolved_named_formula_environment_information_calls = (
        named_formula_environment_information_functions or {}
    )
    resolved_named_function_formula_environment_information_calls = (
        named_function_formula_environment_information_functions or {}
    )
    resolved_tables = structured_tables or {}
    references: list[ParsedReference] = []
    unresolved_range_tokens: list[str] = []
    dynamic_reference_functions: list[str] = []
    external_action_functions: list[str] = []
    python_functions: list[str] = []
    office_custom_function_candidates: list[str] = []
    worksheet_code_resource_registration_functions: list[str] = []
    formula_defined_xlm_registration_functions: list[str] = []
    formula_defined_xlm_evaluation_functions: list[str] = []
    formula_defined_xlm_get_cell_functions: list[str] = []
    formula_defined_xlm_environment_information_functions: list[str] = []
    formula_environment_information_functions: list[str] = []
    formula_environment_information_implicit_cell_reference_count = 0
    formula_environment_information_implicit_sheets_reference_count = 0
    three_d_reference_tokens: list[str] = []
    spill_reference_tokens: list[str] = list(literal_spill_tokens)
    implicit_intersection_tokens: list[str] = list(literal_implicit_intersection_tokens)
    local_variable_indexes = _local_variable_token_indexes(tokens)
    implicit_intersection_replacements = _implicit_intersection_reference_replacements(
        tokens, origin
    )

    def extend_formula_environment_information_signals(
        signals: Sequence[str],
    ) -> None:
        nonlocal formula_environment_information_implicit_cell_reference_count
        nonlocal formula_environment_information_implicit_sheets_reference_count
        for signal in signals:
            if signal == _FORMULA_ENVIRONMENT_INFORMATION_IMPLICIT_CELL_REFERENCE_MARKER:
                formula_environment_information_implicit_cell_reference_count += 1
            elif signal == _FORMULA_ENVIRONMENT_INFORMATION_IMPLICIT_SHEETS_REFERENCE_MARKER:
                formula_environment_information_implicit_sheets_reference_count += 1
            else:
                formula_environment_information_functions.append(signal)

    for position, token in enumerate(tokens):
        if token.type == "OPERAND" and token.subtype == "RANGE":
            if position in local_variable_indexes:
                continue
            if token.value.strip().startswith("@"):
                implicit_intersection_tokens.append(token.value.strip())
            if selected_reference := implicit_intersection_replacements.get(position):
                references.append(selected_reference)
                continue
            reference = parse_reference_token(token.value)
            if reference is not None:
                references.append(reference)
                continue
            three_d_reference = resolve_3d_reference(token.value, sheet_order)
            if three_d_reference is not None:
                references.extend(three_d_reference)
                three_d_reference_tokens.append(token.value)
                continue
            named_key = reference_lookup_key(token.value)
            if named_key in resolved_names:
                references.extend(resolved_names[named_key])
                external_action_functions.extend(
                    resolved_named_formula_external_actions.get(named_key, ())
                )
                office_custom_function_candidates.extend(
                    resolved_named_custom_functions.get(named_key, ())
                )
                worksheet_code_resource_registration_functions.extend(
                    resolved_named_worksheet_code_resource_registrations.get(
                        named_key, ()
                    )
                )
                formula_defined_xlm_registration_functions.extend(
                    resolved_named_formula_defined_xlm_registrations.get(named_key, ())
                )
                formula_defined_xlm_evaluation_functions.extend(
                    resolved_named_formula_defined_xlm_evaluations.get(named_key, ())
                )
                formula_defined_xlm_get_cell_functions.extend(
                    resolved_named_formula_defined_xlm_get_cell_calls.get(named_key, ())
                )
                formula_defined_xlm_environment_information_functions.extend(
                    resolved_named_formula_defined_xlm_environment_information_calls.get(
                        named_key, ()
                    )
                )
                extend_formula_environment_information_signals(
                    resolved_named_formula_environment_information_calls.get(
                        named_key, ()
                    )
                )
                continue
            if named_external_actions := resolved_named_formula_external_actions.get(
                named_key
            ):
                external_action_functions.extend(named_external_actions)
            if named_custom_functions := resolved_named_custom_functions.get(named_key):
                office_custom_function_candidates.extend(named_custom_functions)
            if (
                named_worksheet_code_resource_registrations := (
                    resolved_named_worksheet_code_resource_registrations.get(named_key)
                )
            ):
                worksheet_code_resource_registration_functions.extend(
                    named_worksheet_code_resource_registrations
                )
            if named_formula_defined_xlm_registrations := (
                resolved_named_formula_defined_xlm_registrations.get(named_key)
            ):
                formula_defined_xlm_registration_functions.extend(
                    named_formula_defined_xlm_registrations
                )
            if named_formula_defined_xlm_evaluations := (
                resolved_named_formula_defined_xlm_evaluations.get(named_key)
            ):
                formula_defined_xlm_evaluation_functions.extend(
                    named_formula_defined_xlm_evaluations
                )
            if named_formula_defined_xlm_get_cell_calls := (
                resolved_named_formula_defined_xlm_get_cell_calls.get(named_key)
            ):
                formula_defined_xlm_get_cell_functions.extend(
                    named_formula_defined_xlm_get_cell_calls
                )
            if named_formula_defined_xlm_environment_information_calls := (
                resolved_named_formula_defined_xlm_environment_information_calls.get(
                    named_key
                )
            ):
                formula_defined_xlm_environment_information_functions.extend(
                    named_formula_defined_xlm_environment_information_calls
                )
            if named_formula_environment_information_calls := (
                resolved_named_formula_environment_information_calls.get(named_key)
            ):
                extend_formula_environment_information_signals(
                    named_formula_environment_information_calls
                )
            table_reference = resolve_structured_reference(
                token.value, resolved_tables, origin
            )
            if table_reference is not None:
                references.extend(table_reference)
                continue
            unresolved_range_tokens.append(token.value)
        elif token.type == "FUNC" and token.subtype == "OPEN":
            raw_function_name = token.value.rstrip("(").strip()
            if position not in local_variable_indexes:
                function_key = _function_lookup_key(token)
                if function_key in resolved_named_functions:
                    function_references = resolved_named_functions[function_key]
                    if function_references is None:
                        unresolved_range_tokens.append(token.value.rstrip("(").strip())
                    else:
                        references.extend(function_references)
                    external_action_functions.extend(
                        resolved_named_function_formula_external_actions.get(
                            function_key, ()
                        )
                    )
                    office_custom_function_candidates.extend(
                        resolved_named_function_custom_functions.get(function_key, ())
                    )
                    worksheet_code_resource_registration_functions.extend(
                        resolved_named_function_worksheet_code_resource_registrations.get(
                            function_key, ()
                        )
                    )
                    formula_defined_xlm_registration_functions.extend(
                        resolved_named_function_formula_defined_xlm_registrations.get(
                            function_key, ()
                        )
                    )
                    formula_defined_xlm_evaluation_functions.extend(
                        resolved_named_function_formula_defined_xlm_evaluations.get(
                            function_key, ()
                        )
                    )
                    formula_defined_xlm_get_cell_functions.extend(
                        resolved_named_function_formula_defined_xlm_get_cell_calls.get(
                            function_key, ()
                        )
                    )
                    formula_defined_xlm_environment_information_functions.extend(
                        resolved_named_function_formula_defined_xlm_environment_information_calls.get(
                            function_key, ()
                        )
                    )
                    extend_formula_environment_information_signals(
                        resolved_named_function_formula_environment_information_calls.get(
                            function_key, ()
                        )
                    )
                if (
                    function_key not in resolved_names
                    and function_key not in resolved_named_functions
                    and (
                        custom_function_candidate := _office_custom_function_candidate(
                            token
                        )
                    )
                    is not None
                ):
                    office_custom_function_candidates.append(custom_function_candidate)
                if (
                    function_key not in resolved_names
                    and function_key not in resolved_named_functions
                    and (
                        worksheet_code_resource_registration_function := (
                            _worksheet_code_resource_registration_function(token)
                        )
                    )
                    is not None
                ):
                    worksheet_code_resource_registration_functions.append(
                        worksheet_code_resource_registration_function
                    )
                if (
                    inspect_formula_defined_xlm_registrations
                    and function_key not in resolved_names
                    and function_key not in resolved_named_functions
                    and (
                        formula_defined_xlm_registration_function := (
                            _formula_defined_xlm_registration_function(token)
                        )
                    )
                    is not None
                ):
                    formula_defined_xlm_registration_functions.append(
                        formula_defined_xlm_registration_function
                    )
                if (
                    inspect_formula_defined_xlm_evaluations
                    and function_key not in resolved_names
                    and function_key not in resolved_named_functions
                    and (
                        formula_defined_xlm_evaluation_function := (
                            _formula_defined_xlm_evaluation_function(token)
                        )
                    )
                    is not None
                ):
                    formula_defined_xlm_evaluation_functions.append(
                        formula_defined_xlm_evaluation_function
                    )
                if (
                    inspect_formula_defined_xlm_get_cell_calls
                    and function_key not in resolved_names
                    and function_key not in resolved_named_functions
                    and (
                        formula_defined_xlm_get_cell_function := (
                            _formula_defined_xlm_get_cell_function(token)
                        )
                    )
                    is not None
                ):
                    formula_defined_xlm_get_cell_functions.append(
                        formula_defined_xlm_get_cell_function
                    )
                if (
                    inspect_formula_defined_xlm_environment_information_calls
                    and function_key not in resolved_names
                    and function_key not in resolved_named_functions
                    and (
                        formula_defined_xlm_environment_information_function := (
                            _formula_defined_xlm_environment_information_function(
                                token
                            )
                        )
                    )
                    is not None
                ):
                    formula_defined_xlm_environment_information_functions.append(
                        formula_defined_xlm_environment_information_function
                    )
                if (
                    function_key not in resolved_names
                    and function_key not in resolved_named_functions
                    and (
                        formula_environment_information_function := (
                            _formula_environment_information_function(tokens, position)
                        )
                    )
                    is not None
                ):
                    (
                        function_name,
                        has_implicit_cell_reference,
                        has_implicit_sheets_reference,
                    ) = (
                        formula_environment_information_function
                    )
                    formula_environment_information_functions.append(function_name)
                    if has_implicit_cell_reference:
                        formula_environment_information_implicit_cell_reference_count += 1
                    if has_implicit_sheets_reference:
                        formula_environment_information_implicit_sheets_reference_count += 1
            function_name = _function_name(token)
            if function_name in _DYNAMIC_REFERENCE_FUNCTIONS:
                dynamic_reference_functions.append(function_name)
            if (
                position not in local_variable_indexes
                and function_key not in resolved_names
                and function_key not in resolved_named_functions
                and function_name in _EXTERNAL_ACTION_FUNCTIONS
            ):
                external_action_functions.append(function_name)
            if (
                position not in local_variable_indexes
                and function_name in _PYTHON_FUNCTIONS
            ):
                python_functions.append(function_name)
            if function_name == _SPILL_REFERENCE_FUNCTION:
                spill_reference_tokens.append(raw_function_name)
            if raw_function_name.startswith("@"):
                implicit_intersection_tokens.append(raw_function_name)
            elif function_name == _IMPLICIT_INTERSECTION_FUNCTION:
                implicit_intersection_tokens.append(raw_function_name)
    return FormulaInspection(
        references=tuple(references),
        unresolved_range_tokens=tuple(dict.fromkeys(unresolved_range_tokens)),
        dynamic_reference_functions=tuple(dict.fromkeys(dynamic_reference_functions)),
        external_action_functions=tuple(external_action_functions),
        python_functions=tuple(python_functions),
        office_custom_function_candidates=tuple(office_custom_function_candidates),
        worksheet_code_resource_registration_functions=tuple(
            worksheet_code_resource_registration_functions
        ),
        formula_defined_xlm_registration_functions=tuple(
            formula_defined_xlm_registration_functions
        ),
        formula_defined_xlm_evaluation_functions=tuple(
            formula_defined_xlm_evaluation_functions
        ),
        formula_defined_xlm_get_cell_functions=tuple(
            formula_defined_xlm_get_cell_functions
        ),
        formula_defined_xlm_environment_information_functions=tuple(
            formula_defined_xlm_environment_information_functions
        ),
        formula_environment_information_functions=tuple(
            formula_environment_information_functions
        ),
        formula_environment_information_implicit_cell_reference_count=(
            formula_environment_information_implicit_cell_reference_count
        ),
        formula_environment_information_implicit_sheets_reference_count=(
            formula_environment_information_implicit_sheets_reference_count
        ),
        three_d_reference_tokens=tuple(dict.fromkeys(three_d_reference_tokens)),
        spill_reference_tokens=tuple(dict.fromkeys(spill_reference_tokens)),
        implicit_intersection_tokens=tuple(dict.fromkeys(implicit_intersection_tokens)),
    )


def extract_references(
    formula: str,
    named_references: Mapping[str, Sequence[ParsedReference]] | None = None,
    structured_tables: Mapping[str, StructuredTable] | None = None,
    origin: tuple[str, str] | None = None,
    sheet_order: Sequence[str] | None = None,
) -> list[ParsedReference]:
    """Return A1-style, supplied named-range, and static table references."""
    return list(
        inspect_formula(
            formula, named_references, structured_tables, origin, sheet_order
        ).references
    )


def has_broken_reference(formula: str) -> bool:
    return "#REF!" in formula.upper()
