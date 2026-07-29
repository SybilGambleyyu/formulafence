"""Conservative, single-workbook formula linting."""

from __future__ import annotations

import re
from collections import defaultdict
from dataclasses import dataclass, field
from numbers import Real

from openpyxl.utils.cell import coordinate_to_tuple, get_column_letter

from formulafence.formulas import (
    MAX_EXCEL_COLUMN,
    MAX_EXCEL_ROW,
    parse_reference_token,
)
from formulafence.models import (
    ArrayFormulaRange,
    CellKey,
    Finding,
    FormulaFenceError,
    FormulaLintReport,
    WorkbookSnapshot,
    display_location,
)

DEFAULT_MAX_FORMULA_PATTERN_FINDINGS = 10_000
DEFAULT_MAX_AGGREGATE_OMISSION_GAP_CELLS = 128

_MIN_AGGREGATE_OMISSION_GAP_CELLS = 2
_SIMPLE_NUMERIC_AGGREGATE = re.compile(
    r"^\s*=\s*(?P<function>SUM|AVERAGE|MIN|MAX|COUNT)\s*"
    r"\(\s*(?P<reference>[^()]+?)\s*\)\s*$",
    re.IGNORECASE,
)

_PATTERN_ORIENTATIONS = (
    ("row", 0, 1),
    ("column", 1, 0),
)


@dataclass(frozen=True)
class _FormulaPatternEvidence:
    """One copied-peer pattern supporting a lint finding."""

    orientation: str
    preceding: CellKey
    following: CellKey
    supporting: CellKey


@dataclass
class _FormulaPatternCandidate:
    """One target with one or more independent local pattern signals."""

    kind: str
    evidence: list[_FormulaPatternEvidence] = field(default_factory=list)


@dataclass(frozen=True)
class _AggregateOmissionCandidate:
    """One bounded, contiguous numeric run immediately outside an aggregate."""

    function: str
    orientation: str
    referenced_range: str
    omitted_range: str
    omitted_cell_count: int


def _location_sort_key(location: CellKey) -> tuple[str, int, int]:
    sheet, coordinate = location
    row, column = coordinate_to_tuple(coordinate)
    return sheet.casefold(), row, column


def _offset_location(location: CellKey, row_delta: int, column_delta: int) -> CellKey | None:
    """Return one bounded relative cell coordinate without leaving Excel's grid."""
    sheet, coordinate = location
    row, column = coordinate_to_tuple(coordinate)
    target_row = row + row_delta
    target_column = column + column_delta
    if not (1 <= target_row <= MAX_EXCEL_ROW and 1 <= target_column <= MAX_EXCEL_COLUMN):
        return None
    return sheet, f"{get_column_letter(target_column)}{target_row}"


def _array_ranges_by_sheet(snapshot: WorkbookSnapshot) -> dict[str, tuple[ArrayFormulaRange, ...]]:
    """Group compact array ranges so ordinary pattern evidence never crosses one."""
    grouped: dict[str, list[ArrayFormulaRange]] = defaultdict(list)
    for array_range in (
        *snapshot.legacy_array_formula_ranges,
        *snapshot.dynamic_array_formula_ranges,
    ):
        grouped[array_range.sheet].append(array_range)
    return {
        sheet: tuple(ranges)
        for sheet, ranges in grouped.items()
    }


def _is_array_member(
    snapshot: WorkbookSnapshot,
    location: CellKey,
    array_ranges_by_sheet: dict[str, tuple[ArrayFormulaRange, ...]],
) -> bool:
    """Return whether a cell is an array anchor or observed result member."""
    if (
        location in snapshot.dynamic_array_formula_cells
        or location in snapshot.unclassified_array_formula_cells
    ):
        return True
    return any(
        array_range.contains(location)
        for array_range in array_ranges_by_sheet.get(location[0], ())
    )


def _eligible_formula(
    snapshot: WorkbookSnapshot,
    location: CellKey,
    array_ranges_by_sheet: dict[str, tuple[ArrayFormulaRange, ...]],
) -> bool:
    """Return whether one ordinary formula is safe local-copy evidence."""
    cell = snapshot.cells.get(location)
    return bool(
        cell is not None
        and cell.is_formula
        and cell.value_type == "formula"
        and cell.formula_fingerprint is not None
        and location not in snapshot.tokenization_failure_cells
        and not _is_array_member(snapshot, location, array_ranges_by_sheet)
    )


def _range_intersects_array_territory(
    snapshot: WorkbookSnapshot,
    *,
    sheet: str,
    min_column: int,
    min_row: int,
    max_column: int,
    max_row: int,
    array_ranges_by_sheet: dict[str, tuple[ArrayFormulaRange, ...]],
) -> bool:
    """Return whether a compact local range overlaps any declared array territory."""
    if any(
        min_column <= array_range.max_column
        and array_range.min_column <= max_column
        and min_row <= array_range.max_row
        and array_range.min_row <= max_row
        for array_range in array_ranges_by_sheet.get(sheet, ())
    ):
        return True
    for locations in (
        snapshot.dynamic_array_formula_cells,
        snapshot.unclassified_array_formula_cells,
    ):
        for location in locations:
            if location[0] != sheet:
                continue
            row, column = coordinate_to_tuple(location[1])
            if min_column <= column <= max_column and min_row <= row <= max_row:
                return True
    return False


def _is_numeric_literal(cell: object) -> bool:
    """Return whether one stored ordinary cell is a numeric aggregate input."""
    return bool(
        getattr(cell, "cell_type", None) == "value"
        and isinstance(getattr(cell, "value", None), Real)
        and not isinstance(getattr(cell, "value", None), bool)
    )


def _aggregate_omission_candidate(
    snapshot: WorkbookSnapshot,
    location: CellKey,
    array_ranges_by_sheet: dict[str, tuple[ArrayFormulaRange, ...]],
    *,
    max_gap_cells: int,
) -> _AggregateOmissionCandidate | None:
    """Find one deliberately narrow numeric run omitted by a local aggregate.

    Excel's own error checker warns when an aggregate ends before immediately
    adjacent values.  This portable variant accepts only a pure single-range
    ``SUM``, ``AVERAGE``, ``MIN``, ``MAX``, or ``COUNT`` over a one-dimensional
    local range.  The candidate must sit on the same row or column after that
    range, with a short, fully populated run of numeric literals in between.
    It does not infer an intended formula or evaluate values.
    """
    cell = snapshot.cells[location]
    if not isinstance(cell.formula, str):  # narrowed by _eligible_formula
        return None
    match = _SIMPLE_NUMERIC_AGGREGATE.fullmatch(cell.formula)
    if match is None:
        return None
    reference = parse_reference_token(match.group("reference"))
    if (
        reference is None
        or reference.is_external
        or reference.min_column is None
        or reference.min_row is None
        or reference.max_column is None
        or reference.max_row is None
        or not reference.is_range
        or (
            reference.sheet is not None
            and reference.sheet.casefold() != location[0].casefold()
        )
    ):
        return None
    if (reference.min_column == reference.max_column) == (
        reference.min_row == reference.max_row
    ):
        return None

    sheet, coordinate = location
    formula_row, formula_column = coordinate_to_tuple(coordinate)
    if (
        reference.min_column == reference.max_column
        and formula_column == reference.min_column
        and formula_row > reference.max_row + 1
    ):
        orientation = "column"
        omitted_min_column = omitted_max_column = reference.min_column
        omitted_min_row = reference.max_row + 1
        omitted_max_row = formula_row - 1
    elif (
        reference.min_row == reference.max_row
        and formula_row == reference.min_row
        and formula_column > reference.max_column + 1
    ):
        orientation = "row"
        omitted_min_column = reference.max_column + 1
        omitted_max_column = formula_column - 1
        omitted_min_row = omitted_max_row = reference.min_row
    else:
        return None

    omitted_cell_count = (
        (omitted_max_column - omitted_min_column + 1)
        * (omitted_max_row - omitted_min_row + 1)
    )
    if not _MIN_AGGREGATE_OMISSION_GAP_CELLS <= omitted_cell_count <= max_gap_cells:
        return None
    if _range_intersects_array_territory(
        snapshot,
        sheet=sheet,
        min_column=reference.min_column,
        min_row=reference.min_row,
        max_column=reference.max_column,
        max_row=reference.max_row,
        array_ranges_by_sheet=array_ranges_by_sheet,
    ) or _range_intersects_array_territory(
        snapshot,
        sheet=sheet,
        min_column=omitted_min_column,
        min_row=omitted_min_row,
        max_column=omitted_max_column,
        max_row=omitted_max_row,
        array_ranges_by_sheet=array_ranges_by_sheet,
    ):
        return None

    if orientation == "column":
        omitted_locations = (
            (sheet, f"{get_column_letter(omitted_min_column)}{row}")
            for row in range(omitted_min_row, omitted_max_row + 1)
        )
    else:
        omitted_locations = (
            (sheet, f"{get_column_letter(column)}{omitted_min_row}")
            for column in range(omitted_min_column, omitted_max_column + 1)
        )
    if not all(
        (omitted := snapshot.cells.get(omitted_location)) is not None
        and _is_numeric_literal(omitted)
        for omitted_location in omitted_locations
    ):
        return None

    referenced_start = f"{get_column_letter(reference.min_column)}{reference.min_row}"
    referenced_end = f"{get_column_letter(reference.max_column)}{reference.max_row}"
    omitted_start = f"{get_column_letter(omitted_min_column)}{omitted_min_row}"
    omitted_end = f"{get_column_letter(omitted_max_column)}{omitted_max_row}"
    return _AggregateOmissionCandidate(
        function=match.group("function").upper(),
        orientation=orientation,
        referenced_range=referenced_start
        if referenced_start == referenced_end
        else f"{referenced_start}:{referenced_end}",
        omitted_range=(
            omitted_start
            if omitted_start == omitted_end
            else f"{omitted_start}:{omitted_end}"
        ),
        omitted_cell_count=omitted_cell_count,
    )


def _direct_unlocked_formula_locations(
    snapshot: WorkbookSnapshot,
    array_ranges_by_sheet: dict[str, tuple[ArrayFormulaRange, ...]],
) -> tuple[CellKey, ...]:
    """Return ordinary formulas explicitly unlocked on an active worksheet.

    Excel's error checker warns about unlocked formula cells because they can be
    overwritten.  FormulaFence accepts only direct cell protection assignments
    on an actively protected worksheet.  It deliberately does not try to infer
    an effective style from row, column, default, or allowed-edit-range state:
    those broader cases need a complete precedence model before they are safe
    to present as an actionable review finding.
    """
    protected_sheets = {
        protection.sheet
        for protection in snapshot.sheet_protections
        if protection.enabled and protection.sheet_type == "worksheet"
    }
    locations: list[CellKey] = []
    for assignment in sorted(
        snapshot.cell_protection_assignments,
        key=lambda item: item.sort_key(),
    ):
        if (
            assignment.scope != "cell"
            or assignment.locked
            or assignment.sheet not in protected_sheets
        ):
            continue
        location = (assignment.sheet, assignment.target)
        if _eligible_formula(snapshot, location, array_ranges_by_sheet):
            locations.append(location)
    return tuple(locations)


def _has_incomplete_manual_calculation(snapshot: WorkbookSnapshot) -> bool:
    """Return whether stored formula results may be stale by explicit metadata.

    ``calcCompleted=false`` says that the workbook was not recalculated before
    it was saved.  FormulaFence reports that state only when the workbook also
    explicitly requests manual calculation, where Excel requires a user to
    request recalculation.  It is a configuration risk, not a claim that any
    particular formula result is mathematically wrong.
    """
    return bool(
        any(cell.is_formula for cell in snapshot.cells.values())
        and snapshot.calculation_settings.get("calcMode") == "manual"
        and snapshot.calculation_settings.get("calcCompleted") is False
    )


def _direct_self_reference_locations(
    snapshot: WorkbookSnapshot,
    array_ranges_by_sheet: dict[str, tuple[ArrayFormulaRange, ...]],
) -> tuple[CellKey, ...]:
    """Return exact static self references when OOXML iteration is disabled.

    The workbook ``iterate`` calculation property defaults to false.  A direct
    scalar dependency back to its own ordinary formula cell is therefore a
    circular reference Excel will not attempt to calculate by default.  This
    deliberately does not try to infer indirect cycles or interpret a range,
    dynamic reference, spill, or explicit-intersection expression: each needs
    more evaluation semantics than this lint promises.
    """
    if snapshot.calculation_settings.get("iterate") is True:
        return ()

    locations: list[CellKey] = []
    for location in sorted(snapshot.cells, key=_location_sort_key):
        if not _eligible_formula(snapshot, location, array_ranges_by_sheet):
            continue
        if (
            location in snapshot.dynamic_reference_functions
            or location in snapshot.spill_reference_tokens
            or location in snapshot.implicit_intersection_tokens
        ):
            continue
        if location in snapshot.reverse_dependencies.get(location, ()):
            locations.append(location)
    return tuple(locations)


def _explicit_broken_reference_locations(snapshot: WorkbookSnapshot) -> tuple[CellKey, ...]:
    """Return formula locations whose tokenizer exposed a ``#REF!`` operand.

    Workbook loading records this ledger only for actual error operands, not
    for a matching string literal or quoted worksheet name.  The lint keeps
    that exact lexical fact separate from a cached formula result and does not
    evaluate the formula.
    """
    return tuple(sorted(snapshot.broken_references, key=_location_sort_key))


def _candidate_message(kind: str) -> tuple[str, str, str]:
    """Return the rule, severity, and reviewer-facing message for a pattern kind."""
    if kind == "blank_gap":
        return (
            "FF082",
            "high",
            "A blank cell interrupts a stable copied-formula pattern.",
        )
    if kind == "error_gap":
        return (
            "FF082",
            "high",
            "A stored error cell interrupts a stable copied-formula pattern.",
        )
    if kind == "text_gap":
        return (
            "FF082",
            "low",
            "A text value interrupts a stable copied-formula pattern.",
        )
    if kind == "non_formula_gap":
        return (
            "FF082",
            "medium",
            "A non-formula value interrupts a stable copied-formula pattern.",
        )
    return (
        "FF083",
        "medium",
        "Formula differs from a stable copied-formula pattern.",
    )


def lint_snapshot(
    snapshot: WorkbookSnapshot,
    *,
    max_formula_pattern_findings: int = DEFAULT_MAX_FORMULA_PATTERN_FINDINGS,
    max_aggregate_omission_gap_cells: int = DEFAULT_MAX_AGGREGATE_OMISSION_GAP_CELLS,
) -> FormulaLintReport:
    """Find conservative formula patterns, protection, and calculation risks.

    A target is reported only when its immediate predecessor and successor have
    the same relative formula fingerprint *and* a third contiguous peer repeats
    that fingerprint. This intentionally misses short or ambiguous sequences
    rather than treating a general spreadsheet smell as an error. Separately,
    a pure local numeric aggregate is reported only when it stops before a
    short, contiguous numeric run on the same row or column. It also reports
    an explicitly unlocked formula on a protected sheet, an explicit incomplete
    manual-calculation state for a formula workbook, and a direct static
    self-reference while iteration is disabled, and an explicit broken
    reference operand. It never evaluates formulas, and rejects incomplete
    array metadata before claiming ordinary-cell coverage.
    """
    if max_formula_pattern_findings < 1:
        raise FormulaFenceError("max_formula_pattern_findings must be at least 1.")
    if max_aggregate_omission_gap_cells < _MIN_AGGREGATE_OMISSION_GAP_CELLS:
        raise FormulaFenceError(
            "max_aggregate_omission_gap_cells must be at least "
            f"{_MIN_AGGREGATE_OMISSION_GAP_CELLS}."
        )
    if not snapshot.array_formula_metadata_complete:
        raise FormulaFenceError(
            "Formula lint requires complete array-formula metadata."
        )

    array_ranges_by_sheet = _array_ranges_by_sheet(snapshot)
    candidates: dict[CellKey, _FormulaPatternCandidate] = {}

    for preceding_location in sorted(snapshot.cells, key=_location_sort_key):
        if not _eligible_formula(snapshot, preceding_location, array_ranges_by_sheet):
            continue
        preceding = snapshot.cells[preceding_location]
        fingerprint = preceding.formula_fingerprint
        assert fingerprint is not None  # narrowed by _eligible_formula

        for orientation, row_delta, column_delta in _PATTERN_ORIENTATIONS:
            target_location = _offset_location(
                preceding_location,
                row_delta,
                column_delta,
            )
            following_location = _offset_location(
                preceding_location,
                row_delta * 2,
                column_delta * 2,
            )
            if target_location is None or following_location is None:
                continue
            if not _eligible_formula(snapshot, following_location, array_ranges_by_sheet):
                continue
            following = snapshot.cells[following_location]
            if following.formula_fingerprint != fingerprint:
                continue
            if _is_array_member(snapshot, target_location, array_ranges_by_sheet):
                continue

            target = snapshot.cells.get(target_location)
            if target is None:
                kind = "blank_gap"
            elif target.is_formula:
                if not _eligible_formula(snapshot, target_location, array_ranges_by_sheet):
                    continue
                if target.formula_fingerprint == fingerprint:
                    continue
                kind = "formula_outlier"
            else:
                kind = (
                    "error_gap"
                    if target.cell_type == "error"
                    else "text_gap"
                    if isinstance(target.value, str)
                    else "non_formula_gap"
                )

            support_before = _offset_location(
                preceding_location,
                -row_delta,
                -column_delta,
            )
            support_after = _offset_location(
                following_location,
                row_delta,
                column_delta,
            )
            supporting_location = next(
                (
                    location
                    for location in (support_before, support_after)
                    if location is not None
                    and _eligible_formula(snapshot, location, array_ranges_by_sheet)
                    and snapshot.cells[location].formula_fingerprint == fingerprint
                ),
                None,
            )
            if supporting_location is None:
                continue

            candidate = candidates.get(target_location)
            if candidate is None:
                if len(candidates) >= max_formula_pattern_findings:
                    raise FormulaFenceError(
                        "Formula lint exceeds "
                        f"max_formula_pattern_findings={max_formula_pattern_findings}."
                    )
                candidate = _FormulaPatternCandidate(kind=kind)
                candidates[target_location] = candidate
            elif candidate.kind != kind:  # pragma: no cover - one target has one cell kind
                continue
            candidate.evidence.append(
                _FormulaPatternEvidence(
                    orientation=orientation,
                    preceding=preceding_location,
                    following=following_location,
                    supporting=supporting_location,
                )
            )

    findings: list[Finding] = []
    for location in sorted(candidates, key=_location_sort_key):
        candidate = candidates[location]
        rule_id, severity, message = _candidate_message(candidate.kind)
        findings.append(
            Finding(
                rule_id=rule_id,
                severity=severity,
                message=message,
                location=location,
                details={
                    "pattern_kind": candidate.kind,
                    "pattern_evidence": [
                        {
                            "orientation": evidence.orientation,
                            "preceding_formula": display_location(evidence.preceding),
                            "following_formula": display_location(evidence.following),
                            "supporting_formula": display_location(evidence.supporting),
                        }
                        for evidence in candidate.evidence
                    ],
                },
            )
        )

    for location in sorted(snapshot.cells, key=_location_sort_key):
        if not _eligible_formula(snapshot, location, array_ranges_by_sheet):
            continue
        candidate = _aggregate_omission_candidate(
            snapshot,
            location,
            array_ranges_by_sheet,
            max_gap_cells=max_aggregate_omission_gap_cells,
        )
        if candidate is None:
            continue
        if len(findings) >= max_formula_pattern_findings:
            raise FormulaFenceError(
                "Formula lint exceeds "
                f"max_formula_pattern_findings={max_formula_pattern_findings}."
            )
        findings.append(
            Finding(
                rule_id="FF084",
                severity="medium",
                message=(
                    "A simple numeric aggregate stops before contiguous adjacent "
                    "numeric cells."
                ),
                location=location,
                details={
                    "aggregate_function": candidate.function,
                    "orientation": candidate.orientation,
                    "referenced_range": display_location(
                        (location[0], candidate.referenced_range)
                    ),
                    "omitted_range": display_location((location[0], candidate.omitted_range)),
                    "omitted_cell_count": candidate.omitted_cell_count,
                },
            )
        )
    for location in _direct_unlocked_formula_locations(
        snapshot,
        array_ranges_by_sheet,
    ):
        if len(findings) >= max_formula_pattern_findings:
            raise FormulaFenceError(
                "Formula lint exceeds "
                f"max_formula_pattern_findings={max_formula_pattern_findings}."
            )
        findings.append(
            Finding(
                rule_id="FF085",
                severity="medium",
                message=(
                    "A formula cell is explicitly unlocked on a protected worksheet."
                ),
                location=location,
                details={"protection_scope": "direct_cell"},
            )
        )
    if _has_incomplete_manual_calculation(snapshot):
        if len(findings) >= max_formula_pattern_findings:
            raise FormulaFenceError(
                "Formula lint exceeds "
                f"max_formula_pattern_findings={max_formula_pattern_findings}."
            )
        findings.append(
            Finding(
                rule_id="FF086",
                severity="medium",
                message=(
                    "Workbook contains formulas and was saved with incomplete "
                    "manual calculation."
                ),
                details={
                    "calculation_mode": "manual",
                    "calculation_completed_before_save": False,
                },
            )
        )
    for location in _direct_self_reference_locations(
        snapshot,
        array_ranges_by_sheet,
    ):
        if len(findings) >= max_formula_pattern_findings:
            raise FormulaFenceError(
                "Formula lint exceeds "
                f"max_formula_pattern_findings={max_formula_pattern_findings}."
            )
        findings.append(
            Finding(
                rule_id="FF087",
                severity="high",
                message=(
                    "A formula directly references its own cell while calculation "
                    "iteration is disabled."
                ),
                location=location,
                details={
                    "calculation_iteration_enabled": False,
                    "reference_scope": "direct_static",
                },
            )
        )
    for location in _explicit_broken_reference_locations(snapshot):
        if len(findings) >= max_formula_pattern_findings:
            raise FormulaFenceError(
                "Formula lint exceeds "
                f"max_formula_pattern_findings={max_formula_pattern_findings}."
            )
        findings.append(
            Finding(
                rule_id="FF088",
                severity="critical",
                message="Formula contains an explicit broken #REF! reference.",
                location=location,
            )
        )
    findings.sort(
        key=lambda finding: (
            _location_sort_key(finding.location)
            if finding.location is not None
            else ("", 0, 0),
            finding.rule_id,
        )
    )
    return FormulaLintReport(workbook=snapshot, findings=findings)
