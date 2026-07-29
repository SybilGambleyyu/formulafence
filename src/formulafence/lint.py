"""Conservative, single-workbook formula-pattern linting."""

from __future__ import annotations

from collections import defaultdict
from dataclasses import dataclass, field

from openpyxl.utils.cell import coordinate_to_tuple, get_column_letter

from formulafence.formulas import MAX_EXCEL_COLUMN, MAX_EXCEL_ROW
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
) -> FormulaLintReport:
    """Find high-confidence interruptions inside local copied formula blocks.

    A target is reported only when its immediate predecessor and successor have
    the same relative formula fingerprint *and* a third contiguous peer repeats
    that fingerprint. This intentionally misses short or ambiguous sequences
    rather than treating a general spreadsheet smell as an error. It never
    evaluates formulas, and rejects incomplete array metadata before claiming
    ordinary-cell coverage.
    """
    if max_formula_pattern_findings < 1:
        raise FormulaFenceError("max_formula_pattern_findings must be at least 1.")
    if not snapshot.array_formula_metadata_complete:
        raise FormulaFenceError(
            "Formula-pattern lint requires complete array-formula metadata."
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
                        "Formula-pattern lint exceeds "
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
    return FormulaLintReport(workbook=snapshot, findings=findings)
