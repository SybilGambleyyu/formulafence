"""Semantic workbook diffing, downstream impact, and intrinsic-risk detection."""

from __future__ import annotations

from collections import deque
from collections.abc import Iterable

from openpyxl.utils.cell import coordinate_to_tuple, get_column_letter

from formulafence.models import (
    CellKey,
    CellSnapshot,
    Change,
    DiffReport,
    Finding,
    WorkbookSnapshot,
)

_IMPACT_SAMPLE_SIZE = 20
_IMPACT_NODE_LIMIT = 100_000


def _cell_change_kind(
    before: CellSnapshot | None, after: CellSnapshot | None
) -> tuple[str, str] | None:
    if before is None and after is None:  # pragma: no cover - impossible by caller contract
        return None
    if before is None:
        return ("formula_added", "low") if after and after.is_formula else ("value_added", "low")
    if after is None:
        return ("formula_removed", "high") if before.is_formula else ("value_removed", "low")
    if before.is_formula and not after.is_formula:
        return "formula_to_value", "high"
    if not before.is_formula and after.is_formula:
        return "value_to_formula", "medium"
    if before.is_formula and after.is_formula and before.formula != after.formula:
        return "formula_changed", "medium"
    if before.cell_type != after.cell_type or before.value != after.value:
        return "value_changed", "low"
    return None


def downstream_impact(
    location: CellKey,
    *snapshots: WorkbookSnapshot,
    node_limit: int = _IMPACT_NODE_LIMIT,
) -> tuple[set[CellKey], bool]:
    """Find formula cells affected through explicit dependency paths.

    Range references are checked lazily instead of expanded, so a formula such as
    `SUM(A:A)` does not turn into a million in-memory graph edges.
    """
    queue: deque[CellKey] = deque([location])
    visited: set[CellKey] = {location}
    impacted: set[CellKey] = set()
    truncated = False
    while queue:
        current = queue.popleft()
        dependents: set[CellKey] = set()
        for snapshot in snapshots:
            dependents.update(snapshot.direct_dependents(current))
        for dependent in dependents:
            if dependent in visited:
                continue
            visited.add(dependent)
            impacted.add(dependent)
            if len(visited) >= node_limit:
                truncated = True
                queue.clear()
                break
            queue.append(dependent)
    return impacted, truncated


def _location_sort_key(location: CellKey | None) -> tuple[str, int, int]:
    if location is None:
        return ("", 0, 0)
    sheet, coordinate = location
    row, column = coordinate_to_tuple(coordinate)
    return sheet.casefold(), row, column


def _adjacent(location: CellKey, row_delta: int, column_delta: int) -> CellKey | None:
    sheet, coordinate = location
    row, column = coordinate_to_tuple(coordinate)
    target_row = row + row_delta
    target_column = column + column_delta
    if target_row < 1 or target_column < 1:
        return None
    return sheet, f"{get_column_letter(target_column)}{target_row}"


def _formula_pattern_findings(
    snapshot: WorkbookSnapshot, changed_locations: set[CellKey]
) -> list[Finding]:
    """Detect a changed formula that breaks a stable immediate peer pattern."""
    findings: list[Finding] = []
    for location in sorted(changed_locations, key=_location_sort_key):
        cell = snapshot.cells.get(location)
        if cell is None or not cell.is_formula or cell.formula_fingerprint is None:
            continue
        for orientation, before_delta, after_delta in (
            ("column", (-1, 0), (1, 0)),
            ("row", (0, -1), (0, 1)),
        ):
            previous_location = _adjacent(location, *before_delta)
            next_location = _adjacent(location, *after_delta)
            if previous_location is None or next_location is None:
                continue
            previous = snapshot.cells.get(previous_location)
            following = snapshot.cells.get(next_location)
            if (
                previous is not None
                and following is not None
                and previous.is_formula
                and following.is_formula
                and previous.formula_fingerprint == following.formula_fingerprint
                and previous.formula_fingerprint != cell.formula_fingerprint
            ):
                findings.append(
                    Finding(
                        rule_id="FF006",
                        severity="high",
                        location=location,
                        message=(
                            f"Formula breaks a stable {orientation} pattern between "
                            f"{previous_location[1]} and {next_location[1]}."
                        ),
                        details={
                            "peer_formula_fingerprint": previous.formula_fingerprint,
                            "candidate_formula_fingerprint": cell.formula_fingerprint,
                        },
                    )
                )
                break
    return findings


def _workbook_control_changes(
    before: WorkbookSnapshot, after: WorkbookSnapshot
) -> tuple[list[Change], list[Finding]]:
    changes: list[Change] = []
    findings: list[Finding] = []

    for title in sorted(set(before.sheets) | set(after.sheets), key=str.casefold):
        old_sheet = before.sheets.get(title)
        new_sheet = after.sheets.get(title)
        if old_sheet is None:
            changes.append(
                Change(
                    "sheet_added",
                    None,
                    "medium",
                    details={"sheet": title, "state": new_sheet.state},
                )
            )
        elif new_sheet is None:
            changes.append(
                Change(
                    "sheet_removed",
                    None,
                    "high",
                    details={"sheet": title, "state": old_sheet.state},
                )
            )
        elif old_sheet.state != new_sheet.state:
            changes.append(
                Change(
                    "sheet_visibility_changed",
                    None,
                    "high",
                    details={"sheet": title, "before": old_sheet.state, "after": new_sheet.state},
                )
            )
            findings.append(
                Finding(
                    "FF007",
                    "high",
                    (
                        "Sheet visibility changed from "
                        f"{old_sheet.state} to {new_sheet.state}: {title}."
                    ),
                    details={"sheet": title, "before": old_sheet.state, "after": new_sheet.state},
                )
            )

    for name in sorted(set(before.defined_names) | set(after.defined_names), key=str.casefold):
        old_value = before.defined_names.get(name)
        new_value = after.defined_names.get(name)
        if old_value == new_value:
            continue
        changes.append(
            Change(
                "defined_name_changed",
                None,
                "medium",
                details={"name": name, "before": old_value, "after": new_value},
            )
        )
        findings.append(
            Finding(
                "FF008",
                "medium",
                f"Defined name changed: {name}.",
                details={"before": old_value, "after": new_value},
            )
        )

    if before.calculation_settings != after.calculation_settings:
        changes.append(
            Change(
                "calculation_settings_changed",
                None,
                "high",
                details={
                    "before": before.calculation_settings,
                    "after": after.calculation_settings,
                },
            )
        )
        findings.append(
            Finding(
                "FF009",
                "high",
                "Workbook calculation settings changed.",
                details={
                    "before": before.calculation_settings,
                    "after": after.calculation_settings,
                },
            )
        )

    if before.macro_hash != after.macro_hash:
        changes.append(
            Change(
                "macro_payload_changed",
                None,
                "critical",
                details={"before": before.macro_hash, "after": after.macro_hash},
            )
        )
        findings.append(
            Finding(
                "FF005",
                "critical",
                "VBA macro payload was added, removed, or changed.",
            )
        )
    return changes, findings


def compare_snapshots(before: WorkbookSnapshot, after: WorkbookSnapshot) -> DiffReport:
    """Compare workbook semantics and attach local dependency impact to each edit."""
    changes: list[Change] = []
    findings: list[Finding] = []
    formula_changed_locations: set[CellKey] = set()

    all_locations = sorted(set(before.cells) | set(after.cells), key=_location_sort_key)
    for location in all_locations:
        old_cell = before.cells.get(location)
        new_cell = after.cells.get(location)
        classified = _cell_change_kind(old_cell, new_cell)
        if classified is None:
            continue
        kind, severity = classified
        impact, impact_truncated = downstream_impact(location, before, after)
        details: dict[str, object] = {}
        if impact_truncated:
            details["impact_truncated_at"] = _IMPACT_NODE_LIMIT
        changes.append(
            Change(
                kind=kind,
                location=location,
                severity=severity,
                before=old_cell,
                after=new_cell,
                impact_count=len(impact),
                impacted_cells=tuple(sorted(impact, key=_location_sort_key)[:_IMPACT_SAMPLE_SIZE]),
                details=details,
            )
        )
        if new_cell is not None and new_cell.is_formula:
            formula_changed_locations.add(location)
        if old_cell is not None and old_cell.is_formula:
            formula_changed_locations.add(location)
        if kind == "formula_to_value":
            findings.append(
                Finding(
                    "FF001",
                    "high",
                    "Formula was replaced with a value.",
                    location,
                    details={"impact_count": len(impact)},
                )
            )
        elif kind == "formula_removed":
            findings.append(
                Finding(
                    "FF002",
                    "high",
                    "Formula cell was removed or blanked.",
                    location,
                    details={"impact_count": len(impact)},
                )
            )

    newly_broken = after.broken_references - before.broken_references
    for location in sorted(newly_broken, key=_location_sort_key):
        findings.append(
            Finding("FF003", "critical", "Formula contains a broken #REF! reference.", location)
        )
    new_external = after.external_references - before.external_references
    for location in sorted(new_external, key=_location_sort_key):
        findings.append(
            Finding(
                "FF004",
                "high",
                "Formula introduces an explicit external-workbook reference.",
                location,
            )
        )

    control_changes, control_findings = _workbook_control_changes(before, after)
    changes.extend(control_changes)
    findings.extend(control_findings)
    findings.extend(_formula_pattern_findings(after, formula_changed_locations))

    changes.sort(key=lambda change: (_location_sort_key(change.location), change.kind))
    findings.sort(key=lambda finding: (_location_sort_key(finding.location), finding.rule_id))
    return DiffReport(before=before, after=after, changes=changes, findings=findings)


def report_severities(report: DiffReport, extra_findings: Iterable[Finding] = ()) -> list[str]:
    """Return both risk-finding and semantic-change severities for CLI thresholds."""
    return [
        *(change.severity for change in report.changes),
        *(finding.severity for finding in report.findings),
        *(finding.severity for finding in extra_findings),
    ]
