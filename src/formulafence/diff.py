"""Semantic workbook diffing, downstream impact, and intrinsic-risk detection."""

from __future__ import annotations

from collections import deque
from collections.abc import Iterable, Mapping
from dataclasses import dataclass

from openpyxl.utils.cell import coordinate_to_tuple, get_column_letter

from formulafence.formulas import resolve_3d_reference
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


@dataclass(frozen=True)
class ImpactAnalysis:
    """A deterministic, shortest-path view of explicit downstream dependencies."""

    impacted: frozenset[CellKey]
    paths: dict[CellKey, tuple[CellKey, ...]]
    truncated: bool


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


def analyze_downstream_impact(
    location: CellKey,
    *snapshots: WorkbookSnapshot,
    node_limit: int = _IMPACT_NODE_LIMIT,
) -> ImpactAnalysis:
    """Find formula cells affected through explicit dependency paths.

    Range references are checked lazily instead of expanded, so a formula such as
    `SUM(A:A)` does not turn into a million in-memory graph edges.
    """
    queue: deque[CellKey] = deque([location])
    visited: set[CellKey] = {location}
    impacted: set[CellKey] = set()
    parents: dict[CellKey, CellKey | None] = {location: None}
    truncated = False
    while queue:
        current = queue.popleft()
        dependents: set[CellKey] = set()
        for snapshot in snapshots:
            dependents.update(snapshot.direct_dependents(current))
        for dependent in sorted(dependents, key=_location_sort_key):
            if dependent in visited:
                continue
            visited.add(dependent)
            impacted.add(dependent)
            parents[dependent] = current
            if len(visited) >= node_limit:
                truncated = True
                queue.clear()
                break
            queue.append(dependent)
    paths: dict[CellKey, tuple[CellKey, ...]] = {}
    for target in impacted:
        path: list[CellKey] = []
        current: CellKey | None = target
        while current is not None:
            path.append(current)
            current = parents[current]
        paths[target] = tuple(reversed(path))
    return ImpactAnalysis(frozenset(impacted), paths, truncated)


def downstream_impact(
    location: CellKey,
    *snapshots: WorkbookSnapshot,
    node_limit: int = _IMPACT_NODE_LIMIT,
) -> tuple[set[CellKey], bool]:
    """Backward-compatible shortcut for callers that need counts only."""
    analysis = analyze_downstream_impact(location, *snapshots, node_limit=node_limit)
    return set(analysis.impacted), analysis.truncated


def _serialise_impact_paths(
    targets: Iterable[CellKey], paths: dict[CellKey, tuple[CellKey, ...]]
) -> list[dict[str, object]]:
    from formulafence.models import display_location

    return [
        {
            "target": display_location(target),
            "path": [display_location(step) for step in paths[target]],
        }
        for target in targets
    ]


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


def _new_coverage_items(
    before: Mapping[CellKey, tuple[str, ...]], after: Mapping[CellKey, tuple[str, ...]]
) -> list[tuple[CellKey, tuple[str, ...]]]:
    """Return per-cell static-analysis coverage items newly present in a candidate."""
    additions: list[tuple[CellKey, tuple[str, ...]]] = []
    for location in sorted(after, key=_location_sort_key):
        previous = {item.casefold() for item in before.get(location, ())}
        new_items = tuple(item for item in after[location] if item.casefold() not in previous)
        if new_items:
            additions.append((location, new_items))
    return additions


def _three_d_reference_scope_changes(
    before: WorkbookSnapshot, after: WorkbookSnapshot
) -> tuple[list[Change], list[Finding]]:
    """Find unchanged 3-D formulas whose workbook tab span now resolves differently."""
    changes: list[Change] = []
    findings: list[Finding] = []
    locations = sorted(
        set(before.three_d_reference_tokens) | set(after.three_d_reference_tokens),
        key=_location_sort_key,
    )
    for location in locations:
        before_cell = before.cells.get(location)
        after_cell = after.cells.get(location)
        if before_cell is None or after_cell is None or before_cell.formula != after_cell.formula:
            continue
        tokens = sorted(
            set(before.three_d_reference_tokens.get(location, ()))
            | set(after.three_d_reference_tokens.get(location, ())),
            key=str.casefold,
        )
        changed_references: list[dict[str, object]] = []
        for token in tokens:
            before_references = resolve_3d_reference(token, before.sheet_order) or ()
            after_references = resolve_3d_reference(token, after.sheet_order) or ()
            before_sheets = [reference.sheet for reference in before_references]
            after_sheets = [reference.sheet for reference in after_references]
            if before_sheets != after_sheets:
                changed_references.append(
                    {
                        "token": token,
                        "before_sheets": before_sheets,
                        "after_sheets": after_sheets,
                    }
                )
        if not changed_references:
            continue
        details = {"references": changed_references}
        changes.append(
            Change(
                "three_d_reference_scope_changed",
                location,
                "high",
                details=details,
            )
        )
        findings.append(
            Finding(
                "FF014",
                "high",
                "Worksheet changes altered the scope of a static 3-D reference.",
                location,
                details=details,
            )
        )
    return changes, findings


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

    three_d_changes, three_d_findings = _three_d_reference_scope_changes(before, after)
    changes.extend(three_d_changes)
    findings.extend(three_d_findings)

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

    for name in sorted(set(before.tables) | set(after.tables), key=str.casefold):
        old_table = before.tables.get(name)
        new_table = after.tables.get(name)
        if old_table == new_table:
            continue
        details = {
            "name": name,
            "before": old_table.to_dict() if old_table else None,
            "after": new_table.to_dict() if new_table else None,
        }
        changes.append(
            Change(
                "table_definition_changed",
                None,
                "high",
                details=details,
            )
        )
        findings.append(
            Finding(
                "FF013",
                "high",
                f"Excel table definition changed: {name}.",
                details=details,
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
    new_parser_warnings = sorted(set(after.parser_warnings) - set(before.parser_warnings))
    if new_parser_warnings:
        changes.append(
            Change(
                "parser_coverage_warning_added",
                None,
                "medium",
                details={"warnings": new_parser_warnings},
            )
        )
        findings.append(
            Finding(
                "FF010",
                "medium",
                (
                    "Candidate contains unsupported workbook features; "
                    "inspection coverage may be incomplete."
                ),
                details={"warnings": new_parser_warnings},
            )
        )
    return changes, findings


def _array_formula_semantics_changes(
    before: WorkbookSnapshot, after: WorkbookSnapshot
) -> tuple[list[Change], list[Finding]]:
    """Report changes Excel would otherwise hide behind identical formula text.

    The same formula text can run as an ordinary scalar formula, a legacy CSE
    array formula with a fixed output range, or a dynamic array. Dynamic result
    extents are intentionally excluded because they can change at recalc time.
    """
    changes: list[Change] = []
    findings: list[Finding] = []
    known_modes = {"absent", "value", "ordinary", "legacy_cse", "dynamic"}

    def formula_mode(cell: CellSnapshot | None) -> str:
        if cell is None:
            return "absent"
        if not cell.is_formula:
            return "value"
        return cell.array_formula_kind or "ordinary"

    def details_with_impact(location: CellKey, details: dict[str, object]) -> tuple[
        dict[str, object], int, tuple[CellKey, ...]
    ]:
        impact_analysis = analyze_downstream_impact(location, before, after)
        sampled_impacts = tuple(
            sorted(impact_analysis.impacted, key=_location_sort_key)[:_IMPACT_SAMPLE_SIZE]
        )
        if sampled_impacts:
            details["impact_paths"] = _serialise_impact_paths(
                sampled_impacts, impact_analysis.paths
            )
        if impact_analysis.truncated:
            details["impact_truncated_at"] = _IMPACT_NODE_LIMIT
        return details, len(impact_analysis.impacted), sampled_impacts

    locations = sorted(set(before.cells) | set(after.cells), key=_location_sort_key)
    for location in locations:
        old_cell = before.cells.get(location)
        new_cell = after.cells.get(location)
        old_mode = formula_mode(old_cell)
        new_mode = formula_mode(new_cell)
        if old_mode not in known_modes or new_mode not in known_modes:
            continue
        if old_mode != new_mode and {old_mode, new_mode} & {"legacy_cse", "dynamic"}:
            details, impact_count, impacted_cells = details_with_impact(location, {
                "before": {
                    "mode": old_mode,
                    "output_range": old_cell.array_formula_ref if old_cell else None,
                },
                "after": {
                    "mode": new_mode,
                    "output_range": new_cell.array_formula_ref if new_cell else None,
                },
            })
            changes.append(
                Change(
                    "array_formula_mode_changed",
                    location,
                    "high",
                    impact_count=impact_count,
                    impacted_cells=impacted_cells,
                    details=details,
                )
            )
            findings.append(
                Finding(
                    "FF018",
                    "high",
                    (
                        "Formula changed to or from legacy CSE or dynamic array semantics."
                    ),
                    location,
                    details=details,
                )
            )
            continue
        if (
            old_mode == "legacy_cse"
            and new_mode == "legacy_cse"
            and old_cell is not None
            and new_cell is not None
            and old_cell.array_formula_ref != new_cell.array_formula_ref
        ):
            details, impact_count, impacted_cells = details_with_impact(location, {
                "before_output_range": old_cell.array_formula_ref,
                "after_output_range": new_cell.array_formula_ref,
            })
            changes.append(
                Change(
                    "legacy_array_output_range_changed",
                    location,
                    "high",
                    impact_count=impact_count,
                    impacted_cells=impacted_cells,
                    details=details,
                )
            )
            findings.append(
                Finding(
                    "FF018",
                    "high",
                    "Legacy CSE array formula fixed output range changed.",
                    location,
                    details=details,
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
        impact_analysis = analyze_downstream_impact(location, before, after)
        impact = impact_analysis.impacted
        sampled_impacts = tuple(sorted(impact, key=_location_sort_key)[:_IMPACT_SAMPLE_SIZE])
        details: dict[str, object] = {}
        if sampled_impacts:
            details["impact_paths"] = _serialise_impact_paths(
                sampled_impacts, impact_analysis.paths
            )
        if impact_analysis.truncated:
            details["impact_truncated_at"] = _IMPACT_NODE_LIMIT
        changes.append(
            Change(
                kind=kind,
                location=location,
                severity=severity,
                before=old_cell,
                after=new_cell,
                impact_count=len(impact),
                impacted_cells=sampled_impacts,
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
                    details={
                        "impact_count": len(impact),
                        "impact_paths": details.get("impact_paths", []),
                    },
                )
            )
        elif kind == "formula_removed":
            findings.append(
                Finding(
                    "FF002",
                    "high",
                    "Formula cell was removed or blanked.",
                    location,
                    details={
                        "impact_count": len(impact),
                        "impact_paths": details.get("impact_paths", []),
                    },
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

    for location, tokens in _new_coverage_items(
        before.unresolved_reference_tokens, after.unresolved_reference_tokens
    ):
        details = {"tokens": list(tokens)}
        changes.append(
            Change(
                "unresolved_formula_reference_added",
                location,
                "medium",
                details=details,
            )
        )
        findings.append(
            Finding(
                "FF011",
                "medium",
                "Formula introduces a reference that FormulaFence cannot statically resolve.",
                location,
                details=details,
            )
        )

    for location, functions in _new_coverage_items(
        before.dynamic_reference_functions, after.dynamic_reference_functions
    ):
        details = {"functions": list(functions)}
        changes.append(
            Change(
                "dynamic_formula_reference_added",
                location,
                "medium",
                details=details,
            )
        )
        findings.append(
            Finding(
                "FF012",
                "medium",
                (
                    "Formula introduces a dynamic reference function; "
                    "dependency impact may be incomplete."
                ),
                location,
                details=details,
            )
        )

    for location, tokens in _new_coverage_items(
        before.spill_reference_tokens, after.spill_reference_tokens
    ):
        details = {"tokens": list(tokens)}
        changes.append(
            Change(
                "spill_reference_added",
                location,
                "medium",
                details=details,
            )
        )
        findings.append(
            Finding(
                "FF015",
                "medium",
                (
                    "Formula introduces a dynamic-array spill reference; the anchor is "
                    "traced but dynamic extent and blockers may affect impact."
                ),
                location,
                details=details,
            )
        )

    for location in sorted(
        after.tokenization_failure_cells - before.tokenization_failure_cells,
        key=_location_sort_key,
    ):
        changes.append(
            Change(
                "formula_tokenization_failure_added",
                location,
                "medium",
            )
        )
        findings.append(
            Finding(
                "FF016",
                "medium",
                "Formula could not be tokenized; dependency impact may be incomplete.",
                location,
            )
        )

    for location, tokens in _new_coverage_items(
        before.implicit_intersection_tokens, after.implicit_intersection_tokens
    ):
        details = {"tokens": list(tokens)}
        changes.append(
            Change(
                "implicit_intersection_added",
                location,
                "medium",
                details=details,
            )
        )
        findings.append(
            Finding(
                "FF017",
                "medium",
                (
                    "Formula introduces explicit implicit intersection; it can change which "
                    "cell a range or array contributes."
                ),
                location,
                details=details,
            )
        )

    array_formula_changes, array_formula_findings = _array_formula_semantics_changes(
        before, after
    )
    changes.extend(array_formula_changes)
    findings.extend(array_formula_findings)
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
