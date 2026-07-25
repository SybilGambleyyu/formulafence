"""Semantic workbook diffing, downstream impact, and intrinsic-risk detection."""

from __future__ import annotations

from collections import defaultdict, deque
from collections.abc import Iterable, Mapping
from dataclasses import dataclass

from openpyxl.utils.cell import coordinate_to_tuple, get_column_letter

from formulafence.formulas import resolve_3d_reference
from formulafence.models import (
    CellHyperlinkSnapshot,
    CellKey,
    CellSnapshot,
    Change,
    ChartDefinitionSnapshot,
    ConditionalFormattingExtensionSnapshot,
    ConditionalFormattingSnapshot,
    DataValidationSnapshot,
    DiffReport,
    DigitalSignatureSnapshot,
    ExternalDataConnectionSnapshot,
    ExternalLinkPackageSnapshot,
    FillSnapshot,
    FilterVisibilitySnapshot,
    Finding,
    FontSnapshot,
    FormulaCachedResultSnapshot,
    IgnoredErrorSnapshot,
    LegacyCommentSnapshot,
    NamedSheetViewSnapshot,
    NumberFormatSnapshot,
    OfficeWebAddinSnapshot,
    PivotCacheRefreshSnapshot,
    PivotTableDefinitionSnapshot,
    PowerPivotDataModelSnapshot,
    PowerQuerySnapshot,
    ProtectionCredentialSnapshot,
    ProtectionOpaqueMetadataSnapshot,
    QueryTableRefreshSnapshot,
    RibbonCustomizationSnapshot,
    RichDataSnapshot,
    RichTextRunEntry,
    RichTextRunSnapshot,
    ScenarioManagerSnapshot,
    SlicerTimelineCacheSnapshot,
    ThreadedCommentSnapshot,
    WhatIfDataTableSnapshot,
    WorkbookSnapshot,
    WorksheetDrawingShapeSnapshot,
    WorksheetEmbeddedControlSnapshot,
    WorksheetSparklineSnapshot,
    XlmMacroSheetSnapshot,
    XmlMappingSnapshot,
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


def _credential_material_changed(
    before: ProtectionCredentialSnapshot | None,
    after: ProtectionCredentialSnapshot | None,
) -> bool:
    """Compare private verifier signatures without returning verifier material."""
    return (before.signature if before is not None else None) != (
        after.signature if after is not None else None
    )


def _opaque_protection_metadata_changed(
    before: ProtectionOpaqueMetadataSnapshot | None,
    after: ProtectionOpaqueMetadataSnapshot | None,
) -> bool:
    """Compare private opaque metadata fingerprints without exposing content."""
    return (before.signature if before is not None else None) != (
        after.signature if after is not None else None
    )


def _private_external_data_material_changed(
    before: tuple[str | None, ...],
    after: tuple[str | None, ...],
) -> bool:
    """Compare private source or identity fingerprints without returning them."""
    return before != after


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

    before_validations: dict[str, list[DataValidationSnapshot]] = defaultdict(list)
    after_validations: dict[str, list[DataValidationSnapshot]] = defaultdict(list)
    for validation in before.data_validations:
        before_validations[validation.sheet].append(validation)
    for validation in after.data_validations:
        after_validations[validation.sheet].append(validation)
    for sheet in sorted(set(before_validations) | set(after_validations), key=str.casefold):
        old_validations = tuple(before_validations.get(sheet, ()))
        new_validations = tuple(after_validations.get(sheet, ()))
        if old_validations == new_validations:
            continue
        details = {
            "sheet": sheet,
            "before": [validation.to_dict() for validation in old_validations],
            "after": [validation.to_dict() for validation in new_validations],
        }
        changes.append(
            Change(
                "data_validation_changed",
                None,
                "high",
                details=details,
            )
        )
        findings.append(
            Finding(
                "FF020",
                "high",
                f"Data-validation controls changed on worksheet: {sheet}.",
                details=details,
            )
        )

    before_conditional_formatting: dict[str, list[ConditionalFormattingSnapshot]] = (
        defaultdict(list)
    )
    after_conditional_formatting: dict[str, list[ConditionalFormattingSnapshot]] = (
        defaultdict(list)
    )
    before_conditional_extensions: dict[
        str, list[ConditionalFormattingExtensionSnapshot]
    ] = defaultdict(list)
    after_conditional_extensions: dict[
        str, list[ConditionalFormattingExtensionSnapshot]
    ] = defaultdict(list)
    for rule in before.conditional_formatting:
        before_conditional_formatting[rule.sheet].append(rule)
    for rule in after.conditional_formatting:
        after_conditional_formatting[rule.sheet].append(rule)
    for extension in before.conditional_formatting_extensions:
        before_conditional_extensions[extension.sheet].append(extension)
    for extension in after.conditional_formatting_extensions:
        after_conditional_extensions[extension.sheet].append(extension)
    conditional_sheets = (
        set(before_conditional_formatting)
        | set(after_conditional_formatting)
        | set(before_conditional_extensions)
        | set(after_conditional_extensions)
    )
    for sheet in sorted(conditional_sheets, key=str.casefold):
        old_rules = tuple(before_conditional_formatting.get(sheet, ()))
        new_rules = tuple(after_conditional_formatting.get(sheet, ()))
        old_extensions = tuple(before_conditional_extensions.get(sheet, ()))
        new_extensions = tuple(after_conditional_extensions.get(sheet, ()))
        if old_rules == new_rules and old_extensions == new_extensions:
            continue
        details = {
            "sheet": sheet,
            "before": {
                "rules": [rule.to_dict() for rule in old_rules],
                "extensions": [extension.to_dict() for extension in old_extensions],
            },
            "after": {
                "rules": [rule.to_dict() for rule in new_rules],
                "extensions": [extension.to_dict() for extension in new_extensions],
            },
        }
        changes.append(
            Change(
                "conditional_formatting_changed",
                None,
                "high",
                details=details,
            )
        )
        findings.append(
            Finding(
                "FF021",
                "high",
                f"Conditional-formatting controls changed on worksheet: {sheet}.",
                details=details,
            )
        )

    if before.workbook_protection != after.workbook_protection:
        old_protection = before.workbook_protection
        new_protection = after.workbook_protection
        details: dict[str, object] = {
            "before": old_protection.to_dict() if old_protection is not None else None,
            "after": new_protection.to_dict() if new_protection is not None else None,
        }
        credential_changes = {
            "workbook": _credential_material_changed(
                (
                    old_protection.workbook_credential
                    if old_protection is not None
                    else None
                ),
                (
                    new_protection.workbook_credential
                    if new_protection is not None
                    else None
                ),
            ),
            "revisions": _credential_material_changed(
                (
                    old_protection.revisions_credential
                    if old_protection is not None
                    else None
                ),
                (
                    new_protection.revisions_credential
                    if new_protection is not None
                    else None
                ),
            ),
        }
        if any(credential_changes.values()):
            details["credential_material_changed"] = credential_changes
        if _opaque_protection_metadata_changed(
            old_protection.opaque_metadata if old_protection is not None else None,
            new_protection.opaque_metadata if new_protection is not None else None,
        ):
            details["opaque_metadata_changed"] = True
        changes.append(
            Change(
                "workbook_protection_changed",
                None,
                "high",
                details=details,
            )
        )
        findings.append(
            Finding(
                "FF022",
                "high",
                "Workbook protection controls changed.",
                details=details,
            )
        )

    before_sheet_protections = {
        protection.sheet: protection for protection in before.sheet_protections
    }
    after_sheet_protections = {
        protection.sheet: protection for protection in after.sheet_protections
    }
    for sheet in sorted(
        set(before_sheet_protections) | set(after_sheet_protections), key=str.casefold
    ):
        old_protection = before_sheet_protections.get(sheet)
        new_protection = after_sheet_protections.get(sheet)
        if old_protection == new_protection:
            continue
        details = {
            "sheet": sheet,
            "before": old_protection.to_dict() if old_protection is not None else None,
            "after": new_protection.to_dict() if new_protection is not None else None,
        }
        if _credential_material_changed(
            old_protection.credential if old_protection is not None else None,
            new_protection.credential if new_protection is not None else None,
        ):
            details["credential_material_changed"] = True
        if _opaque_protection_metadata_changed(
            old_protection.opaque_metadata if old_protection is not None else None,
            new_protection.opaque_metadata if new_protection is not None else None,
        ):
            details["opaque_metadata_changed"] = True
        changes.append(
            Change(
                "sheet_protection_changed",
                None,
                "high",
                details=details,
            )
        )
        findings.append(
            Finding(
                "FF022",
                "high",
                f"Sheet protection controls changed: {sheet}.",
                details=details,
            )
        )

    before_ranges: dict[str, list[object]] = defaultdict(list)
    after_ranges: dict[str, list[object]] = defaultdict(list)
    for protected_range in before.protected_ranges:
        before_ranges[protected_range.sheet].append(protected_range)
    for protected_range in after.protected_ranges:
        after_ranges[protected_range.sheet].append(protected_range)
    for sheet in sorted(set(before_ranges) | set(after_ranges), key=str.casefold):
        old_ranges = tuple(before_ranges.get(sheet, ()))
        new_ranges = tuple(after_ranges.get(sheet, ()))
        if old_ranges == new_ranges:
            continue
        details = {
            "sheet": sheet,
            "before": [protected_range.to_dict() for protected_range in old_ranges],
            "after": [protected_range.to_dict() for protected_range in new_ranges],
        }
        if tuple(item.name_signature for item in old_ranges) != tuple(
            item.name_signature for item in new_ranges
        ):
            details["range_name_material_changed"] = True
        if tuple(item.security_descriptor_signature for item in old_ranges) != tuple(
            item.security_descriptor_signature for item in new_ranges
        ):
            details["security_descriptor_material_changed"] = True
        if tuple(item.credential.signature for item in old_ranges) != tuple(
            item.credential.signature for item in new_ranges
        ):
            details["credential_material_changed"] = True
        if tuple(item.opaque_metadata.signature for item in old_ranges) != tuple(
            item.opaque_metadata.signature for item in new_ranges
        ):
            details["opaque_metadata_changed"] = True
        changes.append(
            Change(
                "protected_range_permissions_changed",
                None,
                "high",
                details=details,
            )
        )
        findings.append(
            Finding(
                "FF022",
                "high",
                f"Protected-range permissions changed on worksheet: {sheet}.",
                details=details,
            )
        )

    if before.cell_protection_default != after.cell_protection_default:
        details = {
            "before": (
                before.cell_protection_default.to_dict()
                if before.cell_protection_default is not None
                else None
            ),
            "after": (
                after.cell_protection_default.to_dict()
                if after.cell_protection_default is not None
                else None
            ),
        }
        changes.append(
            Change(
                "cell_protection_default_changed",
                None,
                "high",
                details=details,
            )
        )
        findings.append(
            Finding(
                "FF022",
                "high",
                "Default locked/hidden cell protection changed.",
                details=details,
            )
        )

    before_assignments: dict[str, list[object]] = defaultdict(list)
    after_assignments: dict[str, list[object]] = defaultdict(list)
    for assignment in before.cell_protection_assignments:
        before_assignments[assignment.sheet].append(assignment)
    for assignment in after.cell_protection_assignments:
        after_assignments[assignment.sheet].append(assignment)
    for sheet in sorted(
        set(before_assignments) | set(after_assignments), key=str.casefold
    ):
        old_assignments = tuple(before_assignments.get(sheet, ()))
        new_assignments = tuple(after_assignments.get(sheet, ()))
        if old_assignments == new_assignments:
            continue
        details = {
            "sheet": sheet,
            "before": [assignment.to_dict() for assignment in old_assignments],
            "after": [assignment.to_dict() for assignment in new_assignments],
        }
        changes.append(
            Change(
                "cell_protection_assignments_changed",
                None,
                "high",
                details=details,
            )
        )
        findings.append(
            Finding(
                "FF022",
                "high",
                f"Direct cell protection assignments changed on worksheet: {sheet}.",
                details=details,
            )
        )

    if before.external_data_refresh_settings != after.external_data_refresh_settings:
        details = {
            "before": before.external_data_refresh_settings.to_dict(),
            "after": after.external_data_refresh_settings.to_dict(),
        }
        changes.append(
            Change(
                "external_data_refresh_settings_changed",
                None,
                "high",
                details=details,
            )
        )
        findings.append(
            Finding(
                "FF023",
                "high",
                "Workbook-wide external-data refresh settings changed.",
                details=details,
            )
        )

    if before.external_data_connections != after.external_data_connections:
        old_connections: tuple[ExternalDataConnectionSnapshot, ...] = (
            before.external_data_connections
        )
        new_connections: tuple[ExternalDataConnectionSnapshot, ...] = (
            after.external_data_connections
        )
        details: dict[str, object] = {
            "before": [connection.to_dict() for connection in old_connections],
            "after": [connection.to_dict() for connection in new_connections],
        }
        if _private_external_data_material_changed(
            tuple(connection.identity_signature for connection in old_connections),
            tuple(connection.identity_signature for connection in new_connections),
        ):
            details["identity_material_changed"] = True
        if _private_external_data_material_changed(
            tuple(
                connection.source_configuration_signature
                for connection in old_connections
            ),
            tuple(
                connection.source_configuration_signature
                for connection in new_connections
            ),
        ):
            details["source_configuration_material_changed"] = True
        if _private_external_data_material_changed(
            tuple(connection.opaque_metadata.signature for connection in old_connections),
            tuple(connection.opaque_metadata.signature for connection in new_connections),
        ):
            details["opaque_metadata_changed"] = True
        changes.append(
            Change(
                "external_data_connections_changed",
                None,
                "high",
                details=details,
            )
        )
        findings.append(
            Finding(
                "FF023",
                "high",
                "External-data connection controls changed.",
                details=details,
            )
        )

    if before.query_table_refresh_controls != after.query_table_refresh_controls:
        old_query_tables: tuple[QueryTableRefreshSnapshot, ...] = (
            before.query_table_refresh_controls
        )
        new_query_tables: tuple[QueryTableRefreshSnapshot, ...] = (
            after.query_table_refresh_controls
        )
        details = {
            "before": [control.to_dict() for control in old_query_tables],
            "after": [control.to_dict() for control in new_query_tables],
        }
        if _private_external_data_material_changed(
            tuple(control.identity_signature for control in old_query_tables),
            tuple(control.identity_signature for control in new_query_tables),
        ):
            details["identity_material_changed"] = True
        if _private_external_data_material_changed(
            tuple(control.opaque_metadata.signature for control in old_query_tables),
            tuple(control.opaque_metadata.signature for control in new_query_tables),
        ):
            details["opaque_metadata_changed"] = True
        changes.append(
            Change(
                "query_table_refresh_controls_changed",
                None,
                "high",
                details=details,
            )
        )
        findings.append(
            Finding(
                "FF023",
                "high",
                "Query-table refresh controls changed.",
                details=details,
            )
        )

    if before.pivot_cache_refresh_controls != after.pivot_cache_refresh_controls:
        old_pivot_caches: tuple[PivotCacheRefreshSnapshot, ...] = (
            before.pivot_cache_refresh_controls
        )
        new_pivot_caches: tuple[PivotCacheRefreshSnapshot, ...] = (
            after.pivot_cache_refresh_controls
        )
        details = {
            "before": [control.to_dict() for control in old_pivot_caches],
            "after": [control.to_dict() for control in new_pivot_caches],
        }
        if _private_external_data_material_changed(
            tuple(
                control.source_configuration_signature
                for control in old_pivot_caches
            ),
            tuple(
                control.source_configuration_signature
                for control in new_pivot_caches
            ),
        ):
            details["source_configuration_material_changed"] = True
        if _private_external_data_material_changed(
            tuple(control.opaque_metadata.signature for control in old_pivot_caches),
            tuple(control.opaque_metadata.signature for control in new_pivot_caches),
        ):
            details["opaque_metadata_changed"] = True
        changes.append(
            Change(
                "pivot_cache_refresh_controls_changed",
                None,
                "high",
                details=details,
            )
        )
        findings.append(
            Finding(
                "FF023",
                "high",
                "Pivot-cache refresh controls changed.",
                details=details,
            )
        )

    if before.external_link_packages != after.external_link_packages:
        old_external_links: ExternalLinkPackageSnapshot = before.external_link_packages
        new_external_links: ExternalLinkPackageSnapshot = after.external_link_packages
        details: dict[str, object] = {
            "before": old_external_links.to_dict(),
            "after": new_external_links.to_dict(),
        }
        if old_external_links.source_signature != new_external_links.source_signature:
            details["source_material_changed"] = True
        if old_external_links.definition_signature != new_external_links.definition_signature:
            details["definition_material_changed"] = True
        if (
            old_external_links.cached_material_signature
            != new_external_links.cached_material_signature
        ):
            details["cached_material_changed"] = True
        if (
            old_external_links.opaque_metadata.signature
            != new_external_links.opaque_metadata.signature
        ):
            details["opaque_metadata_changed"] = True
        changes.append(
            Change(
                "external_link_packages_changed",
                None,
                "high",
                details=details,
            )
        )
        findings.append(
            Finding(
                "FF025",
                "high",
                "External-workbook, DDE, or OLE link package controls changed.",
                details=details,
            )
        )

    if before.power_query != after.power_query:
        old_power_query: PowerQuerySnapshot = before.power_query
        new_power_query: PowerQuerySnapshot = after.power_query
        details: dict[str, object] = {
            "before": old_power_query.to_dict(),
            "after": new_power_query.to_dict(),
        }
        if old_power_query.formula_signature != new_power_query.formula_signature:
            details["formula_material_changed"] = True
        if (
            old_power_query.package_configuration_signature
            != new_power_query.package_configuration_signature
        ):
            details["package_configuration_material_changed"] = True
        if (
            old_power_query.metadata_identity_signature
            != new_power_query.metadata_identity_signature
        ):
            details["metadata_identity_material_changed"] = True
        if (
            old_power_query.metadata_control_signature
            != new_power_query.metadata_control_signature
        ):
            details["metadata_control_material_changed"] = True
        if old_power_query.permission_controls != new_power_query.permission_controls:
            details["permission_controls_changed"] = True
        if (
            old_power_query.permission_binding_count
            != new_power_query.permission_binding_count
        ):
            details["permission_binding_presence_changed"] = True
        if (
            old_power_query.opaque_metadata.signature
            != new_power_query.opaque_metadata.signature
        ):
            details["opaque_metadata_changed"] = True
        changes.append(
            Change(
                "power_query_changed",
                None,
                "high",
                details=details,
            )
        )
        findings.append(
            Finding(
                "FF024",
                "high",
                "Power Query formulas or semantic query controls changed.",
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
    if before.xlm_macro_sheets != after.xlm_macro_sheets:
        old_xlm_macro_sheets: XlmMacroSheetSnapshot = before.xlm_macro_sheets
        new_xlm_macro_sheets: XlmMacroSheetSnapshot = after.xlm_macro_sheets
        details: dict[str, object] = {
            "before": old_xlm_macro_sheets.to_dict(),
            "after": new_xlm_macro_sheets.to_dict(),
        }
        if (
            old_xlm_macro_sheets.declaration_signature
            != new_xlm_macro_sheets.declaration_signature
        ):
            details["workbook_binding_changed"] = True
        if (
            old_xlm_macro_sheets.program_signature
            != new_xlm_macro_sheets.program_signature
        ):
            details["macro_program_material_changed"] = True
        if (
            old_xlm_macro_sheets.relationship_signature
            != new_xlm_macro_sheets.relationship_signature
        ):
            details["related_part_relationships_changed"] = True
        if (
            old_xlm_macro_sheets.related_part_payload_signature
            != new_xlm_macro_sheets.related_part_payload_signature
        ):
            details["related_part_payload_material_changed"] = True
        changes.append(
            Change(
                "xlm_macro_sheets_changed",
                None,
                "critical",
                details=details,
            )
        )
        findings.append(
            Finding(
                "FF026",
                "critical",
                "Excel 4.0 / XLM macro-sheet controls changed.",
                details=details,
            )
        )
    if before.ribbon_customization != after.ribbon_customization:
        old_ribbon: RibbonCustomizationSnapshot = before.ribbon_customization
        new_ribbon: RibbonCustomizationSnapshot = after.ribbon_customization
        details = {
            "before": old_ribbon.to_dict(),
            "after": new_ribbon.to_dict(),
        }
        if old_ribbon.declaration_signature != new_ribbon.declaration_signature:
            details["package_binding_changed"] = True
        if old_ribbon.definition_signature != new_ribbon.definition_signature:
            details["ribbon_definition_material_changed"] = True
        if old_ribbon.relationship_signature != new_ribbon.relationship_signature:
            details["image_relationships_changed"] = True
        changes.append(
            Change(
                "ribbon_customization_changed",
                None,
                "critical",
                details=details,
            )
        )
        findings.append(
            Finding(
                "FF027",
                "critical",
                "Office RibbonX customization controls changed.",
                details=details,
            )
        )
    if before.office_web_addins != after.office_web_addins:
        old_addins: OfficeWebAddinSnapshot = before.office_web_addins
        new_addins: OfficeWebAddinSnapshot = after.office_web_addins
        details = {
            "before": old_addins.to_dict(),
            "after": new_addins.to_dict(),
        }
        if old_addins.declaration_signature != new_addins.declaration_signature:
            details["workbook_binding_changed"] = True
        if old_addins.taskpane_signature != new_addins.taskpane_signature:
            details["taskpane_configuration_material_changed"] = True
        if old_addins.web_extension_signature != new_addins.web_extension_signature:
            details["web_extension_definition_material_changed"] = True
        if old_addins.relationship_signature != new_addins.relationship_signature:
            details["related_part_relationships_changed"] = True
        changes.append(
            Change(
                "office_web_addins_changed",
                None,
                "critical",
                details=details,
            )
        )
        findings.append(
            Finding(
                "FF028",
                "critical",
                "Office Web Add-in task-pane controls changed.",
                details=details,
            )
        )
    if before.worksheet_embedded_controls != after.worksheet_embedded_controls:
        old_controls: WorksheetEmbeddedControlSnapshot = before.worksheet_embedded_controls
        new_controls: WorksheetEmbeddedControlSnapshot = after.worksheet_embedded_controls
        details = {
            "before": old_controls.to_dict(),
            "after": new_controls.to_dict(),
        }
        if old_controls.declaration_signature != new_controls.declaration_signature:
            details["worksheet_binding_changed"] = True
        if (
            old_controls.control_definition_signature
            != new_controls.control_definition_signature
        ):
            details["worksheet_control_definition_material_changed"] = True
        if (
            old_controls.active_x_definition_signature
            != new_controls.active_x_definition_signature
        ):
            details["active_x_definition_material_changed"] = True
        if (
            old_controls.form_control_property_signature
            != new_controls.form_control_property_signature
        ):
            details["form_control_property_material_changed"] = True
        if (
            old_controls.legacy_vml_definition_signature
            != new_controls.legacy_vml_definition_signature
        ):
            details["legacy_vml_control_definition_material_changed"] = True
        if (
            old_controls.legacy_vml_relationship_signature
            != new_controls.legacy_vml_relationship_signature
        ):
            details["legacy_vml_related_part_relationships_changed"] = True
        if old_controls.relationship_signature != new_controls.relationship_signature:
            details["related_part_relationships_changed"] = True
        if (
            old_controls.related_part_payload_signature
            != new_controls.related_part_payload_signature
        ):
            details["embedded_payload_material_changed"] = True
        changes.append(
            Change(
                "worksheet_embedded_controls_changed",
                None,
                "critical",
                details=details,
            )
        )
        findings.append(
            Finding(
                "FF029",
                "critical",
                "Worksheet embedded, legacy VML controls, or OLE objects changed.",
                details=details,
            )
        )
    if before.chart_definitions != after.chart_definitions:
        old_charts: ChartDefinitionSnapshot = before.chart_definitions
        new_charts: ChartDefinitionSnapshot = after.chart_definitions
        details: dict[str, object] = {
            "before": old_charts.to_dict(),
            "after": new_charts.to_dict(),
        }
        if old_charts.declaration_signature != new_charts.declaration_signature:
            details["drawing_binding_changed"] = True
        if old_charts.definition_signature != new_charts.definition_signature:
            details["chart_definition_material_changed"] = True
        if old_charts.cached_data_signature != new_charts.cached_data_signature:
            details["cached_series_material_changed"] = True
        if old_charts.user_shape_signature != new_charts.user_shape_signature:
            details["overlay_shape_material_changed"] = True
        if old_charts.relationship_signature != new_charts.relationship_signature:
            details["related_part_relationships_changed"] = True
        if (
            old_charts.related_part_payload_signature
            != new_charts.related_part_payload_signature
        ):
            details["related_part_payload_material_changed"] = True
        changes.append(
            Change(
                "chart_definitions_changed",
                None,
                "high",
                details=details,
            )
        )
        findings.append(
            Finding(
                "FF030",
                "high",
                "Chart definition, cached series data, or overlay shape material changed.",
                details=details,
            )
        )
    if before.pivot_table_definitions != after.pivot_table_definitions:
        old_pivots: PivotTableDefinitionSnapshot = before.pivot_table_definitions
        new_pivots: PivotTableDefinitionSnapshot = after.pivot_table_definitions
        details: dict[str, object] = {
            "before": old_pivots.to_dict(),
            "after": new_pivots.to_dict(),
        }
        if old_pivots.declaration_signature != new_pivots.declaration_signature:
            details["pivot_table_binding_changed"] = True
        if old_pivots.layout_signature != new_pivots.layout_signature:
            details["pivot_table_layout_material_changed"] = True
        if (
            old_pivots.cache_definition_signature
            != new_pivots.cache_definition_signature
        ):
            details["pivot_cache_definition_material_changed"] = True
        if (
            old_pivots.cached_shared_item_signature
            != new_pivots.cached_shared_item_signature
        ):
            details["cached_shared_item_material_changed"] = True
        if old_pivots.relationship_signature != new_pivots.relationship_signature:
            details["related_part_relationships_changed"] = True
        if (
            old_pivots.cache_record_payload_signature
            != new_pivots.cache_record_payload_signature
        ):
            details["cache_record_payload_material_changed"] = True
        changes.append(
            Change(
                "pivot_table_definitions_changed",
                None,
                "high",
                details=details,
            )
        )
        findings.append(
            Finding(
                "FF031",
                "high",
                "PivotTable view, cache-schema, shared-item, or cached-record material changed.",
                details=details,
            )
        )
    if before.slicer_timeline_caches != after.slicer_timeline_caches:
        old_filters: SlicerTimelineCacheSnapshot = before.slicer_timeline_caches
        new_filters: SlicerTimelineCacheSnapshot = after.slicer_timeline_caches
        details: dict[str, object] = {
            "before": old_filters.to_dict(),
            "after": new_filters.to_dict(),
        }
        if old_filters.declaration_signature != new_filters.declaration_signature:
            details["workbook_cache_binding_changed"] = True
        if (
            old_filters.slicer_definition_signature
            != new_filters.slicer_definition_signature
        ):
            details["slicer_filter_state_or_definition_material_changed"] = True
        if (
            old_filters.timeline_definition_signature
            != new_filters.timeline_definition_signature
        ):
            details["timeline_filter_state_or_definition_material_changed"] = True
        if old_filters.relationship_signature != new_filters.relationship_signature:
            details["related_part_relationships_changed"] = True
        changes.append(
            Change(
                "slicer_timeline_cache_definitions_changed",
                None,
                "high",
                details=details,
            )
        )
        findings.append(
            Finding(
                "FF032",
                "high",
                "Slicer or Timeline cache filter state or definition changed.",
                details=details,
            )
        )
    if before.power_pivot_data_model != after.power_pivot_data_model:
        old_data_model: PowerPivotDataModelSnapshot = before.power_pivot_data_model
        new_data_model: PowerPivotDataModelSnapshot = after.power_pivot_data_model
        details: dict[str, object] = {
            "before": old_data_model.to_dict(),
            "after": new_data_model.to_dict(),
        }
        if old_data_model.declaration_signature != new_data_model.declaration_signature:
            details["workbook_data_model_declaration_changed"] = True
        if old_data_model.relationship_signature != new_data_model.relationship_signature:
            details["related_data_model_relationships_changed"] = True
        if old_data_model.payload_signature != new_data_model.payload_signature:
            details["embedded_data_model_payload_changed"] = True
        changes.append(
            Change(
                "power_pivot_data_model_changed",
                None,
                "high",
                details=details,
            )
        )
        findings.append(
            Finding(
                "FF033",
                "high",
                "Embedded Power Pivot/Data Model definition or payload changed.",
                details=details,
            )
        )
    if before.what_if_data_tables != after.what_if_data_tables:
        old_data_tables: WhatIfDataTableSnapshot = before.what_if_data_tables
        new_data_tables: WhatIfDataTableSnapshot = after.what_if_data_tables
        details: dict[str, object] = {
            "before": old_data_tables.to_dict(),
            "after": new_data_tables.to_dict(),
        }
        if old_data_tables.definition_signature != new_data_tables.definition_signature:
            details["data_table_definition_material_changed"] = True
        changes.append(
            Change(
                "what_if_data_tables_changed",
                None,
                "high",
                details=details,
            )
        )
        findings.append(
            Finding(
                "FF034",
                "high",
                "What-If Data Table sensitivity definition or control changed.",
                details=details,
            )
        )
    if before.scenario_manager != after.scenario_manager:
        old_scenarios: ScenarioManagerSnapshot = before.scenario_manager
        new_scenarios: ScenarioManagerSnapshot = after.scenario_manager
        details: dict[str, object] = {
            "before": old_scenarios.to_dict(),
            "after": new_scenarios.to_dict(),
        }
        if old_scenarios.definition_signature != new_scenarios.definition_signature:
            details["scenario_definition_material_changed"] = True
        changes.append(
            Change(
                "scenario_manager_changed",
                None,
                "high",
                details=details,
            )
        )
        findings.append(
            Finding(
                "FF035",
                "high",
                "Excel Scenario Manager definition or stored input set changed.",
                details=details,
            )
        )
    if before.filter_visibility_controls != after.filter_visibility_controls:
        old_controls: FilterVisibilitySnapshot = before.filter_visibility_controls
        new_controls: FilterVisibilitySnapshot = after.filter_visibility_controls
        details: dict[str, object] = {
            "before": old_controls.to_dict(),
            "after": new_controls.to_dict(),
        }
        if old_controls.definition_signature != new_controls.definition_signature:
            details["filter_visibility_definition_material_changed"] = True
        changes.append(
            Change(
                "filter_visibility_controls_changed",
                None,
                "high",
                details=details,
            )
        )
        findings.append(
            Finding(
                "FF036",
                "high",
                "Worksheet/table filter, sort, or row/column-visibility control changed.",
                details=details,
            )
        )
    if before.ignored_error_controls != after.ignored_error_controls:
        old_controls: IgnoredErrorSnapshot = before.ignored_error_controls
        new_controls: IgnoredErrorSnapshot = after.ignored_error_controls
        details: dict[str, object] = {
            "before": old_controls.to_dict(),
            "after": new_controls.to_dict(),
        }
        if old_controls.definition_signature != new_controls.definition_signature:
            details["ignored_error_definition_material_changed"] = True
        changes.append(
            Change(
                "ignored_error_controls_changed",
                None,
                "high",
                details=details,
            )
        )
        findings.append(
            Finding(
                "FF037",
                "high",
                "Excel ignored-error controls changed; review warnings may be suppressed "
                "or restored.",
                details=details,
            )
        )
    if before.named_sheet_views != after.named_sheet_views:
        old_views: NamedSheetViewSnapshot = before.named_sheet_views
        new_views: NamedSheetViewSnapshot = after.named_sheet_views
        details: dict[str, object] = {
            "before": old_views.to_dict(),
            "after": new_views.to_dict(),
        }
        if old_views.definition_signature != new_views.definition_signature:
            details["named_sheet_view_definition_material_changed"] = True
        changes.append(
            Change(
                "named_sheet_views_changed",
                None,
                "high",
                details=details,
            )
        )
        findings.append(
            Finding(
                "FF038",
                "high",
                "Excel Named Sheet View controls changed; alternate filter or sort views "
                "may show a different report.",
                details=details,
            )
        )
    if before.number_format_controls != after.number_format_controls:
        old_controls: NumberFormatSnapshot = before.number_format_controls
        new_controls: NumberFormatSnapshot = after.number_format_controls
        details: dict[str, object] = {
            "before": old_controls.to_dict(),
            "after": new_controls.to_dict(),
        }
        if old_controls.definition_signature != new_controls.definition_signature:
            details["number_format_definition_material_changed"] = True
        changes.append(
            Change(
                "number_format_controls_changed",
                None,
                "high",
                details=details,
            )
        )
        findings.append(
            Finding(
                "FF039",
                "high",
                "Cell number-format controls changed; displayed values may be hidden, "
                "scaled, or reinterpreted.",
                details=details,
            )
        )
    if before.font_controls != after.font_controls:
        old_controls: FontSnapshot = before.font_controls
        new_controls: FontSnapshot = after.font_controls
        details = {
            "before": old_controls.to_dict(),
            "after": new_controls.to_dict(),
        }
        if old_controls.definition_signature != new_controls.definition_signature:
            details["font_definition_material_changed"] = True
        changes.append(
            Change(
                "font_controls_changed",
                None,
                "high",
                details=details,
            )
        )
        findings.append(
            Finding(
                "FF040",
                "high",
                "Cell font controls changed; values or warnings may be made less visible "
                "or misleading.",
                details=details,
            )
        )
    if before.fill_controls != after.fill_controls:
        old_controls: FillSnapshot = before.fill_controls
        new_controls: FillSnapshot = after.fill_controls
        details = {
            "before": old_controls.to_dict(),
            "after": new_controls.to_dict(),
        }
        if old_controls.definition_signature != new_controls.definition_signature:
            details["fill_definition_material_changed"] = True
        changes.append(
            Change(
                "fill_controls_changed",
                None,
                "high",
                details=details,
            )
        )
        findings.append(
            Finding(
                "FF041",
                "high",
                "Cell fill controls changed; values, warnings, or visual classifications "
                "may be made less visible or misleading.",
                details=details,
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


def _dynamic_array_output_reference_changes(
    before: WorkbookSnapshot, after: WorkbookSnapshot
) -> tuple[list[Change], list[Finding]]:
    """Report formulas newly intersecting an observed dynamic spill member.

    The relationship is deliberately keyed by consumer and anchor, rather than
    its observed range text. A dynamic array can resize during recalc, so a
    changed cached extent alone must not act like a fixed-range semantic change.
    """
    changes: list[Change] = []
    findings: list[Finding] = []
    for location in sorted(after.dynamic_array_output_references, key=_location_sort_key):
        previous_anchors = {
            reference.anchor
            for reference in before.dynamic_array_output_references.get(location, ())
        }
        new_references = tuple(
            reference
            for reference in after.dynamic_array_output_references[location]
            if reference.anchor not in previous_anchors
        )
        if not new_references:
            continue
        impact_analysis = analyze_downstream_impact(location, before, after)
        sampled_impacts = tuple(
            sorted(impact_analysis.impacted, key=_location_sort_key)[:_IMPACT_SAMPLE_SIZE]
        )
        details: dict[str, object] = {
            "references": [reference.to_dict() for reference in new_references],
        }
        if sampled_impacts:
            details["impact_paths"] = _serialise_impact_paths(
                sampled_impacts, impact_analysis.paths
            )
        if impact_analysis.truncated:
            details["impact_truncated_at"] = _IMPACT_NODE_LIMIT
        changes.append(
            Change(
                "dynamic_array_output_reference_added",
                location,
                "medium",
                impact_count=len(impact_analysis.impacted),
                impacted_cells=sampled_impacts,
                details=details,
            )
        )
        findings.append(
            Finding(
                "FF019",
                "medium",
                (
                    "Formula newly intersects a non-anchor member of an observed "
                    "dynamic-array spill; the observed output extent can resize."
                ),
                location,
                details=details,
            )
        )
    return changes, findings


def _formula_cached_result_entry_map(
    snapshot: FormulaCachedResultSnapshot,
) -> dict[CellKey, tuple[tuple[str, str], ...]]:
    """Index private result digests without returning them in review evidence."""
    entries: dict[CellKey, list[tuple[str, str]]] = defaultdict(list)
    for entry in snapshot.entries:
        entries[entry.location].append((entry.result_type, entry.result_signature))
    return {
        location: tuple(sorted(values))
        for location, values in entries.items()
    }


def _formula_cached_result_changes(
    before: WorkbookSnapshot,
    after: WorkbookSnapshot,
    semantic_cell_changes: Iterable[Change],
) -> tuple[list[Change], list[Finding]]:
    """Flag result-cache changes that static visible edits cannot explain.

    Stored formula results are not exposed by the normal formula reader. A
    calculation following an ordinary input or formula edit can legitimately
    update them, so explicit downstream dependencies suppress those cases.
    FormulaFence deliberately does not evaluate formulas or infer dynamic,
    volatile, or external inputs: a remaining change is review evidence, not a
    claim that the result is mathematically incorrect.
    """
    old_results = before.formula_cached_results
    new_results = after.formula_cached_results
    if old_results == new_results:
        return [], []

    old_entries = _formula_cached_result_entry_map(old_results)
    new_entries = _formula_cached_result_entry_map(new_results)
    changed_locations = {
        location
        for location in set(old_entries) | set(new_entries)
        if old_entries.get(location) != new_entries.get(location)
    }
    unexplained_locations = set(changed_locations)
    changes_with_locations = tuple(
        change for change in semantic_cell_changes if change.location is not None
    )

    # A changed formula owns its own cache. A value/formula edit can also
    # explain formula caches reachable through FormulaFence's static graph.
    for change in changes_with_locations:
        old_cell = change.before
        new_cell = change.after
        if (old_cell is not None and old_cell.is_formula) or (
            new_cell is not None and new_cell.is_formula
        ):
            unexplained_locations.discard(change.location)

    for change in changes_with_locations:
        if not unexplained_locations:
            break
        impact = analyze_downstream_impact(change.location, before, after)
        unexplained_locations.difference_update(impact.impacted)

    unrecognized_metadata_changed = (
        old_results.unrecognized_cached_result_count
        != new_results.unrecognized_cached_result_count
    )
    if not unexplained_locations and not unrecognized_metadata_changed:
        return [], []

    details: dict[str, object] = {
        "before": old_results.to_dict(),
        "after": new_results.to_dict(),
        "unexplained_cached_result_change_count": len(unexplained_locations),
    }
    if old_results.definition_signature != new_results.definition_signature:
        details["cached_result_material_changed"] = True
    if unrecognized_metadata_changed:
        details["unrecognized_cached_result_metadata_changed"] = True
    change = Change(
        "formula_cached_result_changed",
        None,
        "high",
        details=details,
    )
    finding = Finding(
        "FF042",
        "high",
        (
            "Stored formula results changed without a formula or statically visible "
            "precedent change."
        ),
        details=details,
    )
    return [change], [finding]


def _rich_text_run_entry_map(
    snapshot: RichTextRunSnapshot,
) -> dict[CellKey, RichTextRunEntry]:
    """Index private rich-string presentation records by their effective cells."""
    return {entry.location: entry for entry in snapshot.entries}


def _rich_text_run_controls_changed(
    before: WorkbookSnapshot,
    after: WorkbookSnapshot,
) -> tuple[list[Change], list[Finding]]:
    """Flag character-level presentation changes invisible to normal cell diffs.

    A reader normally exposes the concatenated cell text, not the formatting of
    its individual runs. This check avoids duplicating ordinary text edits:
    changing text inside an unchanged run-property sequence stays a normal cell
    change, while a changed property set or a moved style boundary is guarded.
    """
    old_controls = before.rich_text_runs
    new_controls = after.rich_text_runs
    if old_controls == new_controls:
        return [], []

    old_entries = _rich_text_run_entry_map(old_controls)
    new_entries = _rich_text_run_entry_map(new_controls)
    changed_locations: set[CellKey] = set()
    for location in set(old_entries) | set(new_entries):
        old_entry = old_entries.get(location)
        new_entry = new_entries.get(location)
        if old_entry is None or new_entry is None:
            entry = old_entry or new_entry
            if entry is not None and entry.style_sequence_signature is not None:
                changed_locations.add(location)
            continue
        if old_entry.style_sequence_signature != new_entry.style_sequence_signature:
            changed_locations.add(location)
            continue
        if (
            old_entry.text_signature == new_entry.text_signature
            and old_entry.style_layout_signature != new_entry.style_layout_signature
        ):
            changed_locations.add(location)

    unrecognized_metadata_changed = (
        old_controls.unrecognized_rich_text_count
        != new_controls.unrecognized_rich_text_count
    )
    unrecognized_material_changed = (
        (old_controls.unrecognized_rich_text_count or new_controls.unrecognized_rich_text_count)
        and old_controls.definition_signature != new_controls.definition_signature
    )
    if (
        not changed_locations
        and not unrecognized_metadata_changed
        and not unrecognized_material_changed
    ):
        return [], []

    details: dict[str, object] = {
        "before": old_controls.to_dict(),
        "after": new_controls.to_dict(),
        "rich_text_run_control_change_count": len(changed_locations),
    }
    if old_controls.definition_signature != new_controls.definition_signature:
        details["rich_text_run_definition_material_changed"] = True
    if unrecognized_metadata_changed or unrecognized_material_changed:
        details["unrecognized_rich_text_metadata_changed"] = True
    change = Change(
        "rich_text_run_controls_changed",
        None,
        "high",
        details=details,
    )
    finding = Finding(
        "FF043",
        "high",
        (
            "Rich-text run controls changed; character-level formatting or "
            "phonetic hints may be made less visible or misleading."
        ),
        details=details,
    )
    return [change], [finding]


def _legacy_comment_controls_changed(
    before: WorkbookSnapshot,
    after: WorkbookSnapshot,
) -> tuple[list[Change], list[Finding]]:
    """Flag legacy note changes that ordinary cell diffs cannot expose."""
    old_comments: LegacyCommentSnapshot = before.legacy_comments
    new_comments: LegacyCommentSnapshot = after.legacy_comments
    if old_comments == new_comments:
        return [], []

    details: dict[str, object] = {
        "before": old_comments.to_dict(),
        "after": new_comments.to_dict(),
    }
    if old_comments.declaration_signature != new_comments.declaration_signature:
        details["legacy_comment_binding_changed"] = True
    if old_comments.definition_signature != new_comments.definition_signature:
        details["legacy_comment_definition_material_changed"] = True
    if old_comments.note_shape_signature != new_comments.note_shape_signature:
        details["legacy_note_vml_material_changed"] = True
    if old_comments.relationship_signature != new_comments.relationship_signature:
        details["legacy_note_relationships_changed"] = True
    if (
        old_comments.unrecognized_legacy_comment_count
        != new_comments.unrecognized_legacy_comment_count
        or (
            (
                old_comments.unrecognized_legacy_comment_count
                or new_comments.unrecognized_legacy_comment_count
            )
            and (
                old_comments.definition_signature != new_comments.definition_signature
                or old_comments.note_shape_signature
                != new_comments.note_shape_signature
            )
        )
    ):
        details["unrecognized_legacy_comment_metadata_changed"] = True
    change = Change(
        "legacy_comment_controls_changed",
        None,
        "high",
        details=details,
    )
    finding = Finding(
        "FF046",
        "high",
        (
            "Legacy Excel notes or threaded-comment placeholders changed; review "
            "text, author context, cell association, visibility, or layout may "
            "be altered outside cells."
        ),
        details=details,
    )
    return [change], [finding]


def _cell_hyperlink_controls_changed(
    before: WorkbookSnapshot,
    after: WorkbookSnapshot,
) -> tuple[list[Change], list[Finding]]:
    """Flag stored cell hyperlink changes that ordinary values cannot expose."""
    old_hyperlinks: CellHyperlinkSnapshot = before.cell_hyperlinks
    new_hyperlinks: CellHyperlinkSnapshot = after.cell_hyperlinks
    if old_hyperlinks == new_hyperlinks:
        return [], []

    details: dict[str, object] = {
        "before": old_hyperlinks.to_dict(),
        "after": new_hyperlinks.to_dict(),
    }
    if old_hyperlinks.declaration_signature != new_hyperlinks.declaration_signature:
        details["cell_hyperlink_binding_changed"] = True
    if old_hyperlinks.definition_signature != new_hyperlinks.definition_signature:
        details["cell_hyperlink_definition_material_changed"] = True
    if old_hyperlinks.relationship_signature != new_hyperlinks.relationship_signature:
        details["cell_hyperlink_relationships_changed"] = True
    if (
        old_hyperlinks.unrecognized_cell_hyperlink_count
        != new_hyperlinks.unrecognized_cell_hyperlink_count
        or (
            (
                old_hyperlinks.unrecognized_cell_hyperlink_count
                or new_hyperlinks.unrecognized_cell_hyperlink_count
            )
            and (
                old_hyperlinks.definition_signature
                != new_hyperlinks.definition_signature
                or old_hyperlinks.relationship_signature
                != new_hyperlinks.relationship_signature
            )
        )
    ):
        details["unrecognized_cell_hyperlink_metadata_changed"] = True
    change = Change(
        "cell_hyperlink_controls_changed",
        None,
        "high",
        details=details,
    )
    finding = Finding(
        "FF047",
        "high",
        (
            "Worksheet cell hyperlinks changed; a reviewer may be redirected or "
            "shown a different target outside the ordinary cell value."
        ),
        details=details,
    )
    return [change], [finding]


def _worksheet_sparkline_controls_changed(
    before: WorkbookSnapshot,
    after: WorkbookSnapshot,
) -> tuple[list[Change], list[Finding]]:
    """Flag stored sparkline changes that ordinary cell values cannot expose."""
    old_sparklines: WorksheetSparklineSnapshot = before.worksheet_sparklines
    new_sparklines: WorksheetSparklineSnapshot = after.worksheet_sparklines
    if old_sparklines == new_sparklines:
        return [], []

    details: dict[str, object] = {
        "before": old_sparklines.to_dict(),
        "after": new_sparklines.to_dict(),
    }
    if old_sparklines.binding_signature != new_sparklines.binding_signature:
        details["worksheet_sparkline_bindings_changed"] = True
    if old_sparklines.definition_signature != new_sparklines.definition_signature:
        details["worksheet_sparkline_definition_material_changed"] = True
    if (
        old_sparklines.unrecognized_worksheet_sparkline_count
        != new_sparklines.unrecognized_worksheet_sparkline_count
        or (
            (
                old_sparklines.unrecognized_worksheet_sparkline_count
                or new_sparklines.unrecognized_worksheet_sparkline_count
            )
            and old_sparklines.definition_signature
            != new_sparklines.definition_signature
        )
    ):
        details["unrecognized_worksheet_sparkline_metadata_changed"] = True
    change = Change(
        "worksheet_sparkline_controls_changed",
        None,
        "high",
        details=details,
    )
    finding = Finding(
        "FF048",
        "high",
        (
            "Worksheet sparklines changed; a compact visual trend summary may "
            "now show different source data, output cells, or rendering controls."
        ),
        details=details,
    )
    return [change], [finding]


def _xml_mapping_controls_changed(
    before: WorkbookSnapshot,
    after: WorkbookSnapshot,
) -> tuple[list[Change], list[Finding]]:
    """Flag XML-map changes that can redirect import/export data outside cells."""
    old_mappings: XmlMappingSnapshot = before.xml_mapping_controls
    new_mappings: XmlMappingSnapshot = after.xml_mapping_controls
    if old_mappings == new_mappings:
        return [], []

    details: dict[str, object] = {
        "before": old_mappings.to_dict(),
        "after": new_mappings.to_dict(),
    }
    if old_mappings.declaration_signature != new_mappings.declaration_signature:
        details["xml_mapping_declarations_changed"] = True
    if old_mappings.binding_signature != new_mappings.binding_signature:
        details["xml_mapping_bindings_changed"] = True
    if old_mappings.relationship_signature != new_mappings.relationship_signature:
        details["xml_mapping_relationships_changed"] = True
    if (
        old_mappings.unrecognized_xml_mapping_count
        != new_mappings.unrecognized_xml_mapping_count
        or (
            (
                old_mappings.unrecognized_xml_mapping_count
                or new_mappings.unrecognized_xml_mapping_count
            )
            and (
                old_mappings.declaration_signature
                != new_mappings.declaration_signature
                or old_mappings.binding_signature
                != new_mappings.binding_signature
                or old_mappings.relationship_signature
                != new_mappings.relationship_signature
            )
        )
    ):
        details["unrecognized_xml_mapping_metadata_changed"] = True
    change = Change(
        "xml_mapping_controls_changed",
        None,
        "high",
        details=details,
    )
    finding = Finding(
        "FF049",
        "high",
        (
            "XML-mapped workbook controls changed; a refresh or export may now "
            "route operational data through a different schema, binding, or target."
        ),
        details=details,
    )
    return [change], [finding]


def _digital_signature_controls_changed(
    before: WorkbookSnapshot,
    after: WorkbookSnapshot,
) -> tuple[list[Change], list[Finding]]:
    """Flag package- and VBA-signature envelope changes outside normal cells."""
    old_signatures: DigitalSignatureSnapshot = before.digital_signatures
    new_signatures: DigitalSignatureSnapshot = after.digital_signatures
    if old_signatures == new_signatures:
        return [], []

    details: dict[str, object] = {
        "before": old_signatures.to_dict(),
        "after": new_signatures.to_dict(),
    }
    if (
        old_signatures.package_signature_signature
        != new_signatures.package_signature_signature
    ):
        details["package_signature_material_changed"] = True
    if (
        old_signatures.vba_signature_payload_signature
        != new_signatures.vba_signature_payload_signature
    ):
        details["vba_project_signature_payload_changed"] = True
    if old_signatures.relationship_signature != new_signatures.relationship_signature:
        details["digital_signature_relationships_changed"] = True
    if (
        old_signatures.unrecognized_digital_signature_count
        != new_signatures.unrecognized_digital_signature_count
        or (
            (
                old_signatures.unrecognized_digital_signature_count
                or new_signatures.unrecognized_digital_signature_count
            )
            and (
                old_signatures.package_signature_signature
                != new_signatures.package_signature_signature
                or old_signatures.vba_signature_payload_signature
                != new_signatures.vba_signature_payload_signature
                or old_signatures.relationship_signature
                != new_signatures.relationship_signature
            )
        )
    ):
        details["unrecognized_digital_signature_metadata_changed"] = True
    change = Change(
        "digital_signature_controls_changed",
        None,
        "high",
        details=details,
    )
    finding = Finding(
        "FF050",
        "high",
        (
            "Digital-signature controls changed; package or VBA provenance and "
            "integrity-assurance metadata may have been added, removed, or altered."
        ),
        details=details,
    )
    return [change], [finding]


def _rich_data_controls_changed(
    before: WorkbookSnapshot,
    after: WorkbookSnapshot,
) -> tuple[list[Change], list[Finding]]:
    """Flag rich entity data and related relationship changes outside cells."""
    old_rich_data: RichDataSnapshot = before.rich_data
    new_rich_data: RichDataSnapshot = after.rich_data
    if old_rich_data == new_rich_data:
        return [], []

    details: dict[str, object] = {
        "before": old_rich_data.to_dict(),
        "after": new_rich_data.to_dict(),
    }
    if old_rich_data.definition_signature != new_rich_data.definition_signature:
        details["rich_data_definitions_changed"] = True
    if old_rich_data.value_signature != new_rich_data.value_signature:
        details["rich_data_values_changed"] = True
    if (
        old_rich_data.metadata_binding_signature
        != new_rich_data.metadata_binding_signature
    ):
        details["rich_data_metadata_bindings_changed"] = True
    if old_rich_data.relationship_signature != new_rich_data.relationship_signature:
        details["rich_data_relationships_changed"] = True
    if (
        old_rich_data.unrecognized_rich_data_count
        != new_rich_data.unrecognized_rich_data_count
        or (
            (
                old_rich_data.unrecognized_rich_data_count
                or new_rich_data.unrecognized_rich_data_count
            )
            and (
                old_rich_data.definition_signature
                != new_rich_data.definition_signature
                or old_rich_data.value_signature != new_rich_data.value_signature
                or (
                    old_rich_data.metadata_binding_signature
                    != new_rich_data.metadata_binding_signature
                )
                or (
                    old_rich_data.relationship_signature
                    != new_rich_data.relationship_signature
                )
            )
        )
    ):
        details["unrecognized_rich_data_metadata_changed"] = True
    change = Change(
        "rich_data_controls_changed",
        None,
        "high",
        details=details,
    )
    finding = Finding(
        "FF051",
        "high",
        (
            "Rich data controls changed; workbook entity values, provider-linked "
            "data, external image associations, or metadata bindings may have "
            "been added, removed, or altered."
        ),
        details=details,
    )
    return [change], [finding]


def _threaded_comment_controls_changed(
    before: WorkbookSnapshot,
    after: WorkbookSnapshot,
) -> tuple[list[Change], list[Finding]]:
    """Flag modern-comment changes that ordinary cell diffs cannot expose."""
    old_comments: ThreadedCommentSnapshot = before.threaded_comments
    new_comments: ThreadedCommentSnapshot = after.threaded_comments
    if old_comments == new_comments:
        return [], []

    details: dict[str, object] = {
        "before": old_comments.to_dict(),
        "after": new_comments.to_dict(),
    }
    if old_comments.declaration_signature != new_comments.declaration_signature:
        details["threaded_comment_binding_changed"] = True
    if old_comments.definition_signature != new_comments.definition_signature:
        details["threaded_comment_definition_material_changed"] = True
    if old_comments.person_signature != new_comments.person_signature:
        details["threaded_comment_person_material_changed"] = True
    if old_comments.relationship_signature != new_comments.relationship_signature:
        details["threaded_comment_relationships_changed"] = True
    if (
        old_comments.unrecognized_threaded_comment_count
        != new_comments.unrecognized_threaded_comment_count
        or (
            (
                old_comments.unrecognized_threaded_comment_count
                or new_comments.unrecognized_threaded_comment_count
            )
            and (
                old_comments.definition_signature != new_comments.definition_signature
                or old_comments.person_signature != new_comments.person_signature
            )
        )
    ):
        details["unrecognized_threaded_comment_metadata_changed"] = True
    change = Change(
        "threaded_comment_controls_changed",
        None,
        "high",
        details=details,
    )
    finding = Finding(
        "FF045",
        "high",
        (
            "Modern threaded comments changed; discussion text, replies, resolution, "
            "mentions, or collaborator bindings may be altered outside cells."
        ),
        details=details,
    )
    return [change], [finding]


def _worksheet_drawing_shape_controls_changed(
    before: WorkbookSnapshot,
    after: WorkbookSnapshot,
) -> tuple[list[Change], list[Finding]]:
    """Flag DrawingML shape changes that ordinary cell diffs cannot expose."""
    old_shapes: WorksheetDrawingShapeSnapshot = before.worksheet_drawing_shapes
    new_shapes: WorksheetDrawingShapeSnapshot = after.worksheet_drawing_shapes
    if old_shapes == new_shapes:
        return [], []

    details: dict[str, object] = {
        "before": old_shapes.to_dict(),
        "after": new_shapes.to_dict(),
    }
    if old_shapes.declaration_signature != new_shapes.declaration_signature:
        details["worksheet_drawing_shape_binding_changed"] = True
    if old_shapes.definition_signature != new_shapes.definition_signature:
        details["worksheet_drawing_shape_definition_material_changed"] = True
    if old_shapes.relationship_signature != new_shapes.relationship_signature:
        details["worksheet_drawing_shape_relationships_changed"] = True
    if (
        old_shapes.unrecognized_shape_count != new_shapes.unrecognized_shape_count
        or (
            (old_shapes.unrecognized_shape_count or new_shapes.unrecognized_shape_count)
            and old_shapes.definition_signature != new_shapes.definition_signature
        )
    ):
        details["unrecognized_worksheet_drawing_shape_metadata_changed"] = True
    change = Change(
        "worksheet_drawing_shape_controls_changed",
        None,
        "high",
        details=details,
    )
    finding = Finding(
        "FF044",
        "high",
        (
            "Worksheet DrawingML shape controls changed; text, visual cues, anchors, "
            "or linked actions may be altered outside cells."
        ),
        details=details,
    )
    return [change], [finding]


def compare_snapshots(before: WorkbookSnapshot, after: WorkbookSnapshot) -> DiffReport:
    """Compare workbook semantics and attach local dependency impact to each edit."""
    changes: list[Change] = []
    findings: list[Finding] = []
    formula_changed_locations: set[CellKey] = set()
    semantic_cell_changes: list[Change] = []

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
        change = Change(
            kind=kind,
            location=location,
            severity=severity,
            before=old_cell,
            after=new_cell,
            impact_count=len(impact),
            impacted_cells=sampled_impacts,
            details=details,
        )
        changes.append(change)
        semantic_cell_changes.append(change)
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

    formula_cached_result_changes, formula_cached_result_findings = (
        _formula_cached_result_changes(before, after, semantic_cell_changes)
    )
    changes.extend(formula_cached_result_changes)
    findings.extend(formula_cached_result_findings)

    rich_text_run_changes, rich_text_run_findings = _rich_text_run_controls_changed(
        before,
        after,
    )
    changes.extend(rich_text_run_changes)
    findings.extend(rich_text_run_findings)

    cell_hyperlink_changes, cell_hyperlink_findings = _cell_hyperlink_controls_changed(
        before,
        after,
    )
    changes.extend(cell_hyperlink_changes)
    findings.extend(cell_hyperlink_findings)

    worksheet_sparkline_changes, worksheet_sparkline_findings = (
        _worksheet_sparkline_controls_changed(before, after)
    )
    changes.extend(worksheet_sparkline_changes)
    findings.extend(worksheet_sparkline_findings)

    xml_mapping_changes, xml_mapping_findings = _xml_mapping_controls_changed(
        before,
        after,
    )
    changes.extend(xml_mapping_changes)
    findings.extend(xml_mapping_findings)

    digital_signature_changes, digital_signature_findings = (
        _digital_signature_controls_changed(before, after)
    )
    changes.extend(digital_signature_changes)
    findings.extend(digital_signature_findings)

    rich_data_changes, rich_data_findings = _rich_data_controls_changed(before, after)
    changes.extend(rich_data_changes)
    findings.extend(rich_data_findings)

    legacy_comment_changes, legacy_comment_findings = _legacy_comment_controls_changed(
        before,
        after,
    )
    changes.extend(legacy_comment_changes)
    findings.extend(legacy_comment_findings)

    threaded_comment_changes, threaded_comment_findings = _threaded_comment_controls_changed(
        before,
        after,
    )
    changes.extend(threaded_comment_changes)
    findings.extend(threaded_comment_findings)

    worksheet_drawing_shape_changes, worksheet_drawing_shape_findings = (
        _worksheet_drawing_shape_controls_changed(before, after)
    )
    changes.extend(worksheet_drawing_shape_changes)
    findings.extend(worksheet_drawing_shape_findings)

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
    dynamic_array_reference_changes, dynamic_array_reference_findings = (
        _dynamic_array_output_reference_changes(before, after)
    )
    changes.extend(dynamic_array_reference_changes)
    findings.extend(dynamic_array_reference_findings)
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
