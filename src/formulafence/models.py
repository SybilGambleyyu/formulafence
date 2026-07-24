"""Data structures shared by FormulaFence's parser, diff, and policy layers."""

from __future__ import annotations

from collections.abc import Iterable
from dataclasses import asdict, dataclass, field
from datetime import date, datetime, time
from pathlib import Path
from typing import Any, TypeAlias

CellKey: TypeAlias = tuple[str, str]

SEVERITY_ORDER = {"note": 0, "low": 1, "medium": 2, "high": 3, "critical": 4}


def display_location(location: CellKey | None) -> str | None:
    """Return an Excel-friendly cell location."""
    if location is None:
        return None
    sheet, coordinate = location
    escaped_sheet = sheet.replace("'", "''")
    if any(character in sheet for character in " !'[]"):
        escaped_sheet = f"'{escaped_sheet}'"
    return f"{escaped_sheet}!{coordinate}"


def json_safe_value(value: Any) -> Any:
    """Convert an openpyxl value to a deterministic JSON-compatible representation."""
    if value is None or isinstance(value, (str, int, float, bool)):
        return value
    if isinstance(value, (datetime, date, time)):
        return value.isoformat()
    return str(value)


@dataclass(frozen=True)
class CellSnapshot:
    """The semantic content FormulaFence compares for one non-empty cell."""

    sheet: str
    coordinate: str
    cell_type: str
    value: Any
    value_type: str
    formula: str | None = None
    formula_fingerprint: str | None = None
    array_formula_kind: str | None = None
    array_formula_ref: str | None = None

    @property
    def location(self) -> CellKey:
        return (self.sheet, self.coordinate)

    @property
    def is_formula(self) -> bool:
        return self.cell_type == "formula"

    def to_dict(self) -> dict[str, Any]:
        return {
            "location": display_location(self.location),
            "cell_type": self.cell_type,
            "value": self.value,
            "value_type": self.value_type,
            "formula": self.formula,
            "formula_fingerprint": self.formula_fingerprint,
            "array_formula_kind": self.array_formula_kind,
            "array_formula_ref": self.array_formula_ref,
        }


@dataclass(frozen=True)
class SheetSnapshot:
    title: str
    state: str
    nonempty_cells: int
    formula_cells: int
    max_row: int
    max_column: int

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)


@dataclass(frozen=True)
class TableSnapshot:
    """Inspectable Excel-table metadata that can change formula semantics."""

    name: str
    sheet: str
    ref: str
    columns: tuple[str, ...]
    header_row_count: int
    totals_row_count: int

    def to_dict(self) -> dict[str, Any]:
        return {
            "name": self.name,
            "sheet": self.sheet,
            "ref": self.ref,
            "columns": list(self.columns),
            "header_row_count": self.header_row_count,
            "totals_row_count": self.totals_row_count,
        }


@dataclass(frozen=True)
class DataValidationSnapshot:
    """One compact Excel data-validation control and its effective settings."""

    sheet: str
    ranges: tuple[str, ...]
    validation_type: str
    operator: str
    formula1: str | None
    formula2: str | None
    allow_blank: bool
    dropdown_hidden: bool
    prompts_disabled: bool
    show_input_message: bool
    show_error_message: bool
    error_style: str
    error_title: str | None
    error: str | None
    prompt_title: str | None
    prompt: str | None
    ime_mode: str

    @property
    def target_range_count(self) -> int:
        return len(self.ranges)

    @property
    def criteria_count(self) -> int:
        return sum(value is not None for value in (self.formula1, self.formula2))

    def sort_key(self) -> tuple[object, ...]:
        """Return a writer-order-independent key for deterministic inventories."""
        def optional(value: str | None) -> tuple[bool, str]:
            return value is not None, value or ""

        return (
            self.sheet.casefold(),
            self.ranges,
            self.validation_type,
            self.operator,
            optional(self.formula1),
            optional(self.formula2),
            self.allow_blank,
            self.dropdown_hidden,
            self.prompts_disabled,
            self.show_input_message,
            self.show_error_message,
            self.error_style,
            optional(self.error_title),
            optional(self.error),
            optional(self.prompt_title),
            optional(self.prompt),
            self.ime_mode,
        )

    def to_dict(self) -> dict[str, Any]:
        """Return full local review evidence, including criterion expressions."""
        return {
            "sheet": self.sheet,
            "ranges": [
                display_location((self.sheet, target_range))
                for target_range in self.ranges
            ],
            "type": self.validation_type,
            "operator": self.operator,
            "formula1": self.formula1,
            "formula2": self.formula2,
            "allow_blank": self.allow_blank,
            "dropdown_hidden": self.dropdown_hidden,
            "prompts_disabled": self.prompts_disabled,
            "show_input_message": self.show_input_message,
            "show_error_message": self.show_error_message,
            "error_style": self.error_style,
            "error_title": self.error_title,
            "error": self.error,
            "prompt_title": self.prompt_title,
            "prompt": self.prompt,
            "ime_mode": self.ime_mode,
        }

    def profile_dict(self) -> dict[str, Any]:
        """Return a data-minimising control inventory without messages or formulas."""
        return {
            "sheet": self.sheet,
            "ranges": [
                display_location((self.sheet, target_range))
                for target_range in self.ranges
            ],
            "type": self.validation_type,
            "operator": self.operator,
            "criteria_count": self.criteria_count,
            "allow_blank": self.allow_blank,
            "dropdown_hidden": self.dropdown_hidden,
            "prompts_disabled": self.prompts_disabled,
            "show_input_message": self.show_input_message,
            "show_error_message": self.show_error_message,
            "error_style": self.error_style,
            "has_error_alert_text": bool(self.error_title or self.error),
            "has_input_prompt_text": bool(self.prompt_title or self.prompt),
            "ime_mode": self.ime_mode,
        }


@dataclass(frozen=True)
class XmlFragmentSnapshot:
    """A deterministic, inspectable OOXML fragment used for control evidence."""

    tag: str
    attributes: tuple[tuple[str, str], ...] = ()
    text: str | None = None
    children: tuple[XmlFragmentSnapshot, ...] = ()

    def sort_key(self) -> tuple[object, ...]:
        return (
            self.tag,
            self.attributes,
            self.text is not None,
            self.text or "",
            tuple(child.sort_key() for child in self.children),
        )

    def to_dict(self) -> dict[str, Any]:
        result: dict[str, Any] = {"element": self.tag}
        if self.attributes:
            result["attributes"] = dict(self.attributes)
        if self.text is not None:
            result["text"] = self.text
        if self.children:
            result["children"] = [child.to_dict() for child in self.children]
        return result


@dataclass(frozen=True)
class ConditionalFormattingSnapshot:
    """One worksheet conditional-formatting rule and its effective semantics."""

    sheet: str
    ranges: tuple[str, ...]
    priority: int
    rule_type: str
    operator: str | None
    formulas: tuple[str, ...]
    stop_if_true: bool
    above_average: bool
    percent: bool
    bottom: bool
    rank: int | None
    std_dev: int | None
    equal_average: bool
    text: str | None
    time_period: str | None
    differential_style: XmlFragmentSnapshot | None
    color_scale: XmlFragmentSnapshot | None
    data_bar: XmlFragmentSnapshot | None
    icon_set: XmlFragmentSnapshot | None
    extensions: tuple[XmlFragmentSnapshot, ...] = ()

    @property
    def target_range_count(self) -> int:
        return len(self.ranges)

    @property
    def formula_count(self) -> int:
        return len(self.formulas)

    @property
    def formatting_kinds(self) -> tuple[str, ...]:
        kinds: list[str] = []
        if self.differential_style is not None:
            kinds.append("differential style")
        if self.color_scale is not None:
            kinds.append("color scale")
        if self.data_bar is not None:
            kinds.append("data bar")
        if self.icon_set is not None:
            kinds.append("icon set")
        return tuple(kinds)

    def sort_key(self) -> tuple[object, ...]:
        return (
            self.sheet.casefold(),
            self.priority,
            self.ranges,
            self.rule_type,
            self.operator or "",
            self.formulas,
            self.stop_if_true,
            self.above_average,
            self.percent,
            self.bottom,
            self.rank is not None,
            self.rank or 0,
            self.std_dev is not None,
            self.std_dev or 0,
            self.equal_average,
            self.text is not None,
            self.text or "",
            self.time_period or "",
            self.differential_style is not None,
            self.differential_style.sort_key() if self.differential_style is not None else (),
            self.color_scale is not None,
            self.color_scale.sort_key() if self.color_scale is not None else (),
            self.data_bar is not None,
            self.data_bar.sort_key() if self.data_bar is not None else (),
            self.icon_set is not None,
            self.icon_set.sort_key() if self.icon_set is not None else (),
            tuple(extension.sort_key() for extension in self.extensions),
        )

    def to_dict(self) -> dict[str, Any]:
        """Return full local evidence, including formulas and formatting fragments."""
        return {
            "sheet": self.sheet,
            "ranges": [
                display_location((self.sheet, target_range))
                for target_range in self.ranges
            ],
            "priority": self.priority,
            "type": self.rule_type,
            "operator": self.operator,
            "formulas": list(self.formulas),
            "stop_if_true": self.stop_if_true,
            "above_average": self.above_average,
            "percent": self.percent,
            "bottom": self.bottom,
            "rank": self.rank,
            "std_dev": self.std_dev,
            "equal_average": self.equal_average,
            "text": self.text,
            "time_period": self.time_period,
            "differential_style": (
                self.differential_style.to_dict()
                if self.differential_style is not None
                else None
            ),
            "color_scale": self.color_scale.to_dict() if self.color_scale is not None else None,
            "data_bar": self.data_bar.to_dict() if self.data_bar is not None else None,
            "icon_set": self.icon_set.to_dict() if self.icon_set is not None else None,
            "extensions": [extension.to_dict() for extension in self.extensions],
        }

    def profile_dict(self) -> dict[str, Any]:
        """Return reviewable control metadata without formulas or text criteria."""
        return {
            "sheet": self.sheet,
            "ranges": [
                display_location((self.sheet, target_range))
                for target_range in self.ranges
            ],
            "priority": self.priority,
            "type": self.rule_type,
            "operator": self.operator,
            "formula_count": self.formula_count,
            "has_text_criterion": self.text is not None,
            "stop_if_true": self.stop_if_true,
            "above_average": self.above_average,
            "percent": self.percent,
            "bottom": self.bottom,
            "rank": self.rank,
            "std_dev": self.std_dev,
            "equal_average": self.equal_average,
            "time_period": self.time_period,
            "formatting": list(self.formatting_kinds),
            "extension_count": len(self.extensions),
        }


@dataclass(frozen=True)
class ConditionalFormattingExtensionSnapshot:
    """An OOXML conditional-formatting extension openpyxl may not model."""

    sheet: str
    fragment: XmlFragmentSnapshot

    def sort_key(self) -> tuple[object, ...]:
        return self.sheet.casefold(), self.fragment.sort_key()

    def to_dict(self) -> dict[str, Any]:
        return {"sheet": self.sheet, "extension": self.fragment.to_dict()}

    def profile_dict(self) -> dict[str, Any]:
        return {"sheet": self.sheet, "element": self.fragment.tag}


@dataclass(frozen=True)
class ProtectionCredentialSnapshot:
    """Non-secret evidence that a protection credential is configured.

    OOXML stores legacy password verifiers and, in newer workbooks, password
    hashes and salts alongside the protection settings.  FormulaFence uses a
    private signature to compare that material but never serialises it into a
    profile, change report, or finding.
    """

    has_legacy_verifier: bool = False
    has_modern_verifier: bool = False
    algorithm: str | None = None
    spin_count: int | None = None
    signature: str | None = field(default=None, repr=False)

    @property
    def configured(self) -> bool:
        """Return whether any verifier material was present."""
        return self.has_legacy_verifier or self.has_modern_verifier

    def to_dict(self) -> dict[str, Any]:
        """Return safe verifier metadata without password material."""
        return {
            "configured": self.configured,
            "has_legacy_verifier": self.has_legacy_verifier,
            "has_modern_verifier": self.has_modern_verifier,
            "algorithm": self.algorithm,
            "spin_count": self.spin_count,
        }


@dataclass(frozen=True)
class ProtectionOpaqueMetadataSnapshot:
    """Private comparison evidence for unmodelled protection metadata."""

    count: int = 0
    signature: str | None = field(default=None, repr=False)

    @property
    def present(self) -> bool:
        return self.count > 0

    def to_dict(self) -> dict[str, Any]:
        """Return only the existence and amount of opaque metadata."""
        return {
            "present": self.present,
            "count": self.count,
        }


@dataclass(frozen=True)
class WorkbookProtectionSnapshot:
    """Workbook-structure and revision protection controls."""

    lock_structure: bool
    lock_windows: bool
    lock_revision: bool
    workbook_credential: ProtectionCredentialSnapshot = field(
        default_factory=ProtectionCredentialSnapshot
    )
    revisions_credential: ProtectionCredentialSnapshot = field(
        default_factory=ProtectionCredentialSnapshot
    )
    opaque_metadata: ProtectionOpaqueMetadataSnapshot = field(
        default_factory=ProtectionOpaqueMetadataSnapshot
    )

    @property
    def enabled(self) -> bool:
        """Return whether any workbook-level operation is locked."""
        return self.lock_structure or self.lock_windows or self.lock_revision

    def to_dict(self) -> dict[str, Any]:
        """Return safe, reviewable control settings."""
        return {
            "enabled": self.enabled,
            "lock_structure": self.lock_structure,
            "lock_windows": self.lock_windows,
            "lock_revision": self.lock_revision,
            "workbook_credential": self.workbook_credential.to_dict(),
            "revisions_credential": self.revisions_credential.to_dict(),
            "opaque_metadata": self.opaque_metadata.to_dict(),
        }


@dataclass(frozen=True)
class SheetProtectionSnapshot:
    """One worksheet, dialog-sheet, or chart-sheet protection declaration."""

    sheet: str
    sheet_type: str
    enabled: bool
    locked_actions: tuple[str, ...]
    credential: ProtectionCredentialSnapshot = field(
        default_factory=ProtectionCredentialSnapshot
    )
    opaque_metadata: ProtectionOpaqueMetadataSnapshot = field(
        default_factory=ProtectionOpaqueMetadataSnapshot
    )

    def sort_key(self) -> tuple[object, ...]:
        return self.sheet.casefold(), self.sheet_type, self.locked_actions

    def to_dict(self) -> dict[str, Any]:
        """Return safe, reviewable sheet-protection settings."""
        return {
            "sheet": self.sheet,
            "sheet_type": self.sheet_type,
            "enabled": self.enabled,
            "locked_actions": list(self.locked_actions),
            "credential": self.credential.to_dict(),
            "opaque_metadata": self.opaque_metadata.to_dict(),
        }

    def profile_dict(self) -> dict[str, Any]:
        return self.to_dict()


@dataclass(frozen=True)
class ProtectedRangeSnapshot:
    """A protected worksheet range without its password or identity material."""

    sheet: str
    ranges: tuple[str, ...]
    has_name: bool
    name_signature: str | None = field(default=None, repr=False)
    credential: ProtectionCredentialSnapshot = field(
        default_factory=ProtectionCredentialSnapshot
    )
    has_security_descriptor: bool = False
    security_descriptor_signature: str | None = field(default=None, repr=False)
    opaque_metadata: ProtectionOpaqueMetadataSnapshot = field(
        default_factory=ProtectionOpaqueMetadataSnapshot
    )

    @property
    def target_range_count(self) -> int:
        return len(self.ranges)

    def sort_key(self) -> tuple[object, ...]:
        return (
            self.sheet.casefold(),
            self.ranges,
            self.has_name,
            self.name_signature or "",
            self.credential.configured,
            self.has_security_descriptor,
        )

    def to_dict(self) -> dict[str, Any]:
        """Return safe range permissions without names or credential material."""
        return {
            "sheet": self.sheet,
            "ranges": [
                display_location((self.sheet, target_range))
                for target_range in self.ranges
            ],
            "has_name": self.has_name,
            "credential": self.credential.to_dict(),
            "has_security_descriptor": self.has_security_descriptor,
            "opaque_metadata": self.opaque_metadata.to_dict(),
        }

    def profile_dict(self) -> dict[str, Any]:
        return self.to_dict()


@dataclass(frozen=True)
class CellProtectionDefaultSnapshot:
    """Effective locked/hidden defaults from the workbook's base cell style."""

    locked: bool
    hidden: bool

    def to_dict(self) -> dict[str, bool]:
        return {"locked": self.locked, "hidden": self.hidden}


@dataclass(frozen=True)
class CellProtectionAssignmentSnapshot:
    """One direct cell, row, or column protection-style assignment."""

    sheet: str
    scope: str
    target: str
    locked: bool
    hidden: bool

    def sort_key(self) -> tuple[object, ...]:
        return self.sheet.casefold(), self.scope, self.target.casefold()

    def to_dict(self) -> dict[str, Any]:
        return {
            "sheet": self.sheet,
            "scope": self.scope,
            "target": self.target,
            "locked": self.locked,
            "hidden": self.hidden,
        }

    def profile_dict(self) -> dict[str, Any]:
        return self.to_dict()


@dataclass(frozen=True)
class ExternalDataOpaqueMetadataSnapshot:
    """Private comparison evidence for unmodelled external-data XML or mashups."""

    count: int = 0
    signature: str | None = field(default=None, repr=False)

    @property
    def present(self) -> bool:
        return self.count > 0

    def to_dict(self) -> dict[str, Any]:
        """Return only the existence and amount of opaque metadata."""
        return {
            "present": self.present,
            "count": self.count,
        }


@dataclass(frozen=True)
class ExternalDataRefreshSettingsSnapshot:
    """Workbook-wide controls that can cause external data to refresh."""

    update_links: str = "user_set"
    allow_refresh_query: bool = False
    refresh_all_connections: bool = False
    save_external_link_values: bool = True

    def to_dict(self) -> dict[str, Any]:
        """Return safe, reviewable workbook-wide external-data settings."""
        return {
            "update_links": self.update_links,
            "allow_refresh_query": self.allow_refresh_query,
            "refresh_all_connections": self.refresh_all_connections,
            "save_external_link_values": self.save_external_link_values,
        }


@dataclass(frozen=True)
class ExternalDataConnectionSnapshot:
    """One OOXML external-data connection without its source material."""

    connection_id: int | None
    source_type: str
    deleted: bool = False
    refresh_on_load: bool = False
    refresh_interval_minutes: int | None = None
    background: bool = False
    keep_alive: bool = False
    save_data: bool = False
    save_password: bool = False
    has_source_file: bool = False
    has_connection_file: bool = False
    only_use_connection_file: bool = False
    reconnection_method: str = "as_required"
    credential_method: str = "integrated"
    minimum_refreshable_version: int = 0
    has_single_sign_on_id: bool = False
    awaiting_initial_refresh: bool = False
    has_name: bool = False
    has_description: bool = False
    source_components: tuple[str, ...] = ()
    parameter_count: int = 0
    parameters_refresh_on_change: int = 0
    identity_signature: str | None = field(default=None, repr=False)
    source_configuration_signature: str | None = field(default=None, repr=False)
    opaque_metadata: ExternalDataOpaqueMetadataSnapshot = field(
        default_factory=ExternalDataOpaqueMetadataSnapshot
    )

    def sort_key(self) -> tuple[object, ...]:
        return (
            self.connection_id is None,
            self.connection_id if self.connection_id is not None else -1,
            self.identity_signature or "",
            self.source_configuration_signature or "",
        )

    def to_dict(self) -> dict[str, Any]:
        """Return safe connection controls without source strings or credentials."""
        return {
            "id": self.connection_id,
            "source_type": self.source_type,
            "deleted": self.deleted,
            "refresh_on_load": self.refresh_on_load,
            "refresh_interval_minutes": self.refresh_interval_minutes,
            "background": self.background,
            "keep_alive": self.keep_alive,
            "save_data": self.save_data,
            "save_password": self.save_password,
            "has_source_file": self.has_source_file,
            "has_connection_file": self.has_connection_file,
            "only_use_connection_file": self.only_use_connection_file,
            "reconnection_method": self.reconnection_method,
            "credential_method": self.credential_method,
            "minimum_refreshable_version": self.minimum_refreshable_version,
            "has_single_sign_on_id": self.has_single_sign_on_id,
            "awaiting_initial_refresh": self.awaiting_initial_refresh,
            "has_name": self.has_name,
            "has_description": self.has_description,
            "source_components": list(self.source_components),
            "parameter_count": self.parameter_count,
            "parameters_refresh_on_change": self.parameters_refresh_on_change,
            "opaque_metadata": self.opaque_metadata.to_dict(),
        }

    def profile_dict(self) -> dict[str, Any]:
        return self.to_dict()


@dataclass(frozen=True)
class QueryTableRefreshSnapshot:
    """One query-table refresh control set, tied to its worksheet safely."""

    sheet: str
    connection_id: int | None
    refresh_on_load: bool = False
    background_refresh: bool = True
    refresh_disabled: bool = False
    remove_data_on_save: bool = False
    fill_formulas: bool = False
    connection_edit_disabled: bool = False
    growth_behavior: str = "insert_delete"
    has_name: bool = False
    has_refresh_metadata: bool = False
    identity_signature: str | None = field(default=None, repr=False)
    opaque_metadata: ExternalDataOpaqueMetadataSnapshot = field(
        default_factory=ExternalDataOpaqueMetadataSnapshot
    )

    def sort_key(self) -> tuple[object, ...]:
        return (
            self.sheet.casefold(),
            self.connection_id is None,
            self.connection_id if self.connection_id is not None else -1,
            self.identity_signature or "",
        )

    def to_dict(self) -> dict[str, Any]:
        """Return safe query-table behavior without query names or data."""
        return {
            "sheet": self.sheet,
            "connection_id": self.connection_id,
            "refresh_on_load": self.refresh_on_load,
            "background_refresh": self.background_refresh,
            "refresh_disabled": self.refresh_disabled,
            "remove_data_on_save": self.remove_data_on_save,
            "fill_formulas": self.fill_formulas,
            "connection_edit_disabled": self.connection_edit_disabled,
            "growth_behavior": self.growth_behavior,
            "has_name": self.has_name,
            "has_refresh_metadata": self.has_refresh_metadata,
            "opaque_metadata": self.opaque_metadata.to_dict(),
        }

    def profile_dict(self) -> dict[str, Any]:
        return self.to_dict()


@dataclass(frozen=True)
class PivotCacheRefreshSnapshot:
    """One pivot-cache source and refresh control set without cached data."""

    # OOXML cache IDs are writer-assigned graph handles. Preserve the safe value
    # in a standalone profile, but do not make a harmless renumbering a source
    # or refresh-control change; the PivotTable package scanner compares the
    # normalized cache binding separately.
    cache_id: int | None = field(compare=False)
    source_type: str
    connection_id: int | None = None
    refresh_on_load: bool = False
    background_query: bool = False
    refresh_enabled: bool = True
    save_data: bool = True
    upgrade_on_refresh: bool = False
    source_configuration_signature: str | None = field(default=None, repr=False)
    opaque_metadata: ExternalDataOpaqueMetadataSnapshot = field(
        default_factory=ExternalDataOpaqueMetadataSnapshot
    )

    def sort_key(self) -> tuple[object, ...]:
        return (
            self.source_configuration_signature or "",
            self.source_type,
            self.connection_id is None,
            self.connection_id if self.connection_id is not None else -1,
            self.refresh_on_load,
            self.background_query,
            self.refresh_enabled,
            self.save_data,
            self.upgrade_on_refresh,
            self.opaque_metadata.signature or "",
        )

    def to_dict(self) -> dict[str, Any]:
        """Return safe pivot-cache controls without source or cache contents."""
        return {
            "cache_id": self.cache_id,
            "source_type": self.source_type,
            "connection_id": self.connection_id,
            "refresh_on_load": self.refresh_on_load,
            "background_query": self.background_query,
            "refresh_enabled": self.refresh_enabled,
            "save_data": self.save_data,
            "upgrade_on_refresh": self.upgrade_on_refresh,
            "opaque_metadata": self.opaque_metadata.to_dict(),
        }

    def profile_dict(self) -> dict[str, Any]:
        return self.to_dict()


@dataclass(frozen=True)
class ExternalLinkPackageSnapshot:
    """Safe aggregate of raw OOXML external-workbook, DDE, and OLE links."""

    external_link_count: int = 0
    external_workbook_count: int = 0
    dde_link_count: int = 0
    ole_link_count: int = 0
    unrecognized_link_count: int = 0
    external_workbook_sheet_count: int = 0
    external_defined_name_count: int = 0
    external_workbook_cached_sheet_count: int = 0
    external_workbook_cached_cell_count: int = 0
    external_workbook_cached_refresh_error_count: int = 0
    dde_item_count: int = 0
    dde_advise_item_count: int = 0
    dde_ole_item_count: int = 0
    dde_prefer_picture_item_count: int = 0
    dde_cached_value_count: int = 0
    ole_item_count: int = 0
    ole_advise_item_count: int = 0
    ole_icon_item_count: int = 0
    ole_prefer_picture_item_count: int = 0
    source_signature: str | None = field(default=None, repr=False)
    definition_signature: str | None = field(default=None, repr=False)
    cached_material_signature: str | None = field(default=None, repr=False)
    opaque_metadata: ExternalDataOpaqueMetadataSnapshot = field(
        default_factory=ExternalDataOpaqueMetadataSnapshot
    )

    @property
    def present(self) -> bool:
        return self.external_link_count > 0

    def to_dict(self) -> dict[str, Any]:
        """Return structural link evidence without targets, names, or cached values."""
        return {
            "present": self.present,
            "external_link_count": self.external_link_count,
            "external_workbook_count": self.external_workbook_count,
            "dde_link_count": self.dde_link_count,
            "ole_link_count": self.ole_link_count,
            "unrecognized_link_count": self.unrecognized_link_count,
            "external_workbook_sheet_count": self.external_workbook_sheet_count,
            "external_defined_name_count": self.external_defined_name_count,
            "external_workbook_cached_sheet_count": self.external_workbook_cached_sheet_count,
            "external_workbook_cached_cell_count": self.external_workbook_cached_cell_count,
            "external_workbook_cached_refresh_error_count": (
                self.external_workbook_cached_refresh_error_count
            ),
            "dde_item_count": self.dde_item_count,
            "dde_advise_item_count": self.dde_advise_item_count,
            "dde_ole_item_count": self.dde_ole_item_count,
            "dde_prefer_picture_item_count": self.dde_prefer_picture_item_count,
            "dde_cached_value_count": self.dde_cached_value_count,
            "ole_item_count": self.ole_item_count,
            "ole_advise_item_count": self.ole_advise_item_count,
            "ole_icon_item_count": self.ole_icon_item_count,
            "ole_prefer_picture_item_count": self.ole_prefer_picture_item_count,
            "opaque_metadata": self.opaque_metadata.to_dict(),
        }

    def profile_dict(self) -> dict[str, Any]:
        return self.to_dict()


@dataclass(frozen=True)
class XlmMacroSheetSnapshot:
    """Safe aggregate of raw Excel 4.0 / XLM macro-sheet package parts.

    XLM automation lives in XML macro-sheet parts rather than in the VBA
    binary. The private signatures retain complete program, relationship, and
    direct internal related-part evidence for comparison, while ``to_dict``
    intentionally exposes only structural counts.
    """

    declared_macro_sheet_count: int = 0
    macro_sheet_count: int = 0
    international_macro_sheet_count: int = 0
    unrecognized_macro_sheet_count: int = 0
    hidden_macro_sheet_count: int = 0
    very_hidden_macro_sheet_count: int = 0
    formula_cell_count: int = 0
    related_relationship_count: int = 0
    external_relationship_count: int = 0
    internal_related_part_count: int = 0
    fingerprinted_related_part_count: int = 0
    uninspected_related_part_count: int = 0
    embedded_object_relationship_count: int = 0
    embedded_package_relationship_count: int = 0
    declaration_signature: str | None = field(default=None, repr=False)
    program_signature: str | None = field(default=None, repr=False)
    relationship_signature: str | None = field(default=None, repr=False)
    related_part_payload_signature: str | None = field(default=None, repr=False)

    @property
    def present(self) -> bool:
        return bool(self.declared_macro_sheet_count or self.macro_sheet_count)

    def to_dict(self) -> dict[str, Any]:
        """Return safe macro-sheet inventory without commands or endpoints."""
        return {
            "present": self.present,
            "declared_macro_sheet_count": self.declared_macro_sheet_count,
            "macro_sheet_count": self.macro_sheet_count,
            "international_macro_sheet_count": self.international_macro_sheet_count,
            "unrecognized_macro_sheet_count": self.unrecognized_macro_sheet_count,
            "hidden_macro_sheet_count": self.hidden_macro_sheet_count,
            "very_hidden_macro_sheet_count": self.very_hidden_macro_sheet_count,
            "formula_cell_count": self.formula_cell_count,
            "related_relationship_count": self.related_relationship_count,
            "external_relationship_count": self.external_relationship_count,
            "internal_related_part_count": self.internal_related_part_count,
            "fingerprinted_related_part_count": self.fingerprinted_related_part_count,
            "uninspected_related_part_count": self.uninspected_related_part_count,
            "embedded_object_relationship_count": (
                self.embedded_object_relationship_count
            ),
            "embedded_package_relationship_count": (
                self.embedded_package_relationship_count
            ),
        }

    def profile_dict(self) -> dict[str, Any]:
        return self.to_dict()


@dataclass(frozen=True)
class RibbonCustomizationSnapshot:
    """Safe aggregate of workbook-scoped Office RibbonX customization parts.

    Ribbon XML can bind controls to workbook callbacks while remaining outside
    the VBA project. Private signatures retain complete package and callback
    material for comparison, while ``to_dict`` intentionally exposes only
    structural counts.
    """

    declared_ribbon_part_count: int = 0
    ribbon_part_count: int = 0
    office_2010_ribbon_part_count: int = 0
    unrecognized_ribbon_part_count: int = 0
    control_count: int = 0
    callback_attribute_count: int = 0
    action_callback_count: int = 0
    image_relationship_count: int = 0
    external_relationship_count: int = 0
    declaration_signature: str | None = field(default=None, repr=False)
    definition_signature: str | None = field(default=None, repr=False)
    relationship_signature: str | None = field(default=None, repr=False)

    @property
    def present(self) -> bool:
        return bool(self.declared_ribbon_part_count or self.ribbon_part_count)

    def to_dict(self) -> dict[str, Any]:
        """Return safe RibbonX inventory without names, labels, or callbacks."""
        return {
            "present": self.present,
            "declared_ribbon_part_count": self.declared_ribbon_part_count,
            "ribbon_part_count": self.ribbon_part_count,
            "office_2010_ribbon_part_count": self.office_2010_ribbon_part_count,
            "unrecognized_ribbon_part_count": self.unrecognized_ribbon_part_count,
            "control_count": self.control_count,
            "callback_attribute_count": self.callback_attribute_count,
            "action_callback_count": self.action_callback_count,
            "image_relationship_count": self.image_relationship_count,
            "external_relationship_count": self.external_relationship_count,
        }

    def profile_dict(self) -> dict[str, Any]:
        return self.to_dict()


@dataclass(frozen=True)
class OfficeWebAddinSnapshot:
    """Safe aggregate of document-linked Office Web Add-in task panes.

    OOXML task-pane parts can bind a workbook to an installed Office Web
    Add-in and request auto-show behavior. Private signatures retain the
    add-in reference, property, binding, and relationship material needed for
    comparison; ``to_dict`` intentionally exposes only structural counts.
    """

    declared_taskpane_part_count: int = 0
    taskpane_part_count: int = 0
    web_extension_part_count: int = 0
    unrecognized_part_count: int = 0
    taskpane_count: int = 0
    visible_taskpane_count: int = 0
    locked_taskpane_count: int = 0
    web_extension_reference_count: int = 0
    auto_show_taskpane_count: int = 0
    store_reference_count: int = 0
    alternate_reference_count: int = 0
    binding_count: int = 0
    snapshot_reference_count: int = 0
    related_relationship_count: int = 0
    external_relationship_count: int = 0
    declaration_signature: str | None = field(default=None, repr=False)
    taskpane_signature: str | None = field(default=None, repr=False)
    web_extension_signature: str | None = field(default=None, repr=False)
    relationship_signature: str | None = field(default=None, repr=False)

    @property
    def present(self) -> bool:
        return bool(
            self.declared_taskpane_part_count
            or self.taskpane_part_count
            or self.web_extension_part_count
        )

    def to_dict(self) -> dict[str, Any]:
        """Return task-pane inventory without add-in identities or endpoints."""
        return {
            "present": self.present,
            "declared_taskpane_part_count": self.declared_taskpane_part_count,
            "taskpane_part_count": self.taskpane_part_count,
            "web_extension_part_count": self.web_extension_part_count,
            "unrecognized_part_count": self.unrecognized_part_count,
            "taskpane_count": self.taskpane_count,
            "visible_taskpane_count": self.visible_taskpane_count,
            "locked_taskpane_count": self.locked_taskpane_count,
            "web_extension_reference_count": self.web_extension_reference_count,
            "auto_show_taskpane_count": self.auto_show_taskpane_count,
            "store_reference_count": self.store_reference_count,
            "alternate_reference_count": self.alternate_reference_count,
            "binding_count": self.binding_count,
            "snapshot_reference_count": self.snapshot_reference_count,
            "related_relationship_count": self.related_relationship_count,
            "external_relationship_count": self.external_relationship_count,
        }

    def profile_dict(self) -> dict[str, Any]:
        return self.to_dict()


@dataclass(frozen=True)
class ChartDefinitionSnapshot:
    """Safe aggregate of DrawingML chart definitions and presentation material.

    Excel chart parts sit behind worksheet or chartsheet drawing relationships,
    outside the ordinary cell grid. Private signatures retain chart formulas,
    formatting, cached points, overlay-shape definitions, relationships, and
    bounded direct payload evidence for comparison; ``to_dict`` deliberately
    exposes only structural counts.
    """

    chart_host_sheet_count: int = 0
    chart_drawing_part_count: int = 0
    chart_reference_count: int = 0
    chart_part_count: int = 0
    chart_user_shape_part_count: int = 0
    chart_user_shape_count: int = 0
    chart_type_count: int = 0
    series_count: int = 0
    title_count: int = 0
    data_reference_count: int = 0
    numeric_data_reference_count: int = 0
    string_data_reference_count: int = 0
    literal_data_point_count: int = 0
    cached_data_point_count: int = 0
    pivot_source_count: int = 0
    external_data_reference_count: int = 0
    user_shape_reference_count: int = 0
    related_relationship_count: int = 0
    external_relationship_count: int = 0
    internal_related_part_count: int = 0
    fingerprinted_related_part_count: int = 0
    uninspected_related_part_count: int = 0
    unrecognized_part_count: int = 0
    declaration_signature: str | None = field(default=None, repr=False)
    definition_signature: str | None = field(default=None, repr=False)
    cached_data_signature: str | None = field(default=None, repr=False)
    user_shape_signature: str | None = field(default=None, repr=False)
    relationship_signature: str | None = field(default=None, repr=False)
    related_part_payload_signature: str | None = field(default=None, repr=False)

    @property
    def present(self) -> bool:
        return bool(
            self.chart_host_sheet_count
            or self.chart_drawing_part_count
            or self.chart_reference_count
            or self.chart_part_count
            or self.chart_user_shape_part_count
            or self.unrecognized_part_count
        )

    def to_dict(self) -> dict[str, Any]:
        """Return chart evidence without labels, formulas, values, or targets."""
        return {
            "present": self.present,
            "chart_host_sheet_count": self.chart_host_sheet_count,
            "chart_drawing_part_count": self.chart_drawing_part_count,
            "chart_reference_count": self.chart_reference_count,
            "chart_part_count": self.chart_part_count,
            "chart_user_shape_part_count": self.chart_user_shape_part_count,
            "chart_user_shape_count": self.chart_user_shape_count,
            "chart_type_count": self.chart_type_count,
            "series_count": self.series_count,
            "title_count": self.title_count,
            "data_reference_count": self.data_reference_count,
            "numeric_data_reference_count": self.numeric_data_reference_count,
            "string_data_reference_count": self.string_data_reference_count,
            "literal_data_point_count": self.literal_data_point_count,
            "cached_data_point_count": self.cached_data_point_count,
            "pivot_source_count": self.pivot_source_count,
            "external_data_reference_count": self.external_data_reference_count,
            "user_shape_reference_count": self.user_shape_reference_count,
            "related_relationship_count": self.related_relationship_count,
            "external_relationship_count": self.external_relationship_count,
            "internal_related_part_count": self.internal_related_part_count,
            "fingerprinted_related_part_count": self.fingerprinted_related_part_count,
            "uninspected_related_part_count": self.uninspected_related_part_count,
            "unrecognized_part_count": self.unrecognized_part_count,
        }

    def profile_dict(self) -> dict[str, Any]:
        return self.to_dict()


@dataclass(frozen=True)
class PivotTableDefinitionSnapshot:
    """Safe aggregate of PivotTable view, cache-schema, and cached-data material.

    PivotTable views and their cache packages can change grouping, filtering,
    aggregation, and report presentation outside ordinary formula cells. Private
    signatures retain those definitions and cache payload evidence for comparison;
    ``to_dict`` deliberately exposes only structural counts.
    """

    pivot_table_sheet_count: int = 0
    pivot_table_part_count: int = 0
    pivot_cache_definition_part_count: int = 0
    pivot_cache_records_part_count: int = 0
    pivot_cache_binding_count: int = 0
    layout_location_count: int = 0
    pivot_field_count: int = 0
    row_field_count: int = 0
    column_field_count: int = 0
    page_field_count: int = 0
    data_field_count: int = 0
    filter_count: int = 0
    row_item_count: int = 0
    column_item_count: int = 0
    cache_field_count: int = 0
    shared_item_count: int = 0
    calculated_item_count: int = 0
    calculated_member_count: int = 0
    cache_record_count: int = 0
    related_relationship_count: int = 0
    external_relationship_count: int = 0
    fingerprinted_cache_record_part_count: int = 0
    uninspected_cache_record_part_count: int = 0
    unrecognized_part_count: int = 0
    declaration_signature: str | None = field(default=None, repr=False)
    layout_signature: str | None = field(default=None, repr=False)
    cache_definition_signature: str | None = field(default=None, repr=False)
    cached_shared_item_signature: str | None = field(default=None, repr=False)
    relationship_signature: str | None = field(default=None, repr=False)
    cache_record_payload_signature: str | None = field(default=None, repr=False)

    @property
    def present(self) -> bool:
        return bool(
            self.pivot_table_sheet_count
            or self.pivot_table_part_count
            or self.pivot_cache_definition_part_count
            or self.pivot_cache_records_part_count
            or self.unrecognized_part_count
        )

    def to_dict(self) -> dict[str, Any]:
        """Return PivotTable evidence without values, formulas, names, or targets."""
        return {
            "present": self.present,
            "pivot_table_sheet_count": self.pivot_table_sheet_count,
            "pivot_table_part_count": self.pivot_table_part_count,
            "pivot_cache_definition_part_count": self.pivot_cache_definition_part_count,
            "pivot_cache_records_part_count": self.pivot_cache_records_part_count,
            "pivot_cache_binding_count": self.pivot_cache_binding_count,
            "layout_location_count": self.layout_location_count,
            "pivot_field_count": self.pivot_field_count,
            "row_field_count": self.row_field_count,
            "column_field_count": self.column_field_count,
            "page_field_count": self.page_field_count,
            "data_field_count": self.data_field_count,
            "filter_count": self.filter_count,
            "row_item_count": self.row_item_count,
            "column_item_count": self.column_item_count,
            "cache_field_count": self.cache_field_count,
            "shared_item_count": self.shared_item_count,
            "calculated_item_count": self.calculated_item_count,
            "calculated_member_count": self.calculated_member_count,
            "cache_record_count": self.cache_record_count,
            "related_relationship_count": self.related_relationship_count,
            "external_relationship_count": self.external_relationship_count,
            "fingerprinted_cache_record_part_count": (
                self.fingerprinted_cache_record_part_count
            ),
            "uninspected_cache_record_part_count": self.uninspected_cache_record_part_count,
            "unrecognized_part_count": self.unrecognized_part_count,
        }

    def profile_dict(self) -> dict[str, Any]:
        return self.to_dict()


@dataclass(frozen=True)
class SlicerTimelineCacheSnapshot:
    """Safe aggregate of interactive slicer and Timeline filter-cache material.

    Slicer and Timeline caches can apply report filters to PivotTables or
    tables without changing an ordinary worksheet cell. Private signatures
    retain their declarations, filter state, and source bindings for comparison;
    ``to_dict`` deliberately exposes only structural counts.
    """

    slicer_cache_part_count: int = 0
    timeline_cache_part_count: int = 0
    slicer_workbook_binding_count: int = 0
    timeline_workbook_binding_count: int = 0
    slicer_pivot_cache_binding_count: int = 0
    slicer_table_binding_count: int = 0
    timeline_pivot_cache_binding_count: int = 0
    slicer_pivot_table_binding_count: int = 0
    timeline_pivot_table_binding_count: int = 0
    slicer_item_count: int = 0
    selected_slicer_item_count: int = 0
    timeline_state_count: int = 0
    timeline_filter_count: int = 0
    related_relationship_count: int = 0
    external_relationship_count: int = 0
    unrecognized_part_count: int = 0
    declaration_signature: str | None = field(default=None, repr=False)
    slicer_definition_signature: str | None = field(default=None, repr=False)
    timeline_definition_signature: str | None = field(default=None, repr=False)
    relationship_signature: str | None = field(default=None, repr=False)

    @property
    def present(self) -> bool:
        return bool(
            self.slicer_cache_part_count
            or self.timeline_cache_part_count
            or self.slicer_workbook_binding_count
            or self.timeline_workbook_binding_count
            or self.unrecognized_part_count
        )

    def to_dict(self) -> dict[str, Any]:
        """Return filter-cache evidence without names, values, or targets."""
        return {
            "present": self.present,
            "slicer_cache_part_count": self.slicer_cache_part_count,
            "timeline_cache_part_count": self.timeline_cache_part_count,
            "slicer_workbook_binding_count": self.slicer_workbook_binding_count,
            "timeline_workbook_binding_count": self.timeline_workbook_binding_count,
            "slicer_pivot_cache_binding_count": self.slicer_pivot_cache_binding_count,
            "slicer_table_binding_count": self.slicer_table_binding_count,
            "timeline_pivot_cache_binding_count": self.timeline_pivot_cache_binding_count,
            "slicer_pivot_table_binding_count": self.slicer_pivot_table_binding_count,
            "timeline_pivot_table_binding_count": self.timeline_pivot_table_binding_count,
            "slicer_item_count": self.slicer_item_count,
            "selected_slicer_item_count": self.selected_slicer_item_count,
            "timeline_state_count": self.timeline_state_count,
            "timeline_filter_count": self.timeline_filter_count,
            "related_relationship_count": self.related_relationship_count,
            "external_relationship_count": self.external_relationship_count,
            "unrecognized_part_count": self.unrecognized_part_count,
        }

    def profile_dict(self) -> dict[str, Any]:
        return self.to_dict()


@dataclass(frozen=True)
class WorksheetEmbeddedControlSnapshot:
    """Safe aggregate of worksheet control and OLE package data.

    Worksheet controls can bind persisted ActiveX state, modern and legacy form-
    control formulas, OLE payloads, and external linked objects outside cells and
    the VBA project. Private signatures retain all control definitions,
    relationships, and safely fingerprinted direct payloads for comparison;
    ``to_dict`` deliberately exposes only structural counts.
    """

    control_sheet_count: int = 0
    worksheet_control_count: int = 0
    active_x_part_count: int = 0
    active_x_binary_reference_count: int = 0
    form_control_property_part_count: int = 0
    legacy_vml_drawing_part_count: int = 0
    legacy_vml_control_count: int = 0
    legacy_vml_macro_assignment_count: int = 0
    legacy_vml_cell_link_count: int = 0
    legacy_vml_source_range_count: int = 0
    legacy_vml_camera_source_range_count: int = 0
    control_macro_assignment_count: int = 0
    control_cell_link_count: int = 0
    control_source_range_count: int = 0
    form_control_formula_binding_count: int = 0
    ole_object_count: int = 0
    linked_ole_object_count: int = 0
    auto_load_ole_object_count: int = 0
    auto_update_ole_object_count: int = 0
    related_relationship_count: int = 0
    external_relationship_count: int = 0
    internal_related_part_count: int = 0
    fingerprinted_related_part_count: int = 0
    uninspected_related_part_count: int = 0
    unrecognized_part_count: int = 0
    declaration_signature: str | None = field(default=None, repr=False)
    control_definition_signature: str | None = field(default=None, repr=False)
    active_x_definition_signature: str | None = field(default=None, repr=False)
    form_control_property_signature: str | None = field(default=None, repr=False)
    legacy_vml_definition_signature: str | None = field(default=None, repr=False)
    legacy_vml_relationship_signature: str | None = field(default=None, repr=False)
    relationship_signature: str | None = field(default=None, repr=False)
    related_part_payload_signature: str | None = field(default=None, repr=False)

    @property
    def present(self) -> bool:
        return bool(
            self.control_sheet_count
            or self.worksheet_control_count
            or self.active_x_part_count
            or self.form_control_property_part_count
            or self.legacy_vml_drawing_part_count
            or self.legacy_vml_control_count
            or self.ole_object_count
            or self.unrecognized_part_count
        )

    def to_dict(self) -> dict[str, Any]:
        """Return safe control inventory without identities, formulas, or payloads."""
        return {
            "present": self.present,
            "control_sheet_count": self.control_sheet_count,
            "worksheet_control_count": self.worksheet_control_count,
            "active_x_part_count": self.active_x_part_count,
            "active_x_binary_reference_count": self.active_x_binary_reference_count,
            "form_control_property_part_count": self.form_control_property_part_count,
            "legacy_vml_drawing_part_count": self.legacy_vml_drawing_part_count,
            "legacy_vml_control_count": self.legacy_vml_control_count,
            "legacy_vml_macro_assignment_count": (
                self.legacy_vml_macro_assignment_count
            ),
            "legacy_vml_cell_link_count": self.legacy_vml_cell_link_count,
            "legacy_vml_source_range_count": self.legacy_vml_source_range_count,
            "legacy_vml_camera_source_range_count": (
                self.legacy_vml_camera_source_range_count
            ),
            "control_macro_assignment_count": self.control_macro_assignment_count,
            "control_cell_link_count": self.control_cell_link_count,
            "control_source_range_count": self.control_source_range_count,
            "form_control_formula_binding_count": self.form_control_formula_binding_count,
            "ole_object_count": self.ole_object_count,
            "linked_ole_object_count": self.linked_ole_object_count,
            "auto_load_ole_object_count": self.auto_load_ole_object_count,
            "auto_update_ole_object_count": self.auto_update_ole_object_count,
            "related_relationship_count": self.related_relationship_count,
            "external_relationship_count": self.external_relationship_count,
            "internal_related_part_count": self.internal_related_part_count,
            "fingerprinted_related_part_count": self.fingerprinted_related_part_count,
            "uninspected_related_part_count": self.uninspected_related_part_count,
            "unrecognized_part_count": self.unrecognized_part_count,
        }

    def profile_dict(self) -> dict[str, Any]:
        return self.to_dict()


@dataclass(frozen=True)
class PowerQueryPermissionControlsSnapshot:
    """Safe aggregate controls from Data Mashup permission documents."""

    payload_count: int = 0
    parsed_count: int = 0
    firewall_enabled_count: int = 0
    future_packages_allowed_count: int = 0
    workbook_group_type_count: int = 0
    opaque_metadata: ExternalDataOpaqueMetadataSnapshot = field(
        default_factory=ExternalDataOpaqueMetadataSnapshot
    )
    signature: str | None = field(default=None, repr=False)

    def to_dict(self) -> dict[str, Any]:
        """Return permission-control counts without query or user identity material."""
        return {
            "payload_count": self.payload_count,
            "parsed_count": self.parsed_count,
            "firewall_enabled_count": self.firewall_enabled_count,
            "future_packages_allowed_count": self.future_packages_allowed_count,
            "workbook_group_type_count": self.workbook_group_type_count,
            "opaque_metadata": self.opaque_metadata.to_dict(),
        }


@dataclass(frozen=True)
class PowerQuerySnapshot:
    """Safe inventory of Data Mashup query definitions stored in custom XML."""

    mashup_count: int = 0
    parsed_mashup_count: int = 0
    formula_document_count: int = 0
    package_part_count: int = 0
    embedded_content_part_count: int = 0
    metadata_document_count: int = 0
    metadata_item_count: int = 0
    permission_controls: PowerQueryPermissionControlsSnapshot = field(
        default_factory=PowerQueryPermissionControlsSnapshot
    )
    permission_binding_count: int = 0
    formula_signature: str | None = field(default=None, repr=False)
    package_configuration_signature: str | None = field(default=None, repr=False)
    metadata_identity_signature: str | None = field(default=None, repr=False)
    metadata_control_signature: str | None = field(default=None, repr=False)
    opaque_metadata: ExternalDataOpaqueMetadataSnapshot = field(
        default_factory=ExternalDataOpaqueMetadataSnapshot
    )

    @property
    def present(self) -> bool:
        return self.mashup_count > 0

    def to_dict(self) -> dict[str, Any]:
        """Return structural Power Query evidence without M text or source material."""
        return {
            "present": self.present,
            "mashup_count": self.mashup_count,
            "parsed_mashup_count": self.parsed_mashup_count,
            "formula_document_count": self.formula_document_count,
            "package_part_count": self.package_part_count,
            "embedded_content_part_count": self.embedded_content_part_count,
            "metadata_document_count": self.metadata_document_count,
            "metadata_item_count": self.metadata_item_count,
            "permission_controls": self.permission_controls.to_dict(),
            "permission_binding_count": self.permission_binding_count,
            "opaque_metadata": self.opaque_metadata.to_dict(),
        }

    def profile_dict(self) -> dict[str, Any]:
        return self.to_dict()


@dataclass(frozen=True)
class RangeDependency:
    """A range used by a formula, stored without expanding large Excel ranges."""

    source_sheet: str
    min_column: int
    min_row: int
    max_column: int
    max_row: int
    dependent: CellKey

    def contains(self, location: CellKey) -> bool:
        sheet, coordinate = location
        if sheet != self.source_sheet:
            return False
        from openpyxl.utils.cell import column_index_from_string, coordinate_to_tuple

        row, column = coordinate_to_tuple(coordinate)
        if isinstance(column, str):  # defensive for older openpyxl versions
            column = column_index_from_string(column)
        return (
            self.min_column <= column <= self.max_column
            and self.min_row <= row <= self.max_row
        )

    def intersects(self, other: ArrayFormulaRange) -> bool:
        """Return whether this dependency reads any cell in an array output range."""
        return (
            self.source_sheet == other.sheet
            and self.min_column <= other.max_column
            and other.min_column <= self.max_column
            and self.min_row <= other.max_row
            and other.min_row <= self.max_row
        )


@dataclass(frozen=True)
class ArrayFormulaRange:
    """The declared result range from one array formula.

    Excel stores the formula only at ``anchor``. The remaining cells in ``ref``
    are result members, not independent formulas. For legacy CSE formulas the
    range is fixed; for dynamic arrays it is the currently observed extent.
    FormulaFence preserves either form compactly rather than creating one graph
    node per result cell.
    """

    sheet: str
    anchor: str
    ref: str
    min_column: int
    min_row: int
    max_column: int
    max_row: int

    @property
    def location(self) -> CellKey:
        return (self.sheet, self.anchor)

    @property
    def output_cell_count(self) -> int:
        return (self.max_column - self.min_column + 1) * (
            self.max_row - self.min_row + 1
        )

    @property
    def has_multiple_outputs(self) -> bool:
        return self.output_cell_count > 1

    def contains(self, location: CellKey) -> bool:
        sheet, coordinate = location
        if sheet != self.sheet:
            return False
        from openpyxl.utils.cell import column_index_from_string, coordinate_to_tuple

        row, column = coordinate_to_tuple(coordinate)
        if isinstance(column, str):  # defensive for older openpyxl versions
            column = column_index_from_string(column)
        return (
            self.min_column <= column <= self.max_column
            and self.min_row <= row <= self.max_row
        )

    def intersects_non_anchor(self, dependency: RangeDependency) -> bool:
        """Return whether a dependency intersects any result member after anchor."""
        if not dependency.intersects(self):
            return False
        min_column = max(self.min_column, dependency.min_column)
        max_column = min(self.max_column, dependency.max_column)
        min_row = max(self.min_row, dependency.min_row)
        max_row = min(self.max_row, dependency.max_row)
        if min_column != max_column or min_row != max_row:
            return True
        return (min_column, min_row) != (self.min_column, self.min_row)

    def to_dict(self) -> dict[str, Any]:
        return {
            "anchor": display_location(self.location),
            "ref": display_location((self.sheet, self.ref)),
            "output_cell_count": self.output_cell_count,
        }


@dataclass(frozen=True)
class DynamicArrayOutputReference:
    """One formula reading a non-anchor cell of an observed dynamic spill."""

    anchor: CellKey
    observed_ref: str

    def to_dict(self) -> dict[str, str]:
        sheet, _ = self.anchor
        return {
            "anchor": display_location(self.anchor) or "",
            "observed_range": display_location((sheet, self.observed_ref)) or "",
        }


@dataclass(frozen=True)
class Finding:
    rule_id: str
    severity: str
    message: str
    location: CellKey | None = None
    details: dict[str, Any] = field(default_factory=dict)

    def to_dict(self) -> dict[str, Any]:
        result = {
            "rule_id": self.rule_id,
            "severity": self.severity,
            "message": self.message,
            "location": display_location(self.location),
        }
        if self.details:
            result["details"] = self.details
        return result


@dataclass(frozen=True)
class Change:
    kind: str
    location: CellKey | None
    severity: str
    before: CellSnapshot | None = None
    after: CellSnapshot | None = None
    impact_count: int = 0
    impacted_cells: tuple[CellKey, ...] = ()
    details: dict[str, Any] = field(default_factory=dict)

    def to_dict(self) -> dict[str, Any]:
        return {
            "kind": self.kind,
            "severity": self.severity,
            "location": display_location(self.location),
            "before": self.before.to_dict() if self.before else None,
            "after": self.after.to_dict() if self.after else None,
            "impact_count": self.impact_count,
            "impacted_cells": [display_location(cell) for cell in self.impacted_cells],
            "details": self.details,
        }


@dataclass
class WorkbookSnapshot:
    """Workbook semantics plus the dependency indexes used for impact analysis."""

    path: Path
    sha256: str
    file_type: str
    sheets: dict[str, SheetSnapshot]
    cells: dict[CellKey, CellSnapshot]
    reverse_dependencies: dict[CellKey, set[CellKey]]
    range_dependencies: list[RangeDependency]
    external_references: set[CellKey]
    broken_references: set[CellKey]
    defined_names: dict[str, str]
    macro_hash: str | None
    calculation_settings: dict[str, Any]
    parser_warnings: tuple[str, ...]
    unresolved_reference_tokens: dict[CellKey, tuple[str, ...]] = field(default_factory=dict)
    dynamic_reference_functions: dict[CellKey, tuple[str, ...]] = field(default_factory=dict)
    tables: dict[str, TableSnapshot] = field(default_factory=dict)
    data_validations: tuple[DataValidationSnapshot, ...] = ()
    conditional_formatting: tuple[ConditionalFormattingSnapshot, ...] = ()
    conditional_formatting_extensions: tuple[ConditionalFormattingExtensionSnapshot, ...] = ()
    workbook_protection: WorkbookProtectionSnapshot | None = None
    sheet_protections: tuple[SheetProtectionSnapshot, ...] = ()
    protected_ranges: tuple[ProtectedRangeSnapshot, ...] = ()
    cell_protection_default: CellProtectionDefaultSnapshot | None = None
    cell_protection_assignments: tuple[CellProtectionAssignmentSnapshot, ...] = ()
    external_data_refresh_settings: ExternalDataRefreshSettingsSnapshot = field(
        default_factory=ExternalDataRefreshSettingsSnapshot
    )
    external_data_connections: tuple[ExternalDataConnectionSnapshot, ...] = ()
    query_table_refresh_controls: tuple[QueryTableRefreshSnapshot, ...] = ()
    pivot_cache_refresh_controls: tuple[PivotCacheRefreshSnapshot, ...] = ()
    external_link_packages: ExternalLinkPackageSnapshot = field(
        default_factory=ExternalLinkPackageSnapshot
    )
    xlm_macro_sheets: XlmMacroSheetSnapshot = field(
        default_factory=XlmMacroSheetSnapshot
    )
    ribbon_customization: RibbonCustomizationSnapshot = field(
        default_factory=RibbonCustomizationSnapshot
    )
    office_web_addins: OfficeWebAddinSnapshot = field(
        default_factory=OfficeWebAddinSnapshot
    )
    pivot_table_definitions: PivotTableDefinitionSnapshot = field(
        default_factory=PivotTableDefinitionSnapshot
    )
    slicer_timeline_caches: SlicerTimelineCacheSnapshot = field(
        default_factory=SlicerTimelineCacheSnapshot
    )
    chart_definitions: ChartDefinitionSnapshot = field(
        default_factory=ChartDefinitionSnapshot
    )
    worksheet_embedded_controls: WorksheetEmbeddedControlSnapshot = field(
        default_factory=WorksheetEmbeddedControlSnapshot
    )
    power_query: PowerQuerySnapshot = field(default_factory=PowerQuerySnapshot)
    sheet_order: tuple[str, ...] = ()
    three_d_reference_tokens: dict[CellKey, tuple[str, ...]] = field(default_factory=dict)
    spill_reference_tokens: dict[CellKey, tuple[str, ...]] = field(default_factory=dict)
    implicit_intersection_tokens: dict[CellKey, tuple[str, ...]] = field(
        default_factory=dict
    )
    legacy_array_formula_ranges: tuple[ArrayFormulaRange, ...] = ()
    dynamic_array_formula_cells: set[CellKey] = field(default_factory=set)
    dynamic_array_formula_ranges: tuple[ArrayFormulaRange, ...] = ()
    dynamic_array_output_references: dict[
        CellKey, tuple[DynamicArrayOutputReference, ...]
    ] = field(default_factory=dict)
    unclassified_array_formula_cells: set[CellKey] = field(default_factory=set)
    array_formula_output_dependents: dict[CellKey, set[CellKey]] = field(
        default_factory=dict
    )
    tokenization_failure_cells: set[CellKey] = field(default_factory=set)

    def direct_dependents(self, location: CellKey) -> set[CellKey]:
        dependents = set(self.reverse_dependencies.get(location, set()))
        dependents.update(self.array_formula_output_dependents.get(location, set()))
        for dependency in self.range_dependencies:
            if dependency.contains(location):
                dependents.add(dependency.dependent)
        return dependents

    def summary(self) -> dict[str, Any]:
        return {
            "path": str(self.path),
            "sha256": self.sha256,
            "file_type": self.file_type,
            "sheet_count": len(self.sheets),
            "nonempty_cells": len(self.cells),
            "formula_cells": sum(1 for cell in self.cells.values() if cell.is_formula),
            "defined_names": len(self.defined_names),
            "table_count": len(self.tables),
            "data_validation_rules": len(self.data_validations),
            "data_validation_target_ranges": sum(
                validation.target_range_count for validation in self.data_validations
            ),
            "conditional_formatting_rules": len(self.conditional_formatting),
            "conditional_formatting_target_ranges": sum(
                rule.target_range_count for rule in self.conditional_formatting
            ),
            "conditional_formatting_extensions": len(
                self.conditional_formatting_extensions
            ),
            "workbook_protection_enabled": bool(
                self.workbook_protection and self.workbook_protection.enabled
            ),
            "sheet_protection_controls": len(self.sheet_protections),
            "protected_sheet_count": sum(
                protection.enabled for protection in self.sheet_protections
            ),
            "protected_range_count": len(self.protected_ranges),
            "protected_range_target_ranges": sum(
                protected_range.target_range_count
                for protected_range in self.protected_ranges
            ),
            "cell_protection_assignment_count": len(
                self.cell_protection_assignments
            ),
            "external_data_connection_count": len(self.external_data_connections),
            "external_data_connections_refresh_on_load": sum(
                connection.refresh_on_load
                for connection in self.external_data_connections
            ),
            "query_table_refresh_control_count": len(
                self.query_table_refresh_controls
            ),
            "query_tables_refresh_on_load": sum(
                control.refresh_on_load
                for control in self.query_table_refresh_controls
            ),
            "pivot_cache_refresh_control_count": len(
                self.pivot_cache_refresh_controls
            ),
            "pivot_caches_refresh_on_load": sum(
                control.refresh_on_load
                for control in self.pivot_cache_refresh_controls
            ),
            "external_link_package_count": self.external_link_packages.external_link_count,
            "external_workbook_link_count": self.external_link_packages.external_workbook_count,
            "dde_link_count": self.external_link_packages.dde_link_count,
            "ole_link_count": self.external_link_packages.ole_link_count,
            "xlm_macro_sheet_count": self.xlm_macro_sheets.macro_sheet_count,
            "xlm_macro_formula_cell_count": self.xlm_macro_sheets.formula_cell_count,
            "xlm_related_part_payload_count": (
                self.xlm_macro_sheets.fingerprinted_related_part_count
            ),
            "has_xlm_macro_sheets": self.xlm_macro_sheets.present,
            "ribbon_customization_part_count": self.ribbon_customization.ribbon_part_count,
            "ribbon_callback_attribute_count": (
                self.ribbon_customization.callback_attribute_count
            ),
            "has_ribbon_customization": self.ribbon_customization.present,
            "office_web_addin_taskpane_part_count": (
                self.office_web_addins.taskpane_part_count
            ),
            "office_web_addin_web_extension_part_count": (
                self.office_web_addins.web_extension_part_count
            ),
            "office_web_addin_auto_show_taskpane_count": (
                self.office_web_addins.auto_show_taskpane_count
            ),
            "has_office_web_addins": self.office_web_addins.present,
            "pivot_table_sheet_count": (
                self.pivot_table_definitions.pivot_table_sheet_count
            ),
            "pivot_table_part_count": self.pivot_table_definitions.pivot_table_part_count,
            "pivot_cache_definition_part_count": (
                self.pivot_table_definitions.pivot_cache_definition_part_count
            ),
            "pivot_cache_record_count": self.pivot_table_definitions.cache_record_count,
            "has_pivot_table_definitions": self.pivot_table_definitions.present,
            "slicer_cache_part_count": self.slicer_timeline_caches.slicer_cache_part_count,
            "timeline_cache_part_count": self.slicer_timeline_caches.timeline_cache_part_count,
            "selected_slicer_item_count": (
                self.slicer_timeline_caches.selected_slicer_item_count
            ),
            "has_slicer_timeline_caches": self.slicer_timeline_caches.present,
            "chart_host_sheet_count": self.chart_definitions.chart_host_sheet_count,
            "chart_drawing_part_count": self.chart_definitions.chart_drawing_part_count,
            "chart_part_count": self.chart_definitions.chart_part_count,
            "chart_series_count": self.chart_definitions.series_count,
            "chart_cached_data_point_count": (
                self.chart_definitions.cached_data_point_count
            ),
            "has_chart_definitions": self.chart_definitions.present,
            "worksheet_embedded_control_sheet_count": (
                self.worksheet_embedded_controls.control_sheet_count
            ),
            "worksheet_active_x_part_count": (
                self.worksheet_embedded_controls.active_x_part_count
            ),
            "worksheet_legacy_vml_drawing_part_count": (
                self.worksheet_embedded_controls.legacy_vml_drawing_part_count
            ),
            "worksheet_legacy_vml_control_count": (
                self.worksheet_embedded_controls.legacy_vml_control_count
            ),
            "worksheet_ole_object_count": (
                self.worksheet_embedded_controls.ole_object_count
            ),
            "has_worksheet_embedded_controls": self.worksheet_embedded_controls.present,
            "power_query_mashup_count": self.power_query.mashup_count,
            "power_query_formula_document_count": self.power_query.formula_document_count,
            "power_query_metadata_item_count": self.power_query.metadata_item_count,
            "has_vba": self.macro_hash is not None,
            "external_reference_cells": len(self.external_references),
            "broken_reference_cells": len(self.broken_references),
            "unresolved_reference_cells": len(self.unresolved_reference_tokens),
            "dynamic_reference_cells": len(self.dynamic_reference_functions),
            "three_d_reference_cells": len(self.three_d_reference_tokens),
            "spill_reference_cells": len(self.spill_reference_tokens),
            "implicit_intersection_cells": len(self.implicit_intersection_tokens),
            "legacy_array_formula_cells": len(self.legacy_array_formula_ranges),
            "legacy_array_formula_output_ranges": sum(
                array_range.has_multiple_outputs
                for array_range in self.legacy_array_formula_ranges
            ),
            "legacy_array_formula_output_cells": sum(
                array_range.output_cell_count
                for array_range in self.legacy_array_formula_ranges
            ),
            "dynamic_array_formula_cells": len(self.dynamic_array_formula_cells),
            "dynamic_array_observed_output_ranges": len(
                self.dynamic_array_formula_ranges
            ),
            "dynamic_array_output_reference_cells": len(
                self.dynamic_array_output_references
            ),
            "unclassified_array_formula_cells": len(
                self.unclassified_array_formula_cells
            ),
            "tokenization_failure_cells": len(self.tokenization_failure_cells),
            "parser_warning_count": len(self.parser_warnings),
        }


@dataclass
class DiffReport:
    before: WorkbookSnapshot
    after: WorkbookSnapshot
    changes: list[Change]
    findings: list[Finding]

    def severity_counts(self) -> dict[str, int]:
        counts = {severity: 0 for severity in SEVERITY_ORDER}
        for finding in self.findings:
            counts[finding.severity] = counts.get(finding.severity, 0) + 1
        return {severity: count for severity, count in counts.items() if count}

    @property
    def highest_severity(self) -> str:
        if not self.findings:
            return "note"
        return max(self.findings, key=lambda finding: SEVERITY_ORDER[finding.severity]).severity

    def to_dict(self, extra_findings: Iterable[Finding] = ()) -> dict[str, Any]:
        findings = [*self.findings, *extra_findings]
        counts = {severity: 0 for severity in SEVERITY_ORDER}
        for finding in findings:
            counts[finding.severity] = counts.get(finding.severity, 0) + 1
        highest = (
            max(findings, key=lambda finding: SEVERITY_ORDER[finding.severity]).severity
            if findings
            else "note"
        )
        return {
            "schema_version": "1.0",
            "before": self.before.summary(),
            "after": self.after.summary(),
            "summary": {
                "change_count": len(self.changes),
                "finding_count": len(findings),
                "highest_severity": highest,
                "findings_by_severity": {
                    severity: count for severity, count in counts.items() if count
                },
            },
            "changes": [change.to_dict() for change in self.changes],
            "findings": [finding.to_dict() for finding in findings],
        }


class FormulaFenceError(Exception):
    """Base class for errors that should be presented cleanly by the CLI."""


class WorkbookLoadError(FormulaFenceError):
    """The workbook could not be safely inspected."""


class PolicyError(FormulaFenceError):
    """The policy file is invalid or cannot be interpreted."""
