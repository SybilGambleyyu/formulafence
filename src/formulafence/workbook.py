"""Safe, non-evaluating workbook inspection and dependency indexing."""

from __future__ import annotations

import hashlib
import posixpath
import re
import warnings
from collections import defaultdict
from collections.abc import Mapping
from dataclasses import dataclass, replace
from pathlib import Path
from xml.etree import ElementTree
from zipfile import BadZipFile, ZipFile

from openpyxl import load_workbook
from openpyxl.utils.cell import column_index_from_string, get_column_letter, range_boundaries
from openpyxl.utils.exceptions import InvalidFileException

from formulafence.formulas import (
    ParsedReference,
    StructuredTable,
    formula_fingerprint,
    has_broken_reference,
    inspect_formula,
    lambda_parameter_count,
    parse_reference_token,
    reference_lookup_key,
)
from formulafence.models import (
    ArrayFormulaRange,
    CellKey,
    CellProtectionAssignmentSnapshot,
    CellProtectionDefaultSnapshot,
    CellSnapshot,
    ConditionalFormattingExtensionSnapshot,
    ConditionalFormattingSnapshot,
    DataValidationSnapshot,
    DynamicArrayOutputReference,
    ProtectedRangeSnapshot,
    ProtectionCredentialSnapshot,
    ProtectionOpaqueMetadataSnapshot,
    RangeDependency,
    SheetProtectionSnapshot,
    SheetSnapshot,
    TableSnapshot,
    WorkbookLoadError,
    WorkbookProtectionSnapshot,
    WorkbookSnapshot,
    XmlFragmentSnapshot,
    display_location,
    json_safe_value,
)

_SUPPORTED_SUFFIXES = {".xlsx", ".xlsm"}
_CALCULATION_FIELDS = (
    "calcMode",
    "fullCalcOnLoad",
    "refMode",
    "iterate",
    "iterateCount",
    "iterateDelta",
    "fullPrecision",
    "calcCompleted",
    "calcOnSave",
    "concurrentCalc",
    "concurrentManualCount",
    "forceFullCalc",
)
_SPREADSHEETML_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
_PACKAGE_RELATIONSHIP_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
_DOCUMENT_RELATIONSHIP_NS = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
)
_DYNAMIC_ARRAY_NS = "http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray"
_OFFICE_2010_SPREADSHEET_NS = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"
_EXCEL_2006_MAIN_NS = "http://schemas.microsoft.com/office/excel/2006/main"
_XML_NAMESPACE_PREFIXES = {
    _SPREADSHEETML_NS: "",
    _OFFICE_2010_SPREADSHEET_NS: "x14:",
    _EXCEL_2006_MAIN_NS: "xm:",
}
_GUID_PATTERN = re.compile(
    r"\{[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-"
    r"[0-9a-fA-F]{4}-[0-9a-fA-F]{12}\}"
)


@dataclass(frozen=True)
class _ArrayFormulaMetadata:
    """Raw OOXML metadata needed to distinguish CSE from dynamic arrays."""

    dynamic_cells: set[CellKey]
    unclassified_cells: set[CellKey]
    scanned_sheets: set[str]
    complete: bool
    warnings: tuple[str, ...]


@dataclass(frozen=True)
class _ArrayFormulaClassification:
    """The conservative array-formula classifications for one loaded workbook."""

    kinds: dict[CellKey, str]
    refs: dict[CellKey, str]
    legacy_ranges: tuple[ArrayFormulaRange, ...]
    dynamic_cells: set[CellKey]
    dynamic_ranges: tuple[ArrayFormulaRange, ...]
    unclassified_cells: set[CellKey]
    warnings: tuple[str, ...]


@dataclass(frozen=True)
class _ConditionalFormattingMetadata:
    """Raw worksheet conditional-formatting controls and extension evidence."""

    rules: tuple[ConditionalFormattingSnapshot, ...]
    extensions: tuple[ConditionalFormattingExtensionSnapshot, ...]
    warnings: tuple[str, ...]


@dataclass(frozen=True)
class _ProtectionMetadata:
    """Raw OOXML protection controls retained before libraries can omit them."""

    workbook_protection: WorkbookProtectionSnapshot | None
    sheet_protections: tuple[SheetProtectionSnapshot, ...]
    protected_ranges: tuple[ProtectedRangeSnapshot, ...]
    cell_protection_default: CellProtectionDefaultSnapshot | None
    cell_protection_assignments: tuple[CellProtectionAssignmentSnapshot, ...]
    warnings: tuple[str, ...]


@dataclass(frozen=True)
class _StyleProtection:
    """Effective protection from one cell XF, plus whether it is explicit."""

    locked: bool
    hidden: bool
    explicit: bool = False


_SHEET_PROTECTION_ACTIONS = (
    ("objects", "objects", False),
    ("scenarios", "scenarios", False),
    ("formatCells", "format_cells", True),
    ("formatColumns", "format_columns", True),
    ("formatRows", "format_rows", True),
    ("insertColumns", "insert_columns", True),
    ("insertRows", "insert_rows", True),
    ("insertHyperlinks", "insert_hyperlinks", True),
    ("deleteColumns", "delete_columns", True),
    ("deleteRows", "delete_rows", True),
    ("selectLockedCells", "select_locked_cells", False),
    ("sort", "sort", True),
    ("autoFilter", "auto_filter", True),
    ("pivotTables", "pivot_tables", True),
    ("selectUnlockedCells", "select_unlocked_cells", False),
)
_CHART_SHEET_PROTECTION_ACTIONS = (
    ("content", "content", False),
    ("objects", "objects", False),
)


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as file_handle:
        for block in iter(lambda: file_handle.read(1_048_576), b""):
            digest.update(block)
    return digest.hexdigest()


def _vba_hash(path: Path) -> str | None:
    """Hash the macro payload without loading or executing it."""
    try:
        with ZipFile(path) as archive:
            try:
                payload = archive.read("xl/vbaProject.bin")
            except KeyError:
                return None
    except BadZipFile as error:
        raise WorkbookLoadError(f"{path} is not a valid Office Open XML workbook") from error
    return hashlib.sha256(payload).hexdigest()


def _xml_root(archive: ZipFile, member: str) -> ElementTree.Element:
    """Read one OOXML part without accepting document type declarations."""
    payload = archive.read(member)
    if b"<!DOCTYPE" in payload or b"<!ENTITY" in payload:
        raise ValueError("OOXML metadata contains a document type declaration")
    return ElementTree.fromstring(payload)


def _normalise_package_target(target: str) -> str | None:
    """Turn a workbook relationship target into a safe ZIP member name."""
    candidate = target.lstrip("/") if target.startswith("/") else posixpath.join("xl", target)
    normalised = posixpath.normpath(candidate)
    if normalised == ".." or normalised.startswith("../"):
        return None
    return normalised


def _sheet_xml_parts(archive: ZipFile) -> dict[str, tuple[str, str]]:
    """Map workbook sheet titles to safe OOXML parts and their sheet kind."""
    workbook = _xml_root(archive, "xl/workbook.xml")
    relationships = _xml_root(archive, "xl/_rels/workbook.xml.rels")
    relationship_targets: dict[str, tuple[str, str]] = {}
    relationship_tag = f"{{{_PACKAGE_RELATIONSHIP_NS}}}Relationship"
    for relationship in relationships.findall(relationship_tag):
        relationship_type = relationship.get("Type", "")
        relationship_id = relationship.get("Id")
        target = relationship.get("Target")
        sheet_type = relationship_type.rsplit("/", maxsplit=1)[-1]
        if (
            sheet_type not in {"worksheet", "chartsheet", "dialogsheet"}
            or not relationship_id
            or not target
            or (normalised := _normalise_package_target(target)) is None
        ):
            continue
        relationship_targets[relationship_id] = (normalised, sheet_type)

    sheet_parts: dict[str, tuple[str, str]] = {}
    sheet_tag = f"{{{_SPREADSHEETML_NS}}}sheet"
    relationship_id_attribute = f"{{{_DOCUMENT_RELATIONSHIP_NS}}}id"
    for sheet in workbook.iter(sheet_tag):
        title = sheet.get("name")
        relationship_id = sheet.get(relationship_id_attribute)
        if not title or not relationship_id:
            continue
        if part := relationship_targets.get(relationship_id):
            sheet_parts[title] = part
    return sheet_parts


def _worksheet_xml_paths(archive: ZipFile) -> dict[str, str]:
    """Map standard worksheet titles to their OOXML worksheet parts."""
    return {
        title: member
        for title, (member, sheet_type) in _sheet_xml_parts(archive).items()
        if sheet_type == "worksheet"
    }


def _xml_local_name(tag: str) -> str:
    """Return an XML local name without discarding a namespace elsewhere."""
    return tag.rsplit("}", maxsplit=1)[-1]


def _xml_namespace(tag: str) -> str | None:
    """Return an XML namespace, if ``tag`` is namespace-qualified."""
    if not tag.startswith("{"):
        return None
    return tag[1:].split("}", maxsplit=1)[0]


def _xml_display_name(tag: str) -> str:
    """Render a qualified XML name deterministically for local review evidence."""
    if not tag.startswith("{"):
        return tag
    namespace, local_name = tag[1:].split("}", maxsplit=1)
    if prefix := _XML_NAMESPACE_PREFIXES.get(namespace):
        return f"{prefix}{local_name}"
    if namespace == _SPREADSHEETML_NS:
        return local_name
    return f"{{{namespace}}}{local_name}"


def _normalise_guid(value: str) -> str:
    """Remove writer-specific GUID noise while retaining extension structure."""
    return _GUID_PATTERN.sub("{GUID}", value)


def _normalise_conditional_formula(value: str) -> str:
    """Accept the optional leading-equals spelling used by workbook writers."""
    return value[1:] if value.startswith("=") else value


def _is_conditional_guid_link(
    element: ElementTree.Element,
    attribute: str | None = None,
) -> bool:
    """Return whether an x14 GUID only links a base rule to its extension.

    An ``ext`` element's ``uri`` is a semantic extension-type identifier and
    must never be normalised away. The known ``x14:id`` / ``x14:cfRule@id``
    pair, by contrast, is writer-generated linkage between equivalent base and
    extension rule records.
    """
    if _xml_namespace(element.tag) != _OFFICE_2010_SPREADSHEET_NS:
        return False
    local_name = _xml_local_name(element.tag)
    if attribute is None:
        return local_name == "id"
    return local_name == "cfRule" and _xml_local_name(attribute) == "id"


def _xml_fragment(
    element: ElementTree.Element,
    *,
    normalise_guids: bool = False,
    normalise_formulas: bool = False,
) -> XmlFragmentSnapshot:
    """Capture an XML subtree with deterministic names and attribute ordering."""
    local_name = _xml_local_name(element.tag)
    formula_value = normalise_formulas and local_name in {"f", "formula"}
    value_object = normalise_formulas and local_name == "cfvo" and element.get("type") == "formula"

    def normalise_value(
        value: str,
        *,
        formula: bool = False,
        guid_link: bool = False,
    ) -> str:
        if formula:
            value = _normalise_conditional_formula(value)
        return _normalise_guid(value) if normalise_guids and guid_link else value

    attributes = tuple(
        sorted(
            (
                _xml_display_name(attribute),
                normalise_value(
                    value,
                    formula=value_object and _xml_local_name(attribute) == "val",
                    guid_link=_is_conditional_guid_link(element, attribute),
                ),
            )
            for attribute, value in element.attrib.items()
        )
    )
    children = tuple(
        _xml_fragment(
            child,
            normalise_guids=normalise_guids,
            normalise_formulas=normalise_formulas,
        )
        for child in element
    )
    text = element.text
    if children and text is not None and not text.strip():
        text = None
    if text is not None:
        text = normalise_value(
            text,
            formula=formula_value,
            guid_link=_is_conditional_guid_link(element),
        )
    return XmlFragmentSnapshot(
        tag=_xml_display_name(element.tag),
        attributes=attributes,
        text=text,
        children=children,
    )


def _conditional_ranges(value: str | None) -> tuple[str, ...]:
    """Return an order-independent, compact inventory of an ``sqref`` value."""
    return tuple(sorted({part for part in (value or "").split() if part}, key=str.casefold))


def _conditional_bool(
    value: str | None,
    default: bool,
    attribute: str,
    warnings: set[str],
) -> bool:
    """Read one OOXML boolean and preserve the schema default when omitted."""
    if value is None:
        return default
    lowered = value.casefold()
    if lowered in {"1", "true"}:
        return True
    if lowered in {"0", "false"}:
        return False
    warnings.add(
        "FormulaFence could not interpret a conditional-formatting "
        f"{attribute} boolean; the schema default was used."
    )
    return default


def _conditional_int(
    element: ElementTree.Element,
    attribute: str,
    warnings: set[str],
) -> int | None:
    """Read an optional integer attribute without making malformed XML look safe."""
    value = element.get(attribute)
    if value is None:
        return None
    try:
        return int(value)
    except ValueError:
        warnings.add(
            "FormulaFence could not interpret a conditional-formatting "
            f"{attribute} integer; the affected rule has a coverage gap."
        )
        return None


def _differential_style_fragments(archive: ZipFile) -> tuple[XmlFragmentSnapshot, ...]:
    """Resolve conditional-format ``dxfId`` values without trusting their order."""
    try:
        styles = _xml_root(archive, "xl/styles.xml")
    except KeyError:
        return ()
    dxfs = styles.find(f"{{{_SPREADSHEETML_NS}}}dxfs")
    if dxfs is None:
        return ()
    return tuple(
        _xml_fragment(dxf, normalise_formulas=True)
        for dxf in dxfs.findall(f"{{{_SPREADSHEETML_NS}}}dxf")
    )


def _conditional_rule_snapshot(
    sheet: str,
    ranges: tuple[str, ...],
    rule: ElementTree.Element,
    differential_styles: tuple[XmlFragmentSnapshot, ...],
    warnings: set[str],
) -> ConditionalFormattingSnapshot:
    """Read one base OOXML ``cfRule`` without calculating its formula."""
    raw_priority = _conditional_int(rule, "priority", warnings)
    if raw_priority is None or raw_priority <= 0:
        warnings.add(
            "FormulaFence could not read a valid conditional-formatting priority; "
            "the affected rule has a coverage gap."
        )
        raw_priority = 0

    rule_type = rule.get("type") or "unknown"
    if rule_type == "unknown":
        warnings.add(
            "FormulaFence found a conditional-formatting rule without a type; "
            "the affected rule has a coverage gap."
        )
    formula_tag = f"{{{_SPREADSHEETML_NS}}}formula"
    formulas = tuple(
        _normalise_conditional_formula(formula.text or "")
        for formula in rule.findall(formula_tag)
    )
    differential_style: XmlFragmentSnapshot | None = None
    dxf_id = rule.get("dxfId")
    if dxf_id is not None:
        parsed_dxf_id = _conditional_int(rule, "dxfId", warnings)
        if parsed_dxf_id is None or not 0 <= parsed_dxf_id < len(differential_styles):
            warnings.add(
                "FormulaFence could not resolve a conditional-formatting differential "
                "style; the affected rule has a coverage gap."
            )
            differential_style = XmlFragmentSnapshot(
                "unresolved-dxf", (("id", dxf_id),)
            )
        else:
            differential_style = differential_styles[parsed_dxf_id]

    def component(name: str) -> XmlFragmentSnapshot | None:
        found = rule.find(f"{{{_SPREADSHEETML_NS}}}{name}")
        return (
            _xml_fragment(found, normalise_formulas=True)
            if found is not None
            else None
        )

    extension_list = rule.find(f"{{{_SPREADSHEETML_NS}}}extLst")
    extensions = (
        tuple(
            _xml_fragment(
                extension,
                normalise_guids=True,
                normalise_formulas=True,
            )
            for extension in extension_list
        )
        if extension_list is not None
        else ()
    )
    known_attributes = {
        "type",
        "priority",
        "dxfId",
        "stopIfTrue",
        "aboveAverage",
        "percent",
        "bottom",
        "operator",
        "text",
        "timePeriod",
        "rank",
        "stdDev",
        "equalAverage",
    }
    unmodelled_attributes = tuple(
        sorted(
            (
                _xml_display_name(attribute),
                value,
            )
            for attribute, value in rule.attrib.items()
            if _xml_local_name(attribute) not in known_attributes
        )
    )
    if unmodelled_attributes:
        warnings.add(
            "FormulaFence found unmodelled conditional-formatting rule attributes; "
            "their raw structure is retained for review."
        )
        extensions += (
            XmlFragmentSnapshot("unmodelled-cf-rule-attributes", unmodelled_attributes),
        )
    known_children = {"formula", "colorScale", "dataBar", "iconSet", "extLst"}
    unmodelled_children = tuple(
        _xml_fragment(child, normalise_guids=True, normalise_formulas=True)
        for child in rule
        if _xml_local_name(child.tag) not in known_children
    )
    if unmodelled_children:
        warnings.add(
            "FormulaFence found unmodelled conditional-formatting rule children; "
            "their raw structure is retained for review."
        )
        extensions += unmodelled_children

    return ConditionalFormattingSnapshot(
        sheet=sheet,
        ranges=ranges,
        priority=raw_priority,
        rule_type=rule_type,
        operator=rule.get("operator"),
        formulas=formulas,
        stop_if_true=_conditional_bool(rule.get("stopIfTrue"), False, "stopIfTrue", warnings),
        above_average=_conditional_bool(
            rule.get("aboveAverage"), True, "aboveAverage", warnings
        ),
        percent=_conditional_bool(rule.get("percent"), False, "percent", warnings),
        bottom=_conditional_bool(rule.get("bottom"), False, "bottom", warnings),
        rank=_conditional_int(rule, "rank", warnings),
        std_dev=_conditional_int(rule, "stdDev", warnings),
        equal_average=_conditional_bool(
            rule.get("equalAverage"), False, "equalAverage", warnings
        ),
        text=rule.get("text"),
        time_period=rule.get("timePeriod"),
        differential_style=differential_style,
        color_scale=component("colorScale"),
        data_bar=component("dataBar"),
        icon_set=component("iconSet"),
        extensions=extensions,
    )


def _worksheet_conditional_formatting_extensions(
    sheet: str,
    worksheet: ElementTree.Element,
) -> tuple[ConditionalFormattingExtensionSnapshot, ...]:
    """Keep worksheet-level x14 conditional-formatting extensions inspectable."""
    extension_snapshots: list[ConditionalFormattingExtensionSnapshot] = []
    extension_list_tag = f"{{{_SPREADSHEETML_NS}}}extLst"
    for extension_list in worksheet.findall(extension_list_tag):
        for extension in extension_list:
            if not any(
                _xml_local_name(element.tag)
                in {"conditionalFormattings", "conditionalFormatting", "cfRule"}
                for element in extension.iter()
            ):
                continue
            extension_snapshots.append(
                ConditionalFormattingExtensionSnapshot(
                    sheet=sheet,
                    fragment=_xml_fragment(
                        extension,
                        normalise_guids=True,
                        normalise_formulas=True,
                    ),
                )
            )
    return tuple(
        sorted(extension_snapshots, key=ConditionalFormattingExtensionSnapshot.sort_key)
    )


def _conditional_formatting_metadata(path: Path) -> _ConditionalFormattingMetadata:
    """Read conditional formatting from OOXML before a library can discard extensions."""
    rules_by_sheet: dict[str, list[tuple[int, int, ConditionalFormattingSnapshot]]] = (
        defaultdict(list)
    )
    extensions: list[ConditionalFormattingExtensionSnapshot] = []
    warnings: set[str] = set()
    try:
        with ZipFile(path) as archive:
            differential_styles = _differential_style_fragments(archive)
            conditional_formatting_tag = f"{{{_SPREADSHEETML_NS}}}conditionalFormatting"
            rule_tag = f"{{{_SPREADSHEETML_NS}}}cfRule"
            for sheet, member in _worksheet_xml_paths(archive).items():
                worksheet = _xml_root(archive, member)
                source_order = 0
                for conditional_formatting in worksheet.findall(conditional_formatting_tag):
                    ranges = _conditional_ranges(conditional_formatting.get("sqref"))
                    if not ranges:
                        warnings.add(
                            "FormulaFence found conditional formatting without target ranges; "
                            "the affected control has a coverage gap."
                        )
                        continue
                    for rule in conditional_formatting.findall(rule_tag):
                        snapshot = _conditional_rule_snapshot(
                            sheet,
                            ranges,
                            rule,
                            differential_styles,
                            warnings,
                        )
                        rules_by_sheet[sheet].append(
                            (snapshot.priority, source_order, snapshot)
                        )
                        source_order += 1
                extensions.extend(
                    _worksheet_conditional_formatting_extensions(sheet, worksheet)
                )
    except (BadZipFile, ElementTree.ParseError, KeyError, OSError, ValueError) as error:
        return _ConditionalFormattingMetadata(
            rules=(),
            extensions=(),
            warnings=(
                "FormulaFence could not inspect conditional-formatting OOXML "
                f"({type(error).__name__}); conditional-formatting controls were not compared.",
            ),
        )

    rules: list[ConditionalFormattingSnapshot] = []
    for _sheet, entries in rules_by_sheet.items():
        priorities = [priority for priority, _, _ in entries]
        if len(set(priorities)) != len(priorities):
            warnings.add(
                "FormulaFence found duplicate conditional-formatting priorities on a "
                "worksheet; the affected controls have a coverage gap."
            )
        for normalized_priority, (_, _, snapshot) in enumerate(
            sorted(entries, key=lambda item: item[:2]),
            start=1,
        ):
            rules.append(replace(snapshot, priority=normalized_priority))
    return _ConditionalFormattingMetadata(
        rules=tuple(sorted(rules, key=ConditionalFormattingSnapshot.sort_key)),
        extensions=tuple(
            sorted(extensions, key=ConditionalFormattingExtensionSnapshot.sort_key)
        ),
        warnings=tuple(sorted(warnings)),
    )


def _protection_bool(
    value: str | None,
    default: bool,
    attribute: str,
    warnings: set[str],
) -> bool:
    """Read a protection boolean while retaining its schema default."""
    if value is None:
        return default
    lowered = value.casefold()
    if lowered in {"1", "true", "on"}:
        return True
    if lowered in {"0", "false", "off"}:
        return False
    warnings.add(
        "FormulaFence could not interpret a protection "
        f"{attribute} boolean; the schema default was used."
    )
    return default


def _protection_int(
    element: ElementTree.Element,
    attribute: str,
    warnings: set[str],
    *,
    context: str,
) -> int | None:
    """Read an optional non-negative protection integer conservatively."""
    value = element.get(attribute)
    if value is None:
        return None
    try:
        parsed = int(value)
    except ValueError:
        warnings.add(
            "FormulaFence could not interpret a "
            f"{context} {attribute} integer; the affected protection has a coverage gap."
        )
        return None
    if parsed < 0:
        warnings.add(
            "FormulaFence found a negative "
            f"{context} {attribute} integer; the affected protection has a coverage gap."
        )
        return None
    return parsed


def _private_protection_signature(
    entries: tuple[tuple[str, str], ...],
) -> str | None:
    """Hash sensitive comparison material without retaining it in output models."""
    if not entries:
        return None
    digest = hashlib.sha256()
    for name, value in entries:
        digest.update(name.encode("utf-8"))
        digest.update(b"\0")
        digest.update(value.encode("utf-8"))
        digest.update(b"\0")
    return digest.hexdigest()


def _merge_opaque_protection_metadata(
    *metadata: ProtectionOpaqueMetadataSnapshot,
) -> ProtectionOpaqueMetadataSnapshot:
    """Combine opaque fragments while leaving their contents private."""
    present = tuple(
        item for item in metadata if item.count and item.signature is not None
    )
    if not present:
        return ProtectionOpaqueMetadataSnapshot()
    return ProtectionOpaqueMetadataSnapshot(
        count=sum(item.count for item in present),
        signature=_private_protection_signature(
            tuple(
                (str(index), item.signature or "")
                for index, item in enumerate(present, start=1)
            )
        ),
    )


def _opaque_protection_metadata(
    element: ElementTree.Element,
    *,
    known_attributes: set[str],
    known_children: set[str] = frozenset(),
) -> ProtectionOpaqueMetadataSnapshot:
    """Fingerprint unmodelled protection XML without exposing its contents."""
    entries: list[tuple[str, str]] = []
    for attribute, value in element.attrib.items():
        if _xml_local_name(attribute) not in known_attributes:
            entries.append((f"attribute:{_xml_display_name(attribute)}", value))
    for child in element:
        if _xml_local_name(child.tag) in known_children:
            continue
        entries.append(
            (
                f"child:{_xml_display_name(child.tag)}",
                repr(_xml_fragment(child).sort_key()),
            )
        )
    entries.sort()
    return ProtectionOpaqueMetadataSnapshot(
        count=len(entries),
        signature=_private_protection_signature(tuple(entries)),
    )


def _protection_credential_snapshot(
    element: ElementTree.Element,
    *,
    legacy_attribute: str,
    algorithm_attribute: str,
    hash_attribute: str,
    salt_attribute: str,
    spin_count_attribute: str,
    context: str,
    warnings: set[str],
) -> ProtectionCredentialSnapshot:
    """Record safe verifier metadata while hashing the actual verifier values."""
    attribute_names = (
        legacy_attribute,
        algorithm_attribute,
        hash_attribute,
        salt_attribute,
        spin_count_attribute,
    )
    signature_entries = tuple(
        (attribute, value)
        for attribute in attribute_names
        if (value := element.get(attribute)) is not None
    )
    modern_attributes = (
        algorithm_attribute,
        hash_attribute,
        salt_attribute,
        spin_count_attribute,
    )
    has_modern_verifier = any(element.get(attribute) is not None for attribute in modern_attributes)
    if has_modern_verifier and not (
        element.get(hash_attribute) is not None and element.get(salt_attribute) is not None
    ):
        warnings.add(
            "FormulaFence found incomplete modern "
            f"{context} verifier metadata; its presence is compared but has a coverage gap."
        )
    return ProtectionCredentialSnapshot(
        has_legacy_verifier=element.get(legacy_attribute) is not None,
        has_modern_verifier=has_modern_verifier,
        algorithm=element.get(algorithm_attribute),
        spin_count=_protection_int(
            element,
            spin_count_attribute,
            warnings,
            context=context,
        ),
        signature=_private_protection_signature(signature_entries),
    )


def _workbook_protection_snapshot(
    workbook: ElementTree.Element,
    warnings: set[str],
) -> WorkbookProtectionSnapshot | None:
    """Read workbook protection without passing verifier material to openpyxl."""
    element = workbook.find(f"{{{_SPREADSHEETML_NS}}}workbookProtection")
    if element is None:
        return None
    workbook_credential = _protection_credential_snapshot(
        element,
        legacy_attribute="workbookPassword",
        algorithm_attribute="workbookAlgorithmName",
        hash_attribute="workbookHashValue",
        salt_attribute="workbookSaltValue",
        spin_count_attribute="workbookSpinCount",
        context="workbook-protection",
        warnings=warnings,
    )
    revisions_credential = _protection_credential_snapshot(
        element,
        legacy_attribute="revisionsPassword",
        algorithm_attribute="revisionsAlgorithmName",
        hash_attribute="revisionsHashValue",
        salt_attribute="revisionsSaltValue",
        spin_count_attribute="revisionsSpinCount",
        context="revision-protection",
        warnings=warnings,
    )
    return WorkbookProtectionSnapshot(
        lock_structure=_protection_bool(
            element.get("lockStructure"), False, "lockStructure", warnings
        ),
        lock_windows=_protection_bool(
            element.get("lockWindows"), False, "lockWindows", warnings
        ),
        lock_revision=_protection_bool(
            element.get("lockRevision"), False, "lockRevision", warnings
        ),
        workbook_credential=workbook_credential,
        revisions_credential=revisions_credential,
        opaque_metadata=_opaque_protection_metadata(
            element,
            known_attributes={
                "lockStructure",
                "lockWindows",
                "lockRevision",
                "workbookPassword",
                "workbookAlgorithmName",
                "workbookHashValue",
                "workbookSaltValue",
                "workbookSpinCount",
                "revisionsPassword",
                "revisionsAlgorithmName",
                "revisionsHashValue",
                "revisionsSaltValue",
                "revisionsSpinCount",
            },
        ),
    )


def _sheet_protection_snapshot(
    sheet: str,
    sheet_type: str,
    root: ElementTree.Element,
    warnings: set[str],
) -> SheetProtectionSnapshot | None:
    """Read one sheet-protection declaration with effective action defaults."""
    element = root.find(f"{{{_SPREADSHEETML_NS}}}sheetProtection")
    if element is None:
        return None
    actions = (
        _CHART_SHEET_PROTECTION_ACTIONS
        if sheet_type == "chartsheet"
        else _SHEET_PROTECTION_ACTIONS
    )
    enabled = (
        any(
            _protection_bool(element.get(attribute), default, attribute, warnings)
            for attribute, _name, default in actions
        )
        if sheet_type == "chartsheet"
        else _protection_bool(element.get("sheet"), False, "sheet", warnings)
    )
    locked_actions = (
        tuple(
            name
            for attribute, name, default in actions
            if _protection_bool(element.get(attribute), default, attribute, warnings)
        )
        if enabled
        else ()
    )
    credential = _protection_credential_snapshot(
        element,
        legacy_attribute="password",
        algorithm_attribute="algorithmName",
        hash_attribute="hashValue",
        salt_attribute="saltValue",
        spin_count_attribute="spinCount",
        context="sheet-protection",
        warnings=warnings,
    )
    known_attributes = {
        *(attribute for attribute, _name, _default in actions),
        "password",
        "algorithmName",
        "hashValue",
        "saltValue",
        "spinCount",
    }
    if sheet_type == "chartsheet":
        known_attributes.add("content")
    else:
        known_attributes.add("sheet")
    return SheetProtectionSnapshot(
        sheet=sheet,
        sheet_type=sheet_type,
        enabled=enabled,
        locked_actions=locked_actions,
        credential=credential,
        opaque_metadata=_opaque_protection_metadata(
            element,
            known_attributes=known_attributes,
        ),
    )


def _protected_range_snapshots(
    sheet: str,
    root: ElementTree.Element,
    warnings: set[str],
) -> tuple[tuple[ProtectedRangeSnapshot, ...], ProtectionOpaqueMetadataSnapshot]:
    """Read protected ranges while redacting names and security descriptors."""
    container = root.find(f"{{{_SPREADSHEETML_NS}}}protectedRanges")
    if container is None:
        return (), ProtectionOpaqueMetadataSnapshot()
    snapshots: list[ProtectedRangeSnapshot] = []
    range_tag = f"{{{_SPREADSHEETML_NS}}}protectedRange"
    for protected_range in container.findall(range_tag):
        ranges = _conditional_ranges(protected_range.get("sqref"))
        if not ranges:
            warnings.add(
                "FormulaFence found a protected range without target references; "
                "the affected permission has a coverage gap."
            )
        name = protected_range.get("name")
        if name is None:
            warnings.add(
                "FormulaFence found a protected range without its required name; "
                "the affected permission has a coverage gap."
            )
        security_descriptor = protected_range.get("securityDescriptor")
        credential = _protection_credential_snapshot(
            protected_range,
            legacy_attribute="password",
            algorithm_attribute="algorithmName",
            hash_attribute="hashValue",
            salt_attribute="saltValue",
            spin_count_attribute="spinCount",
            context="protected-range",
            warnings=warnings,
        )
        snapshots.append(
            ProtectedRangeSnapshot(
                sheet=sheet,
                ranges=ranges,
                has_name=name is not None,
                name_signature=_private_protection_signature(
                    (("name", name),) if name is not None else ()
                ),
                credential=credential,
                has_security_descriptor=security_descriptor is not None,
                security_descriptor_signature=_private_protection_signature(
                    (("securityDescriptor", security_descriptor),)
                    if security_descriptor is not None
                    else ()
                ),
                opaque_metadata=_opaque_protection_metadata(
                    protected_range,
                    known_attributes={
                        "name",
                        "sqref",
                        "password",
                        "algorithmName",
                        "hashValue",
                        "saltValue",
                        "spinCount",
                        "securityDescriptor",
                    },
                ),
            )
        )
    return (
        tuple(sorted(snapshots, key=ProtectedRangeSnapshot.sort_key)),
        _opaque_protection_metadata(
            container,
            known_attributes=set(),
            known_children={"protectedRange"},
        ),
    )


def _style_index(
    element: ElementTree.Element,
    attribute: str,
    warnings: set[str],
    *,
    context: str,
) -> int | None:
    """Read a zero-based style index without silently accepting malformed XML."""
    value = element.get(attribute)
    if value is None:
        return None
    try:
        parsed = int(value)
    except ValueError:
        warnings.add(
            "FormulaFence could not interpret a "
            f"{context} style index; the affected cell-protection assignment has a coverage gap."
        )
        return None
    if parsed < 0:
        warnings.add(
            "FormulaFence found a negative "
            f"{context} style index; the affected cell-protection assignment has a coverage gap."
        )
        return None
    return parsed


def _xf_protection(
    xf: ElementTree.Element,
    inherited: _StyleProtection,
    warnings: set[str],
    *,
    context: str,
) -> _StyleProtection:
    """Resolve the protection part of one XF without inspecting formatting data."""
    protection = xf.find(f"{{{_SPREADSHEETML_NS}}}protection")
    if protection is None:
        return inherited
    if (
        any(
            _xml_local_name(attribute) not in {"locked", "hidden"}
            for attribute in protection.attrib
        )
        or list(protection)
    ):
        warnings.add(
            "FormulaFence found unmodelled cell-protection style metadata; "
            "the affected assignment has a coverage gap."
        )
    applies = _protection_bool(
        xf.get("applyProtection"),
        True,
        "applyProtection",
        warnings,
    )
    if not applies:
        return inherited
    return _StyleProtection(
        locked=_protection_bool(
            protection.get("locked"), True, f"{context} locked", warnings
        ),
        hidden=_protection_bool(
            protection.get("hidden"), False, f"{context} hidden", warnings
        ),
        explicit=True,
    )


def _styles_with_protection(
    archive: ZipFile,
    warnings: set[str],
) -> tuple[_StyleProtection, tuple[_StyleProtection, ...]]:
    """Return the base default and effective cell-XF protection table."""
    default = _StyleProtection(locked=True, hidden=False)
    try:
        styles = _xml_root(archive, "xl/styles.xml")
    except KeyError:
        return default, (default,)
    style_xfs = styles.find(f"{{{_SPREADSHEETML_NS}}}cellStyleXfs")
    base_styles: list[_StyleProtection] = []
    if style_xfs is not None:
        for xf in style_xfs.findall(f"{{{_SPREADSHEETML_NS}}}xf"):
            base_styles.append(
                _xf_protection(xf, default, warnings, context="base-cell-style")
            )
    if not base_styles:
        base_styles.append(default)

    cell_xfs = styles.find(f"{{{_SPREADSHEETML_NS}}}cellXfs")
    effective_styles: list[_StyleProtection] = []
    if cell_xfs is not None:
        for xf in cell_xfs.findall(f"{{{_SPREADSHEETML_NS}}}xf"):
            xf_id = _style_index(xf, "xfId", warnings, context="cell-XF")
            inherited = default
            if xf_id is not None:
                if xf_id >= len(base_styles):
                    warnings.add(
                        "FormulaFence found a cell-XF with an unknown base style; "
                        "the affected cell-protection assignment has a coverage gap."
                    )
                else:
                    inherited = base_styles[xf_id]
            effective_styles.append(
                _xf_protection(xf, inherited, warnings, context="cell-style")
            )
    if not effective_styles:
        effective_styles.append(default)
    return effective_styles[0], tuple(effective_styles)


def _style_is_protection_relevant(
    style: _StyleProtection,
    default: _StyleProtection,
) -> bool:
    """Return whether a styled record can alter visible protection behavior."""
    return style.explicit or (style.locked, style.hidden) != (default.locked, default.hidden)


def _style_for_assignment(
    element: ElementTree.Element,
    attribute: str,
    styles: tuple[_StyleProtection, ...],
    warnings: set[str],
    *,
    context: str,
) -> _StyleProtection | None:
    """Resolve a direct style reference, returning ``None`` when absent/invalid."""
    style_index = _style_index(element, attribute, warnings, context=context)
    if style_index is None:
        return None
    if style_index >= len(styles):
        warnings.add(
            "FormulaFence found a "
            f"{context} style index outside the workbook style table; "
            "the affected cell-protection assignment has a coverage gap."
        )
        return None
    return styles[style_index]


def _column_span(
    element: ElementTree.Element,
    warnings: set[str],
) -> tuple[int, int, str] | None:
    """Return a compact column span from a raw ``col`` record."""
    start = _protection_int(element, "min", warnings, context="column-protection")
    end = _protection_int(element, "max", warnings, context="column-protection")
    if start is None or end is None or start < 1 or end < start or end > 16_384:
        warnings.add(
            "FormulaFence could not interpret a column-protection span; "
            "the affected assignment has a coverage gap."
        )
        return None
    try:
        start_column = get_column_letter(start)
        end_column = get_column_letter(end)
    except ValueError:
        warnings.add(
            "FormulaFence found a column-protection span outside Excel's column range; "
            "the affected assignment has a coverage gap."
        )
        return None
    return start, end, f"{start_column}:{end_column}"


def _cell_protection_assignments(
    archive: ZipFile,
    sheet_parts: Mapping[str, tuple[str, str]],
    sheet_protections: tuple[SheetProtectionSnapshot, ...],
    warnings: set[str],
) -> tuple[
    CellProtectionDefaultSnapshot | None,
    tuple[CellProtectionAssignmentSnapshot, ...],
]:
    """Read sparse direct cell/row/column protection assignments from OOXML.

    The inventory deliberately retains raw assignment scopes instead of
    expanding rows, columns, or styled rectangles into cells. It records only
    normal protected sheets, because locked/hidden cell styles are inactive on
    an unprotected sheet.
    """
    protected_sheets = {
        protection.sheet
        for protection in sheet_protections
        if protection.enabled and protection.sheet_type in {"worksheet", "dialogsheet"}
    }
    if not protected_sheets:
        return None, ()
    default, styles = _styles_with_protection(archive, warnings)
    assignments: set[CellProtectionAssignmentSnapshot] = set()
    row_tag = f"{{{_SPREADSHEETML_NS}}}row"
    cell_tag = f"{{{_SPREADSHEETML_NS}}}c"
    column_tag = f"{{{_SPREADSHEETML_NS}}}col"
    for sheet in sorted(protected_sheets, key=str.casefold):
        part = sheet_parts.get(sheet)
        if part is None:
            warnings.add(
                "FormulaFence could not locate OOXML for a protected worksheet; "
                "cell-protection assignments have a coverage gap."
            )
            continue
        member, _sheet_type = part
        worksheet = _xml_root(archive, member)
        raw_rows: list[tuple[int, _StyleProtection]] = []
        raw_columns: list[tuple[int, int, str, _StyleProtection]] = []
        for row in worksheet.iter(row_tag):
            style = _style_for_assignment(
                row, "s", styles, warnings, context="row-protection"
            )
            if style is None:
                continue
            row_number = _protection_int(row, "r", warnings, context="row-protection")
            if row_number is None or row_number < 1:
                warnings.add(
                    "FormulaFence could not interpret a row-protection target; "
                    "the affected assignment has a coverage gap."
                )
                continue
            raw_rows.append((row_number, style))
        for column in worksheet.iter(column_tag):
            style = _style_for_assignment(
                column, "style", styles, warnings, context="column-protection"
            )
            if style is None or (span := _column_span(column, warnings)) is None:
                continue
            start_column, end_column, target = span
            raw_columns.append((start_column, end_column, target, style))

        relevant_rows = {
            row_number
            for row_number, style in raw_rows
            if _style_is_protection_relevant(style, default)
        }
        relevant_columns = [
            column_target
            for _start_column, _end_column, column_target, style in raw_columns
            if _style_is_protection_relevant(style, default)
        ]
        for row_number, style in raw_rows:
            if not _style_is_protection_relevant(style, default) and not relevant_columns:
                continue
            assignments.add(
                CellProtectionAssignmentSnapshot(
                    sheet=sheet,
                    scope="row",
                    target=str(row_number),
                    locked=style.locked,
                    hidden=style.hidden,
                )
            )
        for _start_column, _end_column, target, style in raw_columns:
            if not _style_is_protection_relevant(style, default) and not relevant_rows:
                continue
            assignments.add(
                CellProtectionAssignmentSnapshot(
                    sheet=sheet,
                    scope="column",
                    target=target,
                    locked=style.locked,
                    hidden=style.hidden,
                )
            )

        for cell in worksheet.iter(cell_tag):
            style = _style_for_assignment(
                cell, "s", styles, warnings, context="cell-protection"
            )
            if style is None:
                continue
            coordinate = cell.get("r")
            if not coordinate:
                warnings.add(
                    "FormulaFence found a styled cell without a coordinate; "
                    "the affected cell-protection assignment has a coverage gap."
                )
                continue
            match = re.fullmatch(r"([A-Za-z]+)([1-9][0-9]*)", coordinate)
            if match is None:
                warnings.add(
                    "FormulaFence could not interpret a styled cell coordinate; "
                    "the affected cell-protection assignment has a coverage gap."
                )
                continue
            column_letters, raw_row_number = match.groups()
            row_number = int(raw_row_number)
            try:
                column_number = column_index_from_string(column_letters)
            except ValueError:
                warnings.add(
                    "FormulaFence could not interpret a styled cell column; "
                    "the affected cell-protection assignment has a coverage gap."
                )
                continue
            intersects_relevant_column = any(
                start_column <= column_number <= end_column
                for start_column, end_column, _target, style in raw_columns
                if _style_is_protection_relevant(style, default)
            )
            if not (
                _style_is_protection_relevant(style, default)
                or row_number in relevant_rows
                or intersects_relevant_column
            ):
                continue
            assignments.add(
                CellProtectionAssignmentSnapshot(
                    sheet=sheet,
                    scope="cell",
                    target=coordinate.upper(),
                    locked=style.locked,
                    hidden=style.hidden,
                )
            )
    return (
        CellProtectionDefaultSnapshot(locked=default.locked, hidden=default.hidden),
        tuple(sorted(assignments, key=CellProtectionAssignmentSnapshot.sort_key)),
    )


def _protection_metadata(path: Path) -> _ProtectionMetadata:
    """Read protection controls directly from OOXML before a library drops data."""
    warnings: set[str] = set()
    try:
        with ZipFile(path) as archive:
            workbook = _xml_root(archive, "xl/workbook.xml")
            sheet_parts = _sheet_xml_parts(archive)
            workbook_protection = _workbook_protection_snapshot(workbook, warnings)
            sheet_protections: dict[str, SheetProtectionSnapshot] = {}
            protected_ranges: list[ProtectedRangeSnapshot] = []
            for sheet, (member, sheet_type) in sheet_parts.items():
                root = _xml_root(archive, member)
                sheet_protection = _sheet_protection_snapshot(
                    sheet, sheet_type, root, warnings
                )
                ranges: tuple[ProtectedRangeSnapshot, ...] = ()
                range_container_opaque = ProtectionOpaqueMetadataSnapshot()
                if sheet_type != "chartsheet":
                    ranges, range_container_opaque = _protected_range_snapshots(
                        sheet, root, warnings
                    )
                    protected_ranges.extend(ranges)
                if sheet_protection is not None:
                    sheet_protections[sheet] = replace(
                        sheet_protection,
                        opaque_metadata=_merge_opaque_protection_metadata(
                            sheet_protection.opaque_metadata,
                            range_container_opaque,
                        ),
                    )
                elif range_container_opaque.present:
                    sheet_protections[sheet] = SheetProtectionSnapshot(
                        sheet=sheet,
                        sheet_type=sheet_type,
                        enabled=False,
                        locked_actions=(),
                        opaque_metadata=range_container_opaque,
                    )
            sorted_sheet_protections = tuple(
                sorted(sheet_protections.values(), key=SheetProtectionSnapshot.sort_key)
            )
            cell_default, cell_assignments = _cell_protection_assignments(
                archive,
                sheet_parts,
                sorted_sheet_protections,
                warnings,
            )
    except (BadZipFile, ElementTree.ParseError, KeyError, OSError, ValueError) as error:
        return _ProtectionMetadata(
            workbook_protection=None,
            sheet_protections=(),
            protected_ranges=(),
            cell_protection_default=None,
            cell_protection_assignments=(),
            warnings=(
                "FormulaFence could not inspect protection OOXML "
                f"({type(error).__name__}); protection controls were not compared.",
            ),
        )
    return _ProtectionMetadata(
        workbook_protection=workbook_protection,
        sheet_protections=sorted_sheet_protections,
        protected_ranges=tuple(
            sorted(protected_ranges, key=ProtectedRangeSnapshot.sort_key)
        ),
        cell_protection_default=cell_default,
        cell_protection_assignments=cell_assignments,
        warnings=tuple(sorted(warnings)),
    )


def _dynamic_metadata_indexes(archive: ZipFile) -> set[int]:
    """Return the one-based cell-metadata indexes marked as dynamic arrays."""
    try:
        metadata = _xml_root(archive, "xl/metadata.xml")
    except KeyError:
        return set()

    def tag(name: str) -> str:
        return f"{{{_SPREADSHEETML_NS}}}{name}"

    metadata_types = [
        metadata_type.get("name", "")
        for metadata_type in metadata.findall(f"./{tag('metadataTypes')}/{tag('metadataType')}")
    ]
    dynamic_future_indexes: dict[str, set[int]] = {}
    dynamic_properties_tag = f"{{{_DYNAMIC_ARRAY_NS}}}dynamicArrayProperties"
    for future_metadata in metadata.findall(tag("futureMetadata")):
        name = future_metadata.get("name")
        if not name:
            continue
        dynamic_indexes: set[int] = set()
        for index, metadata_block in enumerate(future_metadata.findall(tag("bk"))):
            if any(
                element.tag == dynamic_properties_tag
                and element.get("fDynamic", "").casefold() in {"1", "true"}
                for element in metadata_block.iter()
            ):
                dynamic_indexes.add(index)
        if dynamic_indexes:
            dynamic_future_indexes[name] = dynamic_indexes

    cell_metadata = metadata.find(tag("cellMetadata"))
    if cell_metadata is None:
        return set()
    dynamic_cell_indexes: set[int] = set()
    for cell_index, metadata_block in enumerate(cell_metadata.findall(tag("bk")), start=1):
        for record in metadata_block.findall(tag("rc")):
            try:
                type_index = int(record.get("t", ""))
                value_index = int(record.get("v", ""))
            except ValueError:
                continue
            if not 1 <= type_index <= len(metadata_types):
                continue
            type_name = metadata_types[type_index - 1]
            if value_index in dynamic_future_indexes.get(type_name, set()):
                dynamic_cell_indexes.add(cell_index)
                break
    return dynamic_cell_indexes


def _array_formula_metadata(path: Path) -> _ArrayFormulaMetadata:
    """Inspect raw markers that openpyxl intentionally does not expose."""
    dynamic_cells: set[CellKey] = set()
    unclassified_cells: set[CellKey] = set()
    scanned_sheets: set[str] = set()
    try:
        with ZipFile(path) as archive:
            dynamic_indexes = _dynamic_metadata_indexes(archive)
            for sheet, member in _worksheet_xml_paths(archive).items():
                worksheet = _xml_root(archive, member)
                scanned_sheets.add(sheet)
                for cell in worksheet.iter(f"{{{_SPREADSHEETML_NS}}}c"):
                    formula = cell.find(f"{{{_SPREADSHEETML_NS}}}f")
                    coordinate = cell.get("r")
                    if (
                        formula is None
                        or formula.get("t") != "array"
                        or not coordinate
                    ):
                        continue
                    try:
                        metadata_index = int(cell.get("cm", "0"))
                    except ValueError:
                        unclassified_cells.add((sheet, coordinate))
                        continue
                    if metadata_index <= 0:
                        continue
                    if metadata_index in dynamic_indexes:
                        dynamic_cells.add((sheet, coordinate))
                    else:
                        unclassified_cells.add((sheet, coordinate))
    except (BadZipFile, ElementTree.ParseError, OSError, ValueError) as error:
        return _ArrayFormulaMetadata(
            dynamic_cells=set(),
            unclassified_cells=set(),
            scanned_sheets=set(),
            complete=False,
            warnings=(
                "FormulaFence could not inspect array-formula OOXML metadata "
                f"({type(error).__name__}); fixed CSE output aliases were not traced.",
            ),
        )
    return _ArrayFormulaMetadata(
        dynamic_cells=dynamic_cells,
        unclassified_cells=unclassified_cells,
        scanned_sheets=scanned_sheets,
        complete=True,
        warnings=(),
    )


def _array_formula_range(sheet: str, cell: object) -> ArrayFormulaRange | None:
    """Return one canonical local output range when an array anchor is valid."""
    value = getattr(cell, "value", None)
    ref = getattr(value, "ref", None)
    if not isinstance(ref, str) or not ref.strip():
        return None
    try:
        min_column, min_row, max_column, max_row = range_boundaries(ref)
    except ValueError:
        return None
    if None in {min_column, min_row, max_column, max_row}:
        return None
    anchor = f"{get_column_letter(min_column)}{min_row}"
    if getattr(cell, "coordinate", "").upper() != anchor:
        return None
    endpoint = f"{get_column_letter(max_column)}{max_row}"
    canonical_ref = anchor if anchor == endpoint else f"{anchor}:{endpoint}"
    return ArrayFormulaRange(
        sheet=sheet,
        anchor=anchor,
        ref=canonical_ref,
        min_column=min_column,
        min_row=min_row,
        max_column=max_column,
        max_row=max_row,
    )


def _is_array_formula(cell: object) -> bool:
    """Return whether openpyxl exposed the cell as an OOXML array formula."""
    value = getattr(cell, "value", None)
    return (
        getattr(cell, "data_type", None) == "f"
        and isinstance(getattr(value, "ref", None), str)
        and _formula_text(value) is not None
    )


def _classify_array_formulas(
    workbook: object, metadata: _ArrayFormulaMetadata
) -> _ArrayFormulaClassification:
    """Classify array formulas without treating dynamic extents as fixed."""
    kinds: dict[CellKey, str] = {}
    refs: dict[CellKey, str] = {}
    legacy_ranges: list[ArrayFormulaRange] = []
    dynamic_cells: set[CellKey] = set()
    dynamic_ranges: list[ArrayFormulaRange] = []
    unclassified_cells: set[CellKey] = set()
    warnings = set(metadata.warnings)
    for worksheet in getattr(workbook, "worksheets", ()):
        for cell in worksheet._cells.values():  # noqa: SLF001 - sparse workbook safety
            if not _is_array_formula(cell):
                continue
            location = (worksheet.title, cell.coordinate)
            if not metadata.complete or worksheet.title not in metadata.scanned_sheets:
                kinds[location] = "unclassified"
                unclassified_cells.add(location)
                continue
            if location in metadata.dynamic_cells:
                kinds[location] = "dynamic"
                dynamic_cells.add(location)
                array_range = _array_formula_range(worksheet.title, cell)
                if array_range is None:
                    warnings.add(
                        "FormulaFence could not read one or more observed dynamic-array "
                        "output ranges; non-anchor output aliases were not traced for "
                        "those anchors."
                    )
                    continue
                refs[location] = array_range.ref
                dynamic_ranges.append(array_range)
                continue
            if location in metadata.unclassified_cells:
                kinds[location] = "unclassified"
                unclassified_cells.add(location)
                continue
            array_range = _array_formula_range(worksheet.title, cell)
            if array_range is None:
                kinds[location] = "unclassified"
                unclassified_cells.add(location)
                continue
            kinds[location] = "legacy_cse"
            refs[location] = array_range.ref
            legacy_ranges.append(array_range)

    if unclassified_cells:
        warnings.add(
            "FormulaFence could not classify one or more OOXML array formulas; "
            "fixed CSE output aliases were not traced for those cells."
        )
    return _ArrayFormulaClassification(
        kinds=kinds,
        refs=refs,
        legacy_ranges=tuple(legacy_ranges),
        dynamic_cells=dynamic_cells,
        dynamic_ranges=tuple(dynamic_ranges),
        unclassified_cells=unclassified_cells,
        warnings=tuple(sorted(warnings)),
    )


def _array_formula_output_dependents(
    legacy_ranges: tuple[ArrayFormulaRange, ...],
    reverse_dependencies: Mapping[CellKey, set[CellKey]],
    range_dependencies: list[RangeDependency],
) -> dict[CellKey, set[CellKey]]:
    """Link a CSE anchor to formulas that read any fixed result member.

    The output range is never expanded into millions of virtual cells. Instead
    each anchor receives direct aliases only to the already-known formula cells
    whose static references intersect that compact range.
    """
    ranges_by_sheet: dict[str, list[ArrayFormulaRange]] = defaultdict(list)
    for array_range in legacy_ranges:
        if array_range.has_multiple_outputs:
            ranges_by_sheet[array_range.sheet].append(array_range)

    dependents: dict[CellKey, set[CellKey]] = defaultdict(set)
    for source, formula_cells in reverse_dependencies.items():
        for array_range in ranges_by_sheet.get(source[0], ()):
            if array_range.contains(source):
                dependents[array_range.location].update(formula_cells)
    for dependency in range_dependencies:
        for array_range in ranges_by_sheet.get(dependency.source_sheet, ()):
            if dependency.intersects(array_range):
                dependents[array_range.location].add(dependency.dependent)
    return {anchor: cells for anchor, cells in dependents.items() if cells}


def _dynamic_array_output_dependents(
    dynamic_ranges: tuple[ArrayFormulaRange, ...],
    reverse_dependencies: Mapping[CellKey, set[CellKey]],
    range_dependencies: list[RangeDependency],
) -> tuple[
    dict[CellKey, set[CellKey]],
    dict[CellKey, tuple[DynamicArrayOutputReference, ...]],
]:
    """Link observed dynamic spill members to their current formula consumers.

    Dynamic-array output ranges may resize during recalculation, so this is not
    a fixed-range assertion. It records only a formula that currently reads at
    least one non-anchor member, without materializing individual spill cells.
    """
    ranges_by_sheet: dict[str, list[ArrayFormulaRange]] = defaultdict(list)
    for array_range in dynamic_ranges:
        if array_range.has_multiple_outputs:
            ranges_by_sheet[array_range.sheet].append(array_range)

    dependents: dict[CellKey, set[CellKey]] = defaultdict(set)
    references: dict[CellKey, set[DynamicArrayOutputReference]] = defaultdict(set)

    def add_reference(array_range: ArrayFormulaRange, dependent: CellKey) -> None:
        if dependent == array_range.location:
            return
        dependents[array_range.location].add(dependent)
        references[dependent].add(
            DynamicArrayOutputReference(
                anchor=array_range.location,
                observed_ref=array_range.ref,
            )
        )

    for source, formula_cells in reverse_dependencies.items():
        for array_range in ranges_by_sheet.get(source[0], ()):
            if source != array_range.location and array_range.contains(source):
                for dependent in formula_cells:
                    add_reference(array_range, dependent)
    for dependency in range_dependencies:
        for array_range in ranges_by_sheet.get(dependency.source_sheet, ()):
            if array_range.intersects_non_anchor(dependency):
                add_reference(array_range, dependency.dependent)

    return (
        {anchor: cells for anchor, cells in dependents.items() if cells},
        {
            dependent: tuple(
                sorted(
                    values,
                    key=lambda reference: (
                        reference.anchor[0].casefold(),
                        reference.anchor[1],
                        reference.observed_ref,
                    ),
                )
            )
            for dependent, values in references.items()
            if values
        },
    )


def _formula_text(value: object) -> str | None:
    if isinstance(value, str):
        return value
    # openpyxl represents array formulas as an object in recent versions.
    text = getattr(value, "text", None)
    return text if isinstance(text, str) else None


def _cell_snapshot(
    sheet: str,
    cell: object,
    *,
    array_formula_kind: str | None = None,
    array_formula_ref: str | None = None,
) -> CellSnapshot:
    coordinate = cell.coordinate
    value = cell.value
    data_type = cell.data_type
    formula = _formula_text(value) if data_type == "f" else None
    if formula is not None:
        return CellSnapshot(
            sheet=sheet,
            coordinate=coordinate,
            cell_type="formula",
            value=formula,
            value_type="formula",
            formula=formula,
            formula_fingerprint=formula_fingerprint(formula, coordinate),
            array_formula_kind=array_formula_kind,
            array_formula_ref=array_formula_ref,
        )
    cell_type = "error" if data_type == "e" else "value"
    return CellSnapshot(
        sheet=sheet,
        coordinate=coordinate,
        cell_type=cell_type,
        value=json_safe_value(value),
        value_type=type(value).__name__,
    )


def _calculation_settings(workbook: object) -> dict[str, object]:
    calculation = getattr(workbook, "calculation", None)
    if calculation is None:
        return {}
    return {
        field: value
        for field in _CALCULATION_FIELDS
        if (value := getattr(calculation, field, None)) is not None
    }


def _definition_text(definition: object) -> str:
    attr_text = getattr(definition, "attr_text", None)
    return str(attr_text if attr_text is not None else definition)


def _defined_names(workbook: object) -> dict[str, str]:
    """Inventory workbook and sheet-scoped names with unambiguous local keys."""
    names = getattr(workbook, "defined_names", {})
    result: dict[str, str] = {}
    try:
        items = names.items()
    except AttributeError:
        items = ()
    for name, definition in items:
        if getattr(definition, "localSheetId", None) is None:
            result[str(name)] = _definition_text(definition)
    for worksheet in getattr(workbook, "worksheets", ()):
        worksheet_names = getattr(worksheet, "defined_names", {})
        try:
            worksheet_items = worksheet_names.items()
        except AttributeError:
            continue
        for name, definition in worksheet_items:
            result[f"{worksheet.title}!{name}"] = _definition_text(definition)
    return result


def _named_destination_reference(
    name: str, sheet: str, coordinate: str
) -> ParsedReference | None:
    """Convert one workbook-defined-name destination into a static range."""
    parsed = parse_reference_token(f"{sheet}!{coordinate}")
    if parsed is None:
        return None
    return ParsedReference(
        parsed.sheet,
        parsed.min_column,
        parsed.min_row,
        parsed.max_column,
        parsed.max_row,
        raw=name,
        is_external=parsed.is_external,
    )


def _definition_references(name: str, definition: object) -> tuple[ParsedReference, ...]:
    """Return only static destinations from an ordinary defined name."""
    try:
        destinations = list(definition.destinations)
    except Exception:  # pragma: no cover - malformed name syntax is workbook-specific
        return ()
    return tuple(
        reference
        for sheet, coordinate in destinations
        if (reference := _named_destination_reference(name, sheet, coordinate)) is not None
    )


@dataclass(frozen=True)
class _FormulaDefinedName:
    """A defined name whose OOXML text is an Excel formula expression."""

    key: str
    formula: str
    scope: str | None


def _formula_defined_name(
    name: str, definition: object, scope: str | None
) -> _FormulaDefinedName | None:
    """Return a formula-valued defined name, excluding ordinary destinations.

    OOXML stores the text of a defined-name formula without its leading equals
    sign, while older writer output commonly retains one. ``DefinedName.type``
    lets us preserve ordinary A1 destinations and normalize either formula
    spelling before sending it through the formula tokenizer.
    """
    attr_text = getattr(definition, "attr_text", None)
    if not isinstance(attr_text, str):
        return None
    formula = attr_text.strip()
    if not formula:
        return None
    if not formula.startswith("="):
        try:
            definition_type = getattr(definition, "type", None)
        except Exception:  # pragma: no cover - malformed name text is workbook-specific
            definition_type = None
        if definition_type == "RANGE":
            return None
        formula = f"={formula}"
    return _FormulaDefinedName(reference_lookup_key(name), formula, scope)


def _qualified_name_key(sheet: str, name: str) -> str:
    """Build a lookup key for a sheet-local name without quote ambiguity."""
    escaped_sheet = sheet.replace("'", "''")
    return reference_lookup_key(f"'{escaped_sheet}'!{name}")


def _named_reference_maps(
    workbook: object,
    structured_tables: Mapping[str, StructuredTable],
    sheet_order: tuple[str, ...],
) -> tuple[
    dict[str, tuple[ParsedReference, ...]],
    dict[str, dict[str, tuple[ParsedReference, ...]]],
    dict[str, tuple[ParsedReference, ...] | None],
    dict[str, dict[str, tuple[ParsedReference, ...] | None]],
]:
    """Build direct, formula-defined, and callable named-LAMBDA maps.

    Formula-valued names are expanded only when every dependency is statically
    visible and internal. Relative references, dynamic functions, unresolved
    tokens, recursive LAMBDAs, external links, and 3-D spans remain unresolved
    at a use site instead of producing a guessed graph edge.
    """
    workbook_names = getattr(workbook, "defined_names", {})
    global_references: dict[str, tuple[ParsedReference, ...]] = {}
    global_formulas: dict[str, _FormulaDefinedName] = {}
    global_lambdas: dict[str, _FormulaDefinedName] = {}
    try:
        workbook_items = workbook_names.items()
    except AttributeError:
        workbook_items = ()
    for name, definition in workbook_items:
        if getattr(definition, "localSheetId", None) is not None:
            continue
        formula_definition = _formula_defined_name(str(name), definition, None)
        if formula_definition is not None:
            target = (
                global_lambdas
                if lambda_parameter_count(formula_definition.formula) is not None
                else global_formulas
            )
            target[formula_definition.key] = formula_definition
            continue
        references = _definition_references(str(name), definition)
        if references:
            global_references[reference_lookup_key(str(name))] = references

    local_references: dict[str, dict[str, tuple[ParsedReference, ...]]] = {}
    local_formulas: dict[str, dict[str, _FormulaDefinedName]] = {}
    local_lambdas: dict[str, dict[str, _FormulaDefinedName]] = {}
    sheet_titles: dict[str, str] = {}
    for worksheet in getattr(workbook, "worksheets", ()):
        scope = worksheet.title.casefold()
        sheet_titles[scope] = worksheet.title
        worksheet_names = getattr(worksheet, "defined_names", {})
        try:
            worksheet_items = worksheet_names.items()
        except AttributeError:
            continue
        sheet_references: dict[str, tuple[ParsedReference, ...]] = {}
        for name, definition in worksheet_items:
            formula_definition = _formula_defined_name(str(name), definition, scope)
            if formula_definition is not None:
                target = (
                    local_lambdas
                    if lambda_parameter_count(formula_definition.formula) is not None
                    else local_formulas
                )
                target.setdefault(scope, {})[formula_definition.key] = formula_definition
                continue
            references = _definition_references(str(name), definition)
            if not references:
                continue
            sheet_references[reference_lookup_key(str(name))] = references
        if sheet_references:
            local_references[scope] = sheet_references

    qualified_formulas = {
        _qualified_name_key(sheet_titles[scope], key): definition
        for scope, definitions in local_formulas.items()
        for key, definition in definitions.items()
    }
    qualified_lambdas = {
        _qualified_name_key(sheet_titles[scope], key): definition
        for scope, definitions in local_lambdas.items()
        for key, definition in definitions.items()
    }
    resolved_definitions: dict[tuple[str | None, str], tuple[ParsedReference, ...]] = {}
    resolving: set[tuple[str | None, str]] = set()
    failed: set[tuple[str | None, str]] = set()

    def identity_for(definition: _FormulaDefinedName) -> tuple[str | None, str]:
        return definition.scope, definition.key

    def cached_references(
        definition: _FormulaDefinedName,
    ) -> tuple[ParsedReference, ...] | None:
        return resolved_definitions.get(identity_for(definition))

    def has_cached_references(definition: _FormulaDefinedName) -> bool:
        return identity_for(definition) in resolved_definitions

    def visible_references(scope: str | None) -> dict[str, tuple[ParsedReference, ...]]:
        references = dict(global_references)
        for key, definition in global_formulas.items():
            if cached := cached_references(definition):
                references[key] = cached
            elif has_cached_references(definition):
                references[key] = ()
        for local_scope, values in local_references.items():
            for key, local_values in values.items():
                references[_qualified_name_key(sheet_titles[local_scope], key)] = local_values
        for local_scope, definitions in local_formulas.items():
            for key, definition in definitions.items():
                if cached := cached_references(definition):
                    references[_qualified_name_key(sheet_titles[local_scope], key)] = cached
                elif has_cached_references(definition):
                    references[_qualified_name_key(sheet_titles[local_scope], key)] = ()
        if scope is not None:
            references.update(local_references.get(scope, {}))
            for key, definition in local_formulas.get(scope, {}).items():
                if cached := cached_references(definition):
                    references[key] = cached
                elif has_cached_references(definition):
                    references[key] = ()
        return references

    def visible_named_functions(
        scope: str | None,
    ) -> dict[str, tuple[ParsedReference, ...] | None]:
        functions: dict[str, tuple[ParsedReference, ...] | None] = {
            key: cached_references(definition)
            for key, definition in global_lambdas.items()
        }
        for local_scope, definitions in local_lambdas.items():
            for key, definition in definitions.items():
                functions[_qualified_name_key(sheet_titles[local_scope], key)] = (
                    cached_references(definition)
                )
        if scope is not None:
            for key, definition in local_lambdas.get(scope, {}).items():
                functions[key] = cached_references(definition)
        return functions

    def named_definition_for(
        token: str, scope: str | None
    ) -> _FormulaDefinedName | None:
        key = reference_lookup_key(token)
        if qualified := qualified_formulas.get(key):
            return qualified
        if qualified := qualified_lambdas.get(key):
            return qualified
        if "!" in key:
            return None
        if scope is not None and (local := local_formulas.get(scope, {}).get(key)):
            return local
        if scope is not None and (local := local_lambdas.get(scope, {}).get(key)):
            return local
        return global_formulas.get(key) or global_lambdas.get(key)

    def resolve_definition(
        definition: _FormulaDefinedName,
    ) -> tuple[ParsedReference, ...] | None:
        identity = identity_for(definition)
        if has_cached_references(definition):
            return cached_references(definition)
        if identity in failed or identity in resolving:
            return None

        resolving.add(identity)
        resolved_references: tuple[ParsedReference, ...] | None = None
        try:
            inspection = inspect_formula(
                definition.formula,
                named_references=visible_references(definition.scope),
                structured_tables=structured_tables,
                sheet_order=sheet_order,
                named_function_references=visible_named_functions(definition.scope),
            )
            can_expand = not (
                has_broken_reference(definition.formula)
                or inspection.tokenization_failed
                or inspection.dynamic_reference_functions
                or inspection.three_d_reference_tokens
                or inspection.spill_reference_tokens
                or inspection.implicit_intersection_tokens
            )
            if can_expand:
                for token in inspection.unresolved_range_tokens:
                    dependency = named_definition_for(token, definition.scope)
                    if dependency is None or resolve_definition(dependency) is None:
                        can_expand = False
                        break
            if can_expand:
                inspection = inspect_formula(
                    definition.formula,
                    named_references=visible_references(definition.scope),
                    structured_tables=structured_tables,
                    sheet_order=sheet_order,
                    named_function_references=visible_named_functions(definition.scope),
                )
                if (
                    inspection.unresolved_range_tokens
                    or inspection.tokenization_failed
                    or inspection.dynamic_reference_functions
                    or inspection.three_d_reference_tokens
                    or inspection.spill_reference_tokens
                    or inspection.implicit_intersection_tokens
                    or any(
                        reference.is_external or reference.sheet is None
                        for reference in inspection.references
                    )
                ):
                    can_expand = False
            if can_expand:
                resolved_references = tuple(dict.fromkeys(inspection.references))
        finally:
            resolving.remove(identity)

        if resolved_references is None:
            failed.add(identity)
            return None
        resolved_definitions[identity] = resolved_references
        return resolved_references

    for definition in global_formulas.values():
        resolve_definition(definition)
    for definition in global_lambdas.values():
        resolve_definition(definition)
    for definitions in local_formulas.values():
        for definition in definitions.values():
            resolve_definition(definition)
    for definitions in local_lambdas.values():
        for definition in definitions.values():
            resolve_definition(definition)

    global_result = dict(global_references)
    for key, definition in global_formulas.items():
        if cached := cached_references(definition):
            global_result[key] = cached
        elif has_cached_references(definition):
            global_result[key] = ()
    for scope, values in local_references.items():
        for key, references in values.items():
            global_result[_qualified_name_key(sheet_titles[scope], key)] = references
    for scope, definitions in local_formulas.items():
        for key, definition in definitions.items():
            if cached := cached_references(definition):
                global_result[_qualified_name_key(sheet_titles[scope], key)] = cached
            elif has_cached_references(definition):
                global_result[_qualified_name_key(sheet_titles[scope], key)] = ()

    local_result: dict[str, dict[str, tuple[ParsedReference, ...]]] = {}
    for scope in set(local_references) | set(local_formulas):
        references = dict(local_references.get(scope, {}))
        for key, definition in local_formulas.get(scope, {}).items():
            if cached := cached_references(definition):
                references[key] = cached
            elif has_cached_references(definition):
                references[key] = ()
        if references:
            local_result[scope] = references

    global_function_result: dict[str, tuple[ParsedReference, ...] | None] = {
        key: cached_references(definition)
        for key, definition in global_lambdas.items()
    }
    for scope, definitions in local_lambdas.items():
        for key, definition in definitions.items():
            global_function_result[_qualified_name_key(sheet_titles[scope], key)] = (
                cached_references(definition)
            )

    local_function_result: dict[
        str, dict[str, tuple[ParsedReference, ...] | None]
    ] = {}
    for scope, definitions in local_lambdas.items():
        functions = {
            key: cached_references(definition) for key, definition in definitions.items()
        }
        if functions:
            local_function_result[scope] = functions
    return global_result, local_result, global_function_result, local_function_result


def _table_columns(
    worksheet: object,
    table: object,
    min_column: int,
    min_row: int,
    max_column: int,
    header_row_count: int,
) -> tuple[str, ...]:
    """Read table column labels, falling back to the inspectable header cells."""
    table_columns = tuple(getattr(table, "tableColumns", ()) or ())
    names = tuple(str(getattr(column, "name", "")) for column in table_columns)
    width = max_column - min_column + 1
    if len(names) == width and all(names):
        return names
    if header_row_count:
        return tuple(
            str(value) if (value := worksheet.cell(min_row, column).value) is not None else ""
            for column in range(min_column, max_column + 1)
        )
    return ()


def _table_snapshots(workbook: object) -> dict[str, TableSnapshot]:
    """Inventory Excel-table definitions that affect structured references."""
    result: dict[str, TableSnapshot] = {}
    for worksheet in getattr(workbook, "worksheets", ()):
        table_list = getattr(worksheet, "tables", {})
        try:
            table_values = table_list.values()
        except AttributeError:
            continue
        for table in table_values:
            name = str(
                getattr(table, "displayName", None)
                or getattr(table, "name", None)
                or ""
            )
            ref = getattr(table, "ref", None)
            if not name or not isinstance(ref, str):
                continue
            try:
                min_column, min_row, max_column, max_row = range_boundaries(ref)
            except ValueError:
                continue
            height = max_row - min_row + 1
            header_rows = min(max(int(getattr(table, "headerRowCount", 1) or 0), 0), height)
            totals_rows = min(
                max(int(getattr(table, "totalsRowCount", 0) or 0), 0), height - header_rows
            )
            result[name] = TableSnapshot(
                name=name,
                sheet=worksheet.title,
                ref=ref,
                columns=_table_columns(
                    worksheet,
                    table,
                    min_column,
                    min_row,
                    max_column,
                    header_rows,
                ),
                header_row_count=header_rows,
                totals_row_count=totals_rows,
            )
    return result


def _validation_formula(value: object) -> str | None:
    """Normalize the optional leading equals sign used by workbook writers."""
    if not isinstance(value, str):
        return None
    return value[1:] if value.startswith("=") else value


def _validation_ranges(validation: object) -> tuple[str, ...]:
    """Return a compact, writer-order-independent validation target inventory."""
    sqref = getattr(validation, "sqref", None)
    try:
        ranges = tuple(str(target_range) for target_range in sqref.ranges)
    except (AttributeError, TypeError):
        raw = str(sqref or "")
        ranges = tuple(raw.split())
    return tuple(sorted(set(ranges), key=str.casefold))


def _validation_bool(validation: object, attribute: str, default: bool) -> bool:
    """Read one OOXML boolean while applying its schema default."""
    value = getattr(validation, attribute, None)
    return default if value is None else bool(value)


def _data_validation_snapshots(workbook: object) -> tuple[DataValidationSnapshot, ...]:
    """Inventory data-entry controls without expanding their applied ranges.

    OOXML defaults are normalized because writers commonly omit values such as
    ``operator=between`` and ``errorStyle=stop`` while other writers serialize
    them explicitly. The control's formulas and messages remain available in a
    local diff but are omitted from profile output.
    """
    snapshots: list[DataValidationSnapshot] = []
    for worksheet in getattr(workbook, "worksheets", ()):
        validation_container = getattr(worksheet, "data_validations", None)
        validation_list = getattr(validation_container, "dataValidation", ())
        prompts_disabled = _validation_bool(validation_container, "disablePrompts", False)
        for validation in validation_list or ():
            ranges = _validation_ranges(validation)
            if not ranges:
                continue
            snapshots.append(
                DataValidationSnapshot(
                    sheet=worksheet.title,
                    ranges=ranges,
                    validation_type=str(getattr(validation, "type", None) or "none"),
                    operator=str(getattr(validation, "operator", None) or "between"),
                    formula1=_validation_formula(getattr(validation, "formula1", None)),
                    formula2=_validation_formula(getattr(validation, "formula2", None)),
                    allow_blank=_validation_bool(validation, "allowBlank", False),
                    dropdown_hidden=_validation_bool(validation, "showDropDown", False),
                    prompts_disabled=prompts_disabled,
                    show_input_message=_validation_bool(
                        validation, "showInputMessage", False
                    ),
                    show_error_message=_validation_bool(
                        validation, "showErrorMessage", False
                    ),
                    error_style=str(getattr(validation, "errorStyle", None) or "stop"),
                    error_title=getattr(validation, "errorTitle", None),
                    error=getattr(validation, "error", None),
                    prompt_title=getattr(validation, "promptTitle", None),
                    prompt=getattr(validation, "prompt", None),
                    ime_mode=str(getattr(validation, "imeMode", None) or "noControl"),
                )
            )
    # ``sqref`` is only a compact list of targets for a rule. Writers may keep
    # identical rules together or split them into separate ``dataValidation``
    # elements, with no change to what Excel applies to any target cell.
    # Canonicalize that grouping before comparing workbooks.
    grouped_ranges: dict[DataValidationSnapshot, set[str]] = defaultdict(set)
    for snapshot in snapshots:
        grouped_ranges[replace(snapshot, ranges=())].update(snapshot.ranges)
    return tuple(
        sorted(
            (
                replace(
                    snapshot,
                    ranges=tuple(sorted(ranges, key=str.casefold)),
                )
                for snapshot, ranges in grouped_ranges.items()
            ),
            key=DataValidationSnapshot.sort_key,
        )
    )


def _structured_table_map(tables: dict[str, TableSnapshot]) -> dict[str, StructuredTable]:
    """Translate stable table inventory records into formula-resolution metadata."""
    return {
        name.casefold(): StructuredTable(
            name=table.name,
            sheet=table.sheet,
            ref=table.ref,
            columns=table.columns,
            header_row_count=table.header_row_count,
            totals_row_count=table.totals_row_count,
        )
        for name, table in tables.items()
    }


def load_snapshot(path: str | Path) -> WorkbookSnapshot:
    """Load a workbook as a semantic snapshot without evaluating its contents."""
    source = Path(path)
    if not source.exists() or not source.is_file():
        raise WorkbookLoadError(f"Workbook does not exist or is not a file: {source}")
    if source.suffix.lower() not in _SUPPORTED_SUFFIXES:
        supported = ", ".join(sorted(_SUPPORTED_SUFFIXES))
        raise WorkbookLoadError(
            f"Unsupported workbook type {source.suffix!r}; supported types: {supported}"
        )

    try:
        with warnings.catch_warnings(record=True) as caught_warnings:
            warnings.simplefilter("always")
            workbook = load_workbook(
                source,
                read_only=False,
                data_only=False,
                keep_vba=False,
                keep_links=False,
                rich_text=False,
            )
    except (BadZipFile, InvalidFileException, OSError, ValueError) as error:
        raise WorkbookLoadError(f"Could not read workbook {source}: {error}") from error
    parser_warnings = {str(warning.message) for warning in caught_warnings}
    has_array_formulas = any(
        _is_array_formula(cell)
        for worksheet in workbook.worksheets
        for cell in worksheet._cells.values()  # noqa: SLF001 - sparse workbook safety
    )
    if has_array_formulas:
        array_formula_classification = _classify_array_formulas(
            workbook, _array_formula_metadata(source)
        )
    else:
        array_formula_classification = _ArrayFormulaClassification(
            kinds={},
            refs={},
            legacy_ranges=(),
            dynamic_cells=set(),
            dynamic_ranges=(),
            unclassified_cells=set(),
            warnings=(),
        )
    parser_warnings.update(array_formula_classification.warnings)
    conditional_formatting_metadata = _conditional_formatting_metadata(source)
    parser_warnings.update(conditional_formatting_metadata.warnings)
    protection_metadata = _protection_metadata(source)
    parser_warnings.update(protection_metadata.warnings)

    sheets: dict[str, SheetSnapshot] = {}
    cells: dict[CellKey, CellSnapshot] = {}
    reverse_dependencies: dict[CellKey, set[CellKey]] = defaultdict(set)
    range_dependencies: list[RangeDependency] = []
    external_references: set[CellKey] = set()
    broken_references: set[CellKey] = set()
    unresolved_reference_tokens: dict[CellKey, tuple[str, ...]] = {}
    dynamic_reference_functions: dict[CellKey, tuple[str, ...]] = {}
    three_d_reference_tokens: dict[CellKey, tuple[str, ...]] = {}
    spill_reference_tokens: dict[CellKey, tuple[str, ...]] = {}
    implicit_intersection_tokens: dict[CellKey, tuple[str, ...]] = {}
    tokenization_failure_cells: set[CellKey] = set()
    tables = _table_snapshots(workbook)
    data_validations = _data_validation_snapshots(workbook)
    structured_tables = _structured_table_map(tables)
    sheet_order = tuple(worksheet.title for worksheet in workbook.worksheets)
    (
        global_named_references,
        local_named_references,
        global_named_functions,
        local_named_functions,
    ) = _named_reference_maps(workbook, structured_tables, sheet_order)

    for worksheet in workbook.worksheets:
        named_references = {
            **global_named_references,
            **local_named_references.get(worksheet.title.casefold(), {}),
        }
        named_functions = {
            **global_named_functions,
            **local_named_functions.get(worksheet.title.casefold(), {}),
        }
        nonempty_cells = 0
        formula_cells = 0
        # _cells lets us avoid traversing a sheet's whole used rectangle when a
        # workbook has a sparse, accidentally enormous dimension.
        worksheet_cells = sorted(
            worksheet._cells.values(),  # noqa: SLF001 - needed for sparse workbook safety
            key=lambda current: (current.row, current.column),
        )
        for cell in worksheet_cells:
            if cell.value is None:
                continue
            location = (worksheet.title, cell.coordinate)
            snapshot = _cell_snapshot(
                worksheet.title,
                cell,
                array_formula_kind=array_formula_classification.kinds.get(location),
                array_formula_ref=array_formula_classification.refs.get(location),
            )
            cells[snapshot.location] = snapshot
            nonempty_cells += 1
            if not snapshot.is_formula or snapshot.formula is None:
                continue

            formula_cells += 1
            if has_broken_reference(snapshot.formula):
                broken_references.add(snapshot.location)
            inspection = inspect_formula(
                snapshot.formula,
                named_references=named_references,
                structured_tables=structured_tables,
                origin=snapshot.location,
                sheet_order=sheet_order,
                named_function_references=named_functions,
            )
            if inspection.unresolved_range_tokens:
                unresolved_reference_tokens[snapshot.location] = inspection.unresolved_range_tokens
            if inspection.tokenization_failed:
                tokenization_failure_cells.add(snapshot.location)
            if inspection.dynamic_reference_functions:
                dynamic_reference_functions[snapshot.location] = (
                    inspection.dynamic_reference_functions
                )
            if inspection.three_d_reference_tokens:
                three_d_reference_tokens[snapshot.location] = inspection.three_d_reference_tokens
            if inspection.spill_reference_tokens:
                spill_reference_tokens[snapshot.location] = inspection.spill_reference_tokens
            if inspection.implicit_intersection_tokens:
                implicit_intersection_tokens[snapshot.location] = (
                    inspection.implicit_intersection_tokens
                )
            for reference in inspection.references:
                if reference.is_external:
                    external_references.add(snapshot.location)
                    continue
                if None in {
                    reference.min_column,
                    reference.min_row,
                    reference.max_column,
                    reference.max_row,
                }:
                    continue
                source_sheet = reference.sheet or worksheet.title
                if reference.is_range:
                    range_dependencies.append(
                        RangeDependency(
                            source_sheet=source_sheet,
                            min_column=reference.min_column,
                            min_row=reference.min_row,
                            max_column=reference.max_column,
                            max_row=reference.max_row,
                            dependent=snapshot.location,
                        )
                    )
                else:
                    source_coordinate = (
                        f"{get_column_letter(reference.min_column)}{reference.min_row}"
                    )
                    reverse_dependencies[(source_sheet, source_coordinate)].add(snapshot.location)

        sheets[worksheet.title] = SheetSnapshot(
            title=worksheet.title,
            state=worksheet.sheet_state,
            nonempty_cells=nonempty_cells,
            formula_cells=formula_cells,
            max_row=worksheet.max_row,
            max_column=worksheet.max_column,
        )

    legacy_array_formula_output_dependents = _array_formula_output_dependents(
        array_formula_classification.legacy_ranges,
        reverse_dependencies,
        range_dependencies,
    )
    (
        dynamic_array_formula_output_dependents,
        dynamic_array_output_references,
    ) = _dynamic_array_output_dependents(
        array_formula_classification.dynamic_ranges,
        reverse_dependencies,
        range_dependencies,
    )
    array_formula_output_dependents = {
        anchor: set(legacy_array_formula_output_dependents.get(anchor, set()))
        | dynamic_array_formula_output_dependents.get(anchor, set())
        for anchor in (
            set(legacy_array_formula_output_dependents)
            | set(dynamic_array_formula_output_dependents)
        )
    }

    return WorkbookSnapshot(
        path=source,
        sha256=sha256_file(source),
        file_type=source.suffix.lower().lstrip("."),
        sheets=sheets,
        cells=cells,
        reverse_dependencies=dict(reverse_dependencies),
        range_dependencies=range_dependencies,
        external_references=external_references,
        broken_references=broken_references,
        unresolved_reference_tokens=unresolved_reference_tokens,
        dynamic_reference_functions=dynamic_reference_functions,
        three_d_reference_tokens=three_d_reference_tokens,
        spill_reference_tokens=spill_reference_tokens,
        implicit_intersection_tokens=implicit_intersection_tokens,
        legacy_array_formula_ranges=array_formula_classification.legacy_ranges,
        dynamic_array_formula_cells=array_formula_classification.dynamic_cells,
        dynamic_array_formula_ranges=array_formula_classification.dynamic_ranges,
        dynamic_array_output_references=dynamic_array_output_references,
        unclassified_array_formula_cells=array_formula_classification.unclassified_cells,
        array_formula_output_dependents=array_formula_output_dependents,
        tokenization_failure_cells=tokenization_failure_cells,
        tables=tables,
        data_validations=data_validations,
        conditional_formatting=conditional_formatting_metadata.rules,
        conditional_formatting_extensions=conditional_formatting_metadata.extensions,
        workbook_protection=protection_metadata.workbook_protection,
        sheet_protections=protection_metadata.sheet_protections,
        protected_ranges=protection_metadata.protected_ranges,
        cell_protection_default=protection_metadata.cell_protection_default,
        cell_protection_assignments=protection_metadata.cell_protection_assignments,
        sheet_order=sheet_order,
        defined_names=_defined_names(workbook),
        macro_hash=_vba_hash(source),
        calculation_settings=_calculation_settings(workbook),
        parser_warnings=tuple(sorted(parser_warnings)),
    )


def profile_snapshot(snapshot: WorkbookSnapshot) -> dict[str, object]:
    """Return a data-minimising inventory suitable for a safe review artifact."""
    return {
        "schema_version": "1.0",
        "workbook": snapshot.summary(),
        "sheets": [sheet.to_dict() for sheet in snapshot.sheets.values()],
        "tables": [
            snapshot.tables[name].to_dict()
            for name in sorted(snapshot.tables, key=str.casefold)
        ],
        "data_validations": [
            validation.profile_dict() for validation in snapshot.data_validations
        ],
        "conditional_formatting": [
            rule.profile_dict() for rule in snapshot.conditional_formatting
        ],
        "conditional_formatting_extensions": [
            extension.profile_dict()
            for extension in snapshot.conditional_formatting_extensions
        ],
        "workbook_protection": (
            snapshot.workbook_protection.to_dict()
            if snapshot.workbook_protection is not None
            else None
        ),
        "sheet_protections": [
            protection.profile_dict() for protection in snapshot.sheet_protections
        ],
        "protected_ranges": [
            protected_range.profile_dict()
            for protected_range in snapshot.protected_ranges
        ],
        "cell_protection_default": (
            snapshot.cell_protection_default.to_dict()
            if snapshot.cell_protection_default is not None
            else None
        ),
        "cell_protection_assignments": [
            assignment.profile_dict()
            for assignment in snapshot.cell_protection_assignments
        ],
        "defined_names": sorted(snapshot.defined_names),
        "calculation_settings": snapshot.calculation_settings,
        "features": {
            "external_reference_cells": [
                f"{sheet}!{coordinate}"
                for sheet, coordinate in sorted(snapshot.external_references)
            ],
            "broken_reference_cells": [
                f"{sheet}!{coordinate}"
                for sheet, coordinate in sorted(snapshot.broken_references)
            ],
            "has_vba": snapshot.macro_hash is not None,
            "parser_warnings": list(snapshot.parser_warnings),
            "unresolved_reference_cells": [
                {
                    "location": display_location(location),
                    "tokens": list(tokens),
                }
                for location, tokens in sorted(snapshot.unresolved_reference_tokens.items())
            ],
            "dynamic_reference_cells": [
                {
                    "location": display_location(location),
                    "functions": list(functions),
                }
                for location, functions in sorted(snapshot.dynamic_reference_functions.items())
            ],
            "three_d_reference_cells": [
                {
                    "location": display_location(location),
                    "tokens": list(tokens),
                }
                for location, tokens in sorted(snapshot.three_d_reference_tokens.items())
            ],
            "spill_reference_cells": [
                {
                    "location": display_location(location),
                    "tokens": list(tokens),
                }
                for location, tokens in sorted(snapshot.spill_reference_tokens.items())
            ],
            "implicit_intersection_cells": [
                {
                    "location": display_location(location),
                    "tokens": list(tokens),
                }
                for location, tokens in sorted(
                    snapshot.implicit_intersection_tokens.items()
                )
            ],
            "legacy_array_formula_ranges": [
                array_range.to_dict()
                for array_range in sorted(
                    snapshot.legacy_array_formula_ranges,
                    key=lambda array_range: (array_range.sheet.casefold(), array_range.anchor),
                )
            ],
            "dynamic_array_formula_cells": [
                display_location(location)
                for location in sorted(snapshot.dynamic_array_formula_cells)
            ],
            "dynamic_array_observed_output_ranges": [
                array_range.to_dict()
                for array_range in sorted(
                    snapshot.dynamic_array_formula_ranges,
                    key=lambda array_range: (array_range.sheet.casefold(), array_range.anchor),
                )
            ],
            "dynamic_array_output_reference_cells": [
                {
                    "location": display_location(location),
                    "references": [reference.to_dict() for reference in references],
                }
                for location, references in sorted(
                    snapshot.dynamic_array_output_references.items()
                )
            ],
            "unclassified_array_formula_cells": [
                display_location(location)
                for location in sorted(snapshot.unclassified_array_formula_cells)
            ],
            "tokenization_failure_cells": [
                display_location(location)
                for location in sorted(snapshot.tokenization_failure_cells)
            ],
        },
    }
