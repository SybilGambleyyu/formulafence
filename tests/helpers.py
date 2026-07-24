from __future__ import annotations

import base64
import io
import struct
from collections.abc import Callable
from pathlib import Path
from xml.etree import ElementTree
from zipfile import ZIP_DEFLATED, ZipFile

from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.chartsheet.protection import ChartsheetProtection
from openpyxl.formatting.rule import (
    CellIsRule,
    ColorScaleRule,
    DataBarRule,
    FormulaRule,
    IconSetRule,
)
from openpyxl.styles import Font, PatternFill, Protection
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.formula import ArrayFormula
from openpyxl.worksheet.table import Table

_DATA_MASHUP_NS = "http://schemas.microsoft.com/DataMashup"
_PACKAGE_RELATIONSHIPS_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
_DOCUMENT_RELATIONSHIPS_NS = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
)
_SPREADSHEETML_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"


def make_model(path: Path) -> Path:
    """Create a small, multi-sheet financial-model-shaped fixture."""
    workbook = Workbook()
    inputs = workbook.active
    inputs.title = "Inputs"
    inputs["A1"] = "Revenue"
    inputs["B2"] = 100
    inputs["B3"] = 150
    inputs["B4"] = 175

    model = workbook.create_sheet("Model")
    model["A1"] = "Calculated revenue"
    model["B2"] = "=Inputs!B2*2"
    model["B3"] = "=Inputs!B3*2"
    model["B4"] = "=Inputs!B4*2"
    model["C2"] = "=B2+10"

    dashboard = workbook.create_sheet("Dashboard")
    dashboard["A1"] = "Headline output"
    dashboard["B12"] = "=Model!C2"

    control = workbook.create_sheet("Control")
    control["A1"] = "Do not hide or alter casually"
    workbook.defined_names.add(DefinedName("HeadlineOutput", attr_text="Dashboard!$B$12"))
    workbook.save(path)
    return path


def make_table_model(path: Path) -> Path:
    """Create a small workbook with static Excel-table references."""
    workbook = Workbook()
    data = workbook.active
    data.title = "Data"
    data.append(["Amount", "Rate", "Value"])
    data.append([10, 0.1, 1])
    data.append([20, 0.2, 4])
    data.append([30, 0.3, 9])
    data.add_table(Table(displayName="Sales", ref="A1:C4"))

    report = workbook.create_sheet("Report")
    report["A1"] = "Table-driven outputs"
    report["B2"] = "=SUM(Sales[Amount])"
    report["B3"] = "=SUM(Sales[[#Data],[Amount]:[Rate]])"
    report["B4"] = "=ROWS(Sales[#All])"
    report["B5"] = "=COUNTA(Sales[[#Headers],[#Data],[Rate]])"
    workbook.save(path)
    return path


def make_data_validation_model(path: Path, *, reverse_status_targets: bool = False) -> Path:
    """Create an input workbook whose data-entry controls carry business rules."""
    workbook = Workbook()
    inputs = workbook.active
    inputs.title = "Inputs"
    inputs["A1"] = "Controlled entries"
    inputs["B1"] = "Status"
    inputs["B2"] = "Draft"
    inputs["C1"] = "Approved amount"
    inputs["C2"] = 25
    inputs["D1"] = "Secondary status"
    inputs["D2"] = "Review"

    limits = workbook.create_sheet("Limits")
    limits["A1"] = "Approved statuses"
    limits["A2"] = "Draft"
    limits["A3"] = "Review"
    limits["A4"] = "Approved"
    limits["B1"] = "Minimum"
    limits["B2"] = 10
    limits["B3"] = 100

    status = DataValidation(
        type="list",
        formula1="=Limits!$A$2:$A$4",
        allow_blank=True,
        showInputMessage=True,
        showErrorMessage=True,
        errorStyle="stop",
        errorTitle="Invalid status",
        error="Choose an approved status.",
        promptTitle="Approved status",
        prompt="Choose a documented status.",
    )
    if reverse_status_targets:
        status.add("D2")
        status.add("B2:B100")
    else:
        status.add("B2:B100")
        status.add("D2")
    inputs.add_data_validation(status)

    amount = DataValidation(
        type="decimal",
        operator="between",
        formula1="=Limits!$B$2",
        formula2="=Limits!$B$3",
        allow_blank=False,
        showErrorMessage=True,
        errorStyle="warning",
        errorTitle="Outside approved range",
        error="Use an amount within the approved range.",
    )
    amount.add("C2:C100")
    inputs.add_data_validation(amount)
    workbook.save(path)
    return path


def make_conditional_formatting_model(path: Path) -> Path:
    """Create visual review controls with overlapping precedence and builtins."""
    workbook = Workbook()
    inputs = workbook.active
    inputs.title = "Inputs"
    inputs["A1"] = "Controlled metric"
    for row, value in enumerate((10, -5, 75, 120), start=2):
        inputs.cell(row, 1, value)

    red_fill = PatternFill(fill_type="solid", fgColor="FFFFC7CE")
    red_font = Font(color="FF9C0006")
    green_fill = PatternFill(fill_type="solid", fgColor="FFC6EFCE")
    green_font = Font(color="FF006100")
    inputs.conditional_formatting.add(
        "A2:A100",
        FormulaRule(
            formula=["$A2<0"],
            stopIfTrue=True,
            fill=red_fill,
            font=red_font,
        ),
    )
    inputs.conditional_formatting.add(
        "A2:A100",
        CellIsRule(
            operator="greaterThan",
            formula=["100"],
            fill=green_fill,
            font=green_font,
        ),
    )
    inputs.conditional_formatting.add(
        "B2:B100",
        ColorScaleRule(
            start_type="min",
            start_color="FFF8696B",
            mid_type="percentile",
            mid_value=50,
            mid_color="FFFFEB84",
            end_type="max",
            end_color="FF63BE7B",
        ),
    )
    inputs.conditional_formatting.add(
        "C2:C100",
        DataBarRule(
            start_type="min",
            start_value=None,
            end_type="max",
            end_value=None,
            color="FF638EC6",
        ),
    )
    inputs.conditional_formatting.add(
        "D2:D100",
        IconSetRule(
            "3TrafficLights1",
            "percent",
            [0, 33, 67],
            showValue=None,
            percent=None,
            reverse=None,
        ),
    )
    workbook.save(path)
    return path


def _rewrite_archive(path: Path, mutate: Callable[[dict[str, bytes]], None], suffix: str) -> Path:
    """Apply a small raw OOXML test mutation without changing ZIP member order."""
    with ZipFile(path) as archive:
        contents = {
            entry.filename: archive.read(entry.filename) for entry in archive.infolist()
        }
    mutate(contents)
    staging = path.with_suffix(suffix)
    with ZipFile(staging, "w", compression=ZIP_DEFLATED) as archive:
        for name, content in contents.items():
            archive.writestr(name, content)
    staging.replace(path)
    return path


def _inputs_worksheet_root(contents: dict[str, bytes]) -> ElementTree.Element:
    """Return the fixture's first worksheet XML root."""
    return ElementTree.fromstring(contents["xl/worksheets/sheet1.xml"])


def _save_inputs_worksheet(contents: dict[str, bytes], root: ElementTree.Element) -> None:
    contents["xl/worksheets/sheet1.xml"] = ElementTree.tostring(
        root,
        encoding="utf-8",
        xml_declaration=True,
    )


def _zip_payload(parts: dict[str, bytes]) -> bytes:
    """Build one in-memory OPC-like package for a Data Mashup test fixture."""
    payload = io.BytesIO()
    with ZipFile(payload, "w", compression=ZIP_DEFLATED) as archive:
        for name, value in parts.items():
            archive.writestr(name, value)
    return payload.getvalue()


def _synthetic_power_query_mashup(
    *,
    formula: str,
    fill_enabled: bool = True,
    firewall_enabled: bool = True,
    volatile_timestamp: str = "2026-07-24T12:00:00.0000000",
    sqmid: str = "11111111-2222-3333-4444-555555555555",
    permission_binding: bytes = b"\x00",
) -> bytes:
    """Return a valid, deliberately confidential Data Mashup XML payload."""
    package_parts = _zip_payload(
        {
            "[Content_Types].xml": b"<Types/>",
            "Config/Package.xml": (
                b"<Package>private-baseline-package-config</Package>"
            ),
            "Formulas/Section1.m": formula.encode("utf-8"),
            "Content/11111111-2222-3333-4444-555555555555": (
                b"private-baseline-embedded-content"
            ),
        }
    )
    metadata_xml = f"""<?xml version="1.0" encoding="utf-8"?>
<LocalPackageMetadataFile xmlns="{_DATA_MASHUP_NS}">
  <Items>
    <Item>
      <ItemLocation>
        <ItemType>Formula</ItemType>
        <ItemPath>Section1/Private revenue query</ItemPath>
      </ItemLocation>
      <StableEntries>
        <Entry Type="FillEnabled" Value="l{int(fill_enabled)}" />
        <Entry Type="FillTarget" Value="sprivate-baseline-target" />
        <Entry Type="FillLastUpdated" Value="d{volatile_timestamp}" />
        <Entry Type="FillErrorMessage" Value="sprivate-refresh-message" />
      </StableEntries>
    </Item>
  </Items>
</LocalPackageMetadataFile>
""".encode()
    metadata_content = _zip_payload({})
    metadata = (
        struct.pack("<II", 0, len(metadata_xml))
        + metadata_xml
        + struct.pack("<I", len(metadata_content))
        + metadata_content
    )
    permissions = f"""<?xml version="1.0" encoding="utf-8"?>
<PermissionList xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <CanEvaluateFuturePackages>false</CanEvaluateFuturePackages>
  <FirewallEnabled>{str(firewall_enabled).lower()}</FirewallEnabled>
  <WorkbookGroupType xsi:nil="true" />
</PermissionList>
""".encode()
    stream = struct.pack("<I", 0)
    for field in (package_parts, permissions, metadata, permission_binding):
        stream += struct.pack("<I", len(field)) + field
    root = ElementTree.Element(
        f"{{{_DATA_MASHUP_NS}}}DataMashup",
        {"sqmid": sqmid},
    )
    root.text = base64.b64encode(stream).decode("ascii")
    return ElementTree.tostring(root, encoding="utf-8", xml_declaration=True)


def _power_query_formula(*, source: str) -> str:
    """Build one M document whose source material must never be emitted."""
    return f'''section Section1;

shared PrivateRevenueQuery = let
    Source = Web.Contents("{source}"),
    Result = Source
in
    Result;
'''


def make_power_query_model(path: Path) -> Path:
    """Create a raw-OOXML workbook containing a Power Query Data Mashup."""
    make_model(path)
    package_relationships = _PACKAGE_RELATIONSHIPS_NS
    document_relationships = _DOCUMENT_RELATIONSHIPS_NS
    spreadsheet = _SPREADSHEETML_NS
    content_types_namespace = (
        "http://schemas.openxmlformats.org/package/2006/content-types"
    )

    def serialize(root: ElementTree.Element) -> bytes:
        return ElementTree.tostring(root, encoding="utf-8", xml_declaration=True)

    def add_override(root: ElementTree.Element, part_name: str, content_type: str) -> None:
        override_tag = f"{{{content_types_namespace}}}Override"
        if any(item.get("PartName") == part_name for item in root.findall(override_tag)):
            return
        ElementTree.SubElement(
            root,
            override_tag,
            {"PartName": part_name, "ContentType": content_type},
        )

    def mutate(contents: dict[str, bytes]) -> None:
        content_types = ElementTree.fromstring(contents["[Content_Types].xml"])
        add_override(
            content_types,
            "/xl/tables/table1.xml",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml",
        )
        add_override(
            content_types,
            "/xl/queryTables/queryTable1.xml",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.queryTable+xml",
        )
        contents["[Content_Types].xml"] = serialize(content_types)

        root_relationships = ElementTree.fromstring(contents["_rels/.rels"])
        relationship_tag = f"{{{package_relationships}}}Relationship"
        ElementTree.SubElement(
            root_relationships,
            relationship_tag,
            {
                "Id": "rIdFencePowerQueryCustomXml",
                "Type": f"{document_relationships}/customXml",
                "Target": "customXml/item1.xml",
            },
        )
        contents["_rels/.rels"] = serialize(root_relationships)

        worksheet = _inputs_worksheet_root(contents)
        table_parts = ElementTree.SubElement(
            worksheet,
            f"{{{spreadsheet}}}tableParts",
            {"count": "1"},
        )
        ElementTree.SubElement(
            table_parts,
            f"{{{spreadsheet}}}tablePart",
            {f"{{{document_relationships}}}id": "rIdFencePowerQueryTable"},
        )
        _save_inputs_worksheet(contents, worksheet)

        worksheet_relationships = ElementTree.Element(
            f"{{{package_relationships}}}Relationships"
        )
        ElementTree.SubElement(
            worksheet_relationships,
            relationship_tag,
            {
                "Id": "rIdFencePowerQueryTable",
                "Type": f"{document_relationships}/table",
                "Target": "../tables/table1.xml",
            },
        )
        contents["xl/worksheets/_rels/sheet1.xml.rels"] = serialize(
            worksheet_relationships
        )

        table = ElementTree.Element(
            f"{{{spreadsheet}}}table",
            {
                "id": "1",
                "name": "ResultTable",
                "displayName": "ResultTable",
                "ref": "A1:B2",
            },
        )
        ElementTree.SubElement(table, f"{{{spreadsheet}}}autoFilter", {"ref": "A1:B2"})
        columns = ElementTree.SubElement(
            table,
            f"{{{spreadsheet}}}tableColumns",
            {"count": "2"},
        )
        ElementTree.SubElement(
            columns,
            f"{{{spreadsheet}}}tableColumn",
            {"id": "1", "name": "Metric"},
        )
        ElementTree.SubElement(
            columns,
            f"{{{spreadsheet}}}tableColumn",
            {"id": "2", "name": "Value"},
        )
        ElementTree.SubElement(
            table,
            f"{{{spreadsheet}}}tableStyleInfo",
            {
                "name": "TableStyleMedium2",
                "showFirstColumn": "0",
                "showLastColumn": "0",
                "showRowStripes": "1",
                "showColumnStripes": "0",
            },
        )
        contents["xl/tables/table1.xml"] = serialize(table)
        table_relationships = ElementTree.Element(
            f"{{{package_relationships}}}Relationships"
        )
        ElementTree.SubElement(
            table_relationships,
            relationship_tag,
            {
                "Id": "rIdFencePowerQueryQueryTable",
                "Type": f"{document_relationships}/queryTable",
                "Target": "../queryTables/queryTable1.xml",
            },
        )
        contents["xl/tables/_rels/table1.xml.rels"] = serialize(table_relationships)

        query_table = ElementTree.Element(
            f"{{{spreadsheet}}}queryTable",
            {
                "name": "ResultTable_ExternalData_1",
                "connectionId": "1",
                "refreshOnLoad": "1",
                "backgroundRefresh": "1",
                "disableRefresh": "0",
                "removeDataOnSave": "0",
                "fillFormulas": "0",
                "disableEdit": "1",
                "growShrinkType": "insertClear",
            },
        )
        contents["xl/queryTables/queryTable1.xml"] = serialize(query_table)
        contents["customXml/item1.xml"] = _synthetic_power_query_mashup(
            formula=_power_query_formula(
                source="https://private.example/private-baseline-power-query-token"
            )
        )

    return _rewrite_archive(path, mutate, ".power-query.tmp.xlsx")


def change_power_query_controls(path: Path) -> Path:
    """Change private M material plus safe Power Query execution controls."""

    def mutate(contents: dict[str, bytes]) -> None:
        contents["customXml/item1.xml"] = _synthetic_power_query_mashup(
            formula=_power_query_formula(
                source="https://private.example/private-candidate-power-query-token"
            ),
            fill_enabled=False,
            firewall_enabled=False,
            sqmid="aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee",
        )

    return _rewrite_archive(path, mutate, ".power-query-change.tmp.xlsx")


def change_power_query_refresh_noise(path: Path) -> Path:
    """Change volatile refresh state and user-bound binding without semantic drift."""

    def mutate(contents: dict[str, bytes]) -> None:
        contents["customXml/item1.xml"] = _synthetic_power_query_mashup(
            formula=_power_query_formula(
                source="https://private.example/private-baseline-power-query-token"
            ),
            volatile_timestamp="2026-07-24T13:45:00.0000000",
            sqmid="99999999-8888-7777-6666-555555555555",
            permission_binding=b"synthetic-user-bound-permission-binding",
        )

    return _rewrite_archive(path, mutate, ".power-query-noise.tmp.xlsx")


def make_external_data_refresh_model(path: Path) -> Path:
    """Create a raw-OOXML fixture covering every external-data refresh layer."""
    make_model(path)
    spreadsheet = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    package_relationships = (
        "http://schemas.openxmlformats.org/package/2006/relationships"
    )
    document_relationships = (
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    )
    content_types_namespace = (
        "http://schemas.openxmlformats.org/package/2006/content-types"
    )

    def serialize(root: ElementTree.Element) -> bytes:
        return ElementTree.tostring(root, encoding="utf-8", xml_declaration=True)

    def add_override(
        root: ElementTree.Element,
        part_name: str,
        content_type: str,
    ) -> None:
        override_tag = f"{{{content_types_namespace}}}Override"
        if any(item.get("PartName") == part_name for item in root.findall(override_tag)):
            return
        ElementTree.SubElement(
            root,
            override_tag,
            {"PartName": part_name, "ContentType": content_type},
        )

    def mutate(contents: dict[str, bytes]) -> None:
        workbook = ElementTree.fromstring(contents["xl/workbook.xml"])
        workbook_properties = workbook.find(f"{{{spreadsheet}}}workbookPr")
        if workbook_properties is None:
            workbook_properties = ElementTree.Element(f"{{{spreadsheet}}}workbookPr")
            workbook.insert(0, workbook_properties)
        workbook_properties.attrib.update(
            {
                "updateLinks": "always",
                "allowRefreshQuery": "1",
                "refreshAllConnections": "1",
                "saveExternalLinkValues": "0",
            }
        )
        pivot_caches = workbook.find(f"{{{spreadsheet}}}pivotCaches")
        if pivot_caches is None:
            pivot_caches = ElementTree.Element(f"{{{spreadsheet}}}pivotCaches")
            calc_properties = workbook.find(f"{{{spreadsheet}}}calcPr")
            insertion_index = (
                list(workbook).index(calc_properties)
                if calc_properties is not None
                else len(workbook)
            )
            workbook.insert(insertion_index, pivot_caches)
        ElementTree.SubElement(
            pivot_caches,
            f"{{{spreadsheet}}}pivotCache",
            {"cacheId": "7", f"{{{document_relationships}}}id": "rIdFencePivot"},
        )
        contents["xl/workbook.xml"] = serialize(workbook)

        workbook_relationships = ElementTree.fromstring(
            contents["xl/_rels/workbook.xml.rels"]
        )
        relationship_tag = f"{{{package_relationships}}}Relationship"
        ElementTree.SubElement(
            workbook_relationships,
            relationship_tag,
            {
                "Id": "rIdFenceConnections",
                "Type": f"{document_relationships}/connections",
                "Target": "connections.xml",
            },
        )
        ElementTree.SubElement(
            workbook_relationships,
            relationship_tag,
            {
                "Id": "rIdFencePivot",
                "Type": f"{document_relationships}/pivotCacheDefinition",
                "Target": "pivotCache/pivotCacheDefinition1.xml",
            },
        )
        contents["xl/_rels/workbook.xml.rels"] = serialize(workbook_relationships)

        content_types = ElementTree.fromstring(contents["[Content_Types].xml"])
        add_override(
            content_types,
            "/xl/connections.xml",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.connections+xml",
        )
        add_override(
            content_types,
            "/xl/queryTables/queryTable1.xml",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.queryTable+xml",
        )
        add_override(
            content_types,
            "/xl/pivotCache/pivotCacheDefinition1.xml",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml",
        )
        contents["[Content_Types].xml"] = serialize(content_types)

        connection_root = ElementTree.Element(f"{{{spreadsheet}}}connections")
        database_connection = ElementTree.SubElement(
            connection_root,
            f"{{{spreadsheet}}}connection",
            {
                "id": "1",
                "name": "synthetic confidential revenue connection",
                "description": "synthetic confidential connection description",
                "type": "5",
                "refreshedVersion": "6",
                "minRefreshableVersion": "3",
                "sourceFile": "C:/private/synthetic-revenue-source.accdb",
                "odcFile": "https://private.example/synthetic-revenue.odc",
                "keepAlive": "1",
                "interval": "60",
                "savePassword": "1",
                "new": "1",
                "onlyUseConnectionFile": "1",
                "reconnectionMethod": "2",
                "background": "1",
                "refreshOnLoad": "1",
                "saveData": "0",
                "credentials": "stored",
                "singleSignOnId": "synthetic-private-sso-identifier",
            },
        )
        ElementTree.SubElement(
            database_connection,
            f"{{{spreadsheet}}}dbPr",
            {
                "connection": "Provider=synthetic;Password=private-baseline-password",
                "command": "select * from confidential_revenue",
                "commandType": "2",
            },
        )
        parameters = ElementTree.SubElement(
            database_connection,
            f"{{{spreadsheet}}}parameters",
            {"count": "1"},
        )
        ElementTree.SubElement(
            parameters,
            f"{{{spreadsheet}}}parameter",
            {
                "name": "synthetic confidential parameter",
                "parameterType": "cell",
                "cell": "Inputs!$B$2",
                "refreshOnChange": "1",
                "string": "synthetic confidential parameter value",
            },
        )
        connection_extensions = ElementTree.SubElement(
            database_connection,
            f"{{{spreadsheet}}}extLst",
        )
        ElementTree.SubElement(
            connection_extensions,
            f"{{{spreadsheet}}}ext",
            {"uri": "urn:synthetic:private-connection-extension"},
        ).text = "synthetic private connection extension payload"
        web_connection = ElementTree.SubElement(
            connection_root,
            f"{{{spreadsheet}}}connection",
            {
                "id": "2",
                "name": "synthetic confidential web connection",
                "type": "4",
                "refreshedVersion": "6",
            },
        )
        ElementTree.SubElement(
            web_connection,
            f"{{{spreadsheet}}}webPr",
            {"url": "https://private.example/synthetic-web-query"},
        )
        contents["xl/connections.xml"] = serialize(connection_root)

        query_table = ElementTree.Element(
            f"{{{spreadsheet}}}queryTable",
            {
                "name": "synthetic confidential query table",
                "connectionId": "1",
                "refreshOnLoad": "1",
                "backgroundRefresh": "0",
                "disableRefresh": "0",
                "removeDataOnSave": "1",
                "fillFormulas": "1",
                "disableEdit": "1",
                "growShrinkType": "overwriteClear",
            },
        )
        query_extensions = ElementTree.SubElement(query_table, f"{{{spreadsheet}}}extLst")
        ElementTree.SubElement(
            query_extensions,
            f"{{{spreadsheet}}}ext",
            {"uri": "urn:synthetic:private-query-extension"},
        ).text = "synthetic private query extension payload"
        contents["xl/queryTables/queryTable1.xml"] = serialize(query_table)

        sheet_relationships = ElementTree.Element(f"{{{package_relationships}}}Relationships")
        ElementTree.SubElement(
            sheet_relationships,
            relationship_tag,
            {
                "Id": "rIdFenceQuery",
                "Type": f"{document_relationships}/queryTable",
                "Target": "../queryTables/queryTable1.xml",
            },
        )
        contents["xl/worksheets/_rels/sheet1.xml.rels"] = serialize(sheet_relationships)

        pivot_cache_definition = ElementTree.Element(
            f"{{{spreadsheet}}}pivotCacheDefinition",
            {
                "refreshOnLoad": "1",
                "backgroundQuery": "1",
                "enableRefresh": "0",
                "saveData": "0",
                "upgradeOnRefresh": "1",
                "refreshedBy": "synthetic confidential refresh identity",
            },
        )
        pivot_source = ElementTree.SubElement(
            pivot_cache_definition,
            f"{{{spreadsheet}}}cacheSource",
            {"type": "external", "connectionId": "2"},
        )
        ElementTree.SubElement(
            pivot_source,
            f"{{{spreadsheet}}}extLst",
        ).text = "synthetic private pivot source payload"
        pivot_extensions = ElementTree.SubElement(
            pivot_cache_definition,
            f"{{{spreadsheet}}}extLst",
        )
        ElementTree.SubElement(
            pivot_extensions,
            f"{{{spreadsheet}}}ext",
            {"uri": "urn:synthetic:private-pivot-extension"},
        ).text = "synthetic private pivot extension payload"
        contents["xl/pivotCache/pivotCacheDefinition1.xml"] = serialize(
            pivot_cache_definition
        )

    return _rewrite_archive(path, mutate, ".external-data.tmp.xlsx")


def change_external_data_refresh_controls(path: Path) -> Path:
    """Change safe controls and private source material in the raw fixture."""
    spreadsheet = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"

    def serialize(root: ElementTree.Element) -> bytes:
        return ElementTree.tostring(root, encoding="utf-8", xml_declaration=True)

    def mutate(contents: dict[str, bytes]) -> None:
        workbook = ElementTree.fromstring(contents["xl/workbook.xml"])
        workbook_properties = workbook.find(f"{{{spreadsheet}}}workbookPr")
        if workbook_properties is None:
            raise ValueError("Fixture does not contain workbook properties")
        workbook_properties.attrib.update(
            {
                "updateLinks": "never",
                "allowRefreshQuery": "0",
                "refreshAllConnections": "0",
                "saveExternalLinkValues": "1",
            }
        )
        contents["xl/workbook.xml"] = serialize(workbook)

        connections = ElementTree.fromstring(contents["xl/connections.xml"])
        connection = connections.find(f"{{{spreadsheet}}}connection")
        if connection is None:
            raise ValueError("Fixture does not contain a connection")
        connection.attrib.update(
            {
                "name": "changed synthetic confidential revenue connection",
                "sourceFile": "C:/private/changed-synthetic-source.accdb",
                "odcFile": "https://private.example/changed-synthetic.odc",
                "singleSignOnId": "changed-synthetic-private-sso-identifier",
                "interval": "15",
                "refreshOnLoad": "0",
                "savePassword": "0",
            }
        )
        database_properties = connection.find(f"{{{spreadsheet}}}dbPr")
        if database_properties is None:
            raise ValueError("Fixture does not contain database connection properties")
        database_properties.set(
            "connection", "Provider=changed;Password=private-candidate-password"
        )
        contents["xl/connections.xml"] = serialize(connections)

        query_table = ElementTree.fromstring(contents["xl/queryTables/queryTable1.xml"])
        query_table.attrib.update(
            {
                "name": "changed synthetic confidential query table",
                "connectionId": "2",
                "refreshOnLoad": "0",
                "fillFormulas": "0",
            }
        )
        contents["xl/queryTables/queryTable1.xml"] = serialize(query_table)

        pivot_cache = ElementTree.fromstring(
            contents["xl/pivotCache/pivotCacheDefinition1.xml"]
        )
        pivot_cache.attrib.update(
            {
                "refreshOnLoad": "0",
                "enableRefresh": "1",
                "saveData": "1",
            }
        )
        pivot_source = pivot_cache.find(f"{{{spreadsheet}}}cacheSource")
        if pivot_source is None:
            raise ValueError("Fixture does not contain a pivot-cache source")
        pivot_source.set("connectionId", "1")
        contents["xl/pivotCache/pivotCacheDefinition1.xml"] = serialize(pivot_cache)

    return _rewrite_archive(path, mutate, ".external-data-change.tmp.xlsx")


def set_external_data_connection_defaults(path: Path, *, explicit: bool) -> Path:
    """Toggle omitted versus explicit defaults on the fixture's web connection."""
    spreadsheet = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    defaults = {
        "keepAlive": "0",
        "interval": "0",
        "reconnectionMethod": "1",
        "minRefreshableVersion": "0",
        "savePassword": "0",
        "new": "0",
        "deleted": "0",
        "onlyUseConnectionFile": "0",
        "background": "0",
        "refreshOnLoad": "0",
        "saveData": "0",
        "credentials": "integrated",
    }

    def mutate(contents: dict[str, bytes]) -> None:
        connections = ElementTree.fromstring(contents["xl/connections.xml"])
        connection_tag = f"{{{spreadsheet}}}connection"
        connection = next(
            (
                item
                for item in connections.findall(connection_tag)
                if item.get("id") == "2"
            ),
            None,
        )
        if connection is None:
            raise ValueError("Fixture does not contain the web connection")
        for attribute, value in defaults.items():
            if explicit:
                connection.set(attribute, value)
            else:
                connection.attrib.pop(attribute, None)
        contents["xl/connections.xml"] = ElementTree.tostring(
            connections,
            encoding="utf-8",
            xml_declaration=True,
        )

    return _rewrite_archive(path, mutate, ".external-data-defaults.tmp.xlsx")


def make_external_link_package_model(path: Path) -> Path:
    """Create external-workbook, DDE, and OLE packages with private payloads."""
    make_model(path)
    spreadsheet = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    package_relationships = (
        "http://schemas.openxmlformats.org/package/2006/relationships"
    )
    document_relationships = (
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    )
    content_types_namespace = (
        "http://schemas.openxmlformats.org/package/2006/content-types"
    )

    def serialize(root: ElementTree.Element) -> bytes:
        return ElementTree.tostring(root, encoding="utf-8", xml_declaration=True)

    def add_override(
        root: ElementTree.Element,
        part_name: str,
        content_type: str,
    ) -> None:
        override_tag = f"{{{content_types_namespace}}}Override"
        if any(item.get("PartName") == part_name for item in root.findall(override_tag)):
            return
        ElementTree.SubElement(
            root,
            override_tag,
            {"PartName": part_name, "ContentType": content_type},
        )

    def mutate(contents: dict[str, bytes]) -> None:
        workbook = ElementTree.fromstring(contents["xl/workbook.xml"])
        external_references = ElementTree.Element(f"{{{spreadsheet}}}externalReferences")
        for relationship_id in (
            "rIdFenceExternalWorkbook",
            "rIdFenceDde",
            "rIdFenceOle",
        ):
            ElementTree.SubElement(
                external_references,
                f"{{{spreadsheet}}}externalReference",
                {f"{{{document_relationships}}}id": relationship_id},
            )
        calc_properties = workbook.find(f"{{{spreadsheet}}}calcPr")
        insertion_index = (
            list(workbook).index(calc_properties)
            if calc_properties is not None
            else len(workbook)
        )
        workbook.insert(insertion_index, external_references)
        contents["xl/workbook.xml"] = serialize(workbook)

        workbook_relationships = ElementTree.fromstring(
            contents["xl/_rels/workbook.xml.rels"]
        )
        relationship_tag = f"{{{package_relationships}}}Relationship"
        for relationship_id, target in (
            ("rIdFenceExternalWorkbook", "externalLinks/externalLink1.xml"),
            ("rIdFenceDde", "externalLinks/externalLink2.xml"),
            ("rIdFenceOle", "externalLinks/externalLink3.xml"),
        ):
            ElementTree.SubElement(
                workbook_relationships,
                relationship_tag,
                {
                    "Id": relationship_id,
                    "Type": f"{document_relationships}/externalLink",
                    "Target": target,
                },
            )
        contents["xl/_rels/workbook.xml.rels"] = serialize(workbook_relationships)

        content_types = ElementTree.fromstring(contents["[Content_Types].xml"])
        for number in (1, 2, 3):
            add_override(
                content_types,
                f"/xl/externalLinks/externalLink{number}.xml",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml",
            )
        contents["[Content_Types].xml"] = serialize(content_types)

        external_book_root = ElementTree.Element(f"{{{spreadsheet}}}externalLink")
        external_book = ElementTree.SubElement(
            external_book_root,
            f"{{{spreadsheet}}}externalBook",
            {f"{{{document_relationships}}}id": "rIdFenceExternalTarget"},
        )
        sheet_names = ElementTree.SubElement(external_book, f"{{{spreadsheet}}}sheetNames")
        for sheet_name in ("private external baseline sheet", "private external scenario"):
            ElementTree.SubElement(
                sheet_names,
                f"{{{spreadsheet}}}sheetName",
                {"val": sheet_name},
            )
        defined_names = ElementTree.SubElement(
            external_book, f"{{{spreadsheet}}}definedNames"
        )
        ElementTree.SubElement(
            defined_names,
            f"{{{spreadsheet}}}definedName",
            {
                "name": "private external baseline defined name",
                "refersTo": "'private external baseline sheet'!$A$1:$A$2",
                "sheetId": "0",
            },
        )
        sheet_data_set = ElementTree.SubElement(
            external_book, f"{{{spreadsheet}}}sheetDataSet"
        )
        sheet_data = ElementTree.SubElement(
            sheet_data_set,
            f"{{{spreadsheet}}}sheetData",
            {"sheetId": "0", "refreshError": "1"},
        )
        row = ElementTree.SubElement(sheet_data, f"{{{spreadsheet}}}row", {"r": "1"})
        for coordinate, value in (
            ("A1", "private external baseline cached value"),
            ("B1", "private external baseline cached amount"),
        ):
            cell = ElementTree.SubElement(
                row,
                f"{{{spreadsheet}}}cell",
                {"r": coordinate, "t": "str"},
            )
            ElementTree.SubElement(cell, f"{{{spreadsheet}}}v").text = value
        contents["xl/externalLinks/externalLink1.xml"] = serialize(external_book_root)

        external_book_relationships = ElementTree.Element(
            f"{{{package_relationships}}}Relationships"
        )
        ElementTree.SubElement(
            external_book_relationships,
            relationship_tag,
            {
                "Id": "rIdFenceExternalTarget",
                "Type": f"{document_relationships}/externalLinkPath",
                "Target": "file:///private/baseline/external-workbook.xlsx",
                "TargetMode": "External",
            },
        )
        contents["xl/externalLinks/_rels/externalLink1.xml.rels"] = serialize(
            external_book_relationships
        )

        dde_root = ElementTree.Element(
            f"{{{spreadsheet}}}externalLink",
        )
        dde_link = ElementTree.SubElement(
            dde_root,
            f"{{{spreadsheet}}}ddeLink",
            {
                "ddeService": "private-baseline-dde-service",
                "ddeTopic": "private-baseline-dde-topic",
            },
        )
        dde_items = ElementTree.SubElement(dde_link, f"{{{spreadsheet}}}ddeItems")
        first_dde_item = ElementTree.SubElement(
            dde_items,
            f"{{{spreadsheet}}}ddeItem",
            {
                "name": "private-baseline-dde-item",
                "ole": "1",
                "advise": "1",
                "preferPic": "0",
            },
        )
        dde_values = ElementTree.SubElement(
            first_dde_item,
            f"{{{spreadsheet}}}values",
            {"rows": "1", "cols": "2"},
        )
        for value in ("private-baseline-dde-value", "private-baseline-dde-second-value"):
            dde_value = ElementTree.SubElement(dde_values, f"{{{spreadsheet}}}value")
            ElementTree.SubElement(dde_value, f"{{{spreadsheet}}}val").text = value
        ElementTree.SubElement(
            dde_items,
            f"{{{spreadsheet}}}ddeItem",
            {
                "name": "private-baseline-dde-picture-item",
                "preferPic": "1",
            },
        )
        contents["xl/externalLinks/externalLink2.xml"] = serialize(dde_root)

        ole_root = ElementTree.Element(f"{{{spreadsheet}}}externalLink")
        ole_link = ElementTree.SubElement(
            ole_root,
            f"{{{spreadsheet}}}oleLink",
            {
                f"{{{document_relationships}}}id": "rIdFenceOleTarget",
                "progId": "private.baseline.ole.program",
            },
        )
        ole_items = ElementTree.SubElement(ole_link, f"{{{spreadsheet}}}oleItems")
        ElementTree.SubElement(
            ole_items,
            f"{{{spreadsheet}}}oleItem",
            {
                "name": "private-baseline-ole-item",
                "icon": "1",
                "advise": "1",
                "preferPic": "0",
            },
        )
        ElementTree.SubElement(
            ole_items,
            f"{{{spreadsheet}}}oleItem",
            {
                "name": "private-baseline-ole-picture-item",
                "preferPic": "1",
            },
        )
        extensions = ElementTree.SubElement(ole_root, f"{{{spreadsheet}}}extLst")
        ElementTree.SubElement(
            extensions,
            f"{{{spreadsheet}}}ext",
            {"uri": "urn:private:baseline:external-link-extension"},
        ).text = "private baseline external-link extension payload"
        contents["xl/externalLinks/externalLink3.xml"] = serialize(ole_root)

        ole_relationships = ElementTree.Element(
            f"{{{package_relationships}}}Relationships"
        )
        ElementTree.SubElement(
            ole_relationships,
            relationship_tag,
            {
                "Id": "rIdFenceOleTarget",
                "Type": f"{document_relationships}/oleObject",
                "Target": "file:///private/baseline/external-object.bin",
                "TargetMode": "External",
            },
        )
        contents["xl/externalLinks/_rels/externalLink3.xml.rels"] = serialize(
            ole_relationships
        )

    return _rewrite_archive(path, mutate, ".external-link.tmp.xlsx")


def change_external_link_package_controls(path: Path) -> Path:
    """Change external source, definition, cache, and opaque package material."""
    spreadsheet = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    package_relationships = (
        "http://schemas.openxmlformats.org/package/2006/relationships"
    )

    def serialize(root: ElementTree.Element) -> bytes:
        return ElementTree.tostring(root, encoding="utf-8", xml_declaration=True)

    def mutate(contents: dict[str, bytes]) -> None:
        external_book_relationships = ElementTree.fromstring(
            contents["xl/externalLinks/_rels/externalLink1.xml.rels"]
        )
        relationship = external_book_relationships.find(
            f"{{{package_relationships}}}Relationship"
        )
        if relationship is None:
            raise ValueError("Fixture does not contain an external-workbook target")
        relationship.set("Target", "file:///private/candidate/external-workbook.xlsx")
        contents["xl/externalLinks/_rels/externalLink1.xml.rels"] = serialize(
            external_book_relationships
        )

        external_book_root = ElementTree.fromstring(
            contents["xl/externalLinks/externalLink1.xml"]
        )
        defined_name = external_book_root.find(
            f"./{{{spreadsheet}}}externalBook/{{{spreadsheet}}}definedNames/"
            f"{{{spreadsheet}}}definedName"
        )
        cached_value = external_book_root.find(
            f"./{{{spreadsheet}}}externalBook/{{{spreadsheet}}}sheetDataSet/"
            f"{{{spreadsheet}}}sheetData/{{{spreadsheet}}}row/"
            f"{{{spreadsheet}}}cell/{{{spreadsheet}}}v"
        )
        if defined_name is None or cached_value is None:
            raise ValueError("Fixture does not contain external-workbook definition data")
        defined_name.set("refersTo", "'private external scenario'!$B$5")
        cached_value.text = "private external candidate cached value"
        contents["xl/externalLinks/externalLink1.xml"] = serialize(external_book_root)

        dde_root = ElementTree.fromstring(contents["xl/externalLinks/externalLink2.xml"])
        dde_link = dde_root.find(f"{{{spreadsheet}}}ddeLink")
        first_dde_item = dde_root.find(
            f"./{{{spreadsheet}}}ddeLink/{{{spreadsheet}}}ddeItems/"
            f"{{{spreadsheet}}}ddeItem"
        )
        dde_value = dde_root.find(
            f"./{{{spreadsheet}}}ddeLink/{{{spreadsheet}}}ddeItems/"
            f"{{{spreadsheet}}}ddeItem/{{{spreadsheet}}}values/"
            f"{{{spreadsheet}}}value/{{{spreadsheet}}}val"
        )
        if dde_link is None or first_dde_item is None or dde_value is None:
            raise ValueError("Fixture does not contain DDE controls")
        dde_link.set("ddeTopic", "private-candidate-dde-topic")
        first_dde_item.set("name", "private-candidate-dde-item")
        first_dde_item.set("advise", "0")
        dde_value.text = "private-candidate-dde-value"
        contents["xl/externalLinks/externalLink2.xml"] = serialize(dde_root)

        ole_root = ElementTree.fromstring(contents["xl/externalLinks/externalLink3.xml"])
        ole_link = ole_root.find(f"{{{spreadsheet}}}oleLink")
        first_ole_item = ole_root.find(
            f"./{{{spreadsheet}}}oleLink/{{{spreadsheet}}}oleItems/"
            f"{{{spreadsheet}}}oleItem"
        )
        extension = ole_root.find(f"./{{{spreadsheet}}}extLst/{{{spreadsheet}}}ext")
        if ole_link is None or first_ole_item is None or extension is None:
            raise ValueError("Fixture does not contain OLE controls")
        ole_link.set("progId", "private.candidate.ole.program")
        first_ole_item.set("name", "private-candidate-ole-item")
        first_ole_item.set("icon", "0")
        extension.set("uri", "urn:private:candidate:external-link-extension")
        extension.text = "private candidate external-link extension payload"
        contents["xl/externalLinks/externalLink3.xml"] = serialize(ole_root)

    return _rewrite_archive(path, mutate, ".external-link-change.tmp.xlsx")


def rebind_external_link_declaration(path: Path) -> Path:
    """Point one workbook declaration at a different existing externalLink part."""
    package_relationships = (
        "http://schemas.openxmlformats.org/package/2006/relationships"
    )

    def mutate(contents: dict[str, bytes]) -> None:
        relationships = ElementTree.fromstring(contents["xl/_rels/workbook.xml.rels"])
        relationship_tag = f"{{{package_relationships}}}Relationship"
        external_workbook = next(
            (
                relationship
                for relationship in relationships.findall(relationship_tag)
                if relationship.get("Id") == "rIdFenceExternalWorkbook"
            ),
            None,
        )
        if external_workbook is None:
            raise ValueError("Fixture does not contain the external-workbook declaration")
        external_workbook.set("Target", "externalLinks/externalLink2.xml")
        contents["xl/_rels/workbook.xml.rels"] = ElementTree.tostring(
            relationships,
            encoding="utf-8",
            xml_declaration=True,
        )

    return _rewrite_archive(path, mutate, ".external-link-rebind.tmp.xlsx")


def renumber_external_link_declaration_relationships(path: Path) -> Path:
    """Rewrite only arbitrary workbook external-link relationship identifiers."""
    package_relationships = (
        "http://schemas.openxmlformats.org/package/2006/relationships"
    )
    document_relationships = (
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    )
    replacements = {
        "rIdFenceExternalWorkbook": "rIdFenceRenumberedWorkbook",
        "rIdFenceDde": "rIdFenceRenumberedDde",
        "rIdFenceOle": "rIdFenceRenumberedOle",
    }

    def mutate(contents: dict[str, bytes]) -> None:
        relationships = ElementTree.fromstring(contents["xl/_rels/workbook.xml.rels"])
        relationship_tag = f"{{{package_relationships}}}Relationship"
        for relationship in relationships.findall(relationship_tag):
            if replacement := replacements.get(relationship.get("Id")):
                relationship.set("Id", replacement)
        contents["xl/_rels/workbook.xml.rels"] = ElementTree.tostring(
            relationships,
            encoding="utf-8",
            xml_declaration=True,
        )

        workbook = ElementTree.fromstring(contents["xl/workbook.xml"])
        reference_tag = (
            "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
            "externalReference"
        )
        relationship_id_attribute = f"{{{document_relationships}}}id"
        for reference in workbook.findall(f".//{reference_tag}"):
            if replacement := replacements.get(reference.get(relationship_id_attribute)):
                reference.set(relationship_id_attribute, replacement)
        contents["xl/workbook.xml"] = ElementTree.tostring(
            workbook,
            encoding="utf-8",
            xml_declaration=True,
        )

    return _rewrite_archive(path, mutate, ".external-link-renumber.tmp.xlsx")


def duplicate_external_link_definition(path: Path) -> Path:
    """Add a second supported definition to exercise the fail-closed parser path."""
    spreadsheet = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"

    def mutate(contents: dict[str, bytes]) -> None:
        root = ElementTree.fromstring(contents["xl/externalLinks/externalLink1.xml"])
        external_book = root.find(f"{{{spreadsheet}}}externalBook")
        if external_book is None:
            raise ValueError("Fixture does not contain an external-workbook definition")
        root.append(
            ElementTree.fromstring(
                ElementTree.tostring(external_book, encoding="utf-8")
            )
        )
        contents["xl/externalLinks/externalLink1.xml"] = ElementTree.tostring(
            root,
            encoding="utf-8",
            xml_declaration=True,
        )

    return _rewrite_archive(path, mutate, ".external-link-duplicate.tmp.xlsx")


def duplicate_external_link_sheet_names(path: Path) -> Path:
    """Repeat a schema-singleton workbook child to test coverage reporting."""
    spreadsheet = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"

    def mutate(contents: dict[str, bytes]) -> None:
        root = ElementTree.fromstring(contents["xl/externalLinks/externalLink1.xml"])
        external_book = root.find(f"{{{spreadsheet}}}externalBook")
        if external_book is None:
            raise ValueError("Fixture does not contain an external-workbook definition")
        sheet_names = external_book.find(f"{{{spreadsheet}}}sheetNames")
        if sheet_names is None:
            raise ValueError("Fixture does not contain external-workbook sheet names")
        external_book.append(
            ElementTree.fromstring(ElementTree.tostring(sheet_names, encoding="utf-8"))
        )
        contents["xl/externalLinks/externalLink1.xml"] = ElementTree.tostring(
            root,
            encoding="utf-8",
            xml_declaration=True,
        )

    return _rewrite_archive(path, mutate, ".external-link-repeated-child.tmp.xlsx")


def make_xlm_macro_sheet_model(path: Path, *, international: bool = False) -> Path:
    """Create a harmless raw XLM macro-sheet fixture that readers commonly omit.

    The macro formula strings are never opened in Excel or evaluated. They are
    only private OOXML test material for FormulaFence's package scanner.
    """
    make_model(path)
    spreadsheet = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    package_relationships = (
        "http://schemas.openxmlformats.org/package/2006/relationships"
    )
    document_relationships = (
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    )
    content_types_namespace = (
        "http://schemas.openxmlformats.org/package/2006/content-types"
    )
    macro_namespace = "http://schemas.microsoft.com/office/excel/2006/main"
    macro_member = (
        "xl/macrosheets/intlsheet1.xml"
        if international
        else "xl/macrosheets/sheet1.xml"
    )
    macro_target = macro_member.removeprefix("xl/")
    macro_kind = "xlIntlMacrosheet" if international else "xlMacrosheet"
    macro_content_type = (
        "application/vnd.ms-excel.intlmacrosheet+xml"
        if international
        else "application/vnd.ms-excel.macrosheet+xml"
    )

    def serialize(root: ElementTree.Element) -> bytes:
        namespace = root.tag[1:].split("}", maxsplit=1)[0]
        if namespace in {
            spreadsheet,
            package_relationships,
            content_types_namespace,
            macro_namespace,
        }:
            ElementTree.register_namespace("", namespace)
        ElementTree.register_namespace("r", document_relationships)
        return ElementTree.tostring(root, encoding="utf-8", xml_declaration=True)

    def mutate(contents: dict[str, bytes]) -> None:
        workbook = ElementTree.fromstring(contents["xl/workbook.xml"])
        sheets = workbook.find(f"{{{spreadsheet}}}sheets")
        if sheets is None:
            raise ValueError("Fixture does not contain workbook sheets")
        sheet_ids = [int(sheet.get("sheetId", "0")) for sheet in sheets]
        ElementTree.SubElement(
            sheets,
            f"{{{spreadsheet}}}sheet",
            {
                "name": "Macro Automation",
                "sheetId": str(max(sheet_ids) + 1),
                "state": "veryHidden",
                f"{{{document_relationships}}}id": "rIdFenceXlmMacro",
            },
        )
        contents["xl/workbook.xml"] = serialize(workbook)

        workbook_relationships = ElementTree.fromstring(
            contents["xl/_rels/workbook.xml.rels"]
        )
        ElementTree.SubElement(
            workbook_relationships,
            f"{{{package_relationships}}}Relationship",
            {
                "Id": "rIdFenceXlmMacro",
                "Type": (
                    "http://schemas.microsoft.com/office/2006/relationships/"
                    f"{macro_kind}"
                ),
                "Target": macro_target,
            },
        )
        contents["xl/_rels/workbook.xml.rels"] = serialize(workbook_relationships)

        content_types = ElementTree.fromstring(contents["[Content_Types].xml"])
        override_tag = f"{{{content_types_namespace}}}Override"
        for override in content_types.findall(override_tag):
            if override.get("PartName") == "/xl/workbook.xml":
                override.set(
                    "ContentType",
                    "application/vnd.ms-excel.sheet.macroEnabled.main+xml",
                )
        ElementTree.SubElement(
            content_types,
            override_tag,
            {"PartName": f"/{macro_member}", "ContentType": macro_content_type},
        )
        contents["[Content_Types].xml"] = serialize(content_types)

        macro_sheet = ElementTree.Element(f"{{{macro_namespace}}}macrosheet")
        sheet_data = ElementTree.SubElement(
            macro_sheet, f"{{{macro_namespace}}}sheetData"
        )
        for row_number, formula, value in (
            (1, 'PRIVATE.XLM("private-baseline-xl-command-argument")', None),
            (2, "RETURN()", "private-baseline-xl-cell-value"),
        ):
            row = ElementTree.SubElement(
                sheet_data,
                f"{{{macro_namespace}}}row",
                {"r": str(row_number)},
            )
            cell = ElementTree.SubElement(
                row,
                f"{{{macro_namespace}}}c",
                {"r": f"A{row_number}"},
            )
            ElementTree.SubElement(cell, f"{{{macro_namespace}}}f").text = formula
            if value is not None:
                ElementTree.SubElement(cell, f"{{{macro_namespace}}}v").text = value
        ole_objects = ElementTree.SubElement(
            macro_sheet, f"{{{macro_namespace}}}oleObjects"
        )
        ElementTree.SubElement(
            ole_objects,
            f"{{{macro_namespace}}}oleObject",
            {
                f"{{{document_relationships}}}id": "rIdFenceEmbeddedObject",
                "progId": "private.baseline.xlm.embedded.object",
            },
        )
        ElementTree.SubElement(
            ole_objects,
            f"{{{macro_namespace}}}oleObject",
            {
                f"{{{document_relationships}}}id": "rIdFenceExternalObject",
                "progId": "private.baseline.xlm.linked.object",
            },
        )
        contents[macro_member] = serialize(macro_sheet)

        macro_relationships = ElementTree.Element(
            f"{{{package_relationships}}}Relationships"
        )
        relationship_tag = f"{{{package_relationships}}}Relationship"
        for relationship_id, relationship_type, target, target_mode in (
            (
                "rIdFenceEmbeddedObject",
                f"{document_relationships}/oleObject",
                "../embeddings/private-baseline-xl-object.bin",
                None,
            ),
            (
                "rIdFenceExternalObject",
                f"{document_relationships}/oleObject",
                "file:///private/baseline-xl-linked-object.bin",
                "External",
            ),
            (
                "rIdFenceEmbeddedPackage",
                f"{document_relationships}/package",
                "../embeddings/private-baseline-xl-package.bin",
                None,
            ),
        ):
            attributes = {
                "Id": relationship_id,
                "Type": relationship_type,
                "Target": target,
            }
            if target_mode is not None:
                attributes["TargetMode"] = target_mode
            ElementTree.SubElement(macro_relationships, relationship_tag, attributes)
        contents[
            "xl/macrosheets/_rels/" + macro_member.rsplit("/", maxsplit=1)[-1] + ".rels"
        ] = serialize(macro_relationships)
        contents["xl/embeddings/private-baseline-xl-object.bin"] = (
            b"private baseline embedded XLM object payload"
        )
        contents["xl/embeddings/private-baseline-xl-package.bin"] = (
            b"private baseline embedded XLM package payload"
        )

    return _rewrite_archive(path, mutate, ".xlm-macro-sheet.tmp.xlsx")


def change_xlm_macro_sheet_controls(path: Path) -> Path:
    """Change private XLM code, binding, and related-part material."""
    macro_namespace = "http://schemas.microsoft.com/office/excel/2006/main"
    package_relationships = (
        "http://schemas.openxmlformats.org/package/2006/relationships"
    )
    macro_member = "xl/macrosheets/sheet1.xml"

    def mutate(contents: dict[str, bytes]) -> None:
        macro_sheet = ElementTree.fromstring(contents[macro_member])
        formula = macro_sheet.find(
            f"./{{{macro_namespace}}}sheetData/{{{macro_namespace}}}row/"
            f"{{{macro_namespace}}}c/{{{macro_namespace}}}f"
        )
        cell_value = macro_sheet.find(
            f"./{{{macro_namespace}}}sheetData/{{{macro_namespace}}}row[2]/"
            f"{{{macro_namespace}}}c/{{{macro_namespace}}}v"
        )
        embedded_object = macro_sheet.find(
            f"./{{{macro_namespace}}}oleObjects/{{{macro_namespace}}}oleObject"
        )
        if formula is None or cell_value is None or embedded_object is None:
            raise ValueError("Fixture does not contain expected XLM macro-sheet controls")
        formula.text = 'PRIVATE.XLM("private-candidate-xl-command-argument")'
        cell_value.text = "private-candidate-xl-cell-value"
        embedded_object.set("progId", "private.candidate.xlm.embedded.object")
        extensions = ElementTree.SubElement(macro_sheet, f"{{{macro_namespace}}}extLst")
        ElementTree.SubElement(
            extensions,
            f"{{{macro_namespace}}}ext",
            {"uri": "urn:private:candidate:xlm-extension"},
        ).text = "private candidate XLM extension payload"
        contents[macro_member] = ElementTree.tostring(
            macro_sheet,
            encoding="utf-8",
            xml_declaration=True,
        )

        relationships_name = "xl/macrosheets/_rels/sheet1.xml.rels"
        relationships = ElementTree.fromstring(contents[relationships_name])
        relationship_tag = f"{{{package_relationships}}}Relationship"
        for relationship in relationships.findall(relationship_tag):
            if relationship.get("Id") == "rIdFenceEmbeddedObject":
                relationship.set(
                    "Target", "../embeddings/private-candidate-xl-object.bin"
                )
            elif relationship.get("Id") == "rIdFenceExternalObject":
                relationship.set(
                    "Target", "file:///private/candidate-xl-linked-object.bin"
                )
        contents[relationships_name] = ElementTree.tostring(
            relationships,
            encoding="utf-8",
            xml_declaration=True,
        )
        contents["xl/embeddings/private-candidate-xl-object.bin"] = (
            b"private candidate embedded XLM object payload"
        )

        workbook = ElementTree.fromstring(contents["xl/workbook.xml"])
        sheet = next(
            (
                item
                for item in workbook.findall(
                    ".//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}sheet"
                )
                if item.get("name") == "Macro Automation"
            ),
            None,
        )
        if sheet is None:
            raise ValueError("Fixture does not contain XLM macro-sheet declaration")
        sheet.set("state", "hidden")
        contents["xl/workbook.xml"] = ElementTree.tostring(
            workbook,
            encoding="utf-8",
            xml_declaration=True,
        )

    return _rewrite_archive(path, mutate, ".xlm-macro-sheet-change.tmp.xlsx")


def change_xlm_macro_sheet_related_part_payload(path: Path) -> Path:
    """Change only a private internal payload reached from an XLM macro sheet."""

    def mutate(contents: dict[str, bytes]) -> None:
        contents["xl/embeddings/private-baseline-xl-object.bin"] = (
            b"private candidate XLM related-part payload only"
        )

    return _rewrite_archive(path, mutate, ".xlm-related-payload-change.tmp.xlsx")


def remove_xlm_macro_sheet_related_part_payload(path: Path) -> Path:
    """Remove one internal XLM related payload to exercise fail-closed coverage."""

    def mutate(contents: dict[str, bytes]) -> None:
        contents.pop("xl/embeddings/private-baseline-xl-object.bin")

    return _rewrite_archive(path, mutate, ".xlm-related-payload-missing.tmp.xlsx")


def renumber_xlm_macro_sheet_relationships(path: Path) -> Path:
    """Rewrite arbitrary XLM relationship ids while retaining their bindings."""
    package_relationships = (
        "http://schemas.openxmlformats.org/package/2006/relationships"
    )
    document_relationships = (
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    )
    replacements = {
        "rIdFenceXlmMacro": "rIdFenceRenumberedXlmMacro",
        "rIdFenceEmbeddedObject": "rIdFenceRenumberedEmbeddedObject",
        "rIdFenceExternalObject": "rIdFenceRenumberedExternalObject",
        "rIdFenceEmbeddedPackage": "rIdFenceRenumberedEmbeddedPackage",
    }

    def mutate(contents: dict[str, bytes]) -> None:
        relationship_tag = f"{{{package_relationships}}}Relationship"
        for name in (
            "xl/_rels/workbook.xml.rels",
            "xl/macrosheets/_rels/sheet1.xml.rels",
        ):
            relationships = ElementTree.fromstring(contents[name])
            for relationship in relationships.findall(relationship_tag):
                if replacement := replacements.get(relationship.get("Id")):
                    relationship.set("Id", replacement)
            contents[name] = ElementTree.tostring(
                relationships,
                encoding="utf-8",
                xml_declaration=True,
            )

        relationship_id_attribute = f"{{{document_relationships}}}id"
        for name in ("xl/workbook.xml", "xl/macrosheets/sheet1.xml"):
            root = ElementTree.fromstring(contents[name])
            for element in root.iter():
                if replacement := replacements.get(element.get(relationship_id_attribute)):
                    element.set(relationship_id_attribute, replacement)
            contents[name] = ElementTree.tostring(
                root,
                encoding="utf-8",
                xml_declaration=True,
            )

    return _rewrite_archive(path, mutate, ".xlm-macro-sheet-renumber.tmp.xlsx")


def rewrite_xlm_macro_sheet_internal_target_spelling(path: Path) -> Path:
    """Rewrite an XLM internal target using an equivalent relative path."""
    package_relationships = (
        "http://schemas.openxmlformats.org/package/2006/relationships"
    )

    def mutate(contents: dict[str, bytes]) -> None:
        relationships_name = "xl/_rels/workbook.xml.rels"
        relationships = ElementTree.fromstring(contents[relationships_name])
        macro_relationship = next(
            (
                relationship
                for relationship in relationships.findall(
                    f"{{{package_relationships}}}Relationship"
                )
                if relationship.get("Type", "").endswith("/xlMacrosheet")
            ),
            None,
        )
        if macro_relationship is None:
            raise ValueError("Fixture does not contain an XLM macro-sheet relationship")
        macro_relationship.set(
            "Target", "./macrosheets/../macrosheets/sheet1.xml"
        )
        contents[relationships_name] = ElementTree.tostring(
            relationships,
            encoding="utf-8",
            xml_declaration=True,
        )

    return _rewrite_archive(path, mutate, ".xlm-macro-sheet-target.tmp.xlsx")


def corrupt_xlm_macro_sheet_root(path: Path) -> Path:
    """Replace the macro-sheet root with an unexpected element for fail-closed tests."""
    macro_namespace = "http://schemas.microsoft.com/office/excel/2006/main"

    def mutate(contents: dict[str, bytes]) -> None:
        macro_sheet = ElementTree.fromstring(contents["xl/macrosheets/sheet1.xml"])
        macro_sheet.tag = f"{{{macro_namespace}}}notMacroSheet"
        contents["xl/macrosheets/sheet1.xml"] = ElementTree.tostring(
            macro_sheet,
            encoding="utf-8",
            xml_declaration=True,
        )

    return _rewrite_archive(path, mutate, ".xlm-macro-sheet-corrupt.tmp.xlsx")


def make_ribbon_customization_model(
    path: Path,
    *,
    office_2010: bool = False,
    compatibility_namespace: bool = False,
) -> Path:
    """Create a harmless raw RibbonX fixture outside ordinary workbook XML.

    The callback names and labels are synthetic private material. The fixture
    is never opened by Office; FormulaFence only reads its bounded package
    parts before the workbook reader can omit them.
    """
    make_model(path)
    package_relationships = (
        "http://schemas.openxmlformats.org/package/2006/relationships"
    )
    document_relationships = (
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    )
    if compatibility_namespace and not office_2010:
        raise ValueError("RibbonX compatibility namespace requires office_2010=True")
    custom_ui_namespace = (
        "http://schemas.microsoft.com/office/2007/10/customui"
        if compatibility_namespace
        else (
            "http://schemas.microsoft.com/office/2009/07/customui"
            if office_2010
            else "http://schemas.microsoft.com/office/2006/01/customui"
        )
    )
    custom_ui_member = (
        "customUI/customUI14.xml" if office_2010 else "customUI/customUI.xml"
    )
    customization_relationship = (
        "http://schemas.microsoft.com/office/2007/relationships/ui/extensibility"
        if office_2010
        else "http://schemas.microsoft.com/office/2006/relationships/ui/extensibility"
    )
    image_relationship = f"{document_relationships}/image"

    def serialize(root: ElementTree.Element) -> bytes:
        namespace = root.tag[1:].split("}", maxsplit=1)[0]
        if namespace in {package_relationships, custom_ui_namespace}:
            ElementTree.register_namespace("", namespace)
        return ElementTree.tostring(root, encoding="utf-8", xml_declaration=True)

    def mutate(contents: dict[str, bytes]) -> None:
        root_relationships = ElementTree.fromstring(contents["_rels/.rels"])
        ElementTree.SubElement(
            root_relationships,
            f"{{{package_relationships}}}Relationship",
            {
                "Id": "rIdFenceRibbonCustomization",
                "Type": customization_relationship,
                "Target": f"/{custom_ui_member}",
            },
        )
        contents["_rels/.rels"] = serialize(root_relationships)

        custom_ui = ElementTree.Element(
            f"{{{custom_ui_namespace}}}customUI",
            {"onLoad": "PrivateBaselineRibbonLoad"},
        )
        ribbon = ElementTree.SubElement(custom_ui, f"{{{custom_ui_namespace}}}ribbon")
        tabs = ElementTree.SubElement(ribbon, f"{{{custom_ui_namespace}}}tabs")
        tab = ElementTree.SubElement(
            tabs,
            f"{{{custom_ui_namespace}}}tab",
            {"id": "FenceTab", "label": "private baseline ribbon tab"},
        )
        group = ElementTree.SubElement(
            tab,
            f"{{{custom_ui_namespace}}}group",
            {"id": "FenceGroup", "label": "private baseline ribbon group"},
        )
        ElementTree.SubElement(
            group,
            f"{{{custom_ui_namespace}}}button",
            {
                "id": "FenceAction",
                "label": "private baseline ribbon action",
                "onAction": "PrivateBaselineRibbonAction",
                "image": "rIdFenceRibbonImage",
            },
        )
        contents[custom_ui_member] = serialize(custom_ui)

        image_member = "customUI/images/private-baseline-ribbon.png"
        part_relationships = ElementTree.Element(
            f"{{{package_relationships}}}Relationships"
        )
        ElementTree.SubElement(
            part_relationships,
            f"{{{package_relationships}}}Relationship",
            {
                "Id": "rIdFenceRibbonImage",
                "Type": image_relationship,
                "Target": "images/private-baseline-ribbon.png",
            },
        )
        relationship_member = (
            "customUI/_rels/"
            f"{custom_ui_member.rsplit('/', maxsplit=1)[-1]}.rels"
        )
        contents[relationship_member] = serialize(part_relationships)
        contents[image_member] = b"private baseline RibbonX image payload"

    return _rewrite_archive(path, mutate, ".ribbon-customization.tmp.xlsx")


def change_ribbon_customization_callback(path: Path) -> Path:
    """Change only a private RibbonX callback name."""

    def mutate(contents: dict[str, bytes]) -> None:
        custom_ui_member = next(
            name
            for name in contents
            if name.startswith("customUI/")
            and name.endswith(".xml")
            and "/_rels/" not in name
        )
        custom_ui = ElementTree.fromstring(contents[custom_ui_member])
        button = next(
            (
                element
                for element in custom_ui.iter()
                if element.tag.rsplit("}", maxsplit=1)[-1] == "button"
            ),
            None,
        )
        if button is None:
            raise ValueError("Fixture does not contain a RibbonX button")
        button.set("onAction", "PrivateCandidateRibbonAction")
        contents[custom_ui_member] = ElementTree.tostring(
            custom_ui,
            encoding="utf-8",
            xml_declaration=True,
        )

    return _rewrite_archive(path, mutate, ".ribbon-callback-change.tmp.xlsx")


def change_ribbon_customization_controls(path: Path) -> Path:
    """Change private RibbonX callbacks, labels, and image relationship material."""
    package_relationships = (
        "http://schemas.openxmlformats.org/package/2006/relationships"
    )

    def mutate(contents: dict[str, bytes]) -> None:
        custom_ui_member = next(
            name
            for name in contents
            if name.startswith("customUI/")
            and name.endswith(".xml")
            and "/_rels/" not in name
        )
        custom_ui = ElementTree.fromstring(contents[custom_ui_member])
        button = next(
            (
                element
                for element in custom_ui.iter()
                if element.tag.rsplit("}", maxsplit=1)[-1] == "button"
            ),
            None,
        )
        if button is None:
            raise ValueError("Fixture does not contain a RibbonX button")
        custom_ui.set("onLoad", "PrivateCandidateRibbonLoad")
        button.set("label", "private candidate ribbon action")
        button.set("onAction", "PrivateCandidateRibbonAction")
        contents[custom_ui_member] = ElementTree.tostring(
            custom_ui,
            encoding="utf-8",
            xml_declaration=True,
        )

        relationship_member = (
            "customUI/_rels/"
            f"{custom_ui_member.rsplit('/', maxsplit=1)[-1]}.rels"
        )
        relationships = ElementTree.fromstring(contents[relationship_member])
        relationship = next(
            (
                item
                for item in relationships.findall(
                    f"{{{package_relationships}}}Relationship"
                )
                if item.get("Id") == "rIdFenceRibbonImage"
            ),
            None,
        )
        if relationship is None:
            raise ValueError("Fixture does not contain a RibbonX image relationship")
        relationship.set("Target", "images/private-candidate-ribbon.png")
        contents[relationship_member] = ElementTree.tostring(
            relationships,
            encoding="utf-8",
            xml_declaration=True,
        )
        contents["customUI/images/private-candidate-ribbon.png"] = (
            b"private candidate RibbonX image payload"
        )

    return _rewrite_archive(path, mutate, ".ribbon-customization-change.tmp.xlsx")


def renumber_ribbon_customization_relationships(path: Path) -> Path:
    """Rewrite RibbonX package IDs while retaining their semantic bindings."""
    package_relationships = (
        "http://schemas.openxmlformats.org/package/2006/relationships"
    )
    replacements = {
        "rIdFenceRibbonCustomization": "rIdFenceRenumberedRibbonCustomization",
        "rIdFenceRibbonImage": "rIdFenceRenumberedRibbonImage",
    }

    def mutate(contents: dict[str, bytes]) -> None:
        root_relationships = ElementTree.fromstring(contents["_rels/.rels"])
        for relationship in root_relationships.findall(
            f"{{{package_relationships}}}Relationship"
        ):
            if replacement := replacements.get(relationship.get("Id")):
                relationship.set("Id", replacement)
        contents["_rels/.rels"] = ElementTree.tostring(
            root_relationships,
            encoding="utf-8",
            xml_declaration=True,
        )

        custom_ui_member = next(
            name
            for name in contents
            if name.startswith("customUI/")
            and name.endswith(".xml")
            and "/_rels/" not in name
        )
        custom_ui = ElementTree.fromstring(contents[custom_ui_member])
        for element in custom_ui.iter():
            if element.get("image") == "rIdFenceRibbonImage":
                element.set("image", "rIdFenceRenumberedRibbonImage")
        contents[custom_ui_member] = ElementTree.tostring(
            custom_ui,
            encoding="utf-8",
            xml_declaration=True,
        )

        relationship_member = (
            "customUI/_rels/"
            f"{custom_ui_member.rsplit('/', maxsplit=1)[-1]}.rels"
        )
        relationships = ElementTree.fromstring(contents[relationship_member])
        for relationship in relationships.findall(
            f"{{{package_relationships}}}Relationship"
        ):
            if replacement := replacements.get(relationship.get("Id")):
                relationship.set("Id", replacement)
        contents[relationship_member] = ElementTree.tostring(
            relationships,
            encoding="utf-8",
            xml_declaration=True,
        )

    return _rewrite_archive(path, mutate, ".ribbon-customization-renumber.tmp.xlsx")


def rewrite_ribbon_customization_internal_target_spelling(path: Path) -> Path:
    """Rewrite a RibbonX package target using an equivalent relative path."""
    package_relationships = (
        "http://schemas.openxmlformats.org/package/2006/relationships"
    )

    def mutate(contents: dict[str, bytes]) -> None:
        relationships = ElementTree.fromstring(contents["_rels/.rels"])
        customization = next(
            (
                relationship
                for relationship in relationships.findall(
                    f"{{{package_relationships}}}Relationship"
                )
                if relationship.get("Type", "").endswith("/ui/extensibility")
            ),
            None,
        )
        if customization is None:
            raise ValueError("Fixture does not contain a RibbonX package declaration")
        customization.set("Target", "./customUI/../customUI/customUI.xml")
        contents["_rels/.rels"] = ElementTree.tostring(
            relationships,
            encoding="utf-8",
            xml_declaration=True,
        )

    return _rewrite_archive(path, mutate, ".ribbon-customization-target.tmp.xlsx")


def corrupt_ribbon_customization_root(path: Path) -> Path:
    """Replace the RibbonX root with an unexpected element for coverage tests."""

    def mutate(contents: dict[str, bytes]) -> None:
        custom_ui_member = next(
            name
            for name in contents
            if name.startswith("customUI/")
            and name.endswith(".xml")
            and "/_rels/" not in name
        )
        custom_ui = ElementTree.fromstring(contents[custom_ui_member])
        namespace = custom_ui.tag[1:].split("}", maxsplit=1)[0]
        custom_ui.tag = f"{{{namespace}}}notCustomUI"
        contents[custom_ui_member] = ElementTree.tostring(
            custom_ui,
            encoding="utf-8",
            xml_declaration=True,
        )

    return _rewrite_archive(path, mutate, ".ribbon-customization-corrupt.tmp.xlsx")


def make_office_web_addin_model(
    path: Path,
    *,
    taskpane_reference_element: str = "webextension",
) -> Path:
    """Create a harmless Office Web Add-in task-pane package fixture.

    The synthetic extension is never opened by Office. FormulaFence only reads
    its bounded OOXML task-pane and web-extension parts before the workbook
    reader can omit them.
    """
    if taskpane_reference_element not in {"webextension", "webextensionref"}:
        raise ValueError("Unsupported task-pane web-extension reference element")
    make_model(path)
    content_types = "http://schemas.openxmlformats.org/package/2006/content-types"
    package_relationships = (
        "http://schemas.openxmlformats.org/package/2006/relationships"
    )
    document_relationships = (
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    )
    taskpanes_namespace = (
        "http://schemas.microsoft.com/office/webextensions/taskpanes/2010/11"
    )
    web_extension_namespace = (
        "http://schemas.microsoft.com/office/webextensions/webextension/2010/11"
    )
    taskpane_relationship = (
        "http://schemas.microsoft.com/office/2011/relationships/webextensiontaskpanes"
    )
    web_extension_relationship = (
        "http://schemas.microsoft.com/office/2011/relationships/webextension"
    )
    image_relationship = f"{document_relationships}/image"

    def serialize(root: ElementTree.Element) -> bytes:
        return ElementTree.tostring(root, encoding="utf-8", xml_declaration=True)

    def mutate(contents: dict[str, bytes]) -> None:
        workbook_relationships = ElementTree.fromstring(
            contents["xl/_rels/workbook.xml.rels"]
        )
        ElementTree.SubElement(
            workbook_relationships,
            f"{{{package_relationships}}}Relationship",
            {
                "Id": "rIdFenceWebTaskpanes",
                "Type": taskpane_relationship,
                "Target": "webextensions/taskpanes.xml",
            },
        )
        contents["xl/_rels/workbook.xml.rels"] = serialize(workbook_relationships)

        types = ElementTree.fromstring(contents["[Content_Types].xml"])
        for member, content_type in (
            (
                "/xl/webextensions/taskpanes.xml",
                "application/vnd.ms-office.webextensiontaskpanes+xml",
            ),
            (
                "/xl/webextensions/webextension1.xml",
                "application/vnd.ms-office.webextension+xml",
            ),
        ):
            ElementTree.SubElement(
                types,
                f"{{{content_types}}}Override",
                {"PartName": member, "ContentType": content_type},
            )
        contents["[Content_Types].xml"] = serialize(types)

        taskpanes = ElementTree.Element(f"{{{taskpanes_namespace}}}taskpanes")
        taskpane = ElementTree.SubElement(
            taskpanes,
            f"{{{taskpanes_namespace}}}taskpane",
            {
                "dockstate": "right",
                "visibility": "1",
                "width": "350",
                "row": "4",
                "locked": "1",
            },
        )
        ElementTree.SubElement(
            taskpane,
            f"{{{taskpanes_namespace}}}{taskpane_reference_element}",
            {f"{{{document_relationships}}}id": "rIdFenceTaskpaneExtension"},
        )
        contents["xl/webextensions/taskpanes.xml"] = serialize(taskpanes)

        taskpane_relationships = ElementTree.Element(
            f"{{{package_relationships}}}Relationships"
        )
        ElementTree.SubElement(
            taskpane_relationships,
            f"{{{package_relationships}}}Relationship",
            {
                "Id": "rIdFenceTaskpaneExtension",
                "Type": web_extension_relationship,
                "Target": "webextension1.xml",
            },
        )
        contents["xl/webextensions/_rels/taskpanes.xml.rels"] = serialize(
            taskpane_relationships
        )

        web_extension = ElementTree.Element(
            f"{{{web_extension_namespace}}}webextension",
            {"id": "{11111111-1111-1111-1111-111111111111}"},
        )
        ElementTree.SubElement(
            web_extension,
            f"{{{web_extension_namespace}}}reference",
            {
                "id": "PrivateBaselineAddin",
                "version": "1.0.0.0",
                "store": "private-baseline-manifest.xml",
                "storeType": "Filesystem",
            },
        )
        alternate_references = ElementTree.SubElement(
            web_extension,
            f"{{{web_extension_namespace}}}alternateReferences",
        )
        ElementTree.SubElement(
            alternate_references,
            f"{{{web_extension_namespace}}}reference",
            {
                "id": "PrivateFallbackAddin",
                "version": "1.0.0.0",
                "store": "private-fallback-manifest.xml",
                "storeType": "Filesystem",
            },
        )
        properties = ElementTree.SubElement(
            web_extension,
            f"{{{web_extension_namespace}}}properties",
        )
        ElementTree.SubElement(
            properties,
            f"{{{web_extension_namespace}}}property",
            {"name": "Office.AutoShowTaskpaneWithDocument", "value": "true"},
        )
        ElementTree.SubElement(
            properties,
            f"{{{web_extension_namespace}}}property",
            {"name": "PrivateAddinBehavior", "value": "private baseline behavior"},
        )
        bindings = ElementTree.SubElement(
            web_extension,
            f"{{{web_extension_namespace}}}bindings",
        )
        ElementTree.SubElement(
            bindings,
            f"{{{web_extension_namespace}}}binding",
            {
                "id": "PrivateBaselineBinding",
                "type": "table",
                "appref": "PrivateBaselineTable",
            },
        )
        ElementTree.SubElement(
            web_extension,
            f"{{{web_extension_namespace}}}snapshot",
            {f"{{{document_relationships}}}embed": "rIdFenceSnapshot"},
        )
        contents["xl/webextensions/webextension1.xml"] = serialize(web_extension)

        extension_relationships = ElementTree.Element(
            f"{{{package_relationships}}}Relationships"
        )
        ElementTree.SubElement(
            extension_relationships,
            f"{{{package_relationships}}}Relationship",
            {
                "Id": "rIdFenceSnapshot",
                "Type": image_relationship,
                "Target": "snapshots/private-baseline.png",
            },
        )
        ElementTree.SubElement(
            extension_relationships,
            f"{{{package_relationships}}}Relationship",
            {
                "Id": "rIdFenceExternalPreview",
                "Type": image_relationship,
                "Target": "https://private.example.invalid/addin-preview.png",
                "TargetMode": "External",
            },
        )
        contents["xl/webextensions/_rels/webextension1.xml.rels"] = serialize(
            extension_relationships
        )
        contents["xl/webextensions/snapshots/private-baseline.png"] = (
            b"private baseline add-in snapshot"
        )

    return _rewrite_archive(path, mutate, ".office-web-addin.tmp.xlsx")


def change_office_web_addin_auto_show(path: Path) -> Path:
    """Disable only the synthetic document auto-show request."""
    web_extension_namespace = (
        "http://schemas.microsoft.com/office/webextensions/webextension/2010/11"
    )

    def mutate(contents: dict[str, bytes]) -> None:
        web_extension = ElementTree.fromstring(
            contents["xl/webextensions/webextension1.xml"]
        )
        auto_show = next(
            (
                property_element
                for property_element in web_extension.iter(
                    f"{{{web_extension_namespace}}}property"
                )
                if property_element.get("name")
                == "Office.AutoShowTaskpaneWithDocument"
            ),
            None,
        )
        if auto_show is None:
            raise ValueError("Fixture does not contain an auto-show property")
        auto_show.set("value", "false")
        contents["xl/webextensions/webextension1.xml"] = ElementTree.tostring(
            web_extension,
            encoding="utf-8",
            xml_declaration=True,
        )

    return _rewrite_archive(path, mutate, ".office-web-addin-autoshow.tmp.xlsx")


def change_office_web_addin_controls(path: Path) -> Path:
    """Change private task-pane configuration and a snapshot relationship target."""
    package_relationships = (
        "http://schemas.openxmlformats.org/package/2006/relationships"
    )
    taskpanes_namespace = (
        "http://schemas.microsoft.com/office/webextensions/taskpanes/2010/11"
    )

    def mutate(contents: dict[str, bytes]) -> None:
        taskpanes = ElementTree.fromstring(contents["xl/webextensions/taskpanes.xml"])
        taskpane = taskpanes.find(f"{{{taskpanes_namespace}}}taskpane")
        if taskpane is None:
            raise ValueError("Fixture does not contain a task pane")
        taskpane.set("visibility", "0")
        taskpane.set("width", "475")
        contents["xl/webextensions/taskpanes.xml"] = ElementTree.tostring(
            taskpanes,
            encoding="utf-8",
            xml_declaration=True,
        )

        relationships_name = "xl/webextensions/_rels/webextension1.xml.rels"
        relationships = ElementTree.fromstring(contents[relationships_name])
        snapshot_relationship = next(
            (
                relationship
                for relationship in relationships.findall(
                    f"{{{package_relationships}}}Relationship"
                )
                if relationship.get("Id") == "rIdFenceSnapshot"
            ),
            None,
        )
        if snapshot_relationship is None:
            raise ValueError("Fixture does not contain a snapshot relationship")
        snapshot_relationship.set("Target", "snapshots/private-candidate.png")
        contents[relationships_name] = ElementTree.tostring(
            relationships,
            encoding="utf-8",
            xml_declaration=True,
        )
        contents["xl/webextensions/snapshots/private-candidate.png"] = (
            b"private candidate add-in snapshot"
        )

    return _rewrite_archive(path, mutate, ".office-web-addin-change.tmp.xlsx")


def renumber_office_web_addin_relationships(path: Path) -> Path:
    """Rewrite task-pane relationship IDs while preserving semantic bindings."""
    package_relationships = (
        "http://schemas.openxmlformats.org/package/2006/relationships"
    )
    document_relationships = (
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    )
    taskpanes_namespace = (
        "http://schemas.microsoft.com/office/webextensions/taskpanes/2010/11"
    )
    web_extension_namespace = (
        "http://schemas.microsoft.com/office/webextensions/webextension/2010/11"
    )
    replacements = {
        "rIdFenceWebTaskpanes": "rIdFenceRenumberedWebTaskpanes",
        "rIdFenceTaskpaneExtension": "rIdFenceRenumberedTaskpaneExtension",
        "rIdFenceSnapshot": "rIdFenceRenumberedSnapshot",
    }

    def replace_relationship_ids(root: ElementTree.Element) -> None:
        for relationship in root.findall(f"{{{package_relationships}}}Relationship"):
            if replacement := replacements.get(relationship.get("Id")):
                relationship.set("Id", replacement)

    def mutate(contents: dict[str, bytes]) -> None:
        for member in (
            "xl/_rels/workbook.xml.rels",
            "xl/webextensions/_rels/taskpanes.xml.rels",
            "xl/webextensions/_rels/webextension1.xml.rels",
        ):
            relationships = ElementTree.fromstring(contents[member])
            replace_relationship_ids(relationships)
            contents[member] = ElementTree.tostring(
                relationships,
                encoding="utf-8",
                xml_declaration=True,
            )

        taskpanes = ElementTree.fromstring(contents["xl/webextensions/taskpanes.xml"])
        for reference in taskpanes.iter():
            if reference.tag not in {
                f"{{{taskpanes_namespace}}}webextension",
                f"{{{taskpanes_namespace}}}webextensionref",
            }:
                continue
            if replacement := replacements.get(
                reference.get(f"{{{document_relationships}}}id")
            ):
                reference.set(f"{{{document_relationships}}}id", replacement)
        contents["xl/webextensions/taskpanes.xml"] = ElementTree.tostring(
            taskpanes,
            encoding="utf-8",
            xml_declaration=True,
        )

        web_extension = ElementTree.fromstring(
            contents["xl/webextensions/webextension1.xml"]
        )
        for snapshot in web_extension.iter(f"{{{web_extension_namespace}}}snapshot"):
            if replacement := replacements.get(
                snapshot.get(f"{{{document_relationships}}}embed")
            ):
                snapshot.set(f"{{{document_relationships}}}embed", replacement)
        contents["xl/webextensions/webextension1.xml"] = ElementTree.tostring(
            web_extension,
            encoding="utf-8",
            xml_declaration=True,
        )

    return _rewrite_archive(path, mutate, ".office-web-addin-renumber.tmp.xlsx")


def rewrite_office_web_addin_internal_target_spelling(path: Path) -> Path:
    """Use equivalent relative target spellings in the Office Web Add-in chain."""
    package_relationships = (
        "http://schemas.openxmlformats.org/package/2006/relationships"
    )

    def mutate(contents: dict[str, bytes]) -> None:
        workbook_relationships_name = "xl/_rels/workbook.xml.rels"
        workbook_relationships = ElementTree.fromstring(
            contents[workbook_relationships_name]
        )
        taskpane_relationship = next(
            (
                relationship
                for relationship in workbook_relationships.findall(
                    f"{{{package_relationships}}}Relationship"
                )
                if relationship.get("Id") == "rIdFenceWebTaskpanes"
            ),
            None,
        )
        if taskpane_relationship is None:
            raise ValueError("Fixture does not contain a workbook task-pane relationship")
        taskpane_relationship.set(
            "Target", "./webextensions/../webextensions/taskpanes.xml"
        )
        contents[workbook_relationships_name] = ElementTree.tostring(
            workbook_relationships,
            encoding="utf-8",
            xml_declaration=True,
        )

        taskpane_relationships_name = "xl/webextensions/_rels/taskpanes.xml.rels"
        taskpane_relationships = ElementTree.fromstring(
            contents[taskpane_relationships_name]
        )
        extension_relationship = next(
            (
                relationship
                for relationship in taskpane_relationships.findall(
                    f"{{{package_relationships}}}Relationship"
                )
                if relationship.get("Id") == "rIdFenceTaskpaneExtension"
            ),
            None,
        )
        if extension_relationship is None:
            raise ValueError("Fixture does not contain a task-pane extension relationship")
        extension_relationship.set("Target", "./webextension1.xml")
        contents[taskpane_relationships_name] = ElementTree.tostring(
            taskpane_relationships,
            encoding="utf-8",
            xml_declaration=True,
        )

        extension_relationships_name = "xl/webextensions/_rels/webextension1.xml.rels"
        extension_relationships = ElementTree.fromstring(
            contents[extension_relationships_name]
        )
        snapshot_relationship = next(
            (
                relationship
                for relationship in extension_relationships.findall(
                    f"{{{package_relationships}}}Relationship"
                )
                if relationship.get("Id") == "rIdFenceSnapshot"
            ),
            None,
        )
        if snapshot_relationship is None:
            raise ValueError("Fixture does not contain a snapshot relationship")
        snapshot_relationship.set(
            "Target", "./snapshots/../snapshots/private-baseline.png"
        )
        contents[extension_relationships_name] = ElementTree.tostring(
            extension_relationships,
            encoding="utf-8",
            xml_declaration=True,
        )

    return _rewrite_archive(path, mutate, ".office-web-addin-target.tmp.xlsx")


def corrupt_office_web_addin_definition_root(path: Path) -> Path:
    """Replace the extension root with an unexpected element for coverage tests."""
    web_extension_namespace = (
        "http://schemas.microsoft.com/office/webextensions/webextension/2010/11"
    )

    def mutate(contents: dict[str, bytes]) -> None:
        web_extension = ElementTree.fromstring(
            contents["xl/webextensions/webextension1.xml"]
        )
        web_extension.tag = f"{{{web_extension_namespace}}}notWebExtension"
        contents["xl/webextensions/webextension1.xml"] = ElementTree.tostring(
            web_extension,
            encoding="utf-8",
            xml_declaration=True,
        )

    return _rewrite_archive(path, mutate, ".office-web-addin-corrupt.tmp.xlsx")


def make_worksheet_embedded_control_model(
    path: Path,
    *,
    alternate_content: bool = False,
) -> Path:
    """Create harmless raw OOXML ActiveX, form-control, and OLE test material.

    The arbitrary binary bytes are never opened by an Office application. They
    exist only to prove FormulaFence's bounded, non-executing package scanner.
    """
    make_model(path)
    content_types = "http://schemas.openxmlformats.org/package/2006/content-types"
    package_relationships = _PACKAGE_RELATIONSHIPS_NS
    document_relationships = _DOCUMENT_RELATIONSHIPS_NS
    spreadsheet = _SPREADSHEETML_NS
    form_control_namespace = (
        "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"
    )
    activex = "http://schemas.microsoft.com/office/2006/activeX"
    markup_compatibility = "http://schemas.openxmlformats.org/markup-compatibility/2006"
    control_relationship = f"{document_relationships}/control"
    control_properties_relationship = f"{document_relationships}/ctrlProp"
    ole_object_relationship = f"{document_relationships}/oleObject"
    image_relationship = f"{document_relationships}/image"
    activex_binary_relationship = (
        "http://schemas.microsoft.com/office/2006/relationships/activeXControlBinary"
    )

    def serialize(root: ElementTree.Element) -> bytes:
        return ElementTree.tostring(root, encoding="utf-8", xml_declaration=True)

    def add_override(
        root: ElementTree.Element,
        part_name: str,
        content_type: str,
    ) -> None:
        ElementTree.SubElement(
            root,
            f"{{{content_types}}}Override",
            {"PartName": part_name, "ContentType": content_type},
        )

    def mutate(contents: dict[str, bytes]) -> None:
        types = ElementTree.fromstring(contents["[Content_Types].xml"])
        add_override(
            types,
            "/xl/activeX/activeX1.xml",
            "application/vnd.ms-office.activeX+xml",
        )
        add_override(
            types,
            "/xl/ctrlProps/ctrlProp1.xml",
            "application/vnd.ms-excel.controlproperties+xml",
        )
        contents["[Content_Types].xml"] = serialize(types)

        worksheet = _inputs_worksheet_root(contents)
        controls = ElementTree.SubElement(worksheet, f"{{{spreadsheet}}}controls")
        active_x_control = ElementTree.Element(
            f"{{{spreadsheet}}}control",
            {
                "shapeId": "1025",
                "name": "PrivateBaselineCommandButton",
                f"{{{document_relationships}}}id": "rIdFenceActiveX",
            },
        )
        ElementTree.SubElement(
            active_x_control,
            f"{{{spreadsheet}}}controlPr",
            {
                "defaultSize": "0",
                f"{{{document_relationships}}}id": "rIdFenceControlPresentation",
            },
        )
        form_control_element = ElementTree.Element(
            f"{{{spreadsheet}}}control",
            {
                "shapeId": "1028",
                "name": "PrivateBaselineFormControl",
                f"{{{document_relationships}}}id": "rIdFenceControlProperties",
            },
        )
        ElementTree.SubElement(
            form_control_element,
            f"{{{spreadsheet}}}controlPr",
            {
                "macro": "PrivateBaselineControlMacro",
                "linkedCell": "Inputs!$B$2",
                "listFillRange": "Inputs!$B$2:$B$4",
            },
        )
        if alternate_content:
            alternate = ElementTree.SubElement(
                controls,
                f"{{{markup_compatibility}}}AlternateContent",
            )
            choice = ElementTree.SubElement(
                alternate,
                f"{{{markup_compatibility}}}Choice",
                {"Requires": "x14"},
            )
            choice.append(active_x_control)
            fallback = ElementTree.SubElement(
                alternate,
                f"{{{markup_compatibility}}}Fallback",
            )
            ElementTree.SubElement(
                fallback,
                f"{{{spreadsheet}}}control",
                {
                    "shapeId": "1025",
                    "name": "PrivateBaselineCommandButton",
                    f"{{{document_relationships}}}id": "rIdFenceActiveX",
                },
            )
            controls.append(form_control_element)
        else:
            controls.extend((active_x_control, form_control_element))

        ole_objects = ElementTree.SubElement(
            worksheet,
            f"{{{spreadsheet}}}oleObjects",
        )
        embedded_ole = ElementTree.SubElement(
            ole_objects,
            f"{{{spreadsheet}}}oleObject",
            {
                "progId": "Private.Baseline.Embedded.Object",
                "shapeId": "1026",
                "autoLoad": "1",
                f"{{{document_relationships}}}id": "rIdFenceEmbeddedOle",
            },
        )
        ElementTree.SubElement(
            embedded_ole,
            f"{{{spreadsheet}}}objectPr",
            {f"{{{document_relationships}}}id": "rIdFenceOlePresentation"},
        )
        ElementTree.SubElement(
            ole_objects,
            f"{{{spreadsheet}}}oleObject",
            {
                "progId": "Private.Baseline.Linked.Object",
                "shapeId": "1027",
                "link": "private-baseline-link-name",
                "oleUpdate": "always",
                f"{{{document_relationships}}}id": "rIdFenceLinkedOle",
            },
        )
        _save_inputs_worksheet(contents, worksheet)

        worksheet_relationships = ElementTree.fromstring(
            contents.get(
                "xl/worksheets/_rels/sheet1.xml.rels",
                (
                    b'<?xml version="1.0" encoding="UTF-8"?>'
                    b'<Relationships xmlns="http://schemas.openxmlformats.org/'
                    b'package/2006/relationships"/>'
                ),
            )
        )
        relationship_tag = f"{{{package_relationships}}}Relationship"
        for relationship_id, relationship_type, target, target_mode in (
            (
                "rIdFenceActiveX",
                control_relationship,
                "../activeX/activeX1.xml",
                None,
            ),
            (
                "rIdFenceControlProperties",
                control_properties_relationship,
                "../ctrlProps/ctrlProp1.xml",
                None,
            ),
            (
                "rIdFenceControlPresentation",
                image_relationship,
                "../media/private-baseline-control.png",
                None,
            ),
            (
                "rIdFenceEmbeddedOle",
                ole_object_relationship,
                "../embeddings/private-baseline-ole.bin",
                None,
            ),
            (
                "rIdFenceLinkedOle",
                ole_object_relationship,
                "file:///private/baseline-linked-ole.bin",
                "External",
            ),
            (
                "rIdFenceOlePresentation",
                image_relationship,
                "../media/private-baseline-ole.png",
                None,
            ),
        ):
            attributes = {
                "Id": relationship_id,
                "Type": relationship_type,
                "Target": target,
            }
            if target_mode is not None:
                attributes["TargetMode"] = target_mode
            ElementTree.SubElement(worksheet_relationships, relationship_tag, attributes)
        contents["xl/worksheets/_rels/sheet1.xml.rels"] = serialize(
            worksheet_relationships
        )

        activex_root = ElementTree.Element(
            f"{{{activex}}}ocx",
            {
                "classid": "{11111111-1111-1111-1111-111111111111}",
                "license": "private-baseline-activex-license",
                "persistence": "persistStreamInit",
                f"{{{document_relationships}}}id": "rIdFenceActiveXBinary",
            },
        )
        contents["xl/activeX/activeX1.xml"] = serialize(activex_root)
        activex_relationships = ElementTree.Element(
            f"{{{package_relationships}}}Relationships"
        )
        ElementTree.SubElement(
            activex_relationships,
            relationship_tag,
            {
                "Id": "rIdFenceActiveXBinary",
                "Type": activex_binary_relationship,
                "Target": "activeX1.bin",
            },
        )
        contents["xl/activeX/_rels/activeX1.xml.rels"] = serialize(
            activex_relationships
        )

        form_control_root = ElementTree.Element(
            f"{{{form_control_namespace}}}formControlPr",
            {
                "objectType": "Drop",
                "fmlaGroup": "Inputs!$B$2",
                "fmlaLink": "Inputs!$B$3",
                "fmlaRange": "Inputs!$B$2:$B$4",
                "fmlaTxbx": "Inputs!$A$1",
            },
        )
        contents["xl/ctrlProps/ctrlProp1.xml"] = serialize(form_control_root)
        contents["xl/activeX/activeX1.bin"] = b"private baseline ActiveX binary payload"
        contents["xl/embeddings/private-baseline-ole.bin"] = (
            b"private baseline embedded OLE payload"
        )
        contents["xl/media/private-baseline-ole.png"] = b"private OLE presentation"
        contents["xl/media/private-baseline-control.png"] = b"private control presentation"

    return _rewrite_archive(path, mutate, ".worksheet-embedded-control.tmp.xlsx")


def change_worksheet_embedded_control_controls(path: Path) -> Path:
    """Change private worksheet-control, ActiveX, and OLE configuration material."""
    package_relationships = _PACKAGE_RELATIONSHIPS_NS
    spreadsheet = _SPREADSHEETML_NS

    def mutate(contents: dict[str, bytes]) -> None:
        worksheet = _inputs_worksheet_root(contents)
        control_properties = next(
            (
                properties
                for properties in worksheet.iter(f"{{{spreadsheet}}}controlPr")
                if properties.get("macro") is not None
            ),
            None,
        )
        embedded_ole = worksheet.find(
            f".//{{{spreadsheet}}}oleObject"
        )
        if control_properties is None or embedded_ole is None:
            raise ValueError("Fixture does not contain worksheet embedded controls")
        control_properties.set("macro", "PrivateCandidateControlMacro")
        control_properties.set("linkedCell", "Inputs!$B$4")
        embedded_ole.set("autoLoad", "0")
        contents["xl/worksheets/sheet1.xml"] = ElementTree.tostring(
            worksheet,
            encoding="utf-8",
            xml_declaration=True,
        )

        activex_root = ElementTree.fromstring(contents["xl/activeX/activeX1.xml"])
        activex_root.set("classid", "{22222222-2222-2222-2222-222222222222}")
        activex_root.set("license", "private-candidate-activex-license")
        contents["xl/activeX/activeX1.xml"] = ElementTree.tostring(
            activex_root,
            encoding="utf-8",
            xml_declaration=True,
        )

        form_control = ElementTree.fromstring(contents["xl/ctrlProps/ctrlProp1.xml"])
        form_control.set("fmlaRange", "Inputs!$B$3:$B$4")
        contents["xl/ctrlProps/ctrlProp1.xml"] = ElementTree.tostring(
            form_control,
            encoding="utf-8",
            xml_declaration=True,
        )

        relationships_name = "xl/worksheets/_rels/sheet1.xml.rels"
        relationships = ElementTree.fromstring(contents[relationships_name])
        relationship = next(
            (
                item
                for item in relationships.findall(
                    f"{{{package_relationships}}}Relationship"
                )
                if item.get("Id") == "rIdFenceEmbeddedOle"
            ),
            None,
        )
        if relationship is None:
            raise ValueError("Fixture does not contain an embedded OLE relationship")
        relationship.set("Target", "../embeddings/private-candidate-ole.bin")
        contents[relationships_name] = ElementTree.tostring(
            relationships,
            encoding="utf-8",
            xml_declaration=True,
        )
        contents["xl/embeddings/private-candidate-ole.bin"] = (
            b"private candidate embedded OLE payload"
        )

    return _rewrite_archive(path, mutate, ".worksheet-embedded-control-change.tmp.xlsx")


def change_worksheet_embedded_control_payload(path: Path) -> Path:
    """Change only an embedded OLE payload to exercise private byte fingerprinting."""

    def mutate(contents: dict[str, bytes]) -> None:
        contents["xl/embeddings/private-baseline-ole.bin"] = (
            b"private candidate changed embedded OLE payload"
        )

    return _rewrite_archive(path, mutate, ".worksheet-embedded-control-payload.tmp.xlsx")


def renumber_worksheet_embedded_control_relationships(path: Path) -> Path:
    """Rewrite only writer-chosen relationship IDs across the control chain."""
    package_relationships = _PACKAGE_RELATIONSHIPS_NS
    document_relationships = _DOCUMENT_RELATIONSHIPS_NS
    replacements = {
        "rIdFenceActiveX": "rIdFenceRenumberedActiveX",
        "rIdFenceControlProperties": "rIdFenceRenumberedControlProperties",
        "rIdFenceControlPresentation": "rIdFenceRenumberedControlPresentation",
        "rIdFenceEmbeddedOle": "rIdFenceRenumberedEmbeddedOle",
        "rIdFenceLinkedOle": "rIdFenceRenumberedLinkedOle",
        "rIdFenceOlePresentation": "rIdFenceRenumberedOlePresentation",
        "rIdFenceActiveXBinary": "rIdFenceRenumberedActiveXBinary",
    }

    def replace_relationship_ids(root: ElementTree.Element) -> None:
        for relationship in root.findall(f"{{{package_relationships}}}Relationship"):
            if replacement := replacements.get(relationship.get("Id")):
                relationship.set("Id", replacement)

    def replace_xml_ids(root: ElementTree.Element) -> None:
        relationship_attribute = f"{{{document_relationships}}}id"
        for element in root.iter():
            if replacement := replacements.get(element.get(relationship_attribute)):
                element.set(relationship_attribute, replacement)

    def mutate(contents: dict[str, bytes]) -> None:
        for member in (
            "xl/worksheets/_rels/sheet1.xml.rels",
            "xl/activeX/_rels/activeX1.xml.rels",
        ):
            relationships = ElementTree.fromstring(contents[member])
            replace_relationship_ids(relationships)
            contents[member] = ElementTree.tostring(
                relationships,
                encoding="utf-8",
                xml_declaration=True,
            )
        for member in ("xl/worksheets/sheet1.xml", "xl/activeX/activeX1.xml"):
            root = ElementTree.fromstring(contents[member])
            replace_xml_ids(root)
            contents[member] = ElementTree.tostring(
                root,
                encoding="utf-8",
                xml_declaration=True,
            )

    return _rewrite_archive(path, mutate, ".worksheet-embedded-control-renumber.tmp.xlsx")


def rewrite_worksheet_embedded_control_internal_target_spelling(path: Path) -> Path:
    """Use equivalent relative spellings without changing control semantics."""
    package_relationships = _PACKAGE_RELATIONSHIPS_NS

    def set_target(
        relationships: ElementTree.Element,
        relationship_id: str,
        target: str,
    ) -> None:
        relationship = next(
            (
                item
                for item in relationships.findall(
                    f"{{{package_relationships}}}Relationship"
                )
                if item.get("Id") == relationship_id
            ),
            None,
        )
        if relationship is None:
            raise ValueError(f"Fixture does not contain relationship {relationship_id}")
        relationship.set("Target", target)

    def mutate(contents: dict[str, bytes]) -> None:
        worksheet_relationships_name = "xl/worksheets/_rels/sheet1.xml.rels"
        worksheet_relationships = ElementTree.fromstring(
            contents[worksheet_relationships_name]
        )
        set_target(
            worksheet_relationships,
            "rIdFenceActiveX",
            "../activeX/./activeX1.xml",
        )
        set_target(
            worksheet_relationships,
            "rIdFenceControlProperties",
            "../ctrlProps/./ctrlProp1.xml",
        )
        set_target(
            worksheet_relationships,
            "rIdFenceControlPresentation",
            "../media/./private-baseline-control.png",
        )
        set_target(
            worksheet_relationships,
            "rIdFenceEmbeddedOle",
            "../embeddings/./private-baseline-ole.bin",
        )
        contents[worksheet_relationships_name] = ElementTree.tostring(
            worksheet_relationships,
            encoding="utf-8",
            xml_declaration=True,
        )

        activex_relationships_name = "xl/activeX/_rels/activeX1.xml.rels"
        activex_relationships = ElementTree.fromstring(
            contents[activex_relationships_name]
        )
        set_target(
            activex_relationships,
            "rIdFenceActiveXBinary",
            "./activeX1.bin",
        )
        contents[activex_relationships_name] = ElementTree.tostring(
            activex_relationships,
            encoding="utf-8",
            xml_declaration=True,
        )

    return _rewrite_archive(path, mutate, ".worksheet-embedded-control-target.tmp.xlsx")


def corrupt_worksheet_embedded_control_activex_root(path: Path) -> Path:
    """Replace an ActiveX root with unexpected XML for fail-closed coverage tests."""
    activex = "http://schemas.microsoft.com/office/2006/activeX"

    def mutate(contents: dict[str, bytes]) -> None:
        root = ElementTree.fromstring(contents["xl/activeX/activeX1.xml"])
        root.tag = f"{{{activex}}}notOcx"
        contents["xl/activeX/activeX1.xml"] = ElementTree.tostring(
            root,
            encoding="utf-8",
            xml_declaration=True,
        )

    return _rewrite_archive(path, mutate, ".worksheet-embedded-control-corrupt.tmp.xlsx")


def make_protection_model(path: Path, *, include_chartsheet: bool = False) -> Path:
    """Create a workbook with operational protection and sparse style controls."""
    workbook = Workbook()
    inputs = workbook.active
    inputs.title = "Inputs"
    inputs["A1"] = "Controlled input"
    inputs["B2"] = 10
    inputs["C2"] = "=B2*2"
    inputs["D2"] = "Column default"
    inputs["B2"].protection = Protection(locked=False)
    inputs["C2"].protection = Protection(hidden=True)
    inputs.row_dimensions[5].protection = Protection(hidden=True)
    inputs.column_dimensions["D"].protection = Protection(locked=False)
    inputs.protection.sheet = True
    inputs.protection.formatCells = False
    inputs.protection.sort = False
    inputs.protection.autoFilter = False
    inputs.protection.selectLockedCells = True
    inputs.protection.set_password("synthetic-sheet-password")

    workbook.security.lockStructure = True
    workbook.security.set_workbook_password("synthetic-workbook-password")
    if include_chartsheet:
        chart = BarChart()
        chart.add_data(
            Reference(inputs, min_col=2, min_row=2, max_row=2), titles_from_data=False
        )
        dashboard = workbook.create_chartsheet("Dashboard")
        dashboard.add_chart(chart)
        dashboard.sheetProtection = ChartsheetProtection(
            content=True,
            objects=True,
            password="synthetic-chart-password",
        )
    workbook.save(path)
    return add_protected_range(path)


def add_protected_range(
    path: Path,
    *,
    name: str = "Synthetic approved inputs",
    sqref: str = "B2:B5",
    password: str = "A1B2",
    security_descriptor: str = "synthetic-security-descriptor",
) -> Path:
    """Add an OOXML protected range that openpyxl intentionally does not model."""
    namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"

    def mutate(contents: dict[str, bytes]) -> None:
        root = _inputs_worksheet_root(contents)
        protected_ranges = root.find(f"{{{namespace}}}protectedRanges")
        if protected_ranges is None:
            protected_ranges = ElementTree.Element(f"{{{namespace}}}protectedRanges")
            protection = root.find(f"{{{namespace}}}sheetProtection")
            insertion_index = (
                list(root).index(protection) + 1
                if protection is not None
                else len(root)
            )
            root.insert(insertion_index, protected_ranges)
        protected_range = ElementTree.SubElement(
            protected_ranges,
            f"{{{namespace}}}protectedRange",
        )
        protected_range.set("name", name)
        protected_range.set("sqref", sqref)
        protected_range.set("password", password)
        protected_range.set("securityDescriptor", security_descriptor)
        _save_inputs_worksheet(contents, root)

    return _rewrite_archive(path, mutate, ".protected-range.tmp.xlsx")


def set_sheet_protection_defaults(path: Path, *, explicit: bool) -> Path:
    """Toggle equivalent omitted versus explicit normal-sheet protection defaults."""
    namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    defaults = {
        "objects": "0",
        "scenarios": "0",
        "formatCells": "1",
        "formatColumns": "1",
        "formatRows": "1",
        "insertColumns": "1",
        "insertRows": "1",
        "insertHyperlinks": "1",
        "deleteColumns": "1",
        "deleteRows": "1",
        "selectLockedCells": "0",
        "sort": "1",
        "autoFilter": "1",
        "pivotTables": "1",
        "selectUnlockedCells": "0",
    }

    def mutate(contents: dict[str, bytes]) -> None:
        root = _inputs_worksheet_root(contents)
        protection = root.find(f"{{{namespace}}}sheetProtection")
        if protection is None:
            raise ValueError("Fixture does not contain sheet protection")
        for attribute, value in defaults.items():
            if explicit:
                protection.set(attribute, value)
            else:
                protection.attrib.pop(attribute, None)
        _save_inputs_worksheet(contents, root)

    return _rewrite_archive(path, mutate, ".protection-defaults.tmp.xlsx")


def set_sheet_protection_modern_verifier(path: Path, hash_value: str) -> Path:
    """Set synthetic SHA-512 verifier fields for redaction and equality tests."""
    namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"

    def mutate(contents: dict[str, bytes]) -> None:
        root = _inputs_worksheet_root(contents)
        protection = root.find(f"{{{namespace}}}sheetProtection")
        if protection is None:
            raise ValueError("Fixture does not contain sheet protection")
        protection.attrib.pop("password", None)
        protection.set("algorithmName", "SHA-512")
        protection.set("hashValue", hash_value)
        protection.set("saltValue", "c3ludGhldGljLXNhbHQ=")
        protection.set("spinCount", "100000")
        _save_inputs_worksheet(contents, root)

    return _rewrite_archive(path, mutate, ".modern-verifier.tmp.xlsx")


def change_protected_range(
    path: Path,
    *,
    sqref: str | None = None,
    name: str | None = None,
    password: str | None = None,
    security_descriptor: str | None = None,
) -> Path:
    """Change one synthetic protected range's non-secret test properties."""
    namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"

    def mutate(contents: dict[str, bytes]) -> None:
        root = _inputs_worksheet_root(contents)
        protected_range = root.find(
            f"{{{namespace}}}protectedRanges/{{{namespace}}}protectedRange"
        )
        if protected_range is None:
            raise ValueError("Fixture does not contain a protected range")
        for attribute, value in (
            ("sqref", sqref),
            ("name", name),
            ("password", password),
            ("securityDescriptor", security_descriptor),
        ):
            if value is not None:
                protected_range.set(attribute, value)
        _save_inputs_worksheet(contents, root)

    return _rewrite_archive(path, mutate, ".protected-range-change.tmp.xlsx")


def reorder_conditional_differential_styles(path: Path) -> Path:
    """Swap ``dxfs`` and their rule ids without changing any visual rule."""
    with ZipFile(path) as archive:
        contents = {
            entry.filename: archive.read(entry.filename) for entry in archive.infolist()
        }
    namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    styles = ElementTree.fromstring(contents["xl/styles.xml"])
    dxfs = styles.find(f"{{{namespace}}}dxfs")
    if dxfs is None or len(dxfs) < 2:
        raise ValueError("Fixture needs at least two conditional differential styles")
    first, second = dxfs[:2]
    dxfs[:2] = [second, first]
    contents["xl/styles.xml"] = ElementTree.tostring(
        styles,
        encoding="utf-8",
        xml_declaration=True,
    )
    rule_tag = f"{{{namespace}}}cfRule"
    for name, content in tuple(contents.items()):
        if not name.startswith("xl/worksheets/"):
            continue
        worksheet = ElementTree.fromstring(content)
        for rule in worksheet.iter(rule_tag):
            if rule.get("dxfId") == "0":
                rule.set("dxfId", "1")
            elif rule.get("dxfId") == "1":
                rule.set("dxfId", "0")
        contents[name] = ElementTree.tostring(
            worksheet,
            encoding="utf-8",
            xml_declaration=True,
        )
    staging = path.with_suffix(".dxf-order.tmp.xlsx")
    with ZipFile(staging, "w", compression=ZIP_DEFLATED) as archive:
        for name, content in contents.items():
            archive.writestr(name, content)
    staging.replace(path)
    return path


def add_conditional_formatting_databar_extension(
    path: Path,
    *,
    guid: str,
    axis_color: str,
    worksheet_extension_uri: str = "{78C0D931-6437-407D-A8EE-F0AAD7539E65}",
) -> Path:
    """Add a minimal Excel-2010 data-bar extension that openpyxl does not model."""
    with ZipFile(path) as archive:
        contents = {
            entry.filename: archive.read(entry.filename) for entry in archive.infolist()
        }
    main = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    x14 = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"
    xm = "http://schemas.microsoft.com/office/excel/2006/main"
    rule_tag = f"{{{main}}}cfRule"
    data_bar_rule: ElementTree.Element | None = None
    worksheet_name: str | None = None
    for name, content in tuple(contents.items()):
        if not name.startswith("xl/worksheets/"):
            continue
        worksheet = ElementTree.fromstring(content)
        data_bar_rule = next(
            (rule for rule in worksheet.iter(rule_tag) if rule.get("type") == "dataBar"),
            None,
        )
        if data_bar_rule is None:
            continue
        worksheet_name = name
        base_extensions = ElementTree.SubElement(data_bar_rule, f"{{{main}}}extLst")
        base_extension = ElementTree.SubElement(
            base_extensions,
            f"{{{main}}}ext",
            {"uri": "{B025F937-C7B1-47D3-B67F-A62EFF666E3E}"},
        )
        ElementTree.SubElement(base_extension, f"{{{x14}}}id").text = guid

        worksheet_extensions = ElementTree.SubElement(worksheet, f"{{{main}}}extLst")
        worksheet_extension = ElementTree.SubElement(
            worksheet_extensions,
            f"{{{main}}}ext",
            {"uri": worksheet_extension_uri},
        )
        conditional_formattings = ElementTree.SubElement(
            worksheet_extension,
            f"{{{x14}}}conditionalFormattings",
        )
        conditional_formatting = ElementTree.SubElement(
            conditional_formattings,
            f"{{{x14}}}conditionalFormatting",
        )
        extension_rule = ElementTree.SubElement(
            conditional_formatting,
            f"{{{x14}}}cfRule",
            {"type": "dataBar", "id": guid},
        )
        extension_data_bar = ElementTree.SubElement(
            extension_rule,
            f"{{{x14}}}dataBar",
            {"minLength": "0", "maxLength": "100", "axisPosition": "middle"},
        )
        ElementTree.SubElement(extension_data_bar, f"{{{x14}}}cfvo", {"type": "autoMin"})
        ElementTree.SubElement(extension_data_bar, f"{{{x14}}}cfvo", {"type": "autoMax"})
        ElementTree.SubElement(extension_data_bar, f"{{{x14}}}axisColor", {"rgb": axis_color})
        ElementTree.SubElement(conditional_formatting, f"{{{xm}}}sqref").text = "C2:C100"
        contents[name] = ElementTree.tostring(
            worksheet,
            encoding="utf-8",
            xml_declaration=True,
        )
        break
    if data_bar_rule is None or worksheet_name is None:
        raise ValueError("Could not find the fixture data-bar rule")
    staging = path.with_suffix(".conditional-extension.tmp.xlsx")
    with ZipFile(staging, "w", compression=ZIP_DEFLATED) as archive:
        for name, content in contents.items():
            archive.writestr(name, content)
    staging.replace(path)
    return path


def make_current_row_table_model(path: Path) -> Path:
    """Create calculated-column formulas using three current-row spellings."""
    workbook = Workbook()
    data = workbook.active
    data.title = "Data"
    data.append(["Sales Amount", "Rate", "Value"])
    data.append([10, 0.1, "=[@[Sales Amount]]*[@Rate]"])
    data.append([20, 0.2, "=[Sales Amount]*[Rate]"])
    data.append([30, 0.3, "=Sales[[#This Row],[Sales Amount]]*Sales[@Rate]"])
    data.add_table(Table(displayName="Sales", ref="A1:C4"))
    data["E1"] = "Adjacent current-row output"
    data["E2"] = "=Sales[[#This Row],[Sales Amount]]"

    report = workbook.create_sheet("Report")
    report["A1"] = "Current-row table output"
    report["B2"] = "=SUM(Sales[Value])"
    workbook.save(path)
    return path


def make_three_d_model(path: Path) -> Path:
    """Create a period-stack workbook with one static 3-D summary formula."""
    workbook = Workbook()
    january = workbook.active
    january.title = "Jan"
    january["A1"] = "Period input"
    january["B2"] = 10

    for title, amount in (("Feb", 20), ("Mar", 30)):
        period = workbook.create_sheet(title)
        period["A1"] = "Period input"
        period["B2"] = amount

    summary = workbook.create_sheet("Summary")
    summary["A1"] = "3-D consolidation"
    summary["B2"] = "=SUM(Jan:Mar!B2)"
    workbook.save(path)
    return path


def make_spill_model(path: Path) -> Path:
    """Create literal and OOXML-style dynamic-array spill-reference callers."""
    workbook = Workbook()
    inputs = workbook.active
    inputs.title = "Inputs"
    inputs["A1"] = "Dynamic-array anchors"
    inputs["B2"] = "=SEQUENCE(3)"
    inputs["B3"] = "=SEQUENCE(2)"

    model = workbook.create_sheet("Model")
    model["A1"] = "Spill-driven calculations"
    model["B2"] = "=SUM(Inputs!B2#)"
    model["B3"] = "=COUNTA(_xlfn.ANCHORARRAY(Inputs!B3))"

    dashboard = workbook.create_sheet("Dashboard")
    dashboard["A1"] = "Spill-driven output"
    dashboard["B2"] = "=Model!B2"
    workbook.save(path)
    return path


def make_implicit_intersection_model(path: Path) -> Path:
    """Create persisted SINGLE() and literal @ consumers of a static input range."""
    workbook = Workbook()
    inputs = workbook.active
    inputs.title = "Inputs"
    inputs["A1"] = "Implicit-intersection inputs"
    inputs["B2"] = 10
    inputs["B3"] = 20
    inputs["B4"] = 30

    model = workbook.create_sheet("Model")
    model["A1"] = "Implicit-intersection calculations"
    model["B2"].value = ArrayFormula(
        ref="B2", text="=_xlfn.SINGLE(Inputs!B2:B4)"
    )
    model["B3"] = "=@Inputs!B2:B4"

    dashboard = workbook.create_sheet("Dashboard")
    dashboard["A1"] = "Implicit-intersection output"
    dashboard["B2"] = "=Model!B2"
    workbook.save(path)
    return path


def make_legacy_array_model(path: Path, array_ref: str = "B1:B3") -> Path:
    """Create a fixed CSE array whose result members feed ordinary formulas."""
    workbook = Workbook()
    inputs = workbook.active
    inputs.title = "Inputs"
    inputs["A1"] = "a"
    inputs["A2"] = "bb"
    inputs["A3"] = "ccc"

    model = workbook.create_sheet("Model")
    model["A1"] = "Legacy CSE array result"
    model["B1"].value = ArrayFormula(ref=array_ref, text="=LEN(Inputs!A1:A3)")
    model["C2"] = "=B2*10"

    dashboard = workbook.create_sheet("Dashboard")
    dashboard["A1"] = "Array output consumers"
    dashboard["B2"] = "=SUM(Model!B2:B3)"
    workbook.save(path)
    return path


def mark_array_formula_dynamic(path: Path, anchor: str = "B1") -> Path:
    """Add the OOXML dynamic-array metadata that openpyxl does not write.

    The small fixture mirrors the publicly documented XlsxWriter serialization:
    an array formula anchor has ``cm=1`` and cell metadata resolves that index
    to an ``XLDAPR`` record with ``fDynamic=1``.
    """
    anchor_bytes = anchor.encode("ascii")
    with ZipFile(path) as archive:
        contents = {
            entry.filename: archive.read(entry.filename) for entry in archive.infolist()
        }

    worksheet_name: str | None = None
    needle = b'<c r="' + anchor_bytes + b'"><f t="array"'
    replacement = b'<c r="' + anchor_bytes + b'" cm="1"><f t="array"'
    for name, content in contents.items():
        if not name.startswith("xl/worksheets/") or needle not in content:
            continue
        if worksheet_name is not None:
            raise ValueError(f"More than one array formula anchor found for {anchor}")
        contents[name] = content.replace(needle, replacement, 1)
        worksheet_name = name
    if worksheet_name is None:
        raise ValueError(f"Could not find array formula anchor {anchor}")

    metadata = (
        b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        b'<metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        b'xmlns:xda="http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray">'
        b'<metadataTypes count="1"><metadataType name="XLDAPR" '
        b'minSupportedVersion="120000" copy="1" pasteAll="1" pasteValues="1" '
        b'merge="1" splitFirst="1" rowColShift="1" clearFormats="1" clearComments="1" '
        b'assign="1" coerce="1" cellMeta="1"/></metadataTypes>'
        b'<futureMetadata name="XLDAPR" count="1"><bk><extLst>'
        b'<ext uri="{bdbb8cdc-fa1e-496e-a857-3c3f30c029c3}">'
        b'<xda:dynamicArrayProperties fDynamic="1" fCollapsed="0"/>'
        b'</ext></extLst></bk></futureMetadata><cellMetadata count="1"><bk>'
        b'<rc t="1" v="0"/></bk></cellMetadata></metadata>'
    )
    content_types = contents["[Content_Types].xml"]
    contents["[Content_Types].xml"] = content_types.replace(
        b"</Types>",
        b'<Override PartName="/xl/metadata.xml" '
        b'ContentType="application/vnd.openxmlformats-officedocument.'
        b'spreadsheetml.sheetMetadata+xml"/></Types>',
        1,
    )
    relationships_name = "xl/_rels/workbook.xml.rels"
    relationships = contents[relationships_name]
    contents[relationships_name] = relationships.replace(
        b"</Relationships>",
        b'<Relationship Id="rIdFormulaFenceDynamicArrayMetadata" '
        b'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/'
        b'sheetMetadata" Target="metadata.xml"/></Relationships>',
        1,
    )
    contents["xl/metadata.xml"] = metadata

    staging = path.with_suffix(".dynamic.tmp.xlsx")
    with ZipFile(staging, "w", compression=ZIP_DEFLATED) as archive:
        for name, content in contents.items():
            archive.writestr(name, content)
    staging.replace(path)
    return path


def mark_array_formula_unclassified(path: Path, anchor: str = "B1") -> Path:
    """Give an array formula an unresolvable OOXML cell-metadata marker."""
    anchor_bytes = anchor.encode("ascii")
    needle = b'<c r="' + anchor_bytes + b'"><f t="array"'
    replacement = b'<c r="' + anchor_bytes + b'" cm="1"><f t="array"'
    staging = path.with_suffix(".array-metadata.tmp.xlsx")
    with ZipFile(path) as source, ZipFile(staging, "w", compression=ZIP_DEFLATED) as target:
        replaced = False
        for entry in source.infolist():
            content = source.read(entry.filename)
            if (
                not replaced
                and entry.filename.startswith("xl/worksheets/")
                and needle in content
            ):
                content = content.replace(needle, replacement, 1)
                replaced = True
            target.writestr(entry.filename, content)
    if not replaced:
        staging.unlink()
        raise ValueError(f"Could not find array formula anchor {anchor}")
    staging.replace(path)
    return path


def make_named_formula_model(path: Path) -> Path:
    """Create global and local names whose definitions contain formulas."""
    workbook = Workbook()
    inputs = workbook.active
    inputs.title = "Inputs"
    inputs["A1"] = "Named-formula inputs"
    inputs["B2"] = 0.1
    inputs["B3"] = 100
    inputs["B4"] = 7

    summary = workbook.create_sheet("Summary")
    summary["A1"] = "Named-formula outputs"
    summary["B2"] = "=DiscountedValue"
    summary["B3"] = "=StaticRate"
    summary["B4"] = "=LocalMetric"

    report = workbook.create_sheet("Report")
    report["A1"] = "Qualified local-name output"
    report["B2"] = "=Summary!LocalMetric"

    workbook.defined_names.add(DefinedName("TaxRate", attr_text="=Inputs!$B$2"))
    workbook.defined_names.add(
        DefinedName(
            "DiscountedValue",
            attr_text="=LET(rate,TaxRate,amount,Inputs!$B$3,amount*(1-rate))",
        )
    )
    workbook.defined_names.add(DefinedName("StaticRate", attr_text="=0.05"))
    workbook.defined_names.add(
        DefinedName("LocalMetric", attr_text="=Inputs!$B$4*2", localSheetId=1)
    )
    workbook.save(path)
    return path


def make_let_model(path: Path) -> Path:
    """Create a model whose calculation uses lexical LET variables."""
    workbook = Workbook()
    inputs = workbook.active
    inputs.title = "Inputs"
    inputs["A1"] = "LET inputs"
    inputs["B2"] = 0.1
    inputs["B3"] = 100

    model = workbook.create_sheet("Model")
    model["A1"] = "LET calculation"
    model["B2"] = "=LET(rate,Inputs!B2,amount,Inputs!B3,amount*(1-rate))"

    dashboard = workbook.create_sheet("Dashboard")
    dashboard["A1"] = "LET output"
    dashboard["B2"] = "=Model!B2"
    workbook.save(path)
    return path


def make_named_lambda_model(path: Path) -> Path:
    """Create reusable named LAMBDAs plus a formula-defined name that calls one."""
    workbook = Workbook()
    inputs = workbook.active
    inputs.title = "Inputs"
    inputs["A1"] = "Named LAMBDA inputs"
    inputs["B2"] = 0.1
    inputs["B3"] = 100
    inputs["B4"] = 32

    model = workbook.create_sheet("Model")
    model["A1"] = "Named LAMBDA calculations"
    model["B2"] = "=ToCelsius(Inputs!B4)"
    model["B3"] = "=AdjustedCelsius(Inputs!B4)"
    model["B4"] = "=NamedAdjustedValue"

    dashboard = workbook.create_sheet("Dashboard")
    dashboard["A1"] = "Named LAMBDA output"
    dashboard["B2"] = "=Model!B4"

    workbook.defined_names.add(
        DefinedName(
            "ToCelsius",
            attr_text="_xlfn.LAMBDA(_xlpm.temp,(5/9)*(_xlpm.temp-32)+Inputs!$B$2)",
        )
    )
    workbook.defined_names.add(
        DefinedName(
            "AdjustedCelsius",
            attr_text="_xlfn.LAMBDA(_xlpm.temp,ToCelsius(_xlpm.temp)+Inputs!$B$3)",
        )
    )
    workbook.defined_names.add(
        DefinedName("NamedAdjustedValue", attr_text="AdjustedCelsius(Inputs!$B$4)")
    )
    workbook.save(path)
    return path


def make_scoped_named_lambda_model(path: Path) -> Path:
    """Create global and worksheet-local LAMBDAs with the same callable name."""
    workbook = Workbook()
    inputs = workbook.active
    inputs.title = "Inputs"
    inputs["B2"] = 2
    inputs["B3"] = 3

    model = workbook.create_sheet("Model")
    model["A2"] = 10
    model["B2"] = "=Scale(A2)"

    report = workbook.create_sheet("Report")
    report["A2"] = 10
    report["B2"] = "=Scale(A2)"
    report["B3"] = "=Model!Scale(A2)"

    workbook.defined_names.add(
        DefinedName("Scale", attr_text="=LAMBDA(value,value*Inputs!$B$2)")
    )
    model.defined_names.add(
        DefinedName(
            "Scale",
            attr_text="=LAMBDA(value,value*Inputs!$B$3)",
            localSheetId=1,
        )
    )
    workbook.save(path)
    return path


def rewrite(path: Path, mutate: Callable[[Workbook], None]) -> Path:
    """Load, mutate, and save a fixture in place."""
    from openpyxl import load_workbook

    workbook = load_workbook(path)
    mutate(workbook)
    workbook.save(path)
    return path
