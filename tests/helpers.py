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
