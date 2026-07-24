from __future__ import annotations

from collections.abc import Callable
from pathlib import Path
from xml.etree import ElementTree
from zipfile import ZIP_DEFLATED, ZipFile

from openpyxl import Workbook
from openpyxl.formatting.rule import (
    CellIsRule,
    ColorScaleRule,
    DataBarRule,
    FormulaRule,
    IconSetRule,
)
from openpyxl.styles import Font, PatternFill
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.formula import ArrayFormula
from openpyxl.worksheet.table import Table


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
