from __future__ import annotations

from collections.abc import Callable
from pathlib import Path

from openpyxl import Workbook
from openpyxl.workbook.defined_name import DefinedName
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
