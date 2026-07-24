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


def rewrite(path: Path, mutate: Callable[[Workbook], None]) -> Path:
    """Load, mutate, and save a fixture in place."""
    from openpyxl import load_workbook

    workbook = load_workbook(path)
    mutate(workbook)
    workbook.save(path)
    return path
