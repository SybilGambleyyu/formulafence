from __future__ import annotations

from collections.abc import Callable
from pathlib import Path

from openpyxl import Workbook
from openpyxl.workbook.defined_name import DefinedName


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


def rewrite(path: Path, mutate: Callable[[Workbook], None]) -> Path:
    """Load, mutate, and save a fixture in place."""
    from openpyxl import load_workbook

    workbook = load_workbook(path)
    mutate(workbook)
    workbook.save(path)
    return path
