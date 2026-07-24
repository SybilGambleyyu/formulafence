from __future__ import annotations

from formulafence.diff import compare_snapshots
from formulafence.workbook import load_snapshot

from .helpers import make_model, rewrite


def test_formula_to_value_traces_cross_sheet_downstream_impact(tmp_path) -> None:
    baseline = make_model(tmp_path / "baseline.xlsx")
    candidate = make_model(tmp_path / "candidate.xlsx")
    rewrite(candidate, lambda workbook: setattr(workbook["Model"]["B2"], "value", 200))

    report = compare_snapshots(load_snapshot(baseline), load_snapshot(candidate))
    change = next(change for change in report.changes if change.location == ("Model", "B2"))

    assert change.kind == "formula_to_value"
    assert change.impact_count == 2
    assert ("Model", "C2") in change.impacted_cells
    assert ("Dashboard", "B12") in change.impacted_cells
    assert any(finding.rule_id == "FF001" for finding in report.findings)


def test_diff_detects_pattern_break_and_static_hazards(tmp_path) -> None:
    baseline = make_model(tmp_path / "baseline.xlsx")
    candidate = make_model(tmp_path / "candidate.xlsx")

    def introduce_risk(workbook) -> None:
        workbook["Model"]["B3"] = "=Inputs!B3*3"
        workbook["Model"]["D2"] = "=[other.xlsx]Sheet1!A1"
        workbook["Model"]["D3"] = "=#REF!"
        workbook["Control"].sheet_state = "hidden"

    rewrite(candidate, introduce_risk)
    report = compare_snapshots(load_snapshot(baseline), load_snapshot(candidate))

    rule_ids = {finding.rule_id for finding in report.findings}
    assert {"FF003", "FF004", "FF006", "FF007"} <= rule_ids


def test_defined_name_change_is_semantic_control_change(tmp_path) -> None:
    baseline = make_model(tmp_path / "baseline.xlsx")
    candidate = make_model(tmp_path / "candidate.xlsx")

    def move_name(workbook) -> None:
        workbook.defined_names["HeadlineOutput"].attr_text = "Model!$C$2"

    rewrite(candidate, move_name)
    report = compare_snapshots(load_snapshot(baseline), load_snapshot(candidate))

    assert any(change.kind == "defined_name_changed" for change in report.changes)
    assert any(finding.rule_id == "FF008" for finding in report.findings)
