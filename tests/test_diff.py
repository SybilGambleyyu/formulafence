from __future__ import annotations

import warnings

from openpyxl.workbook.defined_name import DefinedName

from formulafence.diff import compare_snapshots
from formulafence.output import profile_to_markdown, report_to_markdown
from formulafence.workbook import load_snapshot, profile_snapshot

from .helpers import (
    make_current_row_table_model,
    make_model,
    make_table_model,
    make_three_d_model,
    rewrite,
)


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
    assert change.details["impact_paths"] == [
        {
            "target": "Dashboard!B12",
            "path": ["Model!B2", "Model!C2", "Dashboard!B12"],
        },
        {
            "target": "Model!C2",
            "path": ["Model!B2", "Model!C2"],
        },
    ]
    assert any(finding.rule_id == "FF001" for finding in report.findings)
    assert "`Model!B2` → `Model!C2` → `Dashboard!B12`" in report_to_markdown(report)


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


def test_named_ranges_resolve_into_the_dependency_index(tmp_path) -> None:
    workbook_path = make_model(tmp_path / "named-ranges.xlsx")

    def add_named_range_formulas(workbook) -> None:
        model = workbook["Model"]
        model.defined_names.add(
            DefinedName(
                "LocalInput",
                attr_text="Inputs!$B$2",
                localSheetId=workbook.sheetnames.index("Model"),
            )
        )
        model["D2"] = "=HeadlineOutput*2"
        model["D3"] = "=LocalInput*3"
        workbook["Control"]["B2"] = "=Model!LocalInput*4"

    rewrite(workbook_path, add_named_range_formulas)
    snapshot = load_snapshot(workbook_path)

    assert snapshot.unresolved_reference_tokens == {}
    assert ("Model", "D2") in snapshot.direct_dependents(("Dashboard", "B12"))
    assert ("Model", "D3") in snapshot.direct_dependents(("Inputs", "B2"))
    assert ("Control", "B2") in snapshot.direct_dependents(("Inputs", "B2"))
    assert "Model!LocalInput" in snapshot.defined_names


def test_diff_surfaces_new_static_coverage_gaps(tmp_path) -> None:
    baseline = make_model(tmp_path / "baseline.xlsx")
    candidate = make_model(tmp_path / "candidate.xlsx")
    rewrite(
        candidate,
        lambda workbook: setattr(
            workbook["Model"]["D2"],
            "value",
            '=UnknownMetric+INDIRECT("Inputs!B2")',
        ),
    )

    candidate_snapshot = load_snapshot(candidate)
    profile = profile_snapshot(candidate_snapshot)
    assert profile["features"]["unresolved_reference_cells"] == [
        {"location": "Model!D2", "tokens": ["UnknownMetric"]}
    ]
    assert profile["features"]["dynamic_reference_cells"] == [
        {"location": "Model!D2", "functions": ["INDIRECT"]}
    ]

    report = compare_snapshots(load_snapshot(baseline), candidate_snapshot)
    rule_ids = {finding.rule_id for finding in report.findings}
    change_kinds = {change.kind for change in report.changes}
    assert {"FF011", "FF012"} <= rule_ids
    assert {
        "unresolved_formula_reference_added",
        "dynamic_formula_reference_added",
    } <= change_kinds


def test_external_defined_name_is_tracked_as_an_external_reference(tmp_path) -> None:
    workbook_path = make_model(tmp_path / "external-name.xlsx")

    def add_external_name_formula(workbook) -> None:
        workbook.defined_names.add(
            DefinedName("ExternalInput", attr_text="'[other.xlsx]Inputs'!$B$2")
        )
        workbook["Model"]["D2"] = "=ExternalInput*2"

    rewrite(workbook_path, add_external_name_formula)
    snapshot = load_snapshot(workbook_path)

    assert snapshot.external_references == {("Model", "D2")}
    assert snapshot.unresolved_reference_tokens == {}


def test_static_table_references_feed_dependency_paths_and_profiles(tmp_path) -> None:
    workbook_path = make_table_model(tmp_path / "table.xlsx")
    snapshot = load_snapshot(workbook_path)
    profile = profile_snapshot(snapshot)

    assert snapshot.summary()["table_count"] == 1
    assert snapshot.tables["Sales"].columns == ("Amount", "Rate", "Value")
    assert snapshot.unresolved_reference_tokens == {}
    assert ("Report", "B2") in snapshot.direct_dependents(("Data", "A2"))
    assert ("Report", "B3") in snapshot.direct_dependents(("Data", "B2"))
    assert ("Report", "B4") in snapshot.direct_dependents(("Data", "A1"))
    assert ("Report", "B5") in snapshot.direct_dependents(("Data", "B1"))
    assert profile["tables"] == [
        {
            "name": "Sales",
            "sheet": "Data",
            "ref": "A1:C4",
            "columns": ["Amount", "Rate", "Value"],
            "header_row_count": 1,
            "totals_row_count": 0,
        }
    ]
    assert "## Excel tables" in profile_to_markdown(profile)
    assert "| Sales | Data | A1:C4 | Amount, Rate, Value |" in profile_to_markdown(profile)


def test_table_definition_change_is_a_semantic_control_change(tmp_path) -> None:
    baseline = make_table_model(tmp_path / "baseline.xlsx")
    candidate = make_table_model(tmp_path / "candidate.xlsx")
    rewrite(candidate, lambda workbook: setattr(workbook["Data"].tables["Sales"], "ref", "A1:C3"))

    report = compare_snapshots(load_snapshot(baseline), load_snapshot(candidate))

    assert any(change.kind == "table_definition_changed" for change in report.changes)
    assert any(finding.rule_id == "FF013" for finding in report.findings)


def test_current_row_table_references_trace_only_the_matching_row(tmp_path) -> None:
    baseline = make_current_row_table_model(tmp_path / "baseline.xlsx")
    candidate = make_current_row_table_model(tmp_path / "candidate.xlsx")
    rewrite(candidate, lambda workbook: setattr(workbook["Data"]["A2"], "value", 100))

    baseline_snapshot = load_snapshot(baseline)
    assert baseline_snapshot.unresolved_reference_tokens == {}
    assert baseline_snapshot.direct_dependents(("Data", "A2")) == {
        ("Data", "C2"),
        ("Data", "E2"),
    }
    assert baseline_snapshot.direct_dependents(("Data", "A3")) == {("Data", "C3")}
    assert baseline_snapshot.direct_dependents(("Data", "A4")) == {("Data", "C4")}

    report = compare_snapshots(baseline_snapshot, load_snapshot(candidate))
    change = next(change for change in report.changes if change.location == ("Data", "A2"))

    assert change.impacted_cells == (("Data", "C2"), ("Data", "E2"), ("Report", "B2"))
    assert ("Data", "C3") not in change.impacted_cells
    assert ("Data", "C4") not in change.impacted_cells


def test_three_d_references_trace_every_sheet_in_the_tab_span(tmp_path) -> None:
    baseline = make_three_d_model(tmp_path / "baseline.xlsx")
    candidate = make_three_d_model(tmp_path / "candidate.xlsx")
    rewrite(candidate, lambda workbook: setattr(workbook["Feb"]["B2"], "value", 200))

    baseline_snapshot = load_snapshot(baseline)
    profile = profile_snapshot(baseline_snapshot)
    assert baseline_snapshot.unresolved_reference_tokens == {}
    assert baseline_snapshot.summary()["three_d_reference_cells"] == 1
    for sheet in ("Jan", "Feb", "Mar"):
        assert baseline_snapshot.direct_dependents((sheet, "B2")) == {("Summary", "B2")}
    assert baseline_snapshot.direct_dependents(("Jan:Mar", "B2")) == set()
    assert profile["features"]["three_d_reference_cells"] == [
        {"location": "Summary!B2", "tokens": ["Jan:Mar!B2"]}
    ]
    assert "## 3-D worksheet references" in profile_to_markdown(profile)

    report = compare_snapshots(baseline_snapshot, load_snapshot(candidate))
    change = next(change for change in report.changes if change.location == ("Feb", "B2"))

    assert change.impacted_cells == (("Summary", "B2"),)


def test_three_d_reference_scope_change_is_reported_when_tabs_move(tmp_path) -> None:
    baseline = make_three_d_model(tmp_path / "baseline.xlsx")
    candidate = make_three_d_model(tmp_path / "candidate.xlsx")

    def move_february_after_march(workbook) -> None:
        workbook._sheets = [  # noqa: SLF001 - sheet tab order is the scenario under test
            workbook["Jan"],
            workbook["Mar"],
            workbook["Feb"],
            workbook["Summary"],
        ]

    rewrite(candidate, move_february_after_march)
    report = compare_snapshots(load_snapshot(baseline), load_snapshot(candidate))
    scope_change = next(
        change
        for change in report.changes
        if change.kind == "three_d_reference_scope_changed"
    )

    assert scope_change.location == ("Summary", "B2")
    assert scope_change.details == {
        "references": [
            {
                "token": "Jan:Mar!B2",
                "before_sheets": ["Jan", "Feb", "Mar"],
                "after_sheets": ["Jan", "Mar"],
            }
        ]
    }
    assert any(finding.rule_id == "FF014" for finding in report.findings)


def test_three_d_references_include_a_sheet_inserted_between_tab_endpoints(tmp_path) -> None:
    baseline = make_three_d_model(tmp_path / "baseline.xlsx")
    candidate = make_three_d_model(tmp_path / "candidate.xlsx")

    def insert_period(workbook) -> None:
        inserted = workbook.create_sheet("Feb Extra", 1)
        inserted["A1"] = "Period input"
        inserted["B2"] = 25

    rewrite(candidate, insert_period)
    after = load_snapshot(candidate)
    assert after.direct_dependents(("Feb Extra", "B2")) == {("Summary", "B2")}

    report = compare_snapshots(load_snapshot(baseline), after)
    change = next(change for change in report.changes if change.location == ("Feb Extra", "B2"))

    assert change.kind == "value_added"
    assert change.impacted_cells == (("Summary", "B2"),)
    assert any(finding.rule_id == "FF014" for finding in report.findings)


def test_snapshot_captures_parser_coverage_warnings(tmp_path, monkeypatch) -> None:
    baseline = make_model(tmp_path / "baseline.xlsx")
    import formulafence.workbook as workbook_module

    original_load_workbook = workbook_module.load_workbook

    def noisy_load_workbook(*args, **kwargs):
        warnings.warn("fixture-only unsupported extension", UserWarning, stacklevel=2)
        return original_load_workbook(*args, **kwargs)

    monkeypatch.setattr(workbook_module, "load_workbook", noisy_load_workbook)
    snapshot = load_snapshot(baseline)

    assert snapshot.parser_warnings == ("fixture-only unsupported extension",)


def test_diff_surfaces_new_parser_coverage_warning(tmp_path, monkeypatch) -> None:
    baseline = make_model(tmp_path / "baseline.xlsx")
    candidate = make_model(tmp_path / "candidate.xlsx")
    import formulafence.workbook as workbook_module

    original_load_workbook = workbook_module.load_workbook

    def conditionally_noisy_load_workbook(path, *args, **kwargs):
        if str(path) == str(candidate):
            warnings.warn("candidate-only unsupported extension", UserWarning, stacklevel=2)
        return original_load_workbook(path, *args, **kwargs)

    monkeypatch.setattr(workbook_module, "load_workbook", conditionally_noisy_load_workbook)
    report = compare_snapshots(load_snapshot(baseline), load_snapshot(candidate))

    assert any(finding.rule_id == "FF010" for finding in report.findings)
