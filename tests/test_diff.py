from __future__ import annotations

import json
import warnings

from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation

from formulafence.diff import compare_snapshots
from formulafence.output import profile_to_markdown, report_to_markdown
from formulafence.workbook import load_snapshot, profile_snapshot

from .helpers import (
    add_conditional_formatting_databar_extension,
    make_conditional_formatting_model,
    make_current_row_table_model,
    make_data_validation_model,
    make_implicit_intersection_model,
    make_legacy_array_model,
    make_let_model,
    make_model,
    make_named_formula_model,
    make_named_lambda_model,
    make_scoped_named_lambda_model,
    make_spill_model,
    make_table_model,
    make_three_d_model,
    mark_array_formula_dynamic,
    mark_array_formula_unclassified,
    reorder_conditional_differential_styles,
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


def test_formula_defined_names_expand_nested_and_local_static_dependencies(tmp_path) -> None:
    baseline = make_named_formula_model(tmp_path / "baseline.xlsx")
    candidate = make_named_formula_model(tmp_path / "candidate.xlsx")
    rewrite(candidate, lambda workbook: setattr(workbook["Inputs"]["B2"], "value", 0.2))

    snapshot = load_snapshot(baseline)
    assert snapshot.unresolved_reference_tokens == {}
    assert snapshot.direct_dependents(("Inputs", "B2")) == {("Summary", "B2")}
    assert snapshot.direct_dependents(("Inputs", "B3")) == {("Summary", "B2")}
    assert snapshot.direct_dependents(("Inputs", "B4")) == {
        ("Report", "B2"),
        ("Summary", "B4"),
    }
    assert snapshot.direct_dependents(("Summary", "B3")) == set()
    assert "Summary!LocalMetric" in snapshot.defined_names

    report = compare_snapshots(snapshot, load_snapshot(candidate))
    change = next(change for change in report.changes if change.location == ("Inputs", "B2"))

    assert change.impacted_cells == (("Summary", "B2"),)


def test_formula_defined_name_change_remains_a_semantic_control_change(tmp_path) -> None:
    baseline = make_named_formula_model(tmp_path / "baseline.xlsx")
    candidate = make_named_formula_model(tmp_path / "candidate.xlsx")
    rewrite(
        candidate,
        lambda workbook: setattr(
            workbook.defined_names["TaxRate"], "attr_text", "=Inputs!$B$4"
        ),
    )

    report = compare_snapshots(load_snapshot(baseline), load_snapshot(candidate))

    assert any(change.kind == "defined_name_changed" for change in report.changes)
    assert any(finding.rule_id == "FF008" for finding in report.findings)


def test_formula_defined_names_remain_coverage_gaps_when_not_fully_static(tmp_path) -> None:
    workbook_path = make_named_formula_model(tmp_path / "named-formulas.xlsx")

    def add_unsafe_formula_names(workbook) -> None:
        workbook.defined_names.add(DefinedName("RelativeMetric", attr_text="=B2"))
        workbook.defined_names.add(
            DefinedName("DynamicMetric", attr_text="=OFFSET(Inputs!$B$2,0,0)")
        )
        workbook.defined_names.add(DefinedName("CircularMetricA", attr_text="=CircularMetricB"))
        workbook.defined_names.add(DefinedName("CircularMetricB", attr_text="=CircularMetricA"))
        workbook.defined_names.add(DefinedName("PeriodMetric", attr_text="=SUM(Inputs:Report!B2)"))
        workbook.defined_names.add(DefinedName("SpillMetric", attr_text="=SUM(Inputs!$B$2#)"))
        workbook.defined_names.add(
            DefinedName(
                "SerializedSpillMetric",
                attr_text="=SUM(_xlfn.ANCHORARRAY(Inputs!$B$2))",
            )
        )
        workbook.defined_names.add(
            DefinedName(
                "ImplicitMetric",
                attr_text="=_xlfn.SINGLE(Inputs!$B$2:$B$4)",
            )
        )
        workbook.defined_names.add(
            DefinedName("LiteralImplicitMetric", attr_text="=@Inputs!$B$2:$B$4")
        )
        workbook["Summary"]["B5"] = "=RelativeMetric"
        workbook["Summary"]["B6"] = "=DynamicMetric"
        workbook["Summary"]["B7"] = "=CircularMetricA"
        workbook["Summary"]["B8"] = "=PeriodMetric"
        workbook["Summary"]["B9"] = "=SpillMetric"
        workbook["Summary"]["B10"] = "=SerializedSpillMetric"
        workbook["Summary"]["B11"] = "=ImplicitMetric"
        workbook["Summary"]["B12"] = "=LiteralImplicitMetric"

    rewrite(workbook_path, add_unsafe_formula_names)
    snapshot = load_snapshot(workbook_path)

    assert snapshot.unresolved_reference_tokens == {
        ("Summary", "B5"): ("RelativeMetric",),
        ("Summary", "B6"): ("DynamicMetric",),
        ("Summary", "B7"): ("CircularMetricA",),
        ("Summary", "B8"): ("PeriodMetric",),
        ("Summary", "B9"): ("SpillMetric",),
        ("Summary", "B10"): ("SerializedSpillMetric",),
        ("Summary", "B11"): ("ImplicitMetric",),
        ("Summary", "B12"): ("LiteralImplicitMetric",),
    }
    assert ("Summary", "B5") not in snapshot.direct_dependents(("Inputs", "B2"))
    assert ("Summary", "B6") not in snapshot.direct_dependents(("Inputs", "B2"))
    assert ("Summary", "B7") not in snapshot.direct_dependents(("Inputs", "B2"))
    assert ("Summary", "B8") not in snapshot.direct_dependents(("Inputs", "B2"))
    assert ("Summary", "B9") not in snapshot.direct_dependents(("Inputs", "B2"))
    assert ("Summary", "B10") not in snapshot.direct_dependents(("Inputs", "B2"))
    assert ("Summary", "B11") not in snapshot.direct_dependents(("Inputs", "B2"))
    assert ("Summary", "B12") not in snapshot.direct_dependents(("Inputs", "B2"))


def test_formula_defined_names_expand_supported_static_table_references(tmp_path) -> None:
    workbook_path = make_table_model(tmp_path / "table.xlsx")

    def add_table_formula_name(workbook) -> None:
        workbook.defined_names.add(
            DefinedName("SalesAmount", attr_text="=SUM(Sales[Amount])")
        )
        workbook["Report"]["C2"] = "=SalesAmount"

    rewrite(workbook_path, add_table_formula_name)
    snapshot = load_snapshot(workbook_path)

    assert snapshot.unresolved_reference_tokens == {}
    assert snapshot.direct_dependents(("Data", "A2")) >= {
        ("Report", "B2"),
        ("Report", "B3"),
        ("Report", "C2"),
    }


def test_let_variables_do_not_hide_static_dependency_paths(tmp_path) -> None:
    baseline = make_let_model(tmp_path / "baseline.xlsx")
    candidate = make_let_model(tmp_path / "candidate.xlsx")
    rewrite(candidate, lambda workbook: setattr(workbook["Inputs"]["B2"], "value", 0.2))

    snapshot = load_snapshot(baseline)
    assert snapshot.unresolved_reference_tokens == {}
    assert snapshot.direct_dependents(("Inputs", "B2")) == {("Model", "B2")}
    assert snapshot.direct_dependents(("Inputs", "B3")) == {("Model", "B2")}

    report = compare_snapshots(snapshot, load_snapshot(candidate))
    change = next(change for change in report.changes if change.location == ("Inputs", "B2"))

    assert change.impacted_cells == (("Dashboard", "B2"), ("Model", "B2"))


def test_named_lambdas_expand_static_paths_through_nested_calls_and_named_formulas(
    tmp_path,
) -> None:
    baseline = make_named_lambda_model(tmp_path / "baseline.xlsx")
    candidate = make_named_lambda_model(tmp_path / "candidate.xlsx")
    rewrite(candidate, lambda workbook: setattr(workbook["Inputs"]["B2"], "value", 0.2))

    snapshot = load_snapshot(baseline)
    assert snapshot.unresolved_reference_tokens == {}
    assert snapshot.direct_dependents(("Inputs", "B2")) == {
        ("Model", "B2"),
        ("Model", "B3"),
        ("Model", "B4"),
    }
    assert snapshot.direct_dependents(("Inputs", "B3")) == {
        ("Model", "B3"),
        ("Model", "B4"),
    }
    assert snapshot.direct_dependents(("Inputs", "B4")) == {
        ("Model", "B2"),
        ("Model", "B3"),
        ("Model", "B4"),
    }

    report = compare_snapshots(snapshot, load_snapshot(candidate))
    change = next(change for change in report.changes if change.location == ("Inputs", "B2"))

    assert change.impacted_cells == (
        ("Dashboard", "B2"),
        ("Model", "B2"),
        ("Model", "B3"),
        ("Model", "B4"),
    )


def test_unsafe_or_recursive_named_lambdas_remain_visible_coverage_gaps(tmp_path) -> None:
    baseline = make_model(tmp_path / "baseline.xlsx")
    workbook_path = make_model(tmp_path / "candidate.xlsx")

    def add_unsafe_named_lambdas(workbook) -> None:
        workbook.defined_names.add(
            DefinedName("UnsafeLookup", attr_text="=LAMBDA(address,INDIRECT(address))")
        )
        workbook.defined_names.add(
            DefinedName(
                "RecursiveCount",
                attr_text="=LAMBDA(value,IF(value=0,0,RecursiveCount(value-1)))",
            )
        )
        workbook["Model"]["D2"] = "=UnsafeLookup(Inputs!A1)"
        workbook["Model"]["D3"] = "=RecursiveCount(3)"

    rewrite(workbook_path, add_unsafe_named_lambdas)
    snapshot = load_snapshot(workbook_path)

    assert snapshot.unresolved_reference_tokens == {
        ("Model", "D2"): ("UnsafeLookup",),
        ("Model", "D3"): ("RecursiveCount",),
    }
    assert snapshot.direct_dependents(("Inputs", "A1")) == {("Model", "D2")}
    report = compare_snapshots(load_snapshot(baseline), snapshot)
    assert {
        (finding.rule_id, finding.location)
        for finding in report.findings
    } >= {
        ("FF011", ("Model", "D2")),
        ("FF011", ("Model", "D3")),
    }


def test_named_lambda_calls_follow_worksheet_scope_and_qualified_local_names(tmp_path) -> None:
    snapshot = load_snapshot(make_scoped_named_lambda_model(tmp_path / "scoped.xlsx"))

    assert snapshot.unresolved_reference_tokens == {}
    assert snapshot.direct_dependents(("Inputs", "B2")) == {("Report", "B2")}
    assert snapshot.direct_dependents(("Inputs", "B3")) == {
        ("Model", "B2"),
        ("Report", "B3"),
    }
    assert snapshot.direct_dependents(("Model", "A2")) == {("Model", "B2")}
    assert snapshot.direct_dependents(("Report", "A2")) == {
        ("Report", "B2"),
        ("Report", "B3"),
    }


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


def test_spill_references_trace_anchors_but_remain_explicit_coverage_limits(tmp_path) -> None:
    baseline = make_spill_model(tmp_path / "baseline.xlsx")
    candidate = make_spill_model(tmp_path / "candidate.xlsx")
    rewrite(candidate, lambda workbook: setattr(workbook["Inputs"]["B2"], "value", "=SEQUENCE(4)"))

    snapshot = load_snapshot(baseline)
    profile = profile_snapshot(snapshot)

    assert snapshot.unresolved_reference_tokens == {}
    assert snapshot.tokenization_failure_cells == set()
    assert snapshot.direct_dependents(("Inputs", "B2")) == {("Model", "B2")}
    assert snapshot.direct_dependents(("Inputs", "B3")) == {("Model", "B3")}
    assert snapshot.summary()["spill_reference_cells"] == 2
    assert profile["features"]["spill_reference_cells"] == [
        {"location": "Model!B2", "tokens": ["Inputs!B2#"]},
        {"location": "Model!B3", "tokens": ["_xlfn.ANCHORARRAY"]},
    ]
    assert "## Dynamic-array spill references" in profile_to_markdown(profile)

    report = compare_snapshots(snapshot, load_snapshot(candidate))
    change = next(change for change in report.changes if change.location == ("Inputs", "B2"))

    assert change.impacted_cells == (("Dashboard", "B2"), ("Model", "B2"))


def test_implicit_intersection_traces_the_selected_static_input_and_profiles_it(tmp_path) -> None:
    workbook = make_implicit_intersection_model(tmp_path / "implicit-intersection.xlsx")

    snapshot = load_snapshot(workbook)
    profile = profile_snapshot(snapshot)

    assert snapshot.unresolved_reference_tokens == {}
    assert snapshot.direct_dependents(("Inputs", "B2")) == {("Model", "B2")}
    assert snapshot.direct_dependents(("Inputs", "B3")) == {("Model", "B3")}
    assert snapshot.direct_dependents(("Inputs", "B4")) == set()
    assert snapshot.direct_dependents(("Model", "B2")) == {("Dashboard", "B2")}
    assert snapshot.summary()["implicit_intersection_cells"] == 2
    assert profile["features"]["implicit_intersection_cells"] == [
        {"location": "Model!B2", "tokens": ["_xlfn.SINGLE"]},
        {"location": "Model!B3", "tokens": ["@Inputs!B2:B4"]},
    ]
    assert "## Explicit implicit intersection" in profile_to_markdown(profile)


def test_legacy_cse_array_outputs_connect_input_changes_to_result_consumers(tmp_path) -> None:
    baseline = make_legacy_array_model(tmp_path / "baseline.xlsx")
    candidate = make_legacy_array_model(tmp_path / "candidate.xlsx")
    rewrite(candidate, lambda workbook: setattr(workbook["Inputs"]["A2"], "value", "BBBB"))

    snapshot = load_snapshot(baseline)
    profile = profile_snapshot(snapshot)

    assert snapshot.direct_dependents(("Model", "B1")) == {
        ("Dashboard", "B2"),
        ("Model", "C2"),
    }
    assert snapshot.summary()["legacy_array_formula_cells"] == 1
    assert snapshot.summary()["legacy_array_formula_output_ranges"] == 1
    assert profile["features"]["legacy_array_formula_ranges"] == [
        {
            "anchor": "Model!B1",
            "ref": "Model!B1:B3",
            "output_cell_count": 3,
        }
    ]
    assert "## Legacy CSE array formulas" in profile_to_markdown(profile)

    report = compare_snapshots(snapshot, load_snapshot(candidate))
    change = next(change for change in report.changes if change.location == ("Inputs", "A2"))

    assert change.impacted_cells == (
        ("Dashboard", "B2"),
        ("Model", "B1"),
        ("Model", "C2"),
    )
    assert change.details["impact_paths"] == [
        {
            "target": "Dashboard!B2",
            "path": ["Inputs!A2", "Model!B1", "Dashboard!B2"],
        },
        {"target": "Model!B1", "path": ["Inputs!A2", "Model!B1"]},
        {
            "target": "Model!C2",
            "path": ["Inputs!A2", "Model!B1", "Model!C2"],
        },
    ]


def test_legacy_cse_output_aliases_do_not_expand_a_declared_huge_range(tmp_path) -> None:
    workbook = make_legacy_array_model(tmp_path / "large-cse.xlsx", "B1:XFD1048576")

    snapshot = load_snapshot(workbook)

    assert len(snapshot.cells) == 8
    assert snapshot.summary()["legacy_array_formula_output_cells"] > 1_000_000
    assert snapshot.direct_dependents(("Model", "B1")) == {
        ("Dashboard", "B2"),
        ("Model", "C2"),
    }


def test_dynamic_array_metadata_traces_observed_output_member_consumers(tmp_path) -> None:
    workbook = make_legacy_array_model(tmp_path / "dynamic.xlsx")
    mark_array_formula_dynamic(workbook)

    snapshot = load_snapshot(workbook)
    profile = profile_snapshot(snapshot)

    assert snapshot.legacy_array_formula_ranges == ()
    assert snapshot.dynamic_array_formula_cells == {("Model", "B1")}
    assert snapshot.unclassified_array_formula_cells == set()
    assert snapshot.direct_dependents(("Model", "B1")) == {
        ("Dashboard", "B2"),
        ("Model", "C2"),
    }
    assert snapshot.summary()["dynamic_array_observed_output_ranges"] == 1
    assert snapshot.summary()["dynamic_array_output_reference_cells"] == 2
    assert profile["features"]["dynamic_array_formula_cells"] == ["Model!B1"]
    assert profile["features"]["dynamic_array_observed_output_ranges"] == [
        {
            "anchor": "Model!B1",
            "ref": "Model!B1:B3",
            "output_cell_count": 3,
        }
    ]
    assert profile["features"]["dynamic_array_output_reference_cells"] == [
        {
            "location": "Dashboard!B2",
            "references": [
                {"anchor": "Model!B1", "observed_range": "Model!B1:B3"}
            ],
        },
        {
            "location": "Model!C2",
            "references": [
                {"anchor": "Model!B1", "observed_range": "Model!B1:B3"}
            ],
        },
    ]
    assert "## Dynamic-array formula anchors" in profile_to_markdown(profile)
    assert "observed from this workbook, not fixed" in profile_to_markdown(profile)


def test_dynamic_array_observed_output_aliases_connect_input_changes(tmp_path) -> None:
    baseline = make_legacy_array_model(tmp_path / "baseline.xlsx")
    candidate = make_legacy_array_model(tmp_path / "candidate.xlsx")
    rewrite(candidate, lambda workbook: setattr(workbook["Inputs"]["A2"], "value", "BBBB"))
    mark_array_formula_dynamic(baseline)
    mark_array_formula_dynamic(candidate)

    report = compare_snapshots(load_snapshot(baseline), load_snapshot(candidate))
    change = next(change for change in report.changes if change.location == ("Inputs", "A2"))

    assert change.impacted_cells == (
        ("Dashboard", "B2"),
        ("Model", "B1"),
        ("Model", "C2"),
    )
    assert change.details["impact_paths"] == [
        {
            "target": "Dashboard!B2",
            "path": ["Inputs!A2", "Model!B1", "Dashboard!B2"],
        },
        {"target": "Model!B1", "path": ["Inputs!A2", "Model!B1"]},
        {
            "target": "Model!C2",
            "path": ["Inputs!A2", "Model!B1", "Model!C2"],
        },
    ]


def test_dynamic_array_anchor_references_do_not_create_observed_member_aliases(tmp_path) -> None:
    workbook = make_legacy_array_model(tmp_path / "dynamic-anchor-only.xlsx")

    def use_anchor_only(workbook) -> None:
        workbook["Model"]["C2"] = "=B1*10"
        workbook["Dashboard"]["B2"] = "=Model!B1"

    rewrite(workbook, use_anchor_only)
    mark_array_formula_dynamic(workbook)

    snapshot = load_snapshot(workbook)

    assert snapshot.direct_dependents(("Model", "B1")) == {
        ("Dashboard", "B2"),
        ("Model", "C2"),
    }
    assert snapshot.dynamic_array_output_references == {}


def test_dynamic_array_observed_output_aliases_stay_compact_for_huge_ranges(tmp_path) -> None:
    workbook = make_legacy_array_model(
        tmp_path / "large-dynamic.xlsx", "B1:XFD1048576"
    )
    mark_array_formula_dynamic(workbook)

    snapshot = load_snapshot(workbook)

    assert len(snapshot.cells) == 8
    assert snapshot.summary()["dynamic_array_observed_output_ranges"] == 1
    assert snapshot.dynamic_array_formula_ranges[0].output_cell_count > 1_000_000
    assert snapshot.direct_dependents(("Model", "B1")) == {
        ("Dashboard", "B2"),
        ("Model", "C2"),
    }


def test_unclassified_array_metadata_is_a_visible_coverage_limit(tmp_path) -> None:
    baseline = make_legacy_array_model(tmp_path / "baseline.xlsx")
    workbook = make_legacy_array_model(tmp_path / "unclassified.xlsx")
    mark_array_formula_unclassified(workbook)

    snapshot = load_snapshot(workbook)
    profile = profile_snapshot(snapshot)

    assert snapshot.legacy_array_formula_ranges == ()
    assert snapshot.dynamic_array_formula_cells == set()
    assert snapshot.unclassified_array_formula_cells == {("Model", "B1")}
    assert snapshot.direct_dependents(("Model", "B1")) == set()
    assert snapshot.parser_warnings
    assert profile["features"]["unclassified_array_formula_cells"] == ["Model!B1"]
    assert "fixed-output aliases were not added" in profile_to_markdown(profile)
    report = compare_snapshots(load_snapshot(baseline), snapshot)
    assert any(finding.rule_id == "FF010" for finding in report.findings)


def test_array_formula_mode_and_legacy_output_range_changes_are_semantic(tmp_path) -> None:
    ordinary = make_legacy_array_model(tmp_path / "ordinary.xlsx")
    rewrite(
        ordinary,
        lambda workbook: setattr(
            workbook["Model"]["B1"], "value", "=LEN(Inputs!A1:A3)"
        ),
    )
    cse = make_legacy_array_model(tmp_path / "cse.xlsx")

    mode_report = compare_snapshots(load_snapshot(ordinary), load_snapshot(cse))
    mode_change = next(
        change for change in mode_report.changes if change.kind == "array_formula_mode_changed"
    )
    assert mode_change.location == ("Model", "B1")
    assert mode_change.details["before"] == {"mode": "ordinary", "output_range": None}
    assert mode_change.details["after"] == {
        "mode": "legacy_cse",
        "output_range": "B1:B3",
    }
    assert mode_change.impacted_cells == (("Dashboard", "B2"), ("Model", "C2"))
    assert any(finding.rule_id == "FF018" for finding in mode_report.findings)

    blank = make_legacy_array_model(tmp_path / "blank.xlsx")
    rewrite(blank, lambda workbook: setattr(workbook["Model"]["B1"], "value", None))
    dynamic = make_legacy_array_model(tmp_path / "dynamic.xlsx")
    mark_array_formula_dynamic(dynamic)
    new_dynamic_report = compare_snapshots(load_snapshot(blank), load_snapshot(dynamic))
    new_dynamic_change = next(
        change
        for change in new_dynamic_report.changes
        if change.kind == "array_formula_mode_changed"
    )
    assert new_dynamic_change.details["before"]["mode"] == "absent"
    assert new_dynamic_change.details["after"]["mode"] == "dynamic"
    assert any(finding.rule_id == "FF018" for finding in new_dynamic_report.findings)

    baseline = make_legacy_array_model(tmp_path / "baseline.xlsx", "B1:B3")
    candidate = make_legacy_array_model(tmp_path / "candidate.xlsx", "B1:B4")
    range_report = compare_snapshots(load_snapshot(baseline), load_snapshot(candidate))
    range_change = next(
        change
        for change in range_report.changes
        if change.kind == "legacy_array_output_range_changed"
    )
    assert range_change.location == ("Model", "B1")
    assert range_change.details["before_output_range"] == "B1:B3"
    assert range_change.details["after_output_range"] == "B1:B4"
    assert range_change.impacted_cells == (("Dashboard", "B2"), ("Model", "C2"))
    assert any(finding.rule_id == "FF018" for finding in range_report.findings)


def test_dynamic_array_cached_extent_change_is_not_a_fixed_range_change(tmp_path) -> None:
    baseline = make_legacy_array_model(tmp_path / "baseline.xlsx", "B1:B3")
    candidate = make_legacy_array_model(tmp_path / "candidate.xlsx", "B1:B4")
    mark_array_formula_dynamic(baseline)
    mark_array_formula_dynamic(candidate)

    report = compare_snapshots(load_snapshot(baseline), load_snapshot(candidate))

    assert not {
        change.kind
        for change in report.changes
        if change.kind
        in {"array_formula_mode_changed", "legacy_array_output_range_changed"}
    }
    assert not {finding.rule_id for finding in report.findings if finding.rule_id == "FF018"}
    assert not {finding.rule_id for finding in report.findings if finding.rule_id == "FF019"}


def test_new_dynamic_array_output_member_consumers_emit_ff019(tmp_path) -> None:
    baseline = make_legacy_array_model(tmp_path / "baseline.xlsx")

    def remove_member_consumers(workbook) -> None:
        workbook["Model"]["C2"] = None
        workbook["Dashboard"]["B2"] = None

    rewrite(baseline, remove_member_consumers)
    candidate = make_legacy_array_model(tmp_path / "candidate.xlsx")
    mark_array_formula_dynamic(baseline)
    mark_array_formula_dynamic(candidate)

    report = compare_snapshots(load_snapshot(baseline), load_snapshot(candidate))
    changes = {
        change.location: change
        for change in report.changes
        if change.kind == "dynamic_array_output_reference_added"
    }
    findings = {
        finding.location: finding for finding in report.findings if finding.rule_id == "FF019"
    }

    assert set(changes) == {("Dashboard", "B2"), ("Model", "C2")}
    assert set(findings) == set(changes)
    assert changes[("Model", "C2")].details["references"] == [
        {"anchor": "Model!B1", "observed_range": "Model!B1:B3"}
    ]
    assert "observed dynamic-array spill" in findings[("Model", "C2")].message


def test_dynamic_array_extent_growth_only_flags_new_member_relationships(tmp_path) -> None:
    baseline = make_legacy_array_model(tmp_path / "baseline.xlsx", "B1:B3")
    candidate = make_legacy_array_model(tmp_path / "candidate.xlsx", "B1:B4")

    def add_future_member_consumer(workbook) -> None:
        workbook["Model"]["C4"] = "=B4*10"

    rewrite(baseline, add_future_member_consumer)
    rewrite(candidate, add_future_member_consumer)
    mark_array_formula_dynamic(baseline)
    mark_array_formula_dynamic(candidate)

    report = compare_snapshots(load_snapshot(baseline), load_snapshot(candidate))
    ff019 = [finding for finding in report.findings if finding.rule_id == "FF019"]

    assert [(finding.location, finding.details["references"]) for finding in ff019] == [
        (
            ("Model", "C4"),
            [{"anchor": "Model!B1", "observed_range": "Model!B1:B4"}],
        )
    ]
    assert not {finding.rule_id for finding in report.findings if finding.rule_id == "FF018"}


def test_diff_surfaces_new_spill_and_tokenization_coverage_limits(tmp_path) -> None:
    baseline = make_model(tmp_path / "baseline.xlsx")
    candidate = make_model(tmp_path / "candidate.xlsx")

    def add_coverage_limits(workbook) -> None:
        workbook["Model"]["D2"] = "=SUM(Inputs!B2#)"
        workbook["Model"]["D3"] = "=SUM(Inputs!B2#1)"

    rewrite(candidate, add_coverage_limits)
    candidate_snapshot = load_snapshot(candidate)
    profile = profile_snapshot(candidate_snapshot)
    report = compare_snapshots(load_snapshot(baseline), candidate_snapshot)

    assert profile["features"]["spill_reference_cells"] == [
        {"location": "Model!D2", "tokens": ["Inputs!B2#"]}
    ]
    assert profile["features"]["tokenization_failure_cells"] == ["Model!D3"]
    assert "Formula tokenizer could not inspect `Model!D3`" in profile_to_markdown(profile)
    assert {finding.rule_id for finding in report.findings} >= {"FF015", "FF016"}
    assert {change.kind for change in report.changes} >= {
        "spill_reference_added",
        "formula_tokenization_failure_added",
    }


def test_diff_surfaces_new_implicit_intersection(tmp_path) -> None:
    baseline = make_model(tmp_path / "baseline.xlsx")
    candidate = make_model(tmp_path / "candidate.xlsx")
    rewrite(
        candidate,
        lambda workbook: setattr(
            workbook["Model"]["D2"], "value", "=@Inputs!B2:B4"
        ),
    )

    candidate_snapshot = load_snapshot(candidate)
    report = compare_snapshots(load_snapshot(baseline), candidate_snapshot)

    assert profile_snapshot(candidate_snapshot)["features"]["implicit_intersection_cells"] == [
        {"location": "Model!D2", "tokens": ["@Inputs!B2:B4"]}
    ]
    assert {finding.rule_id for finding in report.findings} >= {"FF017"}
    assert {change.kind for change in report.changes} >= {"implicit_intersection_added"}


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


def test_data_validation_controls_are_profiled_without_exposing_criteria(tmp_path) -> None:
    workbook = make_data_validation_model(tmp_path / "validation.xlsx")

    snapshot = load_snapshot(workbook)
    profile = profile_snapshot(snapshot)
    markdown = profile_to_markdown(profile)

    assert snapshot.summary()["data_validation_rules"] == 2
    assert snapshot.summary()["data_validation_target_ranges"] == 3
    assert profile["data_validations"] == [
        {
            "sheet": "Inputs",
            "ranges": ["Inputs!B2:B100", "Inputs!D2"],
            "type": "list",
            "operator": "between",
            "criteria_count": 1,
            "allow_blank": True,
            "dropdown_hidden": False,
            "prompts_disabled": False,
            "show_input_message": True,
            "show_error_message": True,
            "error_style": "stop",
            "has_error_alert_text": True,
            "has_input_prompt_text": True,
            "ime_mode": "noControl",
        },
        {
            "sheet": "Inputs",
            "ranges": ["Inputs!C2:C100"],
            "type": "decimal",
            "operator": "between",
            "criteria_count": 2,
            "allow_blank": False,
            "dropdown_hidden": False,
            "prompts_disabled": False,
            "show_input_message": False,
            "show_error_message": True,
            "error_style": "warning",
            "has_error_alert_text": True,
            "has_input_prompt_text": False,
            "ime_mode": "noControl",
        },
    ]
    assert "formula1" not in profile["data_validations"][0]
    assert "## Data-validation controls" in markdown
    assert "Limits!$A$2" not in markdown
    assert "Choose an approved status." not in markdown


def test_data_validation_writer_defaults_and_formula_spelling_are_canonical(tmp_path) -> None:
    baseline = make_data_validation_model(tmp_path / "baseline.xlsx")
    candidate = make_data_validation_model(tmp_path / "candidate.xlsx", reverse_status_targets=True)

    def use_omitted_ooxml_defaults(workbook) -> None:
        for validation in workbook["Inputs"].data_validations.dataValidation:
            if validation.formula1:
                validation.formula1 = validation.formula1.removeprefix("=")
            if validation.formula2:
                validation.formula2 = validation.formula2.removeprefix("=")
            validation.operator = None
            if validation.errorStyle == "stop":
                validation.errorStyle = None

    rewrite(candidate, use_omitted_ooxml_defaults)

    report = compare_snapshots(load_snapshot(baseline), load_snapshot(candidate))

    assert not {
        change.kind for change in report.changes if change.kind == "data_validation_changed"
    }
    assert not {finding.rule_id for finding in report.findings if finding.rule_id == "FF020"}


def test_data_validation_equivalent_target_grouping_is_canonical(tmp_path) -> None:
    baseline = make_data_validation_model(tmp_path / "baseline.xlsx")
    candidate = make_data_validation_model(tmp_path / "candidate.xlsx")

    def split_identical_status_control(workbook) -> None:
        inputs = workbook["Inputs"]
        _, amount = inputs.data_validations.dataValidation
        inputs.data_validations.dataValidation.clear()
        for target in ("B2:B100", "D2"):
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
            status.add(target)
            inputs.add_data_validation(status)
        inputs.add_data_validation(amount)

    rewrite(candidate, split_identical_status_control)

    report = compare_snapshots(load_snapshot(baseline), load_snapshot(candidate))

    assert not {
        change.kind for change in report.changes if change.kind == "data_validation_changed"
    }
    assert not {finding.rule_id for finding in report.findings if finding.rule_id == "FF020"}


def test_data_validation_change_is_a_high_risk_semantic_control_change(tmp_path) -> None:
    baseline = make_data_validation_model(tmp_path / "baseline.xlsx")
    candidate = make_data_validation_model(tmp_path / "candidate.xlsx")

    def weaken_amount_control(workbook) -> None:
        amount = workbook["Inputs"].data_validations.dataValidation[1]
        amount.formula2 = "=Limits!$B$2"
        amount.showErrorMessage = False

    rewrite(candidate, weaken_amount_control)

    report = compare_snapshots(load_snapshot(baseline), load_snapshot(candidate))
    change = next(change for change in report.changes if change.kind == "data_validation_changed")
    finding = next(finding for finding in report.findings if finding.rule_id == "FF020")

    assert change.severity == "high"
    assert change.details["sheet"] == "Inputs"
    assert change.details["before"][1]["formula2"] == "Limits!$B$3"
    assert change.details["after"][1]["formula2"] == "Limits!$B$2"
    assert change.details["after"][1]["show_error_message"] is False
    assert finding.location is None


def test_data_validation_global_prompt_disable_is_a_control_change(tmp_path) -> None:
    baseline = make_data_validation_model(tmp_path / "baseline.xlsx")
    candidate = make_data_validation_model(tmp_path / "candidate.xlsx")
    rewrite(
        candidate,
        lambda workbook: setattr(
            workbook["Inputs"].data_validations, "disablePrompts", True
        ),
    )

    report = compare_snapshots(load_snapshot(baseline), load_snapshot(candidate))
    change = next(change for change in report.changes if change.kind == "data_validation_changed")

    assert all(item["prompts_disabled"] is True for item in change.details["after"])
    assert "worksheet prompts disabled" in profile_to_markdown(
        profile_snapshot(load_snapshot(candidate))
    )


def test_data_validation_target_ranges_stay_compact_at_sheet_scale(tmp_path) -> None:
    workbook = make_data_validation_model(tmp_path / "large-validation.xlsx")

    def add_full_column_control(current_workbook) -> None:
        validation = DataValidation(
            type="whole",
            operator="greaterThan",
            formula1="0",
            showErrorMessage=True,
        )
        validation.add("E1:E1048576")
        current_workbook["Inputs"].add_data_validation(validation)

    rewrite(workbook, add_full_column_control)
    snapshot = load_snapshot(workbook)

    assert len(snapshot.cells) < 20
    assert snapshot.summary()["data_validation_rules"] == 3
    assert snapshot.summary()["data_validation_target_ranges"] == 4
    assert any(
        validation.ranges == ("E1:E1048576",)
        for validation in snapshot.data_validations
    )


def test_conditional_formatting_controls_are_profiled_without_exposing_criteria(tmp_path) -> None:
    workbook = make_conditional_formatting_model(tmp_path / "conditional-formatting.xlsx")

    snapshot = load_snapshot(workbook)
    profile = profile_snapshot(snapshot)
    markdown = profile_to_markdown(profile)

    assert snapshot.summary()["conditional_formatting_rules"] == 5
    assert snapshot.summary()["conditional_formatting_target_ranges"] == 5
    assert snapshot.summary()["conditional_formatting_extensions"] == 0
    assert profile["conditional_formatting"][0] == {
        "sheet": "Inputs",
        "ranges": ["Inputs!A2:A100"],
        "priority": 1,
        "type": "expression",
        "operator": None,
        "formula_count": 1,
        "has_text_criterion": False,
        "stop_if_true": True,
        "above_average": True,
        "percent": False,
        "bottom": False,
        "rank": None,
        "std_dev": None,
        "equal_average": False,
        "time_period": None,
        "formatting": ["differential style"],
        "extension_count": 0,
    }
    assert profile["conditional_formatting"][2]["formatting"] == ["color scale"]
    assert profile["conditional_formatting"][3]["formatting"] == ["data bar"]
    assert profile["conditional_formatting"][4]["formatting"] == ["icon set"]
    assert "formulas" not in profile["conditional_formatting"][0]
    assert "$A2<0" not in markdown
    assert "FFFFC7CE" not in markdown
    assert "## Conditional-formatting controls" in markdown


def test_conditional_formatting_defaults_formula_spelling_and_dxf_order_are_canonical(
    tmp_path,
) -> None:
    baseline = make_conditional_formatting_model(tmp_path / "baseline.xlsx")
    candidate = make_conditional_formatting_model(tmp_path / "candidate.xlsx")

    def use_equivalent_writer_spelling(workbook) -> None:
        rules = [
            rule
            for rule_group in workbook["Inputs"].conditional_formatting._cf_rules.values()
            for rule in rule_group
        ]
        for rule in rules:
            rule.priority *= 10
            rule.formula = [f"={formula}" for formula in rule.formula]
            if rule.stopIfTrue is None:
                rule.stopIfTrue = False
            if rule.aboveAverage is None:
                rule.aboveAverage = True
            if rule.percent is None:
                rule.percent = False
            if rule.bottom is None:
                rule.bottom = False
            if rule.equalAverage is None:
                rule.equalAverage = False

    rewrite(candidate, use_equivalent_writer_spelling)
    reorder_conditional_differential_styles(candidate)

    report = compare_snapshots(load_snapshot(baseline), load_snapshot(candidate))

    assert not {
        change.kind for change in report.changes if change.kind == "conditional_formatting_changed"
    }
    assert not {finding.rule_id for finding in report.findings if finding.rule_id == "FF021"}


def test_conditional_formatting_rule_change_and_precedence_change_are_high_risk(tmp_path) -> None:
    baseline = make_conditional_formatting_model(tmp_path / "baseline.xlsx")
    changed_rule = make_conditional_formatting_model(tmp_path / "changed-rule.xlsx")
    changed_precedence = make_conditional_formatting_model(
        tmp_path / "changed-precedence.xlsx"
    )

    def change_rule(workbook) -> None:
        rules = [
            rule
            for rule_group in workbook["Inputs"].conditional_formatting._cf_rules.values()
            for rule in rule_group
        ]
        rules[1].formula = ["90"]
        rules[1].stopIfTrue = True

    def swap_precedence(workbook) -> None:
        rules = [
            rule
            for rule_group in workbook["Inputs"].conditional_formatting._cf_rules.values()
            for rule in rule_group
        ]
        rules[0].priority, rules[1].priority = rules[1].priority, rules[0].priority

    rewrite(changed_rule, change_rule)
    rewrite(changed_precedence, swap_precedence)

    rule_report = compare_snapshots(load_snapshot(baseline), load_snapshot(changed_rule))
    precedence_report = compare_snapshots(
        load_snapshot(baseline), load_snapshot(changed_precedence)
    )
    rule_change = next(
        change for change in rule_report.changes if change.kind == "conditional_formatting_changed"
    )
    precedence_change = next(
        change
        for change in precedence_report.changes
        if change.kind == "conditional_formatting_changed"
    )

    assert rule_change.severity == "high"
    assert rule_change.details["before"]["rules"][1]["formulas"] == ["100"]
    assert rule_change.details["after"]["rules"][1]["formulas"] == ["90"]
    assert rule_change.details["after"]["rules"][1]["stop_if_true"] is True
    assert precedence_change.severity == "high"
    assert [item["type"] for item in precedence_change.details["after"]["rules"][:2]] == [
        "cellIs",
        "expression",
    ]
    assert {finding.rule_id for finding in rule_report.findings} >= {"FF021"}
    assert {finding.rule_id for finding in precedence_report.findings} >= {"FF021"}


def test_conditional_formatting_extensions_are_compared_without_guid_noise(tmp_path) -> None:
    baseline = make_conditional_formatting_model(tmp_path / "baseline.xlsx")
    equivalent = make_conditional_formatting_model(tmp_path / "equivalent.xlsx")
    changed = make_conditional_formatting_model(tmp_path / "changed.xlsx")
    changed_extension_type = make_conditional_formatting_model(
        tmp_path / "changed-extension-type.xlsx"
    )
    add_conditional_formatting_databar_extension(
        baseline,
        guid="{11111111-2222-3333-4444-555555555555}",
        axis_color="FF000000",
    )
    add_conditional_formatting_databar_extension(
        equivalent,
        guid="{AAAAAAAA-BBBB-CCCC-DDDD-EEEEEEEEEEEE}",
        axis_color="FF000000",
    )
    add_conditional_formatting_databar_extension(
        changed,
        guid="{99999999-8888-7777-6666-555555555555}",
        axis_color="FFFF0000",
    )
    add_conditional_formatting_databar_extension(
        changed_extension_type,
        guid="{99999999-8888-7777-6666-555555555555}",
        axis_color="FF000000",
        worksheet_extension_uri="{11111111-2222-3333-4444-555555555555}",
    )

    baseline_snapshot = load_snapshot(baseline)
    profile = profile_snapshot(baseline_snapshot)
    equivalent_report = compare_snapshots(baseline_snapshot, load_snapshot(equivalent))
    changed_report = compare_snapshots(baseline_snapshot, load_snapshot(changed))
    extension_type_report = compare_snapshots(
        baseline_snapshot, load_snapshot(changed_extension_type)
    )

    assert baseline_snapshot.summary()["conditional_formatting_extensions"] == 1
    assert profile["conditional_formatting_extensions"] == [
        {"sheet": "Inputs", "element": "ext"}
    ]
    assert "FF000000" not in json.dumps(profile)
    markdown = profile_to_markdown(profile)
    assert "## Conditional-formatting extension coverage" in markdown
    assert "FF000000" not in markdown
    assert not {
        change.kind
        for change in equivalent_report.changes
        if change.kind == "conditional_formatting_changed"
    }
    changed_change = next(
        change
        for change in changed_report.changes
        if change.kind == "conditional_formatting_changed"
    )
    assert changed_change.details["before"]["extensions"] != changed_change.details["after"][
        "extensions"
    ]
    assert {finding.rule_id for finding in changed_report.findings} >= {"FF021"}
    extension_type_change = next(
        change
        for change in extension_type_report.changes
        if change.kind == "conditional_formatting_changed"
    )
    assert (
        extension_type_change.details["before"]["extensions"][0]["extension"]["attributes"][
            "uri"
        ]
        != extension_type_change.details["after"]["extensions"][0]["extension"]["attributes"][
            "uri"
        ]
    )


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
