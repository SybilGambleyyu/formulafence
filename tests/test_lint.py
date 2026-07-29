"""Tests for conservative single-workbook formula linting."""

from __future__ import annotations

from pathlib import Path

import pytest
from openpyxl import Workbook
from openpyxl.styles import Protection

from formulafence.lint import lint_snapshot
from formulafence.models import (
    ArrayFormulaRange,
    CellProtectionAssignmentSnapshot,
    FormulaFenceError,
)
from formulafence.output import as_json, lint_to_markdown, lint_to_sarif
from formulafence.workbook import load_snapshot

from .helpers import make_formula_cached_result_model


def _snapshot(
    tmp_path: Path,
    cells: dict[str, object],
    *,
    protected: bool = False,
    unlocked_cells: tuple[str, ...] = (),
    calculation_mode: str | None = None,
    calculation_completed: bool | None = None,
    calculation_iteration_enabled: bool | None = None,
):
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = "Model"
    for coordinate, value in cells.items():
        worksheet[coordinate] = value
    for coordinate in unlocked_cells:
        worksheet[coordinate].protection = Protection(locked=False)
    worksheet.protection.sheet = protected
    if calculation_mode is not None:
        workbook.calculation.calcMode = calculation_mode
    if calculation_completed is not None:
        workbook.calculation.calcCompleted = calculation_completed
    if calculation_iteration_enabled is not None:
        workbook.calculation.iterate = calculation_iteration_enabled
    path = tmp_path / "model.xlsx"
    workbook.save(path)
    return load_snapshot(path)


def test_lint_reports_blank_gap_inside_supported_column_copy_block(tmp_path: Path) -> None:
    snapshot = _snapshot(
        tmp_path,
        {
            "A2": 1,
            "A3": 2,
            "A4": 3,
            "A5": 4,
            "B2": "=A2*2",
            "B4": "=A4*2",
            "B5": "=A5*2",
        },
    )

    report = lint_snapshot(snapshot)

    assert len(report.findings) == 1
    finding = report.findings[0]
    assert (finding.rule_id, finding.severity, finding.location) == (
        "FF082",
        "high",
        ("Model", "B3"),
    )
    assert finding.details == {
        "pattern_kind": "blank_gap",
        "pattern_evidence": [
            {
                "orientation": "column",
                "preceding_formula": "Model!B2",
                "following_formula": "Model!B4",
                "supporting_formula": "Model!B5",
            }
        ],
    }


def test_lint_reports_non_formula_gap_inside_supported_row_copy_block(tmp_path: Path) -> None:
    snapshot = _snapshot(
        tmp_path,
        {
            "A2": 1,
            "B2": "=A2*2",
            "C2": 99,
            "D2": "=C2*2",
            "E2": "=D2*2",
        },
    )

    report = lint_snapshot(snapshot)

    assert [
        (finding.rule_id, finding.severity, finding.location)
        for finding in report.findings
    ] == [
        ("FF082", "medium", ("Model", "C2"))
    ]
    assert report.findings[0].details["pattern_kind"] == "non_formula_gap"


def test_lint_treats_textual_formula_exceptions_as_low_severity(tmp_path: Path) -> None:
    snapshot = _snapshot(
        tmp_path,
        {
            "A2": 1,
            "B2": "=A2*2",
            "C2": "n.a.",
            "D2": "=C2*2",
            "E2": "=D2*2",
        },
    )

    report = lint_snapshot(snapshot)

    assert [
        (finding.rule_id, finding.severity, finding.location)
        for finding in report.findings
    ] == [
        ("FF082", "low", ("Model", "C2"))
    ]
    assert report.findings[0].details["pattern_kind"] == "text_gap"


def test_lint_treats_stored_error_interruption_as_high_severity(tmp_path: Path) -> None:
    snapshot = _snapshot(
        tmp_path,
        {
            "A2": 1,
            "B2": "=A2*2",
            "C2": "#N/A",
            "D2": "=C2*2",
            "E2": "=D2*2",
        },
    )

    report = lint_snapshot(snapshot)

    assert [
        (finding.rule_id, finding.severity, finding.location)
        for finding in report.findings
    ] == [
        ("FF082", "high", ("Model", "C2"))
    ]
    assert report.findings[0].details["pattern_kind"] == "error_gap"


def test_lint_reports_formula_outlier_inside_supported_row_copy_block(tmp_path: Path) -> None:
    snapshot = _snapshot(
        tmp_path,
        {
            "A2": 1,
            "B2": "=A2*2",
            "C2": "=B2*3",
            "D2": "=C2*2",
            "E2": "=D2*2",
        },
    )

    report = lint_snapshot(snapshot)

    assert [
        (finding.rule_id, finding.severity, finding.location)
        for finding in report.findings
    ] == [
        ("FF083", "medium", ("Model", "C2"))
    ]
    assert report.findings[0].details["pattern_kind"] == "formula_outlier"


def test_lint_requires_a_third_matching_peer_before_reporting(tmp_path: Path) -> None:
    snapshot = _snapshot(
        tmp_path,
        {
            "A2": 1,
            "A4": 3,
            "B2": "=A2*2",
            "D2": "=C2*2",
        },
    )

    assert lint_snapshot(snapshot).findings == []


def test_lint_accepts_supporting_peer_before_the_interruption(tmp_path: Path) -> None:
    snapshot = _snapshot(
        tmp_path,
        {
            "A2": 1,
            "B2": "=A2*2",
            "C2": "=B2*2",
            "E2": "=D2*2",
        },
    )

    report = lint_snapshot(snapshot)

    assert [(finding.rule_id, finding.location) for finding in report.findings] == [
        ("FF082", ("Model", "D2"))
    ]
    assert report.findings[0].details["pattern_evidence"][0]["supporting_formula"] == "Model!B2"


def test_lint_skips_complete_copy_block(tmp_path: Path) -> None:
    snapshot = _snapshot(
        tmp_path,
        {
            "A2": 1,
            "B2": "=A2*2",
            "C2": "=B2*2",
            "D2": "=C2*2",
            "E2": "=D2*2",
        },
    )

    assert lint_snapshot(snapshot).findings == []


def test_lint_reports_omitted_numeric_column_run_after_simple_aggregate(
    tmp_path: Path,
) -> None:
    snapshot = _snapshot(
        tmp_path,
        {
            "B2": 10,
            "B3": 20,
            "B4": 30,
            "B5": 40,
            "B6": 50,
            "B7": 60,
            "B8": "=SUM($B$2:$B$4)",
        },
    )

    report = lint_snapshot(snapshot)

    assert [
        (finding.rule_id, finding.severity, finding.location)
        for finding in report.findings
    ] == [("FF084", "medium", ("Model", "B8"))]
    assert report.findings[0].details == {
        "aggregate_function": "SUM",
        "orientation": "column",
        "referenced_range": "Model!B2:B4",
        "omitted_range": "Model!B5:B7",
        "omitted_cell_count": 3,
    }


def test_lint_handles_a_local_aggregate_in_excel_last_column(tmp_path: Path) -> None:
    snapshot = _snapshot(
        tmp_path,
        {
            "XFD2": 10,
            "XFD3": 20,
            "XFD4": 30,
            "XFD5": 40,
            "XFD6": 50,
            "XFD7": 60,
            "XFD8": "=SUM(XFD2:XFD4)",
        },
    )

    report = lint_snapshot(snapshot)

    assert [(finding.rule_id, finding.location) for finding in report.findings] == [
        ("FF084", ("Model", "XFD8"))
    ]


def test_lint_reports_omitted_numeric_row_run_after_simple_aggregate(tmp_path: Path) -> None:
    snapshot = _snapshot(
        tmp_path,
        {
            "B2": 10,
            "C2": 20,
            "D2": 30,
            "E2": 40,
            "F2": 50,
            "G2": "=AVERAGE(Model!B2:D2)",
        },
    )

    report = lint_snapshot(snapshot)

    assert [
        (finding.rule_id, finding.location) for finding in report.findings
    ] == [("FF084", ("Model", "G2"))]
    assert report.findings[0].details == {
        "aggregate_function": "AVERAGE",
        "orientation": "row",
        "referenced_range": "Model!B2:D2",
        "omitted_range": "Model!E2:F2",
        "omitted_cell_count": 2,
    }


def test_lint_reports_a_directly_unlocked_formula_on_a_protected_sheet(
    tmp_path: Path,
) -> None:
    snapshot = _snapshot(
        tmp_path,
        {"A2": 10, "B2": "=A2*2"},
        protected=True,
        unlocked_cells=("B2",),
    )

    report = lint_snapshot(snapshot)

    assert [
        (finding.rule_id, finding.severity, finding.location)
        for finding in report.findings
    ] == [("FF085", "medium", ("Model", "B2"))]
    assert report.findings[0].details == {"protection_scope": "direct_cell"}


def test_lint_keeps_non_direct_or_inactive_unlocked_formula_controls_quiet(
    tmp_path: Path,
) -> None:
    inactive = _snapshot(
        tmp_path,
        {"A2": 10, "B2": "=A2*2"},
        unlocked_cells=("B2",),
    )
    assert lint_snapshot(inactive).findings == []

    protected = _snapshot(tmp_path, {"A2": 10, "B2": "=A2*2"}, protected=True)
    protected.cell_protection_assignments = (
        CellProtectionAssignmentSnapshot(
            sheet="Model",
            scope="column",
            target="B:B",
            locked=False,
            hidden=False,
        ),
    )
    assert lint_snapshot(protected).findings == []


def test_lint_reports_explicit_incomplete_manual_formula_calculation(tmp_path: Path) -> None:
    snapshot = _snapshot(
        tmp_path,
        {"A2": 10, "B2": "=A2*2"},
        calculation_mode="manual",
        calculation_completed=False,
    )

    report = lint_snapshot(snapshot)

    assert [
        (finding.rule_id, finding.severity, finding.location)
        for finding in report.findings
    ] == [("FF086", "medium", None)]
    assert report.findings[0].details == {
        "calculation_mode": "manual",
        "calculation_completed_before_save": False,
    }


def test_lint_keeps_completed_or_non_manual_calculation_states_quiet(tmp_path: Path) -> None:
    completed_manual = _snapshot(
        tmp_path,
        {"A2": 10, "B2": "=A2*2"},
        calculation_mode="manual",
        calculation_completed=True,
    )
    incomplete_automatic = _snapshot(
        tmp_path,
        {"A2": 10, "B2": "=A2*2"},
        calculation_mode="auto",
        calculation_completed=False,
    )
    formula_free_manual = _snapshot(
        tmp_path,
        {"A2": 10},
        calculation_mode="manual",
        calculation_completed=False,
    )

    assert lint_snapshot(completed_manual).findings == []
    assert lint_snapshot(incomplete_automatic).findings == []
    assert lint_snapshot(formula_free_manual).findings == []


def test_lint_reports_direct_static_self_reference_when_iteration_is_disabled(
    tmp_path: Path,
) -> None:
    snapshot = _snapshot(tmp_path, {"B2": "=B2+1"})

    report = lint_snapshot(snapshot)

    assert [
        (finding.rule_id, finding.severity, finding.location)
        for finding in report.findings
    ] == [("FF087", "high", ("Model", "B2"))]
    assert report.findings[0].details == {
        "calculation_iteration_enabled": False,
        "reference_scope": "direct_static",
    }


def test_lint_reports_static_multi_cell_cycle_when_iteration_is_disabled(
    tmp_path: Path,
) -> None:
    snapshot = _snapshot(
        tmp_path,
        {"B2": "=C2+1", "C2": "=B2+1", "D2": "=B2+1"},
    )

    report = lint_snapshot(snapshot)

    assert [
        (finding.rule_id, finding.severity, finding.location)
        for finding in report.findings
    ] == [
        ("FF090", "high", ("Model", "B2")),
        ("FF090", "high", ("Model", "C2")),
    ]
    assert [finding.details for finding in report.findings] == [
        {
            "calculation_iteration_enabled": False,
            "reference_scope": "multi_cell_static",
            "cycle_member_count": 2,
        },
        {
            "calculation_iteration_enabled": False,
            "reference_scope": "multi_cell_static",
            "cycle_member_count": 2,
        },
    ]


def test_lint_reports_cross_sheet_static_multi_cell_cycle(tmp_path: Path) -> None:
    workbook = Workbook()
    model = workbook.active
    model.title = "Model"
    inputs = workbook.create_sheet("Inputs")
    model["B2"] = "=Inputs!B2+1"
    inputs["B2"] = "=Model!B2+1"
    path = tmp_path / "cross-sheet-cycle.xlsx"
    workbook.save(path)

    report = lint_snapshot(load_snapshot(path))

    assert [
        (finding.rule_id, finding.location) for finding in report.findings
    ] == [
        ("FF090", ("Inputs", "B2")),
        ("FF090", ("Model", "B2")),
    ]
    assert {finding.details["cycle_member_count"] for finding in report.findings} == {2}


def test_lint_handles_a_deep_static_multi_cell_cycle_iteratively(tmp_path: Path) -> None:
    cycle_size = 1_025
    cells = {
        f"A{row}": f"=A{row + 1}+1" if row < cycle_size else "=A1+1"
        for row in range(1, cycle_size + 1)
    }

    report = lint_snapshot(_snapshot(tmp_path, cells))

    assert len(report.findings) == cycle_size
    assert {finding.rule_id for finding in report.findings} == {"FF090"}
    assert {finding.details["cycle_member_count"] for finding in report.findings} == {
        cycle_size
    }


def test_lint_keeps_iterative_or_ambiguous_multi_cell_cycles_quiet(
    tmp_path: Path,
) -> None:
    iterative = _snapshot(
        tmp_path,
        {"B2": "=C2+1", "C2": "=B2+1"},
        calculation_iteration_enabled=True,
    )
    dynamic = _snapshot(
        tmp_path,
        {"A1": 1, "B2": "=C2+OFFSET(A1,0,0)", "C2": "=B2+1"},
    )
    range_expression = _snapshot(
        tmp_path,
        {"B2": "=SUM(C2:C3)", "C2": "=B2+1", "C3": 1},
    )
    array_territory = _snapshot(tmp_path, {"B2": "=C2+1", "C2": "=B2+1"})
    array_territory.dynamic_array_formula_ranges = (
        ArrayFormulaRange(
            sheet="Model",
            anchor="B2",
            ref="B2:C2",
            min_column=2,
            min_row=2,
            max_column=3,
            max_row=2,
        ),
    )
    three_d = _snapshot(tmp_path, {"B2": "=C2+1", "C2": "=B2+1"})
    three_d.three_d_reference_tokens[("Model", "B2")] = ("Model:Other!C2",)
    spill = _snapshot(tmp_path, {"B2": "=C2+1", "C2": "=B2+1"})
    spill.spill_reference_tokens[("Model", "B2")] = ("C2#",)
    implicit_intersection = _snapshot(tmp_path, {"B2": "=C2+1", "C2": "=B2+1"})
    implicit_intersection.implicit_intersection_tokens[("Model", "B2")] = ("@C2",)
    tokenization_failure = _snapshot(tmp_path, {"B2": "=C2+1", "C2": "=B2+1"})
    tokenization_failure.tokenization_failure_cells.add(("Model", "B2"))

    assert lint_snapshot(iterative).findings == []
    assert lint_snapshot(dynamic).findings == []
    assert lint_snapshot(range_expression).findings == []
    assert lint_snapshot(array_territory).findings == []
    assert lint_snapshot(three_d).findings == []
    assert lint_snapshot(spill).findings == []
    assert lint_snapshot(implicit_intersection).findings == []
    assert lint_snapshot(tokenization_failure).findings == []


def test_lint_deduplicates_direct_self_reference_inside_multi_cell_cycle(
    tmp_path: Path,
) -> None:
    snapshot = _snapshot(tmp_path, {"B2": "=B2+C2", "C2": "=B2+1"})

    report = lint_snapshot(snapshot)
    assert [
        (finding.rule_id, finding.location) for finding in report.findings
    ] == [
        ("FF087", ("Model", "B2")),
        ("FF090", ("Model", "C2")),
    ]
    assert report.findings[1].details == {
        "calculation_iteration_enabled": False,
        "reference_scope": "multi_cell_static",
        "cycle_member_count": 2,
    }


def test_lint_keeps_iteration_and_dynamic_or_range_self_forms_quiet(
    tmp_path: Path,
) -> None:
    iterative = _snapshot(
        tmp_path,
        {"B2": "=B2+1"},
        calculation_iteration_enabled=True,
    )
    dynamic = _snapshot(tmp_path, {"A1": "=OFFSET(A1,1,0)", "A2": 10})
    spill = _snapshot(tmp_path, {"A1": "=A1#"})
    implicit_intersection = _snapshot(tmp_path, {"A1": "=@A1"})
    range_expression = _snapshot(tmp_path, {"A1": "=SUM(A1:A2)", "A2": 10})
    array_territory = _snapshot(tmp_path, {"A1": "=A1+1"})
    array_territory.dynamic_array_formula_ranges = (
        ArrayFormulaRange(
            sheet="Model",
            anchor="A1",
            ref="A1",
            min_column=1,
            min_row=1,
            max_column=1,
            max_row=1,
        ),
    )

    assert lint_snapshot(iterative).findings == []
    assert lint_snapshot(dynamic).findings == []
    assert lint_snapshot(spill).findings == []
    assert lint_snapshot(implicit_intersection).findings == []
    assert lint_snapshot(range_expression).findings == []
    assert lint_snapshot(array_territory).findings == []


def test_lint_reports_an_explicit_broken_reference_operand(tmp_path: Path) -> None:
    snapshot = _snapshot(
        tmp_path,
        {"A2": 10, "B2": "=SUM(A2,#REF!,C2)", "C2": 20},
    )

    report = lint_snapshot(snapshot)

    assert [
        (finding.rule_id, finding.severity, finding.location)
        for finding in report.findings
    ] == [("FF088", "critical", ("Model", "B2"))]
    assert report.findings[0].details == {}


def test_lint_keeps_ref_text_but_not_an_error_operand_quiet(tmp_path: Path) -> None:
    literal = _snapshot(tmp_path, {"B2": '="#REF!"'})

    assert lint_snapshot(literal).findings == []


def test_lint_reports_a_saved_broken_reference_formula_result(tmp_path: Path) -> None:
    snapshot = load_snapshot(
        make_formula_cached_result_model(
            tmp_path / "saved-reference-error.xlsx",
            error_formula="=INDEX(Inputs!A1:A1,2)",
            error_result="#REF!",
        )
    )

    report = lint_snapshot(snapshot)

    assert snapshot.broken_references == set()
    assert [
        (finding.rule_id, finding.severity, finding.location)
        for finding in report.findings
    ] == [("FF089", "high", ("Report", "B5"))]
    assert report.findings[0].details == {}


def test_lint_keeps_other_saved_errors_quiet_and_avoids_static_duplicates(
    tmp_path: Path,
) -> None:
    other_error = load_snapshot(make_formula_cached_result_model(tmp_path / "other.xlsx"))
    static_reference = load_snapshot(
        make_formula_cached_result_model(
            tmp_path / "static-reference.xlsx",
            error_formula="=#REF!",
            error_result="#REF!",
        )
    )

    assert lint_snapshot(other_error).findings == []
    assert [
        (finding.rule_id, finding.severity, finding.location)
        for finding in lint_snapshot(static_reference).findings
    ] == [("FF088", "critical", ("Report", "B5"))]


def test_lint_keeps_short_or_non_numeric_aggregate_gaps_quiet(tmp_path: Path) -> None:
    snapshot = _snapshot(
        tmp_path,
        {
            "B2": 10,
            "B3": 20,
            "B4": 30,
            "B5": "=SUM(B2:B3)",
            "D2": 10,
            "D3": 20,
            "D4": "intentional marker",
            "D5": 30,
            "D6": "=SUM(D2:D3)",
            "F2": 10,
            "F3": 20,
            "F4": 30,
            "F5": 40,
            "F6": "=SUM(F2:F3)+0",
        },
    )

    assert lint_snapshot(snapshot).findings == []


def test_lint_skips_array_territory_and_bounds_aggregate_gap_inspection(
    tmp_path: Path,
) -> None:
    cells = {f"B{row}": row for row in range(2, 132)}
    cells["B132"] = "=MAX(B2:B4)"
    snapshot = _snapshot(tmp_path, cells)
    snapshot.dynamic_array_formula_ranges = (
        ArrayFormulaRange(
            sheet="Model",
            anchor="B5",
            ref="B5:B6",
            min_column=2,
            min_row=5,
            max_column=2,
            max_row=6,
        ),
    )

    assert lint_snapshot(snapshot).findings == []
    snapshot.dynamic_array_formula_ranges = ()
    assert lint_snapshot(snapshot, max_aggregate_omission_gap_cells=64).findings == []
    with pytest.raises(FormulaFenceError, match="max_aggregate_omission_gap_cells"):
        lint_snapshot(snapshot, max_aggregate_omission_gap_cells=1)


def test_lint_deduplicates_crossing_copy_patterns_with_all_evidence(tmp_path: Path) -> None:
    snapshot = _snapshot(
        tmp_path,
        {
            "B2": 1,
            "C2": "=B2",
            "D2": 3,
            "E2": 4,
            "B3": "=B2",
            "D3": "=D2",
            "E3": "=E2",
            "B4": 5,
            "C4": "=B4",
            "D4": 7,
            "E4": 8,
            "B5": 9,
            "C5": "=B5",
        },
    )

    report = lint_snapshot(snapshot)

    assert [(finding.rule_id, finding.location) for finding in report.findings] == [
        ("FF082", ("Model", "C3"))
    ]
    assert [
        evidence["orientation"]
        for evidence in report.findings[0].details["pattern_evidence"]
    ] == ["column", "row"]


def test_lint_never_uses_a_tokenization_failure_as_copy_evidence(tmp_path: Path) -> None:
    snapshot = _snapshot(
        tmp_path,
        {
            "A2": 1,
            "B2": "=A2*2",
            "D2": "=C2*2",
            "E2": "=D2*2",
        },
    )
    snapshot.tokenization_failure_cells.add(("Model", "D2"))

    assert lint_snapshot(snapshot).findings == []


def test_lint_never_treats_declared_array_output_as_an_ordinary_gap(tmp_path: Path) -> None:
    snapshot = _snapshot(
        tmp_path,
        {
            "A2": 1,
            "B2": "=A2*2",
            "D2": "=C2*2",
            "E2": "=D2*2",
        },
    )
    snapshot.dynamic_array_formula_ranges = (
        ArrayFormulaRange(
            sheet="Model",
            anchor="C2",
            ref="C2",
            min_column=3,
            min_row=2,
            max_column=3,
            max_row=2,
        ),
    )

    assert lint_snapshot(snapshot).findings == []


def test_lint_fails_closed_when_array_metadata_is_incomplete(tmp_path: Path) -> None:
    snapshot = _snapshot(tmp_path, {"A1": 1, "B1": "=A1*2"})
    snapshot.array_formula_metadata_complete = False

    with pytest.raises(FormulaFenceError, match="complete array-formula metadata"):
        lint_snapshot(snapshot)


def test_lint_bounds_distinct_pattern_targets(tmp_path: Path) -> None:
    snapshot = _snapshot(
        tmp_path,
        {
            "A2": 1,
            "B2": "=A2*2",
            "D2": "=C2*2",
            "E2": "=D2*2",
            "A4": 1,
            "B4": "=A4*2",
            "D4": "=C4*2",
            "E4": "=D4*2",
        },
    )

    with pytest.raises(FormulaFenceError, match="max_formula_pattern_findings=1"):
        lint_snapshot(snapshot, max_formula_pattern_findings=1)


def test_lint_shares_its_finding_cap_with_aggregate_omissions(tmp_path: Path) -> None:
    snapshot = _snapshot(
        tmp_path,
        {
            "A2": 1,
            "B2": "=A2*2",
            "D2": "=C2*2",
            "E2": "=D2*2",
            "G2": 10,
            "G3": 20,
            "G4": 30,
            "G5": 40,
            "G6": 50,
            "G7": "=SUM(G2:G4)",
        },
    )

    with pytest.raises(FormulaFenceError, match="max_formula_pattern_findings=1"):
        lint_snapshot(snapshot, max_formula_pattern_findings=1)


def test_lint_shares_its_finding_cap_with_direct_unlocked_formulas(
    tmp_path: Path,
) -> None:
    snapshot = _snapshot(
        tmp_path,
        {"A2": 10, "B2": "=A2*2", "C2": "=A2*3"},
        protected=True,
        unlocked_cells=("B2", "C2"),
    )

    with pytest.raises(FormulaFenceError, match="max_formula_pattern_findings=1"):
        lint_snapshot(snapshot, max_formula_pattern_findings=1)


def test_lint_shares_its_finding_cap_with_incomplete_manual_calculation(
    tmp_path: Path,
) -> None:
    snapshot = _snapshot(
        tmp_path,
        {"A2": 10, "B2": "=A2*2"},
        protected=True,
        unlocked_cells=("B2",),
        calculation_mode="manual",
        calculation_completed=False,
    )

    with pytest.raises(FormulaFenceError, match="max_formula_pattern_findings=1"):
        lint_snapshot(snapshot, max_formula_pattern_findings=1)


def test_lint_shares_its_finding_cap_with_direct_self_references(
    tmp_path: Path,
) -> None:
    snapshot = _snapshot(
        tmp_path,
        {"B2": "=B2+1", "C2": "=C2+1"},
    )

    with pytest.raises(FormulaFenceError, match="max_formula_pattern_findings=1"):
        lint_snapshot(snapshot, max_formula_pattern_findings=1)


def test_lint_shares_its_finding_cap_with_multi_cell_static_cycles(
    tmp_path: Path,
) -> None:
    snapshot = _snapshot(
        tmp_path,
        {
            "B2": "=C2+1",
            "C2": "=B2+1",
            "D2": "=E2+1",
            "E2": "=D2+1",
        },
    )

    with pytest.raises(FormulaFenceError, match="max_formula_pattern_findings=1"):
        lint_snapshot(snapshot, max_formula_pattern_findings=1)


def test_lint_shares_its_finding_cap_with_broken_references(tmp_path: Path) -> None:
    snapshot = _snapshot(
        tmp_path,
        {"B2": "=#REF!", "C2": "=IFERROR(#REF!,0)"},
    )

    with pytest.raises(FormulaFenceError, match="max_formula_pattern_findings=1"):
        lint_snapshot(snapshot, max_formula_pattern_findings=1)


def test_lint_renderers_keep_formula_text_out_of_review_artifacts(tmp_path: Path) -> None:
    snapshot = _snapshot(
        tmp_path,
        {
            "A2": 1,
            "B2": "=A2*2",
            "D2": "=C2*2",
            "E2": "=D2*2",
        },
    )
    report = lint_snapshot(snapshot)

    rendered_json = as_json(report.to_dict())
    rendered_markdown = lint_to_markdown(report)
    rendered_sarif = lint_to_sarif(report)

    for rendered in (rendered_json, rendered_markdown, str(rendered_sarif)):
        assert "=A2*2" not in rendered
        assert "FF082" in rendered
    result = rendered_sarif["runs"][0]["results"][0]
    assert result["locations"][0]["logicalLocations"][0]["name"] == "Model!C2"
    assert result["properties"]["pattern_kind"] == "blank_gap"
    with pytest.raises(FormulaFenceError, match="max_report_bytes=1"):
        lint_to_markdown(report, max_bytes=1)


def test_lint_aggregate_renderers_keep_formula_text_out_of_review_artifacts(
    tmp_path: Path,
) -> None:
    snapshot = _snapshot(
        tmp_path,
        {
            "B2": 10,
            "B3": 20,
            "B4": 30,
            "B5": 40,
            "B6": 50,
            "B7": "=SUM(B2:B4)",
        },
    )
    report = lint_snapshot(snapshot)

    rendered_json = as_json(report.to_dict())
    rendered_markdown = lint_to_markdown(report)
    rendered_sarif = lint_to_sarif(report)

    for rendered in (rendered_json, rendered_markdown, str(rendered_sarif)):
        assert "=SUM(B2:B4)" not in rendered
        assert "FF084" in rendered
    assert "## Aggregate range evidence" in rendered_markdown
    assert "`Model!B2:B4`" in rendered_markdown
    assert rendered_sarif["runs"][0]["results"][0]["ruleId"] == "FF084"
    assert rendered_sarif["runs"][0]["tool"]["driver"]["rules"] == [
        {
            "id": "FF084",
            "name": "FF084",
            "shortDescription": {
                "text": "A simple numeric aggregate stops before adjacent numeric cells."
            },
        }
    ]


def test_lint_unlocked_formula_renderers_keep_formula_text_out_of_review_artifacts(
    tmp_path: Path,
) -> None:
    snapshot = _snapshot(
        tmp_path,
        {"A2": 10, "B2": "=A2*2"},
        protected=True,
        unlocked_cells=("B2",),
    )
    report = lint_snapshot(snapshot)

    rendered_json = as_json(report.to_dict())
    rendered_markdown = lint_to_markdown(report)
    rendered_sarif = lint_to_sarif(report)

    for rendered in (rendered_json, rendered_markdown, str(rendered_sarif)):
        assert "=A2*2" not in rendered
        assert "FF085" in rendered
    assert "## Formula protection evidence" in rendered_markdown
    assert rendered_sarif["runs"][0]["results"][0]["ruleId"] == "FF085"
    assert rendered_sarif["runs"][0]["tool"]["driver"]["rules"] == [
        {
            "id": "FF085",
            "name": "FF085",
            "shortDescription": {
                "text": "A formula cell is explicitly unlocked on a protected worksheet."
            },
        }
    ]


def test_lint_incomplete_manual_calculation_renderers_keep_formula_text_out(
    tmp_path: Path,
) -> None:
    snapshot = _snapshot(
        tmp_path,
        {"A2": 10, "B2": "=A2*2"},
        calculation_mode="manual",
        calculation_completed=False,
    )
    report = lint_snapshot(snapshot)

    rendered_json = as_json(report.to_dict())
    rendered_markdown = lint_to_markdown(report)
    rendered_sarif = lint_to_sarif(report)

    for rendered in (rendered_json, rendered_markdown, str(rendered_sarif)):
        assert "=A2*2" not in rendered
        assert "FF086" in rendered
    assert "## Calculation freshness evidence" in rendered_markdown
    result = rendered_sarif["runs"][0]["results"][0]
    assert "locations" not in result
    assert result["properties"] == {
        "severity": "medium",
        "calculation_mode": "manual",
        "calculation_completed_before_save": False,
    }
    assert rendered_sarif["runs"][0]["tool"]["driver"]["rules"] == [
        {
            "id": "FF086",
            "name": "FF086",
            "shortDescription": {
                "text": "A formula workbook was saved with incomplete manual calculation."
            },
        }
    ]


def test_lint_direct_self_reference_renderers_keep_formula_text_out(
    tmp_path: Path,
) -> None:
    snapshot = _snapshot(tmp_path, {"B2": "=B2+1"})
    report = lint_snapshot(snapshot)

    rendered_json = as_json(report.to_dict())
    rendered_markdown = lint_to_markdown(report)
    rendered_sarif = lint_to_sarif(report)

    for rendered in (rendered_json, rendered_markdown, str(rendered_sarif)):
        assert "=B2+1" not in rendered
        assert "FF087" in rendered
    assert "## Static circular-reference evidence" in rendered_markdown
    result = rendered_sarif["runs"][0]["results"][0]
    assert result["locations"][0]["logicalLocations"][0]["name"] == "Model!B2"
    assert result["properties"] == {
        "severity": "high",
        "calculation_iteration_enabled": False,
        "reference_scope": "direct_static",
    }
    assert rendered_sarif["runs"][0]["tool"]["driver"]["rules"] == [
        {
            "id": "FF087",
            "name": "FF087",
            "shortDescription": {
                "text": (
                    "A formula directly references its own cell while calculation "
                    "iteration is disabled."
                )
            },
        }
    ]


def test_lint_multi_cell_static_cycle_renderers_keep_formula_text_out(
    tmp_path: Path,
) -> None:
    snapshot = _snapshot(tmp_path, {"B2": "=C2+1", "C2": "=B2+1"})
    report = lint_snapshot(snapshot)

    rendered_json = as_json(report.to_dict())
    rendered_markdown = lint_to_markdown(report)
    rendered_sarif = lint_to_sarif(report)

    for rendered in (rendered_json, rendered_markdown, str(rendered_sarif)):
        assert "=C2+1" not in rendered
        assert "=B2+1" not in rendered
        assert "FF090" in rendered
    assert "## Static circular-reference evidence" in rendered_markdown
    assert "a static multi-cell dependency component has 2 formula cells" in rendered_markdown
    assert [
        result["locations"][0]["logicalLocations"][0]["name"]
        for result in rendered_sarif["runs"][0]["results"]
    ] == ["Model!B2", "Model!C2"]
    assert {
        result["properties"]["cycle_member_count"]
        for result in rendered_sarif["runs"][0]["results"]
    } == {2}
    assert rendered_sarif["runs"][0]["tool"]["driver"]["rules"] == [
        {
            "id": "FF090",
            "name": "FF090",
            "shortDescription": {
                "text": (
                    "A formula participates in a static multi-cell circular reference "
                    "while calculation iteration is disabled."
                )
            },
        }
    ]


def test_lint_broken_reference_renderers_keep_formula_text_out(tmp_path: Path) -> None:
    snapshot = _snapshot(tmp_path, {"B2": "=IFERROR(#REF!,0)"})
    report = lint_snapshot(snapshot)

    rendered_json = as_json(report.to_dict())
    rendered_markdown = lint_to_markdown(report)
    rendered_sarif = lint_to_sarif(report)

    for rendered in (rendered_json, rendered_markdown, str(rendered_sarif)):
        assert "=IFERROR(#REF!,0)" not in rendered
        assert "FF088" in rendered
    assert "## Explicit broken-reference evidence" in rendered_markdown
    result = rendered_sarif["runs"][0]["results"][0]
    assert result["locations"][0]["logicalLocations"][0]["name"] == "Model!B2"
    assert result["properties"] == {"severity": "critical"}
    assert rendered_sarif["runs"][0]["tool"]["driver"]["rules"] == [
        {
            "id": "FF088",
            "name": "FF088",
            "shortDescription": {
                "text": "Formula contains an explicit broken #REF! reference."
            },
        }
    ]


def test_lint_saved_broken_reference_renderers_keep_cache_data_out(
    tmp_path: Path,
) -> None:
    snapshot = load_snapshot(
        make_formula_cached_result_model(
            tmp_path / "saved-reference-error.xlsx",
            error_formula="=INDEX(Inputs!A1:A1,2)",
            error_result="#REF!",
        )
    )
    report = lint_snapshot(snapshot)

    rendered_json = as_json(report.to_dict())
    rendered_markdown = lint_to_markdown(report)
    rendered_sarif = lint_to_sarif(report)

    for rendered in (rendered_json, rendered_markdown, str(rendered_sarif)):
        assert "=INDEX(Inputs!A1:A1,2)" not in rendered
        assert "#REF!" not in rendered
        assert "FF089" in rendered
    assert "## Saved broken-reference evidence" in rendered_markdown
    result = rendered_sarif["runs"][0]["results"][0]
    assert result["locations"][0]["logicalLocations"][0]["name"] == "Report!B5"
    assert result["properties"] == {"severity": "high"}
    assert rendered_sarif["runs"][0]["tool"]["driver"]["rules"] == [
        {
            "id": "FF089",
            "name": "FF089",
            "shortDescription": {
                "text": "A formula's saved result is a broken-reference error."
            },
        }
    ]
