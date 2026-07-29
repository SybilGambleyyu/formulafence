"""Tests for conservative single-workbook formula-pattern linting."""

from __future__ import annotations

from pathlib import Path

import pytest
from openpyxl import Workbook

from formulafence.lint import lint_snapshot
from formulafence.models import ArrayFormulaRange, FormulaFenceError
from formulafence.output import as_json, lint_to_markdown, lint_to_sarif
from formulafence.workbook import load_snapshot


def _snapshot(tmp_path: Path, cells: dict[str, object]):
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = "Model"
    for coordinate, value in cells.items():
        worksheet[coordinate] = value
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
