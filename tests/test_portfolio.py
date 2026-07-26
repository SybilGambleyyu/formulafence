"""Portfolio comparison contract tests."""

from __future__ import annotations

import json
import shutil
from pathlib import Path

import pytest
from openpyxl import Workbook, load_workbook

from formulafence.cli import main
from formulafence.output import as_json, portfolio_to_markdown, portfolio_to_sarif
from formulafence.policy import parse_policy
from formulafence.portfolio import PortfolioError, compare_portfolios, discover_workbooks

from .helpers import make_model, rewrite


def _copy_workbook(source: Path, destination: Path) -> None:
    destination.parent.mkdir(parents=True, exist_ok=True)
    shutil.copyfile(source, destination)


def _write_workbook(path: Path, sheet_name: str, cells: dict[str, object]) -> Path:
    path.parent.mkdir(parents=True, exist_ok=True)
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = sheet_name
    for coordinate, value in cells.items():
        worksheet[coordinate] = value
    workbook.save(path)
    return path


def _portfolio_pair(tmp_path: Path) -> tuple[Path, Path]:
    baseline = tmp_path / "baseline"
    candidate = tmp_path / "candidate"
    baseline.mkdir()
    candidate.mkdir()
    (baseline / "models").mkdir()
    source = make_model(baseline / "models" / "shared.xlsx")
    _copy_workbook(source, candidate / "models" / "shared.xlsx")
    return baseline, candidate


def _cross_workbook_portfolio_pair(tmp_path: Path) -> tuple[Path, Path]:
    baseline = tmp_path / "baseline"
    candidate = tmp_path / "candidate"
    _write_workbook(baseline / "inputs.xlsx", "Data", {"B2": 100})
    _write_workbook(
        baseline / "calculation.xlsx",
        "Model",
        {
            "D2": "=[inputs.xlsx]Data!B2",
            "E2": "=D2*2",
        },
    )
    _write_workbook(
        baseline / "reports" / "summary.xlsx",
        "Summary",
        {
            "D2": "='..\\[calculation.xlsx]Model'!E2",
            "E2": "=D2*2",
        },
    )
    shutil.copytree(baseline, candidate)
    rewrite(
        candidate / "inputs.xlsx",
        lambda workbook: setattr(workbook["Data"]["B2"], "value", 200),
    )
    return baseline, candidate


def test_portfolio_matches_relative_paths_and_reports_membership_changes(tmp_path: Path) -> None:
    baseline, candidate = _portfolio_pair(tmp_path)

    def replace_formula(workbook) -> None:
        workbook["Model"]["B2"] = 200

    rewrite(candidate / "models" / "shared.xlsx", replace_formula)
    make_model(baseline / "removed.xlsx")
    make_model(candidate / "added.xlsx")

    report = compare_portfolios(baseline, candidate)
    payload = report.to_dict()
    entries = {entry["path"]: entry for entry in payload["workbooks"]}

    assert payload["before"]["workbook_count"] == 2
    assert payload["after"]["workbook_count"] == 2
    assert entries["models/shared.xlsx"]["status"] == "changed"
    assert entries["added.xlsx"]["status"] == "added"
    assert entries["removed.xlsx"]["status"] == "removed"
    assert entries["models/shared.xlsx"]["before"]["path"] == "models/shared.xlsx"
    assert any(
        change["kind"] == "formula_to_value"
        for change in entries["models/shared.xlsx"]["changes"]
    )
    assert {finding["rule_id"] for finding in entries["added.xlsx"]["findings"]} == {
        "FF077"
    }
    assert {finding["rule_id"] for finding in entries["removed.xlsx"]["findings"]} == {
        "FF077"
    }
    assert str(tmp_path) not in json.dumps(payload)


def test_portfolio_policy_applies_per_workbook_and_blocks_membership(tmp_path: Path) -> None:
    baseline, candidate = _portfolio_pair(tmp_path)

    def replace_formula(workbook) -> None:
        workbook["Model"]["B2"] = 200

    rewrite(candidate / "models" / "shared.xlsx", replace_formula)
    make_model(candidate / "added.xlsx")
    policy = parse_policy(
        {
            "version": 1,
            "rules": {
                "no_formula_to_value": True,
                "no_portfolio_membership_changes": True,
            },
        }
    )

    report = compare_portfolios(baseline, candidate, policy=policy)
    by_path = {entry.path: entry for entry in report.workbooks}
    assert {finding.rule_id for finding in by_path["models/shared.xlsx"].findings} >= {
        "FFP001"
    }
    assert {finding.rule_id for finding in by_path["added.xlsx"].findings} == {
        "FF077",
        "FFP077",
    }

    policy_path = tmp_path / "formulafence.yml"
    policy_path.write_text(
        "version: 1\nrules:\n  no_formula_to_value: true\n"
        "  no_portfolio_membership_changes: true\n",
        encoding="utf-8",
    )
    output = tmp_path / "portfolio.json"
    assert (
        main(
            [
                "portfolio",
                str(baseline),
                str(candidate),
                "--policy",
                str(policy_path),
                "--format",
                "json",
                "--output",
                str(output),
            ]
        )
        == 1
    )
    assert "FFP077" in output.read_text(encoding="utf-8")


def test_portfolio_reports_transitive_static_cross_workbook_impacts(tmp_path: Path) -> None:
    baseline, candidate = _cross_workbook_portfolio_pair(tmp_path)
    policy = parse_policy(
        {"version": 1, "rules": {"no_cross_workbook_impacts": True}}
    )

    report = compare_portfolios(baseline, candidate, policy=policy)
    source = next(entry for entry in report.workbooks if entry.path == "inputs.xlsx")
    finding = next(finding for finding in source.findings if finding.rule_id == "FF079")

    assert finding.location == ("Data", "B2")
    assert finding.details["impacted_workbook_count"] == 2
    assert finding.details["impacted_formula_count"] == 4
    assert finding.details["sample_impacts"] == [
        {
            "workbook": "calculation.xlsx",
            "location": "Model!D2",
            "path": [
                {"workbook": "inputs.xlsx", "location": "Data!B2"},
                {"workbook": "calculation.xlsx", "location": "Model!D2"},
            ],
        },
        {
            "workbook": "calculation.xlsx",
            "location": "Model!E2",
            "path": [
                {"workbook": "inputs.xlsx", "location": "Data!B2"},
                {"workbook": "calculation.xlsx", "location": "Model!D2"},
                {"workbook": "calculation.xlsx", "location": "Model!E2"},
            ],
        },
        {
            "workbook": "reports/summary.xlsx",
            "location": "Summary!D2",
            "path": [
                {"workbook": "inputs.xlsx", "location": "Data!B2"},
                {"workbook": "calculation.xlsx", "location": "Model!D2"},
                {"workbook": "calculation.xlsx", "location": "Model!E2"},
                {"workbook": "reports/summary.xlsx", "location": "Summary!D2"},
            ],
        },
        {
            "workbook": "reports/summary.xlsx",
            "location": "Summary!E2",
            "path": [
                {"workbook": "inputs.xlsx", "location": "Data!B2"},
                {"workbook": "calculation.xlsx", "location": "Model!D2"},
                {"workbook": "calculation.xlsx", "location": "Model!E2"},
                {"workbook": "reports/summary.xlsx", "location": "Summary!D2"},
                {"workbook": "reports/summary.xlsx", "location": "Summary!E2"},
            ],
        },
    ]
    assert {finding.rule_id for finding in source.findings} >= {"FF079", "FFP079"}
    assert not report.incomplete

    payload = report.to_dict()
    source_payload = next(
        entry for entry in payload["workbooks"] if entry["path"] == "inputs.xlsx"
    )
    assert source_payload["summary"]["raw_finding_count"] == 1
    assert source_payload["summary"]["policy_finding_count"] == 1
    assert payload["summary"]["cross_workbook_impact_source_count"] == 1
    assert payload["summary"]["cross_workbook_impacted_formula_count"] == 4
    assert not payload["summary"]["cross_workbook_impact_incomplete"]
    markdown = portfolio_to_markdown(report)
    assert "FF079" in markdown
    assert "`inputs.xlsx` `Data!B2` → `calculation.xlsx` `Model!D2`" in markdown
    sarif_results = portfolio_to_sarif(report)["runs"][0]["results"]
    assert any(result["ruleId"] == "FFP079" for result in sarif_results)
    sarif_finding = next(result for result in sarif_results if result["ruleId"] == "FF079")
    assert sarif_finding["locations"][0]["physicalLocation"]["artifactLocation"]["uri"] == (
        "inputs.xlsx"
    )
    assert sarif_finding["locations"][0]["logicalLocations"][0]["name"] == "Data!B2"
    assert sarif_finding["properties"]["sample_impacts"][0]["workbook"] == (
        "calculation.xlsx"
    )


def test_portfolio_resolves_relative_ranges_case_insensitively_without_basename_guessing(
    tmp_path: Path,
) -> None:
    baseline = tmp_path / "baseline"
    candidate = tmp_path / "candidate"
    _write_workbook(
        baseline / "inputs" / "shared.xlsx",
        "Data",
        {"B2": 10, "B3": 20, "B4": 30},
    )
    _write_workbook(
        baseline / "reports" / "summary.xlsx",
        "Summary",
        {"D2": "=SUM('..\\INPUTS\\[SHARED.XLSX]data'!B2:B4)"},
    )
    _write_workbook(baseline / "other" / "shared.xlsx", "Data", {"B3": 999})
    shutil.copytree(baseline, candidate)
    rewrite(
        candidate / "inputs" / "shared.xlsx",
        lambda workbook: setattr(workbook["Data"]["B3"], "value", 21),
    )

    report = compare_portfolios(baseline, candidate)
    source = next(entry for entry in report.workbooks if entry.path == "inputs/shared.xlsx")
    finding = next(finding for finding in source.findings if finding.rule_id == "FF079")

    assert finding.details["impacted_workbook_count"] == 1
    assert finding.details["impacted_formula_count"] == 1
    assert finding.details["sample_impacts"][0]["workbook"] == "reports/summary.xlsx"
    assert finding.details["sample_impacts"][0]["location"] == "Summary!D2"


def test_portfolio_never_guesses_or_discloses_unresolved_external_link_paths(
    tmp_path: Path,
) -> None:
    baseline = tmp_path / "baseline"
    candidate = tmp_path / "candidate"
    private_source = "PRIVATE-EXTERNAL-LINK-PATH"
    _write_workbook(baseline / "inputs.xlsx", "Data", {"B2": 100})
    _write_workbook(
        baseline / "reports" / "summary.xlsx",
        "Summary",
        {
            "D2": "=[inputs.xlsx]Data!B2",
            "E2": f"='C:\\{private_source}\\[inputs.xlsx]Data'!B2",
            "F2": "='..\\..\\[inputs.xlsx]Data'!B2",
            "G2": "='..\\[inputs.xlsx]Data'!B2",
            "H2": "=[inputs.xlsx]ExternalName",
        },
    )
    shutil.copytree(baseline, candidate)
    rewrite(
        candidate / "inputs.xlsx",
        lambda workbook: setattr(workbook["Data"]["B2"], "value", 200),
    )

    report = compare_portfolios(baseline, candidate)
    source = next(entry for entry in report.workbooks if entry.path == "inputs.xlsx")
    rendered = (
        as_json(report.to_dict()),
        portfolio_to_markdown(report),
        as_json(portfolio_to_sarif(report)),
    )

    finding = next(finding for finding in source.findings if finding.rule_id == "FF079")
    assert finding.details["impacted_formula_count"] == 1
    assert finding.details["sample_impacts"][0]["location"] == "Summary!G2"
    assert all(private_source not in value for value in rendered)


def test_portfolio_fails_closed_when_cross_workbook_impact_bound_is_reached(
    tmp_path: Path,
) -> None:
    baseline, candidate = _cross_workbook_portfolio_pair(tmp_path)

    report = compare_portfolios(baseline, candidate, max_link_impact=2)
    source = next(entry for entry in report.workbooks if entry.path == "inputs.xlsx")

    assert report.incomplete
    assert {finding.rule_id for finding in source.findings} >= {"FF079", "FF080"}
    assert next(
        finding for finding in source.findings if finding.rule_id == "FF080"
    ).details == {"max_link_impact": 2}

    output = tmp_path / "portfolio.json"
    assert (
        main(
            [
                "portfolio",
                str(baseline),
                str(candidate),
                "--max-link-impact",
                "2",
                "--format",
                "json",
                "--output",
                str(output),
            ]
        )
        == 2
    )
    assert "FF080" in output.read_text(encoding="utf-8")


def test_portfolio_added_workbook_outputs_keep_contents_private(tmp_path: Path) -> None:
    baseline = tmp_path / "baseline"
    candidate = tmp_path / "candidate"
    baseline.mkdir()
    candidate.mkdir()
    workbook_path = make_model(candidate / "added.xlsx")
    secret = "PORTFOLIO-PRIVATE-CONTENT-DO-NOT-REPORT"
    workbook = load_workbook(workbook_path)
    workbook["Inputs"]["A1"] = secret
    workbook.save(workbook_path)

    report = compare_portfolios(baseline, candidate)
    rendered = (
        as_json(report.to_dict()),
        portfolio_to_markdown(report),
        as_json(portfolio_to_sarif(report)),
    )
    assert all(secret not in value for value in rendered)
    sarif_result = portfolio_to_sarif(report)["runs"][0]["results"][0]
    assert sarif_result["ruleId"] == "FF077"
    assert sarif_result["locations"][0]["physicalLocation"]["artifactLocation"]["uri"] == (
        "added.xlsx"
    )


def test_portfolio_preserves_a_report_for_an_unreadable_workbook(tmp_path: Path) -> None:
    baseline, candidate = _portfolio_pair(tmp_path)
    (candidate / "models" / "shared.xlsx").write_bytes(b"not a workbook")

    report = compare_portfolios(baseline, candidate)
    entry = report.to_dict()["workbooks"][0]
    assert report.incomplete
    assert entry["status"] == "unreadable"
    assert entry["findings"] == [
        {
            "rule_id": "FF078",
            "severity": "critical",
            "message": "Workbook could not be inspected; portfolio comparison is incomplete.",
            "location": None,
            "details": {"unreadable_sides": ["candidate"]},
        }
    ]

    output = tmp_path / "portfolio.md"
    assert (
        main(
            [
                "portfolio",
                str(baseline),
                str(candidate),
                "--output",
                str(output),
            ]
        )
        == 2
    )
    assert output.is_file()
    assert "FF078" in output.read_text(encoding="utf-8")


def test_portfolio_preserves_membership_evidence_for_an_unreadable_new_workbook(
    tmp_path: Path,
) -> None:
    baseline = tmp_path / "baseline"
    candidate = tmp_path / "candidate"
    baseline.mkdir()
    candidate.mkdir()
    (candidate / "new.xlsx").write_bytes(b"not a workbook")
    policy = parse_policy(
        {
            "version": 1,
            "rules": {"no_portfolio_membership_changes": True},
        }
    )

    report = compare_portfolios(baseline, candidate, policy=policy)
    entry = report.to_dict()["workbooks"][0]

    assert report.incomplete
    assert entry["status"] == "unreadable"
    assert len(entry["changes"]) == 1
    assert entry["changes"][0]["kind"] == "workbook_added"
    assert entry["changes"][0]["details"] == {"portfolio_change": "added"}
    assert {finding["rule_id"] for finding in entry["findings"]} == {
        "FF077",
        "FF078",
        "FFP077",
    }


def test_portfolio_cli_refuses_a_report_inside_an_input_directory(tmp_path: Path) -> None:
    baseline, candidate = _portfolio_pair(tmp_path)
    output = baseline / "report.md"

    assert (
        main(
            [
                "portfolio",
                str(baseline),
                str(candidate),
                "--output",
                str(output),
            ]
        )
        == 2
    )
    assert not output.exists()


def test_portfolio_inventory_fails_closed_for_limits_and_unsupported_formats(
    tmp_path: Path,
) -> None:
    baseline, candidate = _portfolio_pair(tmp_path)
    make_model(baseline / "second.xlsx")
    make_model(candidate / "second.xlsx")

    with pytest.raises(PortfolioError, match="max_workbooks=1"):
        compare_portfolios(baseline, candidate, max_workbooks=1)

    (baseline / "legacy.xlsb").write_bytes(b"unsupported")
    with pytest.raises(PortfolioError, match="unsupported spreadsheet files"):
        discover_workbooks(baseline, label="baseline")


def test_portfolio_inventory_ignores_office_lock_files(tmp_path: Path) -> None:
    baseline, _ = _portfolio_pair(tmp_path)
    (baseline / "~$shared.xlsx").write_bytes(b"transient office lock")

    assert set(discover_workbooks(baseline, label="baseline")) == {"models/shared.xlsx"}


def test_portfolio_inventory_refuses_symlinked_paths(tmp_path: Path) -> None:
    baseline, _ = _portfolio_pair(tmp_path)
    (baseline / "linked-models").symlink_to(baseline / "models", target_is_directory=True)

    with pytest.raises(PortfolioError, match="symlinked path"):
        discover_workbooks(baseline, label="baseline")


def test_portfolio_inventory_refuses_case_colliding_paths(tmp_path: Path) -> None:
    baseline, _ = _portfolio_pair(tmp_path)
    _copy_workbook(
        baseline / "models" / "shared.xlsx",
        baseline / "models" / "SHARED.xlsx",
    )

    with pytest.raises(PortfolioError, match="differ only by case"):
        discover_workbooks(baseline, label="baseline")
