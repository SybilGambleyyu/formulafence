"""Rendered-artifact byte-budget checks."""

from __future__ import annotations

from pathlib import Path

import pytest

import formulafence.workbook as workbook_module
from formulafence.diff import compare_snapshots
from formulafence.models import (
    DataValidationSnapshot,
    FormulaFenceError,
    TableSnapshot,
    WorkbookSnapshot,
)
from formulafence.output import (
    as_json,
    profile_to_markdown,
    report_to_html,
    report_to_markdown,
    report_to_sarif,
)
from formulafence.workbook import load_snapshot, profile_snapshot

from .helpers import make_model, rewrite


def test_json_byte_limit_counts_utf8_exactly() -> None:
    payload = {"currency": "€"}
    rendered = as_json(payload)

    assert as_json(payload, max_bytes=len(rendered.encode("utf-8"))) == rendered
    with pytest.raises(FormulaFenceError, match="max_report_bytes=1"):
        as_json(payload, max_bytes=1)


def test_rendered_report_formats_fail_before_an_overage(tmp_path) -> None:
    baseline = make_model(tmp_path / "baseline.xlsx")
    candidate = make_model(tmp_path / "candidate.xlsx")
    rewrite(
        candidate,
        lambda workbook: setattr(workbook["Model"]["B2"], "value", "<&€"),
    )
    report = compare_snapshots(load_snapshot(baseline), load_snapshot(candidate))

    json_report = as_json(report.to_dict())
    markdown_report = report_to_markdown(report)
    html_report = report_to_html(report, max_bytes=1_000_000)
    sarif_report = as_json(report_to_sarif(report))

    assert as_json(report.to_dict(), max_bytes=len(json_report.encode("utf-8"))) == json_report
    assert (
        report_to_markdown(report, max_bytes=len(markdown_report.encode("utf-8")))
        == markdown_report
    )
    assert (
        report_to_html(report, max_bytes=len(html_report.encode("utf-8")))
        == html_report
    )
    assert (
        as_json(report_to_sarif(report), max_bytes=len(sarif_report.encode("utf-8")))
        == sarif_report
    )

    with pytest.raises(FormulaFenceError, match="max_report_bytes=1"):
        as_json(report.to_dict(), max_bytes=1)
    with pytest.raises(FormulaFenceError, match="max_report_bytes=1"):
        report_to_markdown(report, max_bytes=1)
    with pytest.raises(FormulaFenceError, match="max_report_bytes=1"):
        report_to_html(report, max_bytes=1)
    with pytest.raises(FormulaFenceError, match="max_report_bytes=1"):
        as_json(report_to_sarif(report), max_bytes=1)


def test_profile_renderers_respect_the_report_byte_limit(tmp_path) -> None:
    workbook = make_model(tmp_path / "model.xlsx")
    profile = profile_snapshot(load_snapshot(workbook))
    json_profile = as_json(profile)
    markdown_profile = profile_to_markdown(profile)

    assert (
        as_json(profile, max_bytes=len(json_profile.encode("utf-8"))) == json_profile
    )
    assert (
        profile_to_markdown(
            profile,
            max_bytes=len(markdown_profile.encode("utf-8")),
        )
        == markdown_profile
    )

    with pytest.raises(FormulaFenceError, match="max_report_bytes=1"):
        as_json(profile, max_bytes=1)
    with pytest.raises(FormulaFenceError, match="max_report_bytes=1"):
        profile_to_markdown(profile, max_bytes=1)


def test_profile_record_budget_is_exact_and_preflights_before_profile_build(
    monkeypatch,
) -> None:
    snapshot = WorkbookSnapshot(
        path=Path("model.xlsx"),
        sha256="0" * 64,
        file_type="xlsx",
        sheets={},
        cells={},
        reverse_dependencies={},
        range_dependencies=[],
        external_references=set(),
        broken_references=set(),
        defined_names={},
        macro_hash=None,
        calculation_settings={},
        parser_warnings=(),
        tables={
            "Sales": TableSnapshot(
                name="Sales",
                sheet="Model",
                ref="A1:B2",
                columns=("North", "South"),
                header_row_count=1,
                totals_row_count=0,
            )
        },
        data_validations=(
            DataValidationSnapshot(
                sheet="Model",
                ranges=("A1", "B2"),
                validation_type="whole",
                operator="between",
                formula1=None,
                formula2=None,
                allow_blank=False,
                dropdown_hidden=False,
                prompts_disabled=False,
                show_input_message=True,
                show_error_message=True,
                error_style="stop",
                error_title=None,
                error=None,
                prompt_title=None,
                prompt=None,
                ime_mode="noControl",
            ),
        ),
        dynamic_reference_functions={("Model", "A1"): ("INDIRECT",)},
    )

    expected = profile_snapshot(snapshot)
    assert profile_snapshot(snapshot, max_profile_records=8) == expected

    def unexpected_profile_construction(*args, **kwargs):
        raise AssertionError("profile construction ran after the inventory limit failed")

    monkeypatch.setattr(workbook_module, "display_location", unexpected_profile_construction)
    with pytest.raises(FormulaFenceError, match="max_profile_records=7"):
        profile_snapshot(snapshot, max_profile_records=7)
    with pytest.raises(FormulaFenceError, match="must be at least 1"):
        profile_snapshot(snapshot, max_profile_records=0)
