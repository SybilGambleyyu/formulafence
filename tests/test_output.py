"""Rendered-artifact byte-budget checks."""

from __future__ import annotations

import pytest

from formulafence.diff import compare_snapshots
from formulafence.models import FormulaFenceError
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
