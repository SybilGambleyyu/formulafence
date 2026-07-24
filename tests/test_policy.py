from __future__ import annotations

import pytest

from formulafence.diff import compare_snapshots
from formulafence.models import PolicyError
from formulafence.policy import evaluate_policy, parse_policy
from formulafence.workbook import load_snapshot

from .helpers import (
    make_data_validation_model,
    make_legacy_array_model,
    make_model,
    make_table_model,
    make_three_d_model,
    mark_array_formula_dynamic,
    rewrite,
)


def test_policy_fails_formula_override_and_protected_output(tmp_path) -> None:
    baseline = make_model(tmp_path / "baseline.xlsx")
    candidate = make_model(tmp_path / "candidate.xlsx")

    def mutate(workbook) -> None:
        workbook["Model"]["B2"] = 200
        workbook["Dashboard"]["B12"] = 999

    rewrite(candidate, mutate)
    report = compare_snapshots(load_snapshot(baseline), load_snapshot(candidate))
    policy = parse_policy(
        {
            "version": 1,
            "rules": {"no_formula_to_value": True, "max_downstream_impact": 1},
            "protected_cells": ["Dashboard!B12"],
            "allowed_changes": ["Inputs!B2:B3"],
        }
    )

    rule_ids = {finding.rule_id for finding in evaluate_policy(report, policy)}
    assert {"FFP001", "FFP006", "FFP007", "FFP009"} <= rule_ids


def test_policy_rejects_unknown_fields_and_unsafe_selectors() -> None:
    with pytest.raises(PolicyError, match="Unknown rules"):
        parse_policy({"version": 1, "rules": {"no_formula_to_number": True}})
    with pytest.raises(PolicyError, match="sheet-qualified"):
        parse_policy({"version": 1, "protected_cells": ["B12"]})


def test_policy_can_block_new_formula_coverage_gaps(tmp_path) -> None:
    baseline = make_model(tmp_path / "baseline.xlsx")
    candidate = make_model(tmp_path / "candidate.xlsx")
    rewrite(
        candidate,
        lambda workbook: setattr(
            workbook["Model"]["D2"],
            "value",
            '=UnknownMetric+OFFSET(Inputs!B2, 1, 0)',
        ),
    )
    report = compare_snapshots(load_snapshot(baseline), load_snapshot(candidate))
    policy = parse_policy(
        {
            "version": 1,
            "rules": {
                "no_new_unresolved_references": True,
                "no_new_dynamic_references": True,
            },
        }
    )

    rule_ids = {finding.rule_id for finding in evaluate_policy(report, policy)}
    assert {"FFP011", "FFP012"} <= rule_ids


def test_policy_can_block_spill_and_tokenization_coverage_limits(tmp_path) -> None:
    baseline = make_model(tmp_path / "baseline.xlsx")
    candidate = make_model(tmp_path / "candidate.xlsx")

    def add_coverage_limits(workbook) -> None:
        workbook["Model"]["D2"] = "=SUM(Inputs!B2#)"
        workbook["Model"]["D3"] = "=SUM(Inputs!B2#1)"

    rewrite(candidate, add_coverage_limits)
    report = compare_snapshots(load_snapshot(baseline), load_snapshot(candidate))
    policy = parse_policy(
        {
            "version": 1,
            "rules": {
                "no_new_spill_references": True,
                "no_new_tokenization_failures": True,
            },
        }
    )

    rule_ids = {finding.rule_id for finding in evaluate_policy(report, policy)}
    assert {"FFP015", "FFP016"} <= rule_ids


def test_policy_can_block_implicit_intersection(tmp_path) -> None:
    baseline = make_model(tmp_path / "baseline.xlsx")
    candidate = make_model(tmp_path / "candidate.xlsx")
    rewrite(
        candidate,
        lambda workbook: setattr(
            workbook["Model"]["D2"], "value", "=@Inputs!B2:B4"
        ),
    )
    report = compare_snapshots(load_snapshot(baseline), load_snapshot(candidate))
    policy = parse_policy(
        {
            "version": 1,
            "rules": {"no_new_implicit_intersections": True},
        }
    )

    rule_ids = {finding.rule_id for finding in evaluate_policy(report, policy)}
    assert "FFP017" in rule_ids


def test_policy_can_block_array_formula_semantics_changes(tmp_path) -> None:
    baseline = make_legacy_array_model(tmp_path / "baseline.xlsx", "B1:B3")
    candidate = make_legacy_array_model(tmp_path / "candidate.xlsx", "B1:B4")
    report = compare_snapshots(load_snapshot(baseline), load_snapshot(candidate))
    policy = parse_policy(
        {
            "version": 1,
            "rules": {"no_array_formula_semantics_changes": True},
        }
    )

    assert {finding.rule_id for finding in evaluate_policy(report, policy)} >= {"FFP018"}


def test_policy_can_block_new_dynamic_array_output_member_references(tmp_path) -> None:
    baseline = make_legacy_array_model(tmp_path / "baseline.xlsx")

    def remove_member_consumers(workbook) -> None:
        workbook["Model"]["C2"] = None
        workbook["Dashboard"]["B2"] = None

    rewrite(baseline, remove_member_consumers)
    candidate = make_legacy_array_model(tmp_path / "candidate.xlsx")
    mark_array_formula_dynamic(baseline)
    mark_array_formula_dynamic(candidate)
    report = compare_snapshots(load_snapshot(baseline), load_snapshot(candidate))
    policy = parse_policy(
        {
            "version": 1,
            "rules": {"no_new_dynamic_array_output_references": True},
        }
    )

    assert {finding.rule_id for finding in evaluate_policy(report, policy)} >= {"FFP019"}


def test_policy_can_block_table_definition_changes(tmp_path) -> None:
    baseline = make_table_model(tmp_path / "baseline.xlsx")
    candidate = make_table_model(tmp_path / "candidate.xlsx")
    rewrite(candidate, lambda workbook: setattr(workbook["Data"].tables["Sales"], "ref", "A1:C3"))
    report = compare_snapshots(load_snapshot(baseline), load_snapshot(candidate))
    policy = parse_policy(
        {"version": 1, "rules": {"no_table_definition_changes": True}}
    )

    assert {finding.rule_id for finding in evaluate_policy(report, policy)} >= {"FFP013"}


def test_policy_can_block_data_validation_control_changes(tmp_path) -> None:
    baseline = make_data_validation_model(tmp_path / "baseline.xlsx")
    candidate = make_data_validation_model(tmp_path / "candidate.xlsx")
    rewrite(
        candidate,
        lambda workbook: setattr(
            workbook["Inputs"].data_validations.dataValidation[0],
            "showErrorMessage",
            False,
        ),
    )
    report = compare_snapshots(load_snapshot(baseline), load_snapshot(candidate))
    policy = parse_policy(
        {"version": 1, "rules": {"no_data_validation_changes": True}}
    )

    assert {finding.rule_id for finding in evaluate_policy(report, policy)} >= {"FFP020"}


def test_policy_can_block_three_d_reference_scope_changes(tmp_path) -> None:
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
    policy = parse_policy(
        {"version": 1, "rules": {"no_3d_reference_scope_changes": True}}
    )

    assert {finding.rule_id for finding in evaluate_policy(report, policy)} >= {"FFP014"}
