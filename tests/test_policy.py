from __future__ import annotations

import pytest

from formulafence.diff import compare_snapshots
from formulafence.models import PolicyError
from formulafence.policy import evaluate_policy, parse_policy
from formulafence.workbook import load_snapshot

from .helpers import make_model, make_table_model, make_three_d_model, rewrite


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


def test_policy_can_block_table_definition_changes(tmp_path) -> None:
    baseline = make_table_model(tmp_path / "baseline.xlsx")
    candidate = make_table_model(tmp_path / "candidate.xlsx")
    rewrite(candidate, lambda workbook: setattr(workbook["Data"].tables["Sales"], "ref", "A1:C3"))
    report = compare_snapshots(load_snapshot(baseline), load_snapshot(candidate))
    policy = parse_policy(
        {"version": 1, "rules": {"no_table_definition_changes": True}}
    )

    assert {finding.rule_id for finding in evaluate_policy(report, policy)} >= {"FFP013"}


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
