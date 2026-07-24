from __future__ import annotations

import pytest

from formulafence.diff import compare_snapshots
from formulafence.models import PolicyError
from formulafence.policy import evaluate_policy, parse_policy
from formulafence.workbook import load_snapshot

from .helpers import make_model, rewrite


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
