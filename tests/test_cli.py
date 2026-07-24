from __future__ import annotations

import json

from formulafence.cli import main

from .helpers import make_model, rewrite


def test_check_emits_sarif_and_fails_for_a_policy_violation(tmp_path) -> None:
    baseline = make_model(tmp_path / "baseline.xlsx")
    candidate = make_model(tmp_path / "candidate.xlsx")
    rewrite(candidate, lambda workbook: setattr(workbook["Model"]["B2"], "value", 200))
    policy = tmp_path / "formulafence.yml"
    policy.write_text(
        "version: 1\nrules:\n  no_formula_to_value: true\n",
        encoding="utf-8",
    )
    sarif = tmp_path / "result.sarif"

    result = main(
        [
            "check",
            str(baseline),
            str(candidate),
            "--policy",
            str(policy),
            "--format",
            "sarif",
            "--output",
            str(sarif),
        ]
    )

    assert result == 1
    payload = json.loads(sarif.read_text(encoding="utf-8"))
    assert payload["version"] == "2.1.0"
    assert any(item["ruleId"] == "FFP001" for item in payload["runs"][0]["results"])


def test_profile_does_not_expose_cell_values(tmp_path) -> None:
    workbook = make_model(tmp_path / "model.xlsx")
    output = tmp_path / "profile.json"

    assert main(["profile", str(workbook), "--format", "json", "--output", str(output)]) == 0

    profile = output.read_text(encoding="utf-8")
    assert "Calculated revenue" not in profile
    assert '"formula_cells"' in profile
