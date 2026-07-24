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
    results = payload["runs"][0]["results"]
    assert any(item["ruleId"] == "FFP001" for item in results)
    formula_override = next(item for item in results if item["ruleId"] == "FF001")
    assert formula_override["properties"]["impact_paths"] == [
        {"path": ["Model!B2", "Model!C2", "Dashboard!B12"], "target": "Dashboard!B12"},
        {"path": ["Model!B2", "Model!C2"], "target": "Model!C2"},
    ]


def test_profile_does_not_expose_cell_values(tmp_path) -> None:
    workbook = make_model(tmp_path / "model.xlsx")
    output = tmp_path / "profile.json"

    assert main(["profile", str(workbook), "--format", "json", "--output", str(output)]) == 0

    profile = output.read_text(encoding="utf-8")
    assert "Calculated revenue" not in profile
    assert '"formula_cells"' in profile


def test_init_policy_includes_modern_formula_coverage_controls(tmp_path) -> None:
    policy = tmp_path / "formulafence.yml"

    assert main(["init", str(policy)]) == 0

    content = policy.read_text(encoding="utf-8")
    assert "no_new_spill_references: true" in content
    assert "no_new_dynamic_array_output_references: true" in content
    assert "no_new_implicit_intersections: true" in content
    assert "no_array_formula_semantics_changes: true" in content
    assert "no_data_validation_changes: true" in content
    assert "no_conditional_formatting_changes: true" in content
    assert "no_protection_changes: true" in content
    assert "no_external_data_connection_changes: true" in content
    assert "no_external_link_package_changes: true" in content
    assert "no_xlm_macro_sheet_changes: true" in content
    assert "no_ribbon_customization_changes: true" in content
    assert "no_office_web_addin_changes: true" in content
    assert "no_power_query_changes: true" in content
    assert "no_new_tokenization_failures: true" in content
