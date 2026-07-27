from __future__ import annotations

import json

from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation

from formulafence.cli import main

from .helpers import (
    change_formula_dde_link_input,
    change_formula_external_action_input,
    change_formula_external_action_target,
    make_formula_dde_link_model,
    make_formula_external_action_model,
    make_model,
    rewrite,
)


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


def test_cli_can_redact_external_workbook_link_material_from_shared_reports(tmp_path) -> None:
    baseline_marker = "PRIVATE-SHARED-BASELINE"
    candidate_marker = "PRIVATE-SHARED-CANDIDATE"

    def add_external_link_surfaces(workbook, marker: str) -> None:
        source = f"C:\\{marker}\\[Source.xlsx]Inputs"
        workbook["Model"]["D2"] = f"='{source}'!$B$2"
        workbook.defined_names.add(
            DefinedName("ExternalNamedLimit", attr_text=f"'{source}'!$C$2")
        )
        validation = DataValidation(
            type="whole",
            operator="greaterThan",
            formula1=f"='{source}'!$D$2",
        )
        validation.add("E2")
        workbook["Model"].add_data_validation(validation)

    def external_model(path, marker: str):
        workbook = make_model(path)
        rewrite(workbook, lambda model: add_external_link_surfaces(model, marker))
        return workbook

    baseline = external_model(tmp_path / "baseline.xlsx", baseline_marker)
    candidate = external_model(tmp_path / "candidate.xlsx", candidate_marker)
    default_json = tmp_path / "default.json"
    assert (
        main(
            [
                "diff",
                str(baseline),
                str(candidate),
                "--format",
                "json",
                "--output",
                str(default_json),
            ]
        )
        == 0
    )
    default_rendered = default_json.read_text(encoding="utf-8")
    assert baseline_marker in default_rendered
    assert candidate_marker in default_rendered

    for report_format, suffix in (("json", "json"), ("markdown", "md"), ("sarif", "sarif")):
        output = tmp_path / f"redacted.{suffix}"
        assert (
            main(
                [
                    "diff",
                    str(baseline),
                    str(candidate),
                    "--format",
                    report_format,
                    "--redact-external-workbook-links",
                    "--output",
                    str(output),
                ]
            )
            == 0
        )
        rendered = output.read_text(encoding="utf-8")
        assert baseline_marker not in rendered
        assert candidate_marker not in rendered
        if report_format == "markdown":
            assert "External-workbook link material:** redacted for sharing" in rendered
        else:
            assert "external-workbook link material redacted" in rendered

    policy = tmp_path / "formulafence.yml"
    policy.write_text(
        "version: 1\nrules:\n  no_external_workbook_link_surface_changes: true\n",
        encoding="utf-8",
    )
    policy_output = tmp_path / "redacted-policy.sarif"
    assert (
        main(
            [
                "check",
                str(baseline),
                str(candidate),
                "--policy",
                str(policy),
                "--format",
                "sarif",
                "--redact-external-workbook-links",
                "--output",
                str(policy_output),
            ]
        )
        == 1
    )
    policy_rendered = policy_output.read_text(encoding="utf-8")
    assert "FFP081" in policy_rendered
    assert baseline_marker not in policy_rendered
    assert candidate_marker not in policy_rendered

    baseline_directory = tmp_path / "baseline-portfolio"
    candidate_directory = tmp_path / "candidate-portfolio"
    baseline_directory.mkdir()
    candidate_directory.mkdir()
    external_model(baseline_directory / "model.xlsx", baseline_marker)
    external_model(candidate_directory / "model.xlsx", candidate_marker)
    for report_format, suffix in (("json", "json"), ("markdown", "md"), ("sarif", "sarif")):
        output = tmp_path / f"redacted-portfolio.{suffix}"
        assert (
            main(
                [
                    "portfolio",
                    str(baseline_directory),
                    str(candidate_directory),
                    "--format",
                    report_format,
                    "--redact-external-workbook-links",
                    "--output",
                    str(output),
                ]
            )
            == 0
        )
        rendered = output.read_text(encoding="utf-8")
        assert baseline_marker not in rendered
        assert candidate_marker not in rendered
        if report_format == "markdown":
            assert "External-workbook link material:** redacted for sharing" in rendered
        else:
            assert "external-workbook link material redacted" in rendered


def test_cli_can_redact_formula_external_action_and_dde_material_from_shared_reports(
    tmp_path,
) -> None:
    action_baseline_marker = "PRIVATE-LINK-BASELINE"
    action_candidate_marker = "PRIVATE-LINK-CANDIDATE"
    baseline = make_formula_external_action_model(tmp_path / "baseline.xlsx")
    candidate = make_formula_external_action_model(tmp_path / "candidate.xlsx")
    change_formula_external_action_target(candidate)

    default_json = tmp_path / "default.json"
    assert (
        main(
            [
                "diff",
                str(baseline),
                str(candidate),
                "--format",
                "json",
                "--output",
                str(default_json),
            ]
        )
        == 0
    )
    default_rendered = default_json.read_text(encoding="utf-8")
    assert action_baseline_marker in default_rendered
    assert action_candidate_marker in default_rendered

    for report_format, suffix in (("json", "json"), ("markdown", "md"), ("sarif", "sarif")):
        output = tmp_path / f"action-redacted.{suffix}"
        assert (
            main(
                [
                    "diff",
                    str(baseline),
                    str(candidate),
                    "--format",
                    report_format,
                    "--redact-formula-external-actions",
                    "--output",
                    str(output),
                ]
            )
            == 0
        )
        rendered = output.read_text(encoding="utf-8")
        assert action_baseline_marker not in rendered
        assert action_candidate_marker not in rendered
        assert "FF064" in rendered
        if report_format == "markdown":
            assert "Formula external-action / DDE material:** redacted for sharing" in rendered
        elif report_format == "json":
            assert "formula external-action material redacted" in rendered

    action_input_baseline = make_formula_external_action_model(
        tmp_path / "action-input-baseline.xlsx"
    )
    action_input_candidate = make_formula_external_action_model(
        tmp_path / "action-input-candidate.xlsx"
    )
    change_formula_external_action_input(action_input_candidate)
    action_input_output = tmp_path / "action-input-redacted.json"
    assert (
        main(
            [
                "diff",
                str(action_input_baseline),
                str(action_input_candidate),
                "--format",
                "json",
                "--redact-formula-external-actions",
                "--output",
                str(action_input_output),
            ]
        )
        == 0
    )
    action_input_rendered = action_input_output.read_text(encoding="utf-8")
    assert "PRIVATE-REFERENCED-LINK-BASELINE" not in action_input_rendered
    assert "PRIVATE-REFERENCED-LINK-CANDIDATE" not in action_input_rendered
    assert "formula external-action material redacted" in action_input_rendered

    dde_baseline = make_formula_dde_link_model(tmp_path / "dde-baseline.xlsx")
    dde_candidate = make_formula_dde_link_model(tmp_path / "dde-candidate.xlsx")
    change_formula_dde_link_input(dde_candidate)
    dde_output = tmp_path / "dde-redacted.json"
    assert (
        main(
            [
                "diff",
                str(dde_baseline),
                str(dde_candidate),
                "--format",
                "json",
                "--redact-formula-external-actions",
                "--output",
                str(dde_output),
            ]
        )
        == 0
    )
    dde_rendered = dde_output.read_text(encoding="utf-8")
    assert "PRIVATE-DDE-INPUT-BASELINE" not in dde_rendered
    assert "PRIVATE-DDE-INPUT-CANDIDATE" not in dde_rendered
    assert "FF074" in dde_rendered

    policy = tmp_path / "formulafence.yml"
    policy.write_text(
        "version: 1\nrules:\n  no_formula_external_action_changes: true\n",
        encoding="utf-8",
    )
    policy_output = tmp_path / "policy-redacted.json"
    assert (
        main(
            [
                "check",
                str(baseline),
                str(candidate),
                "--policy",
                str(policy),
                "--format",
                "json",
                "--redact-formula-external-actions",
                "--output",
                str(policy_output),
            ]
        )
        == 1
    )
    policy_rendered = policy_output.read_text(encoding="utf-8")
    assert "FF064" in policy_rendered
    assert "FFP064" in policy_rendered
    assert action_baseline_marker not in policy_rendered
    assert action_candidate_marker not in policy_rendered

    baseline_directory = tmp_path / "baseline-portfolio"
    candidate_directory = tmp_path / "candidate-portfolio"
    baseline_directory.mkdir()
    candidate_directory.mkdir()
    make_formula_external_action_model(baseline_directory / "model.xlsx")
    portfolio_candidate = make_formula_external_action_model(
        candidate_directory / "model.xlsx"
    )
    change_formula_external_action_target(portfolio_candidate)
    portfolio_output = tmp_path / "portfolio-redacted.json"
    assert (
        main(
            [
                "portfolio",
                str(baseline_directory),
                str(candidate_directory),
                "--format",
                "json",
                "--redact-formula-external-actions",
                "--output",
                str(portfolio_output),
            ]
        )
        == 0
    )
    portfolio_rendered = portfolio_output.read_text(encoding="utf-8")
    assert action_baseline_marker not in portfolio_rendered
    assert action_candidate_marker not in portfolio_rendered
    assert "formula external-action material redacted" in portfolio_rendered


def test_cli_refuses_to_overwrite_an_input_workbook(tmp_path) -> None:
    baseline = make_model(tmp_path / "baseline.xlsx")
    candidate = make_model(tmp_path / "candidate.xlsx")
    original = baseline.read_bytes()

    assert (
        main(
            [
                "diff",
                str(baseline),
                str(candidate),
                "--output",
                str(baseline),
            ]
        )
        == 2
    )
    assert baseline.read_bytes() == original


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
    assert "no_external_workbook_link_surface_changes: true" in content
    assert "no_external_relationship_changes: true" in content
    assert "no_formula_external_action_changes: true" in content
    assert "no_formula_dde_link_changes: true" in content
    assert "no_python_in_excel_changes: true" in content
    assert "no_office_custom_function_changes: true" in content
    assert "no_unqualified_runtime_function_changes: true" in content
    assert "no_worksheet_code_resource_registration_changes: true" in content
    assert "no_formula_defined_xlm_registration_changes: true" in content
    assert "no_formula_defined_xlm_evaluation_changes: true" in content
    assert "no_formula_defined_xlm_action_changes: true" in content
    assert "no_formula_defined_xlm_get_cell_changes: true" in content
    assert "no_formula_defined_xlm_environment_information_changes: true" in content
    assert "no_formula_environment_information_changes: true" in content
    assert "no_xlm_macro_sheet_changes: true" in content
    assert "no_xlm_automatic_macro_binding_changes: true" in content
    assert "no_ribbon_customization_changes: true" in content
    assert "no_office_web_addin_changes: true" in content
    assert "no_pivot_table_definition_changes: true" in content
    assert "no_slicer_timeline_cache_changes: true" in content
    assert "no_power_pivot_data_model_changes: true" in content
    assert "no_what_if_data_table_changes: true" in content
    assert "no_scenario_manager_changes: true" in content
    assert "no_filter_visibility_changes: true" in content
    assert "no_ignored_error_changes: true" in content
    assert "no_named_sheet_view_changes: true" in content
    assert "no_number_format_changes: true" in content
    assert "no_cell_font_changes: true" in content
    assert "no_cell_fill_changes: true" in content
    assert "no_workbook_theme_changes: true" in content
    assert "no_cell_alignment_changes: true" in content
    assert "no_cell_border_changes: true" in content
    assert "no_worksheet_dimension_changes: true" in content
    assert "no_worksheet_display_control_changes: true" in content
    assert "no_worksheet_print_layout_changes: true" in content
    assert "no_formula_cached_result_changes: true" in content
    assert "no_rich_text_run_changes: true" in content
    assert "no_cell_hyperlink_changes: true" in content
    assert "no_worksheet_sparkline_changes: true" in content
    assert "no_xml_mapping_changes: true" in content
    assert "no_digital_signature_changes: true" in content
    assert "no_rich_data_changes: true" in content
    assert "no_custom_data_store_changes: true" in content
    assert "no_legacy_comment_changes: true" in content
    assert "no_threaded_comment_changes: true" in content
    assert "no_worksheet_drawing_shape_changes: true" in content
    assert "no_worksheet_image_changes: true" in content
    assert "no_worksheet_embedded_control_changes: true" in content
    assert "no_power_query_changes: true" in content
    assert "no_portfolio_membership_changes: true" in content
    assert "no_cross_workbook_impacts: true" in content
    assert "no_new_tokenization_failures: true" in content
