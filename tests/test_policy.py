from __future__ import annotations

import pytest

from formulafence.diff import compare_snapshots
from formulafence.models import PolicyError
from formulafence.policy import evaluate_policy, parse_policy
from formulafence.workbook import load_snapshot

from .helpers import (
    change_chart_definition_material,
    change_external_data_refresh_controls,
    change_external_link_package_controls,
    change_fill_definition,
    change_filter_visibility_criterion,
    change_filter_visibility_hidden_column,
    change_font_definition,
    change_formula_cached_result,
    change_ignored_error_target,
    change_legacy_vml_control_controls,
    change_named_sheet_view_criterion,
    change_number_format_code,
    change_office_web_addin_auto_show,
    change_pivot_table_definition_material,
    change_power_pivot_data_model_payload,
    change_power_query_controls,
    change_ribbon_customization_callback,
    change_rich_text_run_color,
    change_scenario_manager_input_value,
    change_slicer_timeline_filter_material,
    change_threaded_comment_reply,
    change_what_if_data_table_input,
    change_worksheet_drawing_shape_presentation,
    change_worksheet_embedded_control_controls,
    change_xlm_macro_sheet_controls,
    change_zero_dimension_visibility_controls,
    make_chart_definition_model,
    make_conditional_formatting_model,
    make_data_validation_model,
    make_external_data_refresh_model,
    make_external_link_package_model,
    make_fill_model,
    make_filter_visibility_model,
    make_font_model,
    make_formula_cached_result_model,
    make_ignored_error_model,
    make_legacy_array_model,
    make_legacy_vml_control_model,
    make_model,
    make_named_sheet_view_model,
    make_number_format_model,
    make_office_web_addin_model,
    make_pivot_table_definition_model,
    make_power_pivot_data_model,
    make_power_query_model,
    make_protection_model,
    make_ribbon_customization_model,
    make_rich_text_run_model,
    make_scenario_manager_model,
    make_slicer_timeline_cache_model,
    make_table_model,
    make_threaded_comment_model,
    make_three_d_model,
    make_what_if_data_table_model,
    make_worksheet_drawing_shape_model,
    make_worksheet_embedded_control_model,
    make_xlm_macro_sheet_model,
    make_zero_dimension_visibility_model,
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


def test_policy_can_block_conditional_formatting_control_changes(tmp_path) -> None:
    baseline = make_conditional_formatting_model(tmp_path / "baseline.xlsx")
    candidate = make_conditional_formatting_model(tmp_path / "candidate.xlsx")

    def weaken_visual_control(workbook) -> None:
        rules = [
            rule
            for rule_group in workbook["Inputs"].conditional_formatting._cf_rules.values()
            for rule in rule_group
        ]
        rules[0].stopIfTrue = False

    rewrite(candidate, weaken_visual_control)
    report = compare_snapshots(load_snapshot(baseline), load_snapshot(candidate))
    policy = parse_policy(
        {"version": 1, "rules": {"no_conditional_formatting_changes": True}}
    )

    assert {finding.rule_id for finding in evaluate_policy(report, policy)} >= {"FFP021"}


def test_policy_can_block_protection_control_changes(tmp_path) -> None:
    baseline = make_protection_model(tmp_path / "baseline.xlsx")
    candidate = make_protection_model(tmp_path / "candidate.xlsx")

    def remove_structure_lock(workbook) -> None:
        workbook.security.lockStructure = False

    rewrite(candidate, remove_structure_lock)
    report = compare_snapshots(load_snapshot(baseline), load_snapshot(candidate))
    policy = parse_policy({"version": 1, "rules": {"no_protection_changes": True}})

    assert {finding.rule_id for finding in evaluate_policy(report, policy)} >= {"FFP022"}


def test_policy_can_block_external_data_connection_changes(tmp_path) -> None:
    baseline = make_external_data_refresh_model(tmp_path / "baseline.xlsx")
    candidate = make_external_data_refresh_model(tmp_path / "candidate.xlsx")
    change_external_data_refresh_controls(candidate)

    report = compare_snapshots(load_snapshot(baseline), load_snapshot(candidate))
    policy = parse_policy(
        {"version": 1, "rules": {"no_external_data_connection_changes": True}}
    )

    assert {finding.rule_id for finding in evaluate_policy(report, policy)} >= {"FFP023"}


def test_policy_can_block_external_link_package_changes(tmp_path) -> None:
    baseline = make_external_link_package_model(tmp_path / "baseline.xlsx")
    candidate = make_external_link_package_model(tmp_path / "candidate.xlsx")
    change_external_link_package_controls(candidate)

    report = compare_snapshots(load_snapshot(baseline), load_snapshot(candidate))
    policy = parse_policy(
        {"version": 1, "rules": {"no_external_link_package_changes": True}}
    )

    assert {finding.rule_id for finding in evaluate_policy(report, policy)} >= {"FFP025"}


def test_policy_can_block_xlm_macro_sheet_changes(tmp_path) -> None:
    baseline = make_xlm_macro_sheet_model(tmp_path / "baseline.xlsm")
    candidate = make_xlm_macro_sheet_model(tmp_path / "candidate.xlsm")
    change_xlm_macro_sheet_controls(candidate)

    report = compare_snapshots(load_snapshot(baseline), load_snapshot(candidate))
    policy = parse_policy(
        {"version": 1, "rules": {"no_xlm_macro_sheet_changes": True}}
    )

    assert {finding.rule_id for finding in evaluate_policy(report, policy)} >= {"FFP026"}


def test_policy_can_block_ribbon_customization_changes(tmp_path) -> None:
    baseline = make_ribbon_customization_model(tmp_path / "baseline.xlsx")
    candidate = make_ribbon_customization_model(tmp_path / "candidate.xlsx")
    change_ribbon_customization_callback(candidate)

    report = compare_snapshots(load_snapshot(baseline), load_snapshot(candidate))
    policy = parse_policy(
        {"version": 1, "rules": {"no_ribbon_customization_changes": True}}
    )

    assert {finding.rule_id for finding in evaluate_policy(report, policy)} >= {"FFP027"}


def test_policy_can_block_office_web_addin_changes(tmp_path) -> None:
    baseline = make_office_web_addin_model(tmp_path / "baseline.xlsx")
    candidate = make_office_web_addin_model(tmp_path / "candidate.xlsx")
    change_office_web_addin_auto_show(candidate)

    report = compare_snapshots(load_snapshot(baseline), load_snapshot(candidate))
    policy = parse_policy(
        {"version": 1, "rules": {"no_office_web_addin_changes": True}}
    )

    assert {finding.rule_id for finding in evaluate_policy(report, policy)} >= {"FFP028"}


def test_policy_can_block_chart_definition_changes(tmp_path) -> None:
    baseline = make_chart_definition_model(tmp_path / "baseline.xlsx")
    candidate = make_chart_definition_model(tmp_path / "candidate.xlsx")
    change_chart_definition_material(candidate)

    report = compare_snapshots(load_snapshot(baseline), load_snapshot(candidate))
    policy = parse_policy(
        {"version": 1, "rules": {"no_chart_definition_changes": True}}
    )

    assert {finding.rule_id for finding in evaluate_policy(report, policy)} >= {"FFP030"}


def test_policy_can_block_pivot_table_definition_changes(tmp_path) -> None:
    baseline = make_pivot_table_definition_model(tmp_path / "baseline.xlsx")
    candidate = make_pivot_table_definition_model(tmp_path / "candidate.xlsx")
    change_pivot_table_definition_material(candidate)

    report = compare_snapshots(load_snapshot(baseline), load_snapshot(candidate))
    policy = parse_policy(
        {"version": 1, "rules": {"no_pivot_table_definition_changes": True}}
    )

    assert {finding.rule_id for finding in evaluate_policy(report, policy)} >= {"FFP031"}


def test_policy_can_block_slicer_timeline_cache_changes(tmp_path) -> None:
    baseline = make_slicer_timeline_cache_model(tmp_path / "baseline.xlsx")
    candidate = make_slicer_timeline_cache_model(tmp_path / "candidate.xlsx")
    change_slicer_timeline_filter_material(candidate)

    report = compare_snapshots(load_snapshot(baseline), load_snapshot(candidate))
    policy = parse_policy(
        {"version": 1, "rules": {"no_slicer_timeline_cache_changes": True}}
    )

    assert {finding.rule_id for finding in evaluate_policy(report, policy)} >= {"FFP032"}


def test_policy_can_block_power_pivot_data_model_changes(tmp_path) -> None:
    baseline = make_power_pivot_data_model(tmp_path / "baseline.xlsx")
    candidate = make_power_pivot_data_model(tmp_path / "candidate.xlsx")
    change_power_pivot_data_model_payload(candidate)

    report = compare_snapshots(load_snapshot(baseline), load_snapshot(candidate))
    policy = parse_policy(
        {"version": 1, "rules": {"no_power_pivot_data_model_changes": True}}
    )

    assert {finding.rule_id for finding in evaluate_policy(report, policy)} >= {"FFP033"}


def test_policy_can_block_what_if_data_table_changes(tmp_path) -> None:
    baseline = make_what_if_data_table_model(tmp_path / "baseline.xlsx")
    candidate = make_what_if_data_table_model(tmp_path / "candidate.xlsx")
    change_what_if_data_table_input(candidate)

    report = compare_snapshots(load_snapshot(baseline), load_snapshot(candidate))
    policy = parse_policy(
        {"version": 1, "rules": {"no_what_if_data_table_changes": True}}
    )

    assert {finding.rule_id for finding in evaluate_policy(report, policy)} >= {"FFP034"}


def test_policy_can_block_scenario_manager_changes(tmp_path) -> None:
    baseline = make_scenario_manager_model(tmp_path / "baseline.xlsx")
    candidate = make_scenario_manager_model(tmp_path / "candidate.xlsx")
    change_scenario_manager_input_value(candidate)

    report = compare_snapshots(load_snapshot(baseline), load_snapshot(candidate))
    policy = parse_policy(
        {"version": 1, "rules": {"no_scenario_manager_changes": True}}
    )

    assert {finding.rule_id for finding in evaluate_policy(report, policy)} >= {"FFP035"}


def test_policy_can_block_filter_visibility_changes(tmp_path) -> None:
    baseline = make_filter_visibility_model(tmp_path / "baseline.xlsx")
    candidate = make_filter_visibility_model(tmp_path / "candidate.xlsx")
    column_candidate = make_filter_visibility_model(tmp_path / "column-candidate.xlsx")
    change_filter_visibility_criterion(candidate)
    change_filter_visibility_hidden_column(column_candidate)

    report = compare_snapshots(load_snapshot(baseline), load_snapshot(candidate))
    column_report = compare_snapshots(
        load_snapshot(baseline),
        load_snapshot(column_candidate),
    )
    policy = parse_policy(
        {"version": 1, "rules": {"no_filter_visibility_changes": True}}
    )

    assert {finding.rule_id for finding in evaluate_policy(report, policy)} >= {"FFP036"}
    assert {finding.rule_id for finding in evaluate_policy(column_report, policy)} >= {
        "FFP036"
    }


def test_policy_can_block_zero_dimension_visibility_changes(tmp_path) -> None:
    baseline = make_zero_dimension_visibility_model(tmp_path / "baseline.xlsx")
    candidate = make_zero_dimension_visibility_model(tmp_path / "candidate.xlsx")
    change_zero_dimension_visibility_controls(candidate)

    report = compare_snapshots(load_snapshot(baseline), load_snapshot(candidate))
    policy = parse_policy(
        {"version": 1, "rules": {"no_filter_visibility_changes": True}}
    )

    assert {finding.rule_id for finding in evaluate_policy(report, policy)} >= {
        "FFP036"
    }


def test_policy_can_block_ignored_error_changes(tmp_path) -> None:
    baseline = make_ignored_error_model(tmp_path / "baseline.xlsx")
    candidate = make_ignored_error_model(tmp_path / "candidate.xlsx")
    change_ignored_error_target(candidate)

    report = compare_snapshots(load_snapshot(baseline), load_snapshot(candidate))
    policy = parse_policy(
        {"version": 1, "rules": {"no_ignored_error_changes": True}}
    )

    assert {finding.rule_id for finding in evaluate_policy(report, policy)} >= {"FFP037"}


def test_policy_can_block_named_sheet_view_changes(tmp_path) -> None:
    baseline = make_named_sheet_view_model(tmp_path / "baseline.xlsx")
    candidate = make_named_sheet_view_model(tmp_path / "candidate.xlsx")
    change_named_sheet_view_criterion(candidate)

    report = compare_snapshots(load_snapshot(baseline), load_snapshot(candidate))
    policy = parse_policy(
        {"version": 1, "rules": {"no_named_sheet_view_changes": True}}
    )

    assert {finding.rule_id for finding in evaluate_policy(report, policy)} >= {"FFP038"}


def test_policy_can_block_number_format_changes(tmp_path) -> None:
    baseline = make_number_format_model(tmp_path / "baseline.xlsx")
    candidate = make_number_format_model(tmp_path / "candidate.xlsx")
    change_number_format_code(candidate)

    report = compare_snapshots(load_snapshot(baseline), load_snapshot(candidate))
    policy = parse_policy(
        {"version": 1, "rules": {"no_number_format_changes": True}}
    )

    assert {finding.rule_id for finding in evaluate_policy(report, policy)} >= {"FFP039"}


def test_policy_can_block_cell_font_changes(tmp_path) -> None:
    baseline = make_font_model(tmp_path / "baseline.xlsx")
    candidate = make_font_model(tmp_path / "candidate.xlsx")
    change_font_definition(candidate)

    report = compare_snapshots(load_snapshot(baseline), load_snapshot(candidate))
    policy = parse_policy(
        {"version": 1, "rules": {"no_cell_font_changes": True}}
    )

    assert {finding.rule_id for finding in evaluate_policy(report, policy)} >= {"FFP040"}


def test_policy_can_block_cell_fill_changes(tmp_path) -> None:
    baseline = make_fill_model(tmp_path / "baseline.xlsx")
    candidate = make_fill_model(tmp_path / "candidate.xlsx")
    change_fill_definition(candidate)

    report = compare_snapshots(load_snapshot(baseline), load_snapshot(candidate))
    policy = parse_policy(
        {"version": 1, "rules": {"no_cell_fill_changes": True}}
    )

    assert {finding.rule_id for finding in evaluate_policy(report, policy)} >= {"FFP041"}


def test_policy_can_block_formula_cached_result_changes(tmp_path) -> None:
    baseline = make_formula_cached_result_model(tmp_path / "baseline.xlsx")
    candidate = make_formula_cached_result_model(tmp_path / "candidate.xlsx")
    change_formula_cached_result(candidate)

    report = compare_snapshots(load_snapshot(baseline), load_snapshot(candidate))
    policy = parse_policy(
        {"version": 1, "rules": {"no_formula_cached_result_changes": True}}
    )

    assert {finding.rule_id for finding in evaluate_policy(report, policy)} >= {"FFP042"}


def test_policy_can_block_rich_text_run_changes(tmp_path) -> None:
    baseline = make_rich_text_run_model(tmp_path / "baseline.xlsx")
    candidate = make_rich_text_run_model(tmp_path / "candidate.xlsx")
    change_rich_text_run_color(candidate)

    report = compare_snapshots(load_snapshot(baseline), load_snapshot(candidate))
    policy = parse_policy(
        {"version": 1, "rules": {"no_rich_text_run_changes": True}}
    )

    assert {finding.rule_id for finding in evaluate_policy(report, policy)} >= {"FFP043"}


def test_policy_can_block_threaded_comment_changes(tmp_path) -> None:
    baseline = make_threaded_comment_model(tmp_path / "baseline.xlsx")
    candidate = make_threaded_comment_model(tmp_path / "candidate.xlsx")
    change_threaded_comment_reply(candidate)

    report = compare_snapshots(load_snapshot(baseline), load_snapshot(candidate))
    policy = parse_policy(
        {"version": 1, "rules": {"no_threaded_comment_changes": True}}
    )

    assert {finding.rule_id for finding in evaluate_policy(report, policy)} >= {"FFP045"}


def test_policy_can_block_worksheet_drawing_shape_changes(tmp_path) -> None:
    baseline = make_worksheet_drawing_shape_model(tmp_path / "baseline.xlsx")
    candidate = make_worksheet_drawing_shape_model(tmp_path / "candidate.xlsx")
    change_worksheet_drawing_shape_presentation(candidate)

    report = compare_snapshots(load_snapshot(baseline), load_snapshot(candidate))
    policy = parse_policy(
        {"version": 1, "rules": {"no_worksheet_drawing_shape_changes": True}}
    )

    assert {finding.rule_id for finding in evaluate_policy(report, policy)} >= {"FFP044"}


def test_policy_can_block_worksheet_embedded_control_changes(tmp_path) -> None:
    baseline = make_worksheet_embedded_control_model(tmp_path / "baseline.xlsx")
    candidate = make_worksheet_embedded_control_model(tmp_path / "candidate.xlsx")
    change_worksheet_embedded_control_controls(candidate)

    report = compare_snapshots(load_snapshot(baseline), load_snapshot(candidate))
    policy = parse_policy(
        {
            "version": 1,
            "rules": {"no_worksheet_embedded_control_changes": True},
        }
    )

    assert {finding.rule_id for finding in evaluate_policy(report, policy)} >= {"FFP029"}


def test_policy_can_block_legacy_vml_control_changes(tmp_path) -> None:
    baseline = make_legacy_vml_control_model(tmp_path / "baseline.xlsx")
    candidate = make_legacy_vml_control_model(tmp_path / "candidate.xlsx")
    change_legacy_vml_control_controls(candidate)

    report = compare_snapshots(load_snapshot(baseline), load_snapshot(candidate))
    policy = parse_policy(
        {
            "version": 1,
            "rules": {"no_worksheet_embedded_control_changes": True},
        }
    )

    assert {finding.rule_id for finding in evaluate_policy(report, policy)} >= {"FFP029"}


def test_policy_can_block_power_query_changes(tmp_path) -> None:
    baseline = make_power_query_model(tmp_path / "baseline.xlsx")
    candidate = make_power_query_model(tmp_path / "candidate.xlsx")
    change_power_query_controls(candidate)

    report = compare_snapshots(load_snapshot(baseline), load_snapshot(candidate))
    policy = parse_policy({"version": 1, "rules": {"no_power_query_changes": True}})

    assert {finding.rule_id for finding in evaluate_policy(report, policy)} >= {"FFP024"}


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
