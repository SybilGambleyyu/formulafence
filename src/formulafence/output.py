"""Deterministic JSON, Markdown, and SARIF renderers for review artifacts."""

from __future__ import annotations

import json
from collections.abc import Iterable
from typing import Any

from formulafence import __version__
from formulafence.models import DiffReport, Finding, display_location
from formulafence.portfolio import PortfolioReport


def as_json(payload: dict[str, Any]) -> str:
    return json.dumps(payload, indent=2, sort_keys=True, ensure_ascii=False) + "\n"


def _markdown_escape(value: object) -> str:
    return str(value).replace("|", "\\|").replace("\n", "<br>")


def _credential_summary(credential: dict[str, Any]) -> str:
    """Render safe protection-verifier metadata without credential material."""
    if not credential["configured"]:
        return "none"
    kinds: list[str] = []
    if credential["has_legacy_verifier"]:
        kinds.append("legacy verifier")
    if credential["has_modern_verifier"]:
        algorithm = credential["algorithm"] or "modern"
        kinds.append(f"{algorithm} verifier")
    if credential["spin_count"] is not None:
        kinds.append(f"{credential['spin_count']} iterations")
    return "; ".join(kinds)


def profile_to_markdown(profile: dict[str, Any]) -> str:
    workbook = profile["workbook"]
    lines = [
        "# FormulaFence workbook profile",
        "",
        f"- **Workbook:** `{workbook['path']}`",
        f"- **SHA-256:** `{workbook['sha256']}`",
        f"- **Sheets:** {workbook['sheet_count']}",
        f"- **Non-empty cells:** {workbook['nonempty_cells']}",
        f"- **Formula cells:** {workbook['formula_cells']}",
        f"- **Tables:** {workbook['table_count']}",
        f"- **Data-validation rules:** {workbook['data_validation_rules']}",
        (
            "- **Data-validation target ranges:** "
            f"{workbook['data_validation_target_ranges']}"
        ),
        (
            "- **Conditional-formatting rules:** "
            f"{workbook['conditional_formatting_rules']}"
        ),
        (
            "- **Conditional-formatting target ranges:** "
            f"{workbook['conditional_formatting_target_ranges']}"
        ),
        (
            "- **Conditional-formatting extension fragments:** "
            f"{workbook['conditional_formatting_extensions']}"
        ),
        (
            "- **Workbook protection active:** "
            f"{'yes' if workbook['workbook_protection_enabled'] else 'no'}"
        ),
        f"- **Sheet-protection declarations:** {workbook['sheet_protection_controls']}",
        f"- **Protected sheets:** {workbook['protected_sheet_count']}",
        f"- **Protected ranges:** {workbook['protected_range_count']}",
        (
            "- **Protected-range target references:** "
            f"{workbook['protected_range_target_ranges']}"
        ),
        (
            "- **Direct cell-protection assignments:** "
            f"{workbook['cell_protection_assignment_count']}"
        ),
        f"- **External-data connections:** {workbook['external_data_connection_count']}",
        f"- **External-link package parts:** {workbook['external_link_package_count']}",
        (
            "- **Static external-workbook link surfaces / endpoints:** "
            f"{workbook['external_workbook_link_surface_count']} / "
            f"{workbook['external_workbook_link_surface_reference_count']}"
        ),
        f"- **DDE links:** {workbook['dde_link_count']}",
        f"- **OLE links:** {workbook['ole_link_count']}",
        (
            "- **Package external relationships (hyperlink / image):** "
            f"{workbook['package_external_relationship_count']} "
            f"({workbook['package_external_hyperlink_relationship_count']} / "
            f"{workbook['package_external_image_relationship_count']})"
        ),
        (
            "- **Scenario Manager sheets / scenarios / stored inputs:** "
            f"{workbook['scenario_manager_sheet_count']} / "
            f"{workbook['scenario_manager_scenario_count']} / "
            f"{workbook['scenario_manager_input_cell_count']}"
        ),
        (
            "- **Filter declarations / hidden rows / zero-height rows / hidden columns / "
            "zero-width columns:** "
            f"{workbook['filter_visibility_auto_filter_count']} / "
            f"{workbook['filter_visibility_hidden_row_count']} / "
            f"{workbook['filter_visibility_zero_height_row_count']} / "
            f"{workbook['filter_visibility_hidden_column_count']} / "
            f"{workbook['filter_visibility_zero_width_column_count']}"
        ),
        (
            "- **Cell number-format controls / custom assignments:** "
            f"{workbook['number_format_assignment_count']} / "
            f"{workbook['number_format_custom_assignment_count']}"
        ),
        f"- **Cell font controls:** {workbook['font_assignment_count']}",
        f"- **Cell fill controls:** {workbook['fill_assignment_count']}",
        f"- **Cell alignment controls:** {workbook['alignment_assignment_count']}",
        (
            "- **Workbook theme parts / image parts:** "
            f"{workbook['workbook_theme_part_count']} / "
            f"{workbook['workbook_theme_image_part_count']}"
        ),
        (
            "- **Python-in-Excel parts / PY formula cells / stored scripts:** "
            f"{workbook['python_in_excel_part_count']} / "
            f"{workbook['python_in_excel_formula_cell_count']} / "
            f"{workbook['python_in_excel_script_count']}"
        ),
        (
            "- **Namespaced custom-function formula cells / calls / namespaces:** "
            f"{workbook['namespaced_custom_function_formula_cell_count']} / "
            f"{workbook['namespaced_custom_function_call_count']} / "
            f"{workbook['namespaced_custom_function_namespace_count']}"
        ),
        (
            "- **Unqualified runtime-function formula cells / calls / named definitions:** "
            f"{workbook['unqualified_runtime_function_formula_cell_count']} / "
            f"{workbook['unqualified_runtime_function_call_count']} / "
            f"{workbook['unqualified_runtime_function_defined_name_count']}"
        ),
        f"- **XLM macro-sheet parts:** {workbook['xlm_macro_sheet_count']}",
        (
            "- **XLM automatic-macro bindings:** "
            f"{workbook['xlm_automatic_macro_binding_count']}"
        ),
        (
            "- **XLM macro-sheet formula cells:** "
            f"{workbook['xlm_macro_formula_cell_count']}"
        ),
        (
            "- **Fingerprinted XLM related parts:** "
            f"{workbook['xlm_related_part_payload_count']}"
        ),
        (
            "- **RibbonX customization parts:** "
            f"{workbook['ribbon_customization_part_count']}"
        ),
        (
            "- **RibbonX callback attributes:** "
            f"{workbook['ribbon_callback_attribute_count']}"
        ),
        (
            "- **Office Web Add-in task-pane parts:** "
            f"{workbook['office_web_addin_taskpane_part_count']}"
        ),
        (
            "- **Office Web Add-in definition parts:** "
            f"{workbook['office_web_addin_web_extension_part_count']}"
        ),
        (
            "- **Office Web Add-ins requesting auto-show:** "
            f"{workbook['office_web_addin_auto_show_taskpane_count']}"
        ),
        (
            "- **Office Web Add-in worksheet bindings:** "
            f"{workbook['office_web_addin_worksheet_binding_count']}"
        ),
        (
            "- **Office Web Add-in in-content references:** "
            f"{workbook['office_web_addin_in_content_reference_count']}"
        ),
        f"- **PivotTable host sheets:** {workbook['pivot_table_sheet_count']}",
        f"- **PivotTable parts:** {workbook['pivot_table_part_count']}",
        (
            "- **Pivot cache-definition parts:** "
            f"{workbook['pivot_cache_definition_part_count']}"
        ),
        f"- **Declared PivotTable cache records:** {workbook['pivot_cache_record_count']}",
        f"- **Chart host sheets:** {workbook['chart_host_sheet_count']}",
        f"- **Chart parts:** {workbook['chart_part_count']}",
        f"- **Office 2016+ ChartEx parts:** {workbook['chart_ex_part_count']}",
        f"- **Cached chart data points:** {workbook['chart_cached_data_point_count']}",
        (
            "- **Worksheet DrawingML shapes / text-bearing shapes / graphic frames / SmartArt:** "
            f"{workbook['worksheet_drawing_shape_count']} / "
            f"{workbook['worksheet_drawing_text_shape_count']} / "
            f"{workbook['worksheet_drawing_graphic_frame_count']} / "
            f"{workbook['worksheet_drawing_diagram_frame_count']}"
        ),
        (
            "- **Worksheet images (anchored / backgrounds / header-footer):** "
            f"{workbook['worksheet_anchored_picture_count']} / "
            f"{workbook['worksheet_background_image_count']} / "
            f"{workbook['worksheet_header_footer_image_count']}"
        ),
        (
            "- **Worksheet control-bearing sheets:** "
            f"{workbook['worksheet_embedded_control_sheet_count']}"
        ),
        (
            "- **Worksheet ActiveX parts:** "
            f"{workbook['worksheet_active_x_part_count']}"
        ),
        (
            "- **Worksheet legacy VML drawing parts:** "
            f"{workbook['worksheet_legacy_vml_drawing_part_count']}"
        ),
        (
            "- **Worksheet legacy VML controls:** "
            f"{workbook['worksheet_legacy_vml_control_count']}"
        ),
        (
            "- **Worksheet OLE objects:** "
            f"{workbook['worksheet_ole_object_count']}"
        ),
        (
            "- **Connections refreshing on open:** "
            f"{workbook['external_data_connections_refresh_on_load']}"
        ),
        (
            "- **Query-table refresh controls:** "
            f"{workbook['query_table_refresh_control_count']}"
        ),
        (
            "- **Query tables refreshing on open:** "
            f"{workbook['query_tables_refresh_on_load']}"
        ),
        (
            "- **Pivot-cache refresh controls:** "
            f"{workbook['pivot_cache_refresh_control_count']}"
        ),
        (
            "- **Pivot caches refreshing on open:** "
            f"{workbook['pivot_caches_refresh_on_load']}"
        ),
        f"- **Power Query Data Mashup parts:** {workbook['power_query_mashup_count']}",
        (
            "- **Power Query formula documents:** "
            f"{workbook['power_query_formula_document_count']}"
        ),
        (
            "- **Power Query metadata items:** "
            f"{workbook['power_query_metadata_item_count']}"
        ),
        (
            "- **Custom XML state-store parts:** "
            f"{workbook['custom_xml_part_count']}"
        ),
        (
            "- **Custom binary-data state-store parts:** "
            f"{workbook['custom_data_part_count']}"
        ),
        (
            "- **Custom document properties:** "
            f"{workbook['document_custom_property_count']}"
        ),
        f"- **3-D reference formulas:** {workbook['three_d_reference_cells']}",
        f"- **Spill-reference formulas:** {workbook['spill_reference_cells']}",
        (
            "- **Implicit-intersection formulas:** "
            f"{workbook['implicit_intersection_cells']}"
        ),
        f"- **Legacy CSE array formulas:** {workbook['legacy_array_formula_cells']}",
        (
            "- **Fixed CSE output ranges:** "
            f"{workbook['legacy_array_formula_output_ranges']}"
        ),
        f"- **Dynamic array formulas:** {workbook['dynamic_array_formula_cells']}",
        (
            "- **Observed dynamic-array output ranges:** "
            f"{workbook['dynamic_array_observed_output_ranges']}"
        ),
        (
            "- **Formulas reading observed dynamic output members:** "
            f"{workbook['dynamic_array_output_reference_cells']}"
        ),
        (
            "- **Unclassified array formulas:** "
            f"{workbook['unclassified_array_formula_cells']}"
        ),
        f"- **Formula tokenizer failures:** {workbook['tokenization_failure_cells']}",
        f"- **VBA payload:** {'present' if workbook['has_vba'] else 'absent'}",
        "",
        "## Sheets",
        "",
        "| Sheet | State | Non-empty cells | Formula cells | Used range |",
        "| --- | --- | ---: | ---: | --- |",
    ]
    for sheet in profile["sheets"]:
        safe_sheet = {key: _markdown_escape(value) for key, value in sheet.items()}
        lines.append(
            (
                "| {title} | {state} | {nonempty_cells} | {formula_cells} | "
                "{max_column} × {max_row} |"
            ).format(**safe_sheet)
        )
    if profile["tables"]:
        lines.extend(
            [
                "",
                "## Excel tables",
                "",
                "| Table | Sheet | Range | Columns |",
                "| --- | --- | --- | --- |",
            ]
        )
        for table in profile["tables"]:
            lines.append(
                "| {name} | {sheet} | {ref} | {columns} |".format(
                    name=_markdown_escape(table["name"]),
                    sheet=_markdown_escape(table["sheet"]),
                    ref=_markdown_escape(table["ref"]),
                    columns=_markdown_escape(", ".join(table["columns"])),
                )
            )
    if profile["data_validations"]:
        lines.extend(
            [
                "",
                "## Data-validation controls",
                "",
                "| Sheet | Applies to | Rule | Criteria | Behavior |",
                "| --- | --- | --- | ---: | --- |",
            ]
        )
        for validation in profile["data_validations"]:
            rule = validation["type"]
            if validation["type"] not in {"list", "custom", "none"}:
                rule = f"{validation['type']} / {validation['operator']}"
            behavior: list[str] = []
            if validation["allow_blank"]:
                behavior.append("blanks allowed")
            if validation["dropdown_hidden"]:
                behavior.append("dropdown hidden")
            if validation["prompts_disabled"]:
                behavior.append("worksheet prompts disabled")
            elif validation["show_input_message"]:
                behavior.append("input prompt")
            if validation["show_error_message"]:
                behavior.append(f"{validation['error_style']} alert")
            if validation["has_input_prompt_text"] and not validation["prompts_disabled"]:
                behavior.append("prompt text")
            if validation["has_error_alert_text"]:
                behavior.append("alert text")
            lines.append(
                "| {sheet} | {ranges} | {rule} | "
                "{criteria_count} | {behavior} |".format(
                    sheet=_markdown_escape(validation["sheet"]),
                    ranges=_markdown_escape(
                        ", ".join(validation["ranges"])
                    ),
                    rule=_markdown_escape(rule),
                    criteria_count=validation["criteria_count"],
                    behavior=_markdown_escape(
                        "; ".join(behavior) if behavior else "default behavior"
                    ),
                )
            )
        lines.append(
            "Profiles omit validation formulas and prompt/error text; the full local "
            "before/after settings appear only in a change report."
        )
    if profile["conditional_formatting"]:
        lines.extend(
            [
                "",
                "## Conditional-formatting controls",
                "",
                "| Sheet | Applies to | Priority | Rule | Formula(s) | Behavior | Formatting |",
                "| --- | --- | ---: | --- | ---: | --- | --- |",
            ]
        )
        for rule in profile["conditional_formatting"]:
            behavior: list[str] = []
            if rule["stop_if_true"]:
                behavior.append("stop if true")
            if rule["type"] == "top10":
                rank = rule["rank"] if rule["rank"] is not None else "unspecified"
                direction = "bottom" if rule["bottom"] else "top"
                qualifier = "%" if rule["percent"] else ""
                behavior.append(f"{direction} {rank}{qualifier}")
            if rule["type"] == "aboveAverage":
                direction = "above" if rule["above_average"] else "below"
                behavior.append(f"{direction} average")
                if rule["equal_average"]:
                    behavior.append("includes average")
                if rule["std_dev"] is not None:
                    behavior.append(f"{rule['std_dev']} standard deviations")
            if rule["time_period"] is not None:
                behavior.append(rule["time_period"])
            if rule["has_text_criterion"]:
                behavior.append("text criterion")
            if rule["extension_count"]:
                behavior.append(f"{rule['extension_count']} extension fragment(s)")
            lines.append(
                "| {sheet} | {ranges} | {priority} | {rule_type} | {formula_count} | "
                "{behavior} | {formatting} |".format(
                    sheet=_markdown_escape(rule["sheet"]),
                    ranges=_markdown_escape(", ".join(rule["ranges"])),
                    priority=rule["priority"],
                    rule_type=_markdown_escape(rule["type"]),
                    formula_count=rule["formula_count"],
                    behavior=_markdown_escape(
                        "; ".join(behavior) if behavior else "default behavior"
                    ),
                    formatting=_markdown_escape(
                        ", ".join(rule["formatting"]) if rule["formatting"] else "none"
                    ),
                )
            )
        lines.append(
            "Profiles omit conditional-format formulas, text criteria, and raw style "
            "or extension XML; full local before/after evidence appears only in a change report."
        )
    if profile["conditional_formatting_extensions"]:
        lines.extend(
            [
                "",
                "## Conditional-formatting extension coverage",
                "",
                "| Sheet | OOXML extension |",
                "| --- | --- |",
            ]
        )
        for extension in profile["conditional_formatting_extensions"]:
            lines.append(
                "| {sheet} | {element} |".format(
                    sheet=_markdown_escape(extension["sheet"]),
                    element=_markdown_escape(extension["element"]),
                )
            )
        lines.append(
            "Extension structure is compared locally but intentionally omitted from the profile."
        )
    if profile["workbook_protection"] is not None:
        protection = profile["workbook_protection"]
        locked_operations = [
            label
            for key, label in (
                ("lock_structure", "workbook structure"),
                ("lock_windows", "workbook windows"),
                ("lock_revision", "revisions"),
            )
            if protection[key]
        ]
        lines.extend(
            [
                "",
                "## Workbook protection",
                "",
                f"- **Active:** {'yes' if protection['enabled'] else 'no'}",
                (
                    "- **Locked operations:** "
                    + (", ".join(locked_operations) if locked_operations else "none")
                ),
                (
                    "- **Workbook verifier:** "
                    + _credential_summary(protection["workbook_credential"])
                ),
                (
                    "- **Revision verifier:** "
                    + _credential_summary(protection["revisions_credential"])
                ),
            ]
        )
        if protection["opaque_metadata"]["present"]:
            lines.append(
                "- **Unmodelled protection metadata:** "
                f"{protection['opaque_metadata']['count']} item(s)"
            )
    if profile["sheet_protections"]:
        lines.extend(
            [
                "",
                "## Sheet protection controls",
                "",
                "| Sheet | Type | Active | Locked actions | Verifier |",
                "| --- | --- | --- | --- | --- |",
            ]
        )
        for protection in profile["sheet_protections"]:
            lines.append(
                "| {sheet} | {sheet_type} | {enabled} | {actions} | {credential} |".format(
                    sheet=_markdown_escape(protection["sheet"]),
                    sheet_type=_markdown_escape(protection["sheet_type"]),
                    enabled="yes" if protection["enabled"] else "no",
                    actions=_markdown_escape(
                        ", ".join(protection["locked_actions"])
                        if protection["locked_actions"]
                        else "none"
                    ),
                    credential=_markdown_escape(
                        _credential_summary(protection["credential"])
                    ),
                )
            )
        lines.append(
            "Password verifier values, hashes, salts, and opaque protection XML are not included."
        )
    if profile["protected_ranges"]:
        lines.extend(
            [
                "",
                "## Protected ranges",
                "",
                "| Sheet | Applies to | Named | Verifier | Identity restriction |",
                "| --- | --- | --- | --- | --- |",
            ]
        )
        for protected_range in profile["protected_ranges"]:
            lines.append(
                "| {sheet} | {ranges} | {has_name} | {credential} | {descriptor} |".format(
                    sheet=_markdown_escape(protected_range["sheet"]),
                    ranges=_markdown_escape(
                        ", ".join(protected_range["ranges"])
                        if protected_range["ranges"]
                        else "missing target"
                    ),
                    has_name="yes" if protected_range["has_name"] else "no",
                    credential=_markdown_escape(
                        _credential_summary(protected_range["credential"])
                    ),
                    descriptor=(
                        "present"
                        if protected_range["has_security_descriptor"]
                        else "none"
                    ),
                )
            )
        lines.append(
            "Range names and security-descriptor contents are intentionally omitted from profiles."
        )
    if profile["cell_protection_default"] is not None:
        default = profile["cell_protection_default"]
        lines.extend(
            [
                "",
                "## Cell protection defaults",
                "",
                (
                    "- **Locked by default:** "
                    + ("yes" if default["locked"] else "no")
                ),
                (
                    "- **Formulas hidden by default:** "
                    + ("yes" if default["hidden"] else "no")
                ),
            ]
        )
    if profile["cell_protection_assignments"]:
        lines.extend(
            [
                "",
                "## Direct cell-protection assignments",
                "",
                "| Sheet | Scope | Target | Locked | Formulas hidden |",
                "| --- | --- | --- | --- | --- |",
            ]
        )
        for assignment in profile["cell_protection_assignments"]:
            lines.append(
                "| {sheet} | {scope} | {target} | {locked} | {hidden} |".format(
                    sheet=_markdown_escape(assignment["sheet"]),
                    scope=_markdown_escape(assignment["scope"]),
                    target=_markdown_escape(assignment["target"]),
                    locked="yes" if assignment["locked"] else "no",
                    hidden="yes" if assignment["hidden"] else "no",
                )
        )
        lines.append(
            "Assignments retain their serialized cell, row, or column scope; "
            "they are not expanded into cells."
        )
    external_settings = profile["external_data_refresh_settings"]
    external_connections = profile["external_data_connections"]
    query_tables = profile["query_table_refresh_controls"]
    pivot_caches = profile["pivot_cache_refresh_controls"]
    external_link_packages = profile["external_link_packages"]
    external_workbook_link_surfaces = profile["external_workbook_link_surfaces"]
    external_relationships = profile["external_relationships"]
    formula_external_actions = profile["formula_external_actions"]
    formula_dde_links = profile["formula_dde_links"]
    python_in_excel = profile["python_in_excel"]
    office_custom_functions = profile["office_custom_functions"]
    unqualified_runtime_functions = profile["unqualified_runtime_functions"]
    worksheet_code_resource_registrations = profile[
        "worksheet_code_resource_registrations"
    ]
    formula_defined_xlm_registrations = profile[
        "formula_defined_xlm_registrations"
    ]
    formula_defined_xlm_evaluations = profile[
        "formula_defined_xlm_evaluations"
    ]
    formula_defined_xlm_actions = profile["formula_defined_xlm_actions"]
    formula_defined_xlm_get_cell_calls = profile[
        "formula_defined_xlm_get_cell_calls"
    ]
    formula_defined_xlm_environment_information_calls = profile[
        "formula_defined_xlm_environment_information_calls"
    ]
    formula_environment_information_calls = profile[
        "formula_environment_information_calls"
    ]
    environment_information_formula_cell_count = (
        formula_defined_xlm_environment_information_calls[
            "environment_information_formula_cell_count"
        ]
    )
    environment_information_function_count = (
        formula_defined_xlm_environment_information_calls[
            "environment_information_function_count"
        ]
    )
    environment_information_defined_name_count = (
        formula_defined_xlm_environment_information_calls[
            "environment_information_defined_name_count"
        ]
    )
    native_environment_information_formula_cell_count = (
        formula_environment_information_calls[
            "environment_information_formula_cell_count"
        ]
    )
    native_environment_information_function_count = (
        formula_environment_information_calls["environment_information_function_count"]
    )
    native_environment_information_defined_name_count = (
        formula_environment_information_calls[
            "environment_information_defined_name_count"
        ]
    )
    implicit_cell_reference_function_count = formula_environment_information_calls[
        "implicit_cell_reference_function_count"
    ]
    implicit_sheets_reference_function_count = formula_environment_information_calls[
        "implicit_sheets_reference_function_count"
    ]
    sheet_function_count = formula_environment_information_calls["sheet_function_count"]
    sheets_function_count = formula_environment_information_calls[
        "sheets_function_count"
    ]
    formula_external_action_call_count = sum(
        (
            formula_external_actions["hyperlink_function_count"],
            formula_external_actions["webservice_function_count"],
            formula_external_actions["image_function_count"],
            formula_external_actions["rtd_function_count"],
            formula_external_actions["stockhistory_function_count"],
            formula_external_actions["cube_function_count"],
        )
    )
    has_nondefault_external_settings = external_settings != {
        "update_links": "user_set",
        "allow_refresh_query": False,
        "refresh_all_connections": False,
        "save_external_link_values": True,
    }
    if has_nondefault_external_settings or external_connections or query_tables or pivot_caches:
        update_links = {
            "user_set": "user choice",
            "always": "always",
            "never": "never",
            "unrecognized": "unrecognized",
        }.get(external_settings["update_links"], "unrecognized")
        lines.extend(
            [
                "",
                "## External-data refresh controls",
                "",
                f"- **External-workbook link updates on open:** {update_links}",
                (
                    "- **Query-table refresh allowed:** "
                    + ("yes" if external_settings["allow_refresh_query"] else "no")
                ),
                (
                    "- **Refresh all connections on open:** "
                    + ("yes" if external_settings["refresh_all_connections"] else "no")
                ),
                (
                    "- **Cache external-workbook values on save:** "
                    + (
                        "yes"
                        if external_settings["save_external_link_values"]
                        else "no"
                    )
                ),
            ]
        )
    if external_connections:
        lines.extend(
            [
                "",
                "## External-data connections",
                "",
                "| Connection | Source | Automatic refresh | Runtime | Data / authentication | "
                "Source signals |",
                "| --- | --- | --- | --- | --- | --- |",
            ]
        )
        for connection in external_connections:
            automatic: list[str] = []
            if connection["refresh_on_load"]:
                automatic.append("on open")
            if connection["refresh_interval_minutes"] is not None:
                automatic.append(
                    f"every {connection['refresh_interval_minutes']} min"
                )
            if connection["parameters_refresh_on_change"]:
                automatic.append(
                    "parameter change "
                    f"({connection['parameters_refresh_on_change']})"
                )
            runtime = ["background" if connection["background"] else "foreground"]
            if connection["keep_alive"]:
                runtime.append("keep alive")
            data_behavior = [
                "cache data" if connection["save_data"] else "do not cache data",
                "password saved" if connection["save_password"] else "password not saved",
                f"credentials: {connection['credential_method']}",
            ]
            source_signals: list[str] = []
            if connection["deleted"]:
                source_signals.append("deleted")
            if connection["has_source_file"]:
                source_signals.append("source file")
            if connection["has_connection_file"]:
                source_signals.append("connection file")
                source_signals.append(
                    f"reload: {connection['reconnection_method']}"
                )
            if connection["only_use_connection_file"]:
                source_signals.append("connection-file only")
            if connection["has_single_sign_on_id"]:
                source_signals.append("SSO id")
            if connection["awaiting_initial_refresh"]:
                source_signals.append("initial refresh pending")
            if connection["source_components"]:
                source_signals.append(
                    ", ".join(connection["source_components"])
                )
            if connection["parameter_count"]:
                source_signals.append(
                    f"{connection['parameter_count']} parameter(s)"
                )
            if connection["opaque_metadata"]["present"]:
                source_signals.append(
                    f"{connection['opaque_metadata']['count']} unmodelled XML item(s)"
                )
            identifier = (
                f"#{connection['id']}"
                if connection["id"] is not None
                else "unknown"
            )
            lines.append(
                "| {identifier} | {source} | {automatic} | {runtime} | {data} | {signals} |".format(
                    identifier=_markdown_escape(identifier),
                    source=_markdown_escape(connection["source_type"]),
                    automatic=_markdown_escape(
                        "; ".join(automatic) if automatic else "manual"
                    ),
                    runtime=_markdown_escape("; ".join(runtime)),
                    data=_markdown_escape("; ".join(data_behavior)),
                    signals=_markdown_escape(
                        "; ".join(source_signals) if source_signals else "none"
                    ),
                )
            )
        lines.append(
            "Connection names, descriptions, paths, URLs, commands, parameter values, "
            "SSO identifiers, and opaque XML are intentionally omitted."
        )
    if query_tables:
        lines.extend(
            [
                "",
                "## Query-table refresh controls",
                "",
                "| Sheet | Connection | Automatic refresh | Behavior |",
                "| --- | --- | --- | --- |",
            ]
        )
        for query_table in query_tables:
            behavior = [
                "background" if query_table["background_refresh"] else "foreground",
                f"growth: {query_table['growth_behavior']}",
            ]
            if query_table["refresh_disabled"]:
                behavior.append("refresh disabled")
            if query_table["fill_formulas"]:
                behavior.append("fill adjacent formulas")
            if query_table["remove_data_on_save"]:
                behavior.append("remove data on save")
            if query_table["connection_edit_disabled"]:
                behavior.append("connection editing disabled")
            if query_table["has_refresh_metadata"]:
                behavior.append("refresh metadata")
            if query_table["opaque_metadata"]["present"]:
                behavior.append(
                    f"{query_table['opaque_metadata']['count']} unmodelled XML item(s)"
                )
            identifier = (
                f"#{query_table['connection_id']}"
                if query_table["connection_id"] is not None
                else "unknown"
            )
            lines.append(
                "| {sheet} | {identifier} | {automatic} | {behavior} |".format(
                    sheet=_markdown_escape(query_table["sheet"]),
                    identifier=_markdown_escape(identifier),
                    automatic=("on open" if query_table["refresh_on_load"] else "manual"),
                    behavior=_markdown_escape("; ".join(behavior)),
                )
            )
        lines.append("Query-table names and refreshed data are intentionally omitted.")
    if pivot_caches:
        lines.extend(
            [
                "",
                "## Pivot-cache refresh controls",
                "",
                "| Cache | Source | Connection | Automatic refresh | Behavior |",
                "| --- | --- | --- | --- | --- |",
            ]
        )
        for pivot_cache in pivot_caches:
            behavior = [
                "background" if pivot_cache["background_query"] else "foreground",
                (
                    "refresh enabled"
                    if pivot_cache["refresh_enabled"]
                    else "refresh disabled"
                ),
                "cache data" if pivot_cache["save_data"] else "do not cache data",
            ]
            if pivot_cache["upgrade_on_refresh"]:
                behavior.append("upgrade on refresh")
            if pivot_cache["opaque_metadata"]["present"]:
                behavior.append(
                    f"{pivot_cache['opaque_metadata']['count']} unmodelled XML item(s)"
                )
            cache_identifier = (
                f"#{pivot_cache['cache_id']}"
                if pivot_cache["cache_id"] is not None
                else "unknown"
            )
            connection_identifier = (
                f"#{pivot_cache['connection_id']}"
                if pivot_cache["connection_id"] is not None
                else "none"
            )
            lines.append(
                "| {cache} | {source} | {connection} | {automatic} | {behavior} |".format(
                    cache=_markdown_escape(cache_identifier),
                    source=_markdown_escape(pivot_cache["source_type"]),
                    connection=_markdown_escape(connection_identifier),
                    automatic=("on open" if pivot_cache["refresh_on_load"] else "manual"),
                    behavior=_markdown_escape("; ".join(behavior)),
                )
            )
        lines.append(
            "Pivot-cache source details, cached records, and raw extension XML "
            "are intentionally omitted."
        )
    if external_link_packages["present"]:
        lines.extend(
            [
                "",
                "## External-link packages",
                "",
                (
                    "- **Package parts:** "
                    f"{external_link_packages['external_link_count']} "
                    f"({external_link_packages['external_workbook_count']} workbook, "
                    f"{external_link_packages['dde_link_count']} DDE, "
                    f"{external_link_packages['ole_link_count']} OLE)"
                ),
                (
                    "- **External workbook references:** "
                    f"{external_link_packages['external_workbook_sheet_count']} sheet name(s), "
                    f"{external_link_packages['external_defined_name_count']} defined name(s)"
                ),
                (
                    "- **Cached external-workbook data:** "
                    f"{external_link_packages['external_workbook_cached_cell_count']} cell(s) "
                    f"across {external_link_packages['external_workbook_cached_sheet_count']} "
                    "sheet cache(s)"
                ),
                (
                    "- **Cached external-workbook refresh errors:** "
                    f"{external_link_packages['external_workbook_cached_refresh_error_count']}"
                ),
                (
                    "- **DDE items:** "
                    f"{external_link_packages['dde_item_count']} "
                    f"({external_link_packages['dde_advise_item_count']} advise, "
                    f"{external_link_packages['dde_ole_item_count']} OLE, "
                    f"{external_link_packages['dde_prefer_picture_item_count']} picture)"
                ),
                (
                    "- **OLE items:** "
                    f"{external_link_packages['ole_item_count']} "
                    f"({external_link_packages['ole_advise_item_count']} advise, "
                    f"{external_link_packages['ole_icon_item_count']} icon, "
                    f"{external_link_packages['ole_prefer_picture_item_count']} picture)"
                ),
            ]
        )
        if external_link_packages["unrecognized_link_count"]:
            lines.append(
                "- **Unrecognized package definitions:** "
                f"{external_link_packages['unrecognized_link_count']}"
            )
        if external_link_packages["opaque_metadata"]["present"]:
            lines.append(
                "- **Unmodelled package metadata:** "
                f"{external_link_packages['opaque_metadata']['count']} item(s)"
            )
        lines.append(
            "External workbook targets, sheet and defined names, DDE services/topics/items, "
            "OLE program and item names, and cached values are intentionally omitted."
        )
    if external_workbook_link_surfaces["present"]:
        lines.extend(
            [
                "",
                "## Static external-workbook link surfaces",
                "",
                (
                    "- **Surfaces / endpoints:** "
                    f"{external_workbook_link_surfaces['surface_count']} / "
                    f"{external_workbook_link_surfaces['external_reference_count']}"
                ),
                (
                    "- **Cell formulas / defined names / data validation / chart formulas:** "
                    f"{external_workbook_link_surfaces['cell_formula_surface_count']} / "
                    f"{external_workbook_link_surfaces['defined_name_surface_count']} / "
                    f"{external_workbook_link_surfaces['data_validation_surface_count']} / "
                    f"{external_workbook_link_surfaces['chart_formula_surface_count']}"
                ),
            ]
        )
        if external_workbook_link_surfaces["opaque_chart_part_count"]:
            lines.append(
                "- **Chart parts with unavailable link-surface coverage:** "
                f"{external_workbook_link_surfaces['opaque_chart_part_count']}"
            )
        lines.append(
            "External workbook targets, source paths, names, formulas, and surface "
            "identities are compared privately and intentionally omitted."
        )
    if external_relationships["present"]:
        lines.extend(
            [
                "",
                "## Package-wide external relationships",
                "",
                (
                    "- **Relationship parts / sources / targets:** "
                    f"{external_relationships['external_relationship_part_count']} / "
                    f"{external_relationships['external_relationship_source_count']} / "
                    f"{external_relationships['external_relationship_count']}"
                ),
                (
                    "- **Hyperlink / image / other targets:** "
                    f"{external_relationships['external_hyperlink_relationship_count']} / "
                    f"{external_relationships['external_image_relationship_count']} / "
                    f"{external_relationships['external_other_relationship_count']}"
                ),
            ]
        )
        if external_relationships["unrecognized_relationship_count"]:
            lines.append(
                "- **Unrecognized or uninspected relationship metadata:** "
                f"{external_relationships['unrecognized_relationship_count']}"
            )
        lines.append(
            "Relationship source parts, types, identifiers, targets, and raw XML are "
            "compared privately and intentionally omitted."
        )
    if formula_external_actions["present"]:
        lines.extend(
            [
                "",
                "## Formula external-action and data-provider surfaces",
                "",
                (
                    "- **Formula cells / function calls / formula-defined names:** "
                    f"{formula_external_actions['formula_external_action_cell_count']} / "
                    f"{formula_external_action_call_count} / "
                    f"{formula_external_actions['action_defined_name_count']}"
                ),
                (
                    "- **HYPERLINK / WEBSERVICE / IMAGE / RTD:** "
                    f"{formula_external_actions['hyperlink_function_count']} / "
                    f"{formula_external_actions['webservice_function_count']} / "
                    f"{formula_external_actions['image_function_count']} / "
                    f"{formula_external_actions['rtd_function_count']}"
                ),
                (
                    "- **STOCKHISTORY / Cube functions:** "
                    f"{formula_external_actions['stockhistory_function_count']} / "
                    f"{formula_external_actions['cube_function_count']}"
                ),
                (
                    "FormulaFence inventories stored calls in cells, formula-defined names, "
                    "and named LAMBDAs. Formula cells, name identities, arguments, destinations, "
                    "connections, queries, provider names, and results are compared privately and "
                    "intentionally omitted; no formula is evaluated or provider contacted."
                ),
            ]
        )
    if formula_dde_links["present"]:
        lines.extend(
            [
                "",
                "## Direct DDE-style formula links",
                "",
                (
                    "- **Formula cells / DDE links / formula-defined names:** "
                    f"{formula_dde_links['dde_formula_cell_count']} / "
                    f"{formula_dde_links['dde_link_count']} / "
                    f"{formula_dde_links['dde_defined_name_count']}"
                ),
                (
                    "FormulaFence inventories only the explicit lexical "
                    "application|topic!item form (including a terminal command form) "
                    "in worksheet formulas, "
                    "formula-defined names, and named LAMBDAs. Services, topics, "
                    "items, cells, formulas, and name identities are compared privately "
                    "and intentionally omitted."
                ),
                (
                    "No formula is evaluated and no DDE server is looked up, started, "
                    "contacted, or sent a command. Package-level externalLink DDE/OLE "
                    "metadata remains a separate inventory because it can exist without "
                    "a direct formula link."
                ),
            ]
        )
    if python_in_excel["present"]:
        lines.extend(
            [
                "",
                "## Python in Excel code",
                "",
                (
                    "- **Package parts / PY formula cells / PY calls:** "
                    f"{python_in_excel['python_part_count']} / "
                    f"{python_in_excel['python_formula_cell_count']} / "
                    f"{python_in_excel['python_function_count']}"
                ),
                (
                    "- **Stored scripts / environment definitions / initializations:** "
                    f"{python_in_excel['python_script_count']} / "
                    f"{python_in_excel['python_environment_definition_count']} / "
                    f"{python_in_excel['python_initialization_count']}"
                ),
            ]
        )
        if python_in_excel["unrecognized_python_in_excel_count"]:
            lines.append(
                "- **Unrecognized or uninspected Python metadata:** "
                f"{python_in_excel['unrecognized_python_in_excel_count']}"
            )
        lines.append(
            "Package totals include each stored 2023 Python and 2022 PythonScripts "
            "part; FormulaFence does not assume they agree. Python code, environment "
            "identifiers, script indexes, and raw XML are compared privately and "
            "intentionally omitted. Ordinary semantic diffs retain changed PY formulas "
            "and values by design; no code is loaded or run and no cloud runtime is "
            "contacted."
        )
    if office_custom_functions["present"]:
        lines.extend(
            [
                "",
                "## Namespaced custom-function calls",
                "",
                (
                    "- **Formula cells / calls / namespaces:** "
                    f"{office_custom_functions['namespaced_custom_function_formula_cell_count']} / "
                    f"{office_custom_functions['namespaced_custom_function_call_count']} / "
                    f"{office_custom_functions['namespaced_custom_function_namespace_count']}"
                ),
                (
                    "FormulaFence inventories direct namespaced callable candidates that are "
                    "not known native Excel functions or workbook-defined names, then "
                    "propagates candidates in formula-defined names and named LAMBDAs. "
                    "Function names, cells, formulas, and arguments are compared privately "
                    "and intentionally omitted; no formula is evaluated, add-in loaded, or "
                    "network request made."
                ),
                (
                    "A matching call is not proof that an Office Add-in is installed or "
                    "runnable: its manifest, code, and runtime are outside the workbook. "
                    "Unqualified VBA, COM, and XLL UDF calls are outside this boundary."
                ),
            ]
        )
    if unqualified_runtime_functions["present"]:
        runtime_formula_cell_count = unqualified_runtime_functions[
            "unqualified_runtime_function_formula_cell_count"
        ]
        runtime_call_count = unqualified_runtime_functions[
            "unqualified_runtime_function_call_count"
        ]
        runtime_definition_count = unqualified_runtime_functions[
            "unqualified_runtime_function_defined_name_count"
        ]
        lines.extend(
            [
                "",
                "## Unqualified runtime-function candidates",
                "",
                (
                    "- **Formula cells / calls / relevant named definitions:** "
                    f"{runtime_formula_cell_count} / {runtime_call_count} / "
                    f"{runtime_definition_count}"
                ),
                (
                    "FormulaFence inventories bare callable identifiers that are not "
                    "known native Excel functions or workbook-defined names, then "
                    "propagates candidates through formula-defined names and named "
                    "LAMBDAs. Candidate names, cells, formulas, and arguments are "
                    "compared privately and intentionally omitted."
                ),
                (
                    "A candidate is not proof that any runtime is installed or can run: "
                    "it may resolve to VBA, COM/Automation, an XLL, or another "
                    "registered provider. FormulaFence does not evaluate formulas, "
                    "resolve a provider, load code, or inspect the host environment."
                ),
            ]
        )
    if worksheet_code_resource_registrations["present"]:
        lines.extend(
            [
                "",
                "## Worksheet code-resource registrations",
                "",
                (
                    "- **Formula cells / REGISTER.ID calls / formula-defined names:** "
                    f"{worksheet_code_resource_registrations['registration_formula_cell_count']} / "
                    f"{worksheet_code_resource_registrations['register_id_function_count']} / "
                    f"{worksheet_code_resource_registrations['registration_defined_name_count']}"
                ),
                (
                    "FormulaFence inventories stored worksheet and formula-defined "
                    "REGISTER.ID calls, including named "
                    "LAMBDAs. Module paths, procedure names, type strings, cells, "
                    "formulas, and arguments are compared privately and intentionally "
                    "omitted; no formula is evaluated and no DLL or code resource is loaded."
                ),
                (
                    "This is distinct from XLM macro-sheet CALL/REGISTER behavior, "
                    "which remains covered by the raw XLM macro-sheet boundary below."
                ),
            ]
        )
    if formula_defined_xlm_registrations["present"]:
        lines.extend(
            [
                "",
                "## Formula-defined XLM registrations",
                "",
                (
                    "- **Invoking formula cells / REGISTER calls / formula-defined names:** "
                    f"{formula_defined_xlm_registrations['registration_formula_cell_count']} / "
                    f"{formula_defined_xlm_registrations['register_function_count']} / "
                    f"{formula_defined_xlm_registrations['registration_defined_name_count']}"
                ),
                (
                    "FormulaFence inventories XLM REGISTER calls stored in "
                    "formula-defined names and named LAMBDAs, then records cells that "
                    "statically invoke them. Module paths, procedure names, type strings, "
                    "cells, formulas, and arguments are compared privately and intentionally "
                    "omitted; no formula is evaluated, no macro is run, and no DLL or XLL is "
                    "loaded."
                ),
                (
                    "Direct worksheet REGISTER calls and raw XLM macro-sheet parts are "
                    "outside this narrow boundary; macro-sheet content remains covered by "
                    "the raw XLM macro-sheet boundary below."
                ),
            ]
        )
    if formula_defined_xlm_evaluations["present"]:
        lines.extend(
            [
                "",
                "## Formula-defined XLM expression evaluation",
                "",
                (
                    "- **Invoking formula cells / EVALUATE calls / formula-defined names:** "
                    f"{formula_defined_xlm_evaluations['evaluation_formula_cell_count']} / "
                    f"{formula_defined_xlm_evaluations['evaluate_function_count']} / "
                    f"{formula_defined_xlm_evaluations['evaluation_defined_name_count']}"
                ),
                (
                    "FormulaFence inventories XLM EVALUATE calls stored in "
                    "formula-defined names and named LAMBDAs, then records cells that "
                    "statically invoke them. Expressions, cells, formulas, arguments, and "
                    "name identities are compared privately and intentionally omitted; no "
                    "formula or text expression is evaluated."
                ),
                (
                    "Only a stored call's statically visible argument edge is traced. "
                    "Formula text parsed by EVALUATE is not re-tokenized, so dependencies "
                    "inside that text remain an explicit coverage limit. Direct worksheet "
                    "EVALUATE calls and raw XLM macro-sheet parts are outside this narrow "
                    "boundary."
                ),
            ]
        )
    if formula_defined_xlm_actions["present"]:
        lines.extend(
            [
                "",
                "## Formula-defined XLM actions and event dispatch",
                "",
                (
                    "- **Invoking formula cells / selected action calls / "
                    "formula-defined names:** "
                    f"{formula_defined_xlm_actions['action_formula_cell_count']} / "
                    f"{formula_defined_xlm_actions['action_function_count']} / "
                    f"{formula_defined_xlm_actions['action_defined_name_count']}"
                ),
                (
                    "FormulaFence inventories selected XLM CALL, EXEC, EXECUTE, "
                    "RUN, SEND.KEYS, and ON.* action or event-dispatch calls stored "
                    "in formula-defined names and named LAMBDAs, then records cells "
                    "that statically invoke them. Function spelling, cells, formulas, "
                    "arguments, action targets, and name identities are compared "
                    "privately and intentionally omitted."
                ),
                (
                    "No formula is evaluated and no macro, program, DLL entry point, "
                    "DDE command, or event handler is resolved or run. Direct worksheet "
                    "action calls and raw XLM macro-sheet parts are outside this narrow "
                    "boundary."
                ),
            ]
        )
    if formula_defined_xlm_get_cell_calls["present"]:
        lines.extend(
            [
                "",
                "## Formula-defined XLM GET.CELL information",
                "",
                (
                    "- **Invoking formula cells / GET.CELL calls / formula-defined names:** "
                    f"{formula_defined_xlm_get_cell_calls['get_cell_formula_cell_count']} / "
                    f"{formula_defined_xlm_get_cell_calls['get_cell_function_count']} / "
                    f"{formula_defined_xlm_get_cell_calls['get_cell_defined_name_count']}"
                ),
                (
                    "FormulaFence inventories XLM GET.CELL calls stored in "
                    "formula-defined names and named LAMBDAs, then records cells that "
                    "statically invoke them. Information types, references, cells, "
                    "formulas, arguments, and name identities are compared privately and "
                    "intentionally omitted; no formula is evaluated."
                ),
                (
                    "Only ordinary static argument edges are traced. FormulaFence does "
                    "not determine which information type is requested, resolve dynamic "
                    "references, or simulate formatting, display, protection, comments, "
                    "or other Excel state. Direct worksheet GET.CELL calls and raw XLM "
                    "macro-sheet parts are outside this narrow boundary."
                ),
            ]
        )
    if formula_defined_xlm_environment_information_calls["present"]:
        lines.extend(
            [
                "",
                "## Formula-defined XLM environment information",
                "",
                (
                    "- **Invoking formula cells / information calls / formula-defined names:** "
                    f"{environment_information_formula_cell_count} / "
                    f"{environment_information_function_count} / "
                    f"{environment_information_defined_name_count}"
                ),
                (
                    "FormulaFence inventories selected XLM GET.WORKBOOK, "
                    "GET.WORKSPACE, and GET.DOCUMENT calls stored in "
                    "formula-defined names and named LAMBDAs, then records cells "
                    "that statically invoke them. Information types, references, "
                    "formulas, arguments, locations, and name identities are "
                    "compared privately and intentionally omitted; no formula is "
                    "evaluated."
                ),
                (
                    "Only ordinary static argument edges are traced. FormulaFence "
                    "does not determine which information type is requested, "
                    "resolve dynamic references, or simulate workbook, workspace, "
                    "document, client, add-in, printer, or other Excel state. "
                    "Direct worksheet calls and raw XLM macro-sheet parts are "
                    "outside this narrow boundary."
                ),
            ]
        )
    if formula_environment_information_calls["present"]:
        lines.extend(
            [
                "",
                "## Native CELL, INFO, SHEET, and SHEETS information",
                "",
                (
                    "- **Formula cells / native calls / formula-defined names:** "
                    f"{native_environment_information_formula_cell_count} / "
                    f"{native_environment_information_function_count} / "
                    f"{native_environment_information_defined_name_count}"
                ),
                (
                    "- **CELL calls without an explicit reference:** "
                    f"{implicit_cell_reference_function_count}"
                ),
                (
                    "- **SHEET calls / SHEETS calls:** "
                    f"{sheet_function_count} / {sheets_function_count}"
                ),
                (
                    "- **SHEETS calls without an explicit reference:** "
                    f"{implicit_sheets_reference_function_count}"
                ),
                (
                    "FormulaFence inventories native CELL, INFO, SHEET, and SHEETS "
                    "calls in worksheet formulas, formula-defined names, and named "
                    "LAMBDAs. Information types, references, formulas, arguments, "
                    "locations, and name identities are compared privately and "
                    "intentionally omitted; no formula is evaluated."
                ),
                (
                    "CELL calls without a stored reference are counted because Excel "
                    "may use the current selection at calculation time. FormulaFence "
                    "also compares the private OOXML workbook tab catalog when SHEET "
                    "or an omitted-reference SHEETS call is present, including hidden "
                    "and non-worksheet tabs. It does not determine an information type, "
                    "resolve dynamic arguments, or simulate file, folder, client, "
                    "workspace, selection, or workbook state."
                ),
            ]
        )
    xlm_macro_sheets = profile["xlm_macro_sheets"]
    if xlm_macro_sheets["present"]:
        lines.extend(
            [
                "",
                "## Excel 4.0 / XLM macro sheets",
                "",
                (
                    "- **Workbook declarations:** "
                    f"{xlm_macro_sheets['declared_macro_sheet_count']}"
                ),
                (
                    "- **Package parts:** "
                    f"{xlm_macro_sheets['macro_sheet_count']} "
                    f"({xlm_macro_sheets['international_macro_sheet_count']} international)"
                ),
                (
                    "- **Macro formula cells:** "
                    f"{xlm_macro_sheets['formula_cell_count']}"
                ),
                (
                    "- **Hidden parts:** "
                    f"{xlm_macro_sheets['hidden_macro_sheet_count']} hidden, "
                    f"{xlm_macro_sheets['very_hidden_macro_sheet_count']} very hidden"
                ),
                (
                    "- **Related package relationships:** "
                    f"{xlm_macro_sheets['related_relationship_count']} "
                    f"({xlm_macro_sheets['external_relationship_count']} external, "
                    f"{xlm_macro_sheets['embedded_object_relationship_count']} OLE object, "
                    f"{xlm_macro_sheets['embedded_package_relationship_count']} package)"
                ),
                (
                    "- **Internal related parts:** "
                    f"{xlm_macro_sheets['internal_related_part_count']} "
                    f"({xlm_macro_sheets['fingerprinted_related_part_count']} "
                    "fingerprinted)"
                ),
            ]
        )
        if xlm_macro_sheets["uninspected_related_part_count"]:
            lines.append(
                "- **Uninspected related parts:** "
                f"{xlm_macro_sheets['uninspected_related_part_count']}"
            )
        if xlm_macro_sheets["unrecognized_macro_sheet_count"]:
            lines.append(
                "- **Unrecognized or uninspected macro-sheet parts:** "
                f"{xlm_macro_sheets['unrecognized_macro_sheet_count']}"
            )
        lines.append(
            "XLM commands, cell values, relationship targets, and direct internal "
            "related-part contents are compared privately and intentionally omitted."
        )
    xlm_automatic_macro_bindings = profile["xlm_automatic_macro_bindings"]
    if xlm_automatic_macro_bindings["present"]:
        lines.extend(
            [
                "",
                "## XLM automatic-macro bindings",
                "",
                (
                    "- **Workbook bindings:** "
                    f"{xlm_automatic_macro_bindings['automatic_macro_binding_count']}"
                ),
                (
                    "- **Auto_Open / Auto_Close / Auto_Activate / Auto_Deactivate:** "
                    f"{xlm_automatic_macro_bindings['auto_open_binding_count']} / "
                    f"{xlm_automatic_macro_bindings['auto_close_binding_count']} / "
                    f"{xlm_automatic_macro_bindings['auto_activate_binding_count']} / "
                    f"{xlm_automatic_macro_bindings['auto_deactivate_binding_count']}"
                ),
                (
                    "Automatic-macro name spellings, target cells, and stored definitions "
                    "are compared privately and intentionally omitted."
                ),
            ]
        )
    ribbon_customization = profile["ribbon_customization"]
    if ribbon_customization["present"]:
        lines.extend(
            [
                "",
                "## Office RibbonX customization",
                "",
                (
                    "- **Package declarations:** "
                    f"{ribbon_customization['declared_ribbon_part_count']}"
                ),
                (
                    "- **Customization parts:** "
                    f"{ribbon_customization['ribbon_part_count']} "
                    f"({ribbon_customization['office_2010_ribbon_part_count']} Office 2010)"
                ),
                f"- **Controls:** {ribbon_customization['control_count']}",
                (
                    "- **Callback attributes:** "
                    f"{ribbon_customization['callback_attribute_count']} "
                    f"({ribbon_customization['action_callback_count']} onAction)"
                ),
                (
                    "- **Image relationships:** "
                    f"{ribbon_customization['image_relationship_count']}"
                ),
                (
                    "- **External Ribbon relationships:** "
                    f"{ribbon_customization['external_relationship_count']}"
                ),
            ]
        )
        if ribbon_customization["unrecognized_ribbon_part_count"]:
            lines.append(
                "- **Unrecognized or uninspected customization parts:** "
                f"{ribbon_customization['unrecognized_ribbon_part_count']}"
            )
        lines.append(
            "Ribbon XML, control names and labels, callback names, and image targets are "
            "compared privately and intentionally omitted."
        )
    office_web_addins = profile["office_web_addins"]
    if office_web_addins["present"]:
        lines.extend(
            [
                "",
                "## Office Web Add-ins",
                "",
                (
                    "- **Workbook task-pane declarations:** "
                    f"{office_web_addins['declared_taskpane_part_count']}"
                ),
                (
                    "- **Task-pane parts:** "
                    f"{office_web_addins['taskpane_part_count']}"
                ),
                (
                    "- **Web-extension definition parts:** "
                    f"{office_web_addins['web_extension_part_count']}"
                ),
                f"- **Task panes:** {office_web_addins['taskpane_count']}",
                (
                    "- **Visible task panes:** "
                    f"{office_web_addins['visible_taskpane_count']}"
                ),
                (
                    "- **Locked task panes:** "
                    f"{office_web_addins['locked_taskpane_count']}"
                ),
                (
                    "- **Task-pane web-extension references:** "
                    f"{office_web_addins['web_extension_reference_count']}"
                ),
                (
                    "- **Auto-show task-pane requests:** "
                    f"{office_web_addins['auto_show_taskpane_count']}"
                ),
                (
                    "- **Store references:** "
                    f"{office_web_addins['store_reference_count']} "
                    f"({office_web_addins['alternate_reference_count']} alternate)"
                ),
                f"- **Bindings:** {office_web_addins['binding_count']}",
                (
                    "- **Worksheet binding sheets / bindings:** "
                    f"{office_web_addins['worksheet_binding_sheet_count']} / "
                    f"{office_web_addins['worksheet_binding_count']}"
                ),
                (
                    "- **In-content drawing parts / references / definitions:** "
                    f"{office_web_addins['in_content_drawing_part_count']} / "
                    f"{office_web_addins['in_content_web_extension_reference_count']} / "
                    f"{office_web_addins['in_content_web_extension_part_count']}"
                ),
                (
                    "- **Snapshot relationships:** "
                    f"{office_web_addins['snapshot_reference_count']}"
                ),
                (
                    "- **Related package relationships:** "
                    f"{office_web_addins['related_relationship_count']} "
                    f"({office_web_addins['external_relationship_count']} external)"
                ),
            ]
        )
        if office_web_addins["unrecognized_part_count"]:
            lines.append(
                "- **Unrecognized or uninspected parts/bindings:** "
                f"{office_web_addins['unrecognized_part_count']}"
            )
        lines.append(
            "Add-in identities, store references, properties, bindings, worksheet formulas, "
            "snapshots, frame XML, and relationship targets are compared privately and "
            "intentionally omitted."
        )
    pivot_table_definitions = profile["pivot_table_definitions"]
    if pivot_table_definitions["present"]:
        lines.extend(
            [
                "",
                "## PivotTable definitions and cached report material",
                "",
                (
                    "- **PivotTable host sheets / parts:** "
                    f"{pivot_table_definitions['pivot_table_sheet_count']} / "
                    f"{pivot_table_definitions['pivot_table_part_count']}"
                ),
                (
                    "- **Cache-definition / cache-record parts:** "
                    f"{pivot_table_definitions['pivot_cache_definition_part_count']} / "
                    f"{pivot_table_definitions['pivot_cache_records_part_count']}"
                ),
                (
                    "- **PivotTable-to-cache bindings:** "
                    f"{pivot_table_definitions['pivot_cache_binding_count']}"
                ),
                (
                    "- **Layout locations:** "
                    f"{pivot_table_definitions['layout_location_count']}"
                ),
                (
                    "- **Pivot / row / column / page / data fields:** "
                    f"{pivot_table_definitions['pivot_field_count']} / "
                    f"{pivot_table_definitions['row_field_count']} / "
                    f"{pivot_table_definitions['column_field_count']} / "
                    f"{pivot_table_definitions['page_field_count']} / "
                    f"{pivot_table_definitions['data_field_count']}"
                ),
                (
                    "- **Filters / row items / column items:** "
                    f"{pivot_table_definitions['filter_count']} / "
                    f"{pivot_table_definitions['row_item_count']} / "
                    f"{pivot_table_definitions['column_item_count']}"
                ),
                (
                    "- **Cache fields / shared items:** "
                    f"{pivot_table_definitions['cache_field_count']} / "
                    f"{pivot_table_definitions['shared_item_count']}"
                ),
                (
                    "- **Calculated cache items / members:** "
                    f"{pivot_table_definitions['calculated_item_count']} / "
                    f"{pivot_table_definitions['calculated_member_count']}"
                ),
                (
                    "- **Declared cache records:** "
                    f"{pivot_table_definitions['cache_record_count']}"
                ),
                (
                    "- **Related package relationships:** "
                    f"{pivot_table_definitions['related_relationship_count']} "
                    f"({pivot_table_definitions['external_relationship_count']} external)"
                ),
                (
                    "- **Fingerprinted cache-record parts:** "
                    f"{pivot_table_definitions['fingerprinted_cache_record_part_count']}"
                ),
            ]
        )
        if pivot_table_definitions["uninspected_cache_record_part_count"]:
            lines.append(
                "- **Uninspected cache-record parts:** "
                f"{pivot_table_definitions['uninspected_cache_record_part_count']}"
            )
        if pivot_table_definitions["unrecognized_part_count"]:
            lines.append(
                "- **Unrecognized or uninspected PivotTable parts/bindings:** "
                f"{pivot_table_definitions['unrecognized_part_count']}"
            )
        lines.append(
            "PivotTable names, source ranges, fields, item values, formulas, cache records, "
            "relationship targets, XML, and payload bytes are compared privately and "
            "intentionally omitted."
        )
    slicer_timeline_caches = profile["slicer_timeline_caches"]
    if slicer_timeline_caches["present"]:
        lines.extend(
            [
                "",
                "## Slicer and Timeline cache filter state",
                "",
                (
                    "- **Slicer / Timeline cache parts:** "
                    f"{slicer_timeline_caches['slicer_cache_part_count']} / "
                    f"{slicer_timeline_caches['timeline_cache_part_count']}"
                ),
                (
                    "- **Workbook cache bindings (slicer / Timeline):** "
                    f"{slicer_timeline_caches['slicer_workbook_binding_count']} / "
                    f"{slicer_timeline_caches['timeline_workbook_binding_count']}"
                ),
                (
                    "- **Pivot-cache bindings (slicer / Timeline):** "
                    f"{slicer_timeline_caches['slicer_pivot_cache_binding_count']} / "
                    f"{slicer_timeline_caches['timeline_pivot_cache_binding_count']}"
                ),
                (
                    "- **Table-slicer bindings:** "
                    f"{slicer_timeline_caches['slicer_table_binding_count']}"
                ),
                (
                    "- **Filtered PivotTable views (slicer / Timeline):** "
                    f"{slicer_timeline_caches['slicer_pivot_table_binding_count']} / "
                    f"{slicer_timeline_caches['timeline_pivot_table_binding_count']}"
                ),
                (
                    "- **Slicer items / selected items:** "
                    f"{slicer_timeline_caches['slicer_item_count']} / "
                    f"{slicer_timeline_caches['selected_slicer_item_count']}"
                ),
                (
                    "- **Timeline states / filters:** "
                    f"{slicer_timeline_caches['timeline_state_count']} / "
                    f"{slicer_timeline_caches['timeline_filter_count']}"
                ),
                (
                    "- **Related package relationships:** "
                    f"{slicer_timeline_caches['related_relationship_count']} "
                    f"({slicer_timeline_caches['external_relationship_count']} external)"
                ),
            ]
        )
        if slicer_timeline_caches["unrecognized_part_count"]:
            lines.append(
                "- **Unrecognized or uninspected cache parts/bindings:** "
                f"{slicer_timeline_caches['unrecognized_part_count']}"
            )
        lines.append(
            "Slicer and Timeline cache names, source fields, selected values, filter ranges, "
            "PivotTable names, relationship targets, and XML are compared privately and "
            "intentionally omitted."
        )
    power_pivot_data_model = profile["power_pivot_data_model"]
    if power_pivot_data_model["present"]:
        lines.extend(
            [
                "",
                "## Power Pivot Data Model",
                "",
                (
                    "- **Embedded model parts / workbook bindings:** "
                    f"{power_pivot_data_model['data_model_part_count']} / "
                    f"{power_pivot_data_model['workbook_binding_count']}"
                ),
                (
                    "- **Workbook Data Model declarations:** "
                    f"{power_pivot_data_model['data_model_declaration_count']}"
                ),
                (
                    "- **Model tables / relationships:** "
                    f"{power_pivot_data_model['model_table_count']} / "
                    f"{power_pivot_data_model['model_relationship_count']}"
                ),
                (
                    "- **Fingerprinted / uninspected model payloads:** "
                    f"{power_pivot_data_model['fingerprinted_data_part_count']} / "
                    f"{power_pivot_data_model['uninspected_data_part_count']}"
                ),
                (
                    "- **Related package relationships:** "
                    f"{power_pivot_data_model['related_relationship_count']} "
                    f"({power_pivot_data_model['external_relationship_count']} external)"
                ),
            ]
        )
        if power_pivot_data_model["unrecognized_part_count"]:
            lines.append(
                "- **Unrecognized or uninspected model parts/bindings:** "
                f"{power_pivot_data_model['unrecognized_part_count']}"
            )
        lines.append(
            "Power Pivot/Data Model table names, relationships, connection details, DAX, "
            "stored values, relationship targets, and raw payload bytes are compared "
            "privately and intentionally omitted."
        )
    what_if_data_tables = profile["what_if_data_tables"]
    if what_if_data_tables["present"]:
        lines.extend(
            [
                "",
                "## What-If Data Tables",
                "",
                f"- **Sensitivity-table masters:** {what_if_data_tables['data_table_count']}",
                (
                    "- **One-variable / two-variable tables:** "
                    f"{what_if_data_tables['one_variable_data_table_count']} / "
                    f"{what_if_data_tables['two_variable_data_table_count']}"
                ),
                (
                    "- **One-variable row / column orientation:** "
                    f"{what_if_data_tables['one_variable_row_oriented_count']} / "
                    f"{what_if_data_tables['one_variable_column_oriented_count']}"
                ),
                (
                    "- **Declared output cells:** "
                    f"{what_if_data_tables['declared_output_cell_count']}"
                ),
                (
                    "- **Recalculation requested / deleted input references:** "
                    f"{what_if_data_tables['recalculation_requested_count']} / "
                    f"{what_if_data_tables['deleted_input_reference_count']}"
                ),
            ]
        )
        if what_if_data_tables["unrecognized_data_table_count"]:
            lines.append(
                "- **Unrecognized or malformed Data Table masters:** "
                f"{what_if_data_tables['unrecognized_data_table_count']}"
            )
        lines.append(
            "Data Table output ranges, input-cell references, and formula metadata are "
            "compared privately and intentionally omitted. Cached scenario-output cells "
            "remain ordinary cell values under the regular cell-diff boundary."
        )
    scenario_manager = profile["scenario_manager"]
    if scenario_manager["present"]:
        lines.extend(
            [
                "",
                "## Scenario Manager",
                "",
                (
                    "- **Scenario-bearing worksheets / scenarios / stored inputs:** "
                    f"{scenario_manager['scenario_sheet_count']} / "
                    f"{scenario_manager['scenario_count']} / "
                    f"{scenario_manager['input_cell_count']}"
                ),
                (
                    "- **Locked / hidden scenarios:** "
                    f"{scenario_manager['locked_scenario_count']} / "
                    f"{scenario_manager['hidden_scenario_count']}"
                ),
                (
                    "- **Scenarios with comments / users:** "
                    f"{scenario_manager['scenario_with_comment_count']} / "
                    f"{scenario_manager['scenario_with_user_count']}"
                ),
                (
                    "- **Summary references / current selections / shown selections:** "
                    f"{scenario_manager['summary_reference_count']} / "
                    f"{scenario_manager['current_scenario_selection_count']} / "
                    f"{scenario_manager['shown_scenario_selection_count']}"
                ),
                (
                    "- **Deleted / undone / formatted stored inputs:** "
                    f"{scenario_manager['deleted_input_cell_count']} / "
                    f"{scenario_manager['undone_input_cell_count']} / "
                    f"{scenario_manager['formatted_input_cell_count']}"
                ),
            ]
        )
        if scenario_manager["unrecognized_scenario_count"]:
            lines.append(
                "- **Unrecognized or malformed scenario declarations:** "
                f"{scenario_manager['unrecognized_scenario_count']}"
            )
        lines.append(
            "Scenario names, comments, user metadata, input values, input references, and "
            "summary references are compared privately and intentionally omitted."
        )
    filter_visibility_controls = profile["filter_visibility_controls"]
    if filter_visibility_controls["present"]:
        lines.extend(
            [
                "",
                "## Filter, sort, and visibility controls",
                "",
                (
                    "- **Worksheet / table AutoFilters:** "
                    f"{filter_visibility_controls['worksheet_auto_filter_count']} / "
                    f"{filter_visibility_controls['table_auto_filter_count']}"
                ),
                (
                    "- **Filter columns / criterion groups:** "
                    f"{filter_visibility_controls['filter_column_count']} / "
                    f"{filter_visibility_controls['filter_criterion_count']}"
                ),
                (
                    "- **Sort states / conditions:** "
                    f"{filter_visibility_controls['sort_state_count']} / "
                    f"{filter_visibility_controls['sort_condition_count']}"
                ),
                (
                    "- **Default-hidden / default-zero-height / default-zero-width sheets:** "
                    f"{filter_visibility_controls['default_hidden_sheet_count']} / "
                    f"{filter_visibility_controls['default_zero_height_sheet_count']} / "
                    f"{filter_visibility_controls['default_zero_width_sheet_count']}"
                ),
                (
                    "- **Hidden / zero-height rows / hidden / zero-width columns:** "
                    f"{filter_visibility_controls['hidden_row_count']} / "
                    f"{filter_visibility_controls['zero_height_row_count']} / "
                    f"{filter_visibility_controls['hidden_column_count']} / "
                    f"{filter_visibility_controls['zero_width_column_count']}"
                ),
                (
                    "- **Outlined rows / columns:** "
                    f"{filter_visibility_controls['outlined_row_count']} / "
                    f"{filter_visibility_controls['outlined_column_count']}"
                ),
                (
                    "- **Collapsed rows / columns / visible-row overrides:** "
                    f"{filter_visibility_controls['collapsed_row_count']} / "
                    f"{filter_visibility_controls['collapsed_column_count']} / "
                    f"{filter_visibility_controls['visible_row_override_count']}"
                ),
            ]
        )
        if filter_visibility_controls["unrecognized_control_count"]:
            lines.append(
                "- **Unrecognized or malformed visibility controls:** "
                f"{filter_visibility_controls['unrecognized_control_count']}"
            )
        lines.append(
            "Filter criteria, selected values, table names, sort keys, custom lists, "
            "and row/column ranges are compared privately and intentionally omitted."
        )
    ignored_error_controls = profile["ignored_error_controls"]
    if ignored_error_controls["present"]:
        lines.extend(
            [
                "",
                "## Ignored Excel error-checking controls",
                "",
                (
                    "- **Worksheets / standard containers / Office 2010 extension containers:** "
                    f"{ignored_error_controls['worksheet_count']} / "
                    f"{ignored_error_controls['standard_container_count']} / "
                    f"{ignored_error_controls['extension_container_count']}"
                ),
                (
                    "- **Suppressed warning rules / target ranges:** "
                    f"{ignored_error_controls['ignored_error_rule_count']} / "
                    f"{ignored_error_controls['target_range_count']}"
                ),
                (
                    "- **Evaluation / inconsistent-formula / omitted-range warnings:** "
                    f"{ignored_error_controls['evaluation_error_count']} / "
                    f"{ignored_error_controls['inconsistent_formula_count']} / "
                    f"{ignored_error_controls['formula_range_omission_count']}"
                ),
                (
                    "- **Unlocked-formula / empty-reference / list-validation / calculated-column "
                    "warnings:** "
                    f"{ignored_error_controls['unlocked_formula_count']} / "
                    f"{ignored_error_controls['empty_cell_reference_count']} / "
                    f"{ignored_error_controls['list_data_validation_count']} / "
                    f"{ignored_error_controls['calculated_column_count']}"
                ),
                (
                    "- **Numbers-stored-as-text / two-digit-text-year warnings:** "
                    f"{ignored_error_controls['number_stored_as_text_count']} / "
                    f"{ignored_error_controls['two_digit_text_year_count']}"
                ),
            ]
        )
        if ignored_error_controls["unrecognized_ignored_error_count"]:
            lines.append(
                "- **Unrecognized or malformed ignored-error controls:** "
                f"{ignored_error_controls['unrecognized_ignored_error_count']}"
            )
        lines.append(
            "Target ranges and individual warning suppressions are compared privately and "
            "intentionally omitted."
        )
    named_sheet_views = profile["named_sheet_views"]
    if named_sheet_views["present"]:
        lines.extend(
            [
                "",
                "## Excel Named Sheet Views",
                "",
                (
                    "- **Worksheets / relationship-backed parts / named views:** "
                    f"{named_sheet_views['worksheet_count']} / "
                    f"{named_sheet_views['part_count']} / "
                    f"{named_sheet_views['named_sheet_view_count']}"
                ),
                (
                    "- **Alternate filters / column filters / criterion groups:** "
                    f"{named_sheet_views['named_filter_count']} / "
                    f"{named_sheet_views['column_filter_count']} / "
                    f"{named_sheet_views['filter_criterion_count']}"
                ),
                (
                    "- **Sort rules / conditions:** "
                    f"{named_sheet_views['sort_rule_count']} / "
                    f"{named_sheet_views['sort_condition_count']}"
                ),
            ]
        )
        if named_sheet_views["unrecognized_named_sheet_view_count"]:
            lines.append(
                "- **Unrecognized or malformed Named Sheet View controls:** "
                f"{named_sheet_views['unrecognized_named_sheet_view_count']}"
            )
        lines.append(
            "View names, IDs, criteria, target ranges, table bindings, and sort keys are "
            "compared privately and intentionally omitted."
        )
    custom_workbook_views = profile["custom_workbook_views"]
    if custom_workbook_views["present"]:
        lines.extend(
            [
                "",
                "## Legacy Excel Custom Views",
                "",
                (
                    "- **Workbook views / per-sheet views / sheets with alternate views:** "
                    f"{custom_workbook_views['custom_workbook_view_count']} / "
                    f"{custom_workbook_views['custom_sheet_view_count']} / "
                    f"{custom_workbook_views['custom_view_sheet_count']}"
                ),
                (
                    "- **Per-sheet views with hidden rows-or-columns / filters / print "
                    "settings / display settings:** "
                    f"{custom_workbook_views['hidden_row_or_column_view_count']} / "
                    f"{custom_workbook_views['filtered_view_count']} / "
                    f"{custom_workbook_views['print_setting_view_count']} / "
                    f"{custom_workbook_views['display_setting_view_count']}"
                ),
            ]
        )
        if custom_workbook_views["unrecognized_custom_view_count"]:
            lines.append(
                "- **Unrecognized, malformed, or incompletely linked Custom Views:** "
                f"{custom_workbook_views['unrecognized_custom_view_count']}"
            )
        lines.append(
            "View names, GUIDs, sheet bindings, ranges, filters, print settings, and raw "
            "Custom View XML are compared privately and intentionally omitted."
        )
    table_style_controls = profile["table_style_controls"]
    if table_style_controls["present"]:
        lines.extend(
            [
                "",
                "## Excel Table Style Controls",
                "",
                (
                    "- **Style declarations / styled tables / custom styles / custom "
                    "style elements:** "
                    f"{table_style_controls['table_style_info_count']} / "
                    f"{table_style_controls['styled_table_count']} / "
                    f"{table_style_controls['custom_table_style_count']} / "
                    f"{table_style_controls['custom_table_style_element_count']}"
                ),
                (
                    "- **Tables using custom styles / row stripes / column stripes / "
                    "emphasized columns:** "
                    f"{table_style_controls['custom_style_applied_table_count']} / "
                    f"{table_style_controls['row_striped_table_count']} / "
                    f"{table_style_controls['column_striped_table_count']} / "
                    f"{table_style_controls['emphasized_column_table_count']}"
                ),
                (
                    "- **Tables with direct Dxf formats / direct Dxf assignments / "
                    "named cell-style assignments:** "
                    f"{table_style_controls['table_direct_dxf_table_count']} / "
                    f"{table_style_controls['table_direct_dxf_assignment_count']} / "
                    f"{table_style_controls['table_named_cell_style_assignment_count']}"
                ),
            ]
        )
        if table_style_controls["unrecognized_table_style_count"]:
            lines.append(
                "- **Unrecognized, malformed, or unresolved Table Style controls:** "
                f"{table_style_controls['unrecognized_table_style_count']}"
            )
        lines.append(
            "Table names, custom and named cell-style names, differential formats, "
            "colours, and raw XML are compared privately and intentionally omitted."
        )
    shared_workbook_revisions = profile["shared_workbook_revisions"]
    if shared_workbook_revisions["present"]:
        lines.extend(
            [
                "",
                "## Legacy Shared-Workbook Revision History",
                "",
                (
                    "- **Revision header parts / headers / log parts / log entries:** "
                    f"{shared_workbook_revisions['revision_header_part_count']} / "
                    f"{shared_workbook_revisions['revision_header_count']} / "
                    f"{shared_workbook_revisions['revision_log_part_count']} / "
                    f"{shared_workbook_revisions['revision_log_entry_count']}"
                ),
                (
                    "- **Shared workbooks / tracked histories / history-enabled / "
                    "retained histories / protected histories:** "
                    f"{shared_workbook_revisions['shared_workbook_enabled_count']} / "
                    f"{shared_workbook_revisions['track_revisions_enabled_count']} / "
                    f"{shared_workbook_revisions['revision_history_enabled_count']} / "
                    f"{shared_workbook_revisions['keep_change_history_enabled_count']} / "
                    f"{shared_workbook_revisions['revision_history_protected_count']}"
                ),
            ]
        )
        if shared_workbook_revisions["unrecognized_shared_workbook_revision_count"]:
            lines.append(
                "- **Unrecognized, malformed, unresolved, or bounded revision controls:** "
                f"{shared_workbook_revisions['unrecognized_shared_workbook_revision_count']}"
            )
        lines.append(
            "Revision values, locations, authors, timestamps, comments, GUIDs, "
            "relationship identifiers, and raw XML are compared privately and omitted."
        )
    number_format_controls = profile["number_format_controls"]
    if number_format_controls["present"]:
        lines.extend(
            [
                "",
                "## Cell number-format controls",
                "",
                (
                    "- **Default overrides / direct-cell assignments / row assignments / "
                    "column assignments:** "
                    f"{number_format_controls['default_format_override_count']} / "
                    f"{number_format_controls['cell_format_assignment_count']} / "
                    f"{number_format_controls['row_format_assignment_count']} / "
                    f"{number_format_controls['column_format_assignment_count']}"
                ),
                (
                    "- **Built-in / custom format assignments:** "
                    f"{number_format_controls['built_in_format_assignment_count']} / "
                    f"{number_format_controls['custom_format_assignment_count']}"
                ),
            ]
        )
        if number_format_controls["unrecognized_number_format_count"]:
            lines.append(
                "- **Unrecognized or malformed number-format controls:** "
                f"{number_format_controls['unrecognized_number_format_count']}"
            )
        lines.append(
            "Number-format codes, style indexes, and cell/row/column targets are compared "
            "privately and intentionally omitted."
        )
    font_controls = profile["font_controls"]
    if font_controls["present"]:
        lines.extend(
            [
                "",
                "## Cell font controls",
                "",
                (
                    "- **Default definition / direct-cell assignments / row assignments / "
                    "column assignments:** "
                    f"{font_controls['default_font_definition_count']} / "
                    f"{font_controls['cell_font_assignment_count']} / "
                    f"{font_controls['row_font_assignment_count']} / "
                    f"{font_controls['column_font_assignment_count']}"
                ),
            ]
        )
        if font_controls["unrecognized_font_count"]:
            lines.append(
                "- **Unrecognized or malformed font controls:** "
                f"{font_controls['unrecognized_font_count']}"
            )
        lines.append(
            "Font definitions, style indexes, and cell/row/column targets are compared "
            "privately and intentionally omitted."
        )
    fill_controls = profile["fill_controls"]
    if fill_controls["present"]:
        lines.extend(
            [
                "",
                "## Cell fill controls",
                "",
                (
                    "- **Default definition / direct-cell assignments / row assignments / "
                    "column assignments:** "
                    f"{fill_controls['default_fill_definition_count']} / "
                    f"{fill_controls['cell_fill_assignment_count']} / "
                    f"{fill_controls['row_fill_assignment_count']} / "
                    f"{fill_controls['column_fill_assignment_count']}"
                ),
            ]
        )
        if fill_controls["unrecognized_fill_count"]:
            lines.append(
                "- **Unrecognized or malformed fill controls:** "
                f"{fill_controls['unrecognized_fill_count']}"
            )
        lines.append(
            "Fill definitions, style indexes, and cell/row/column targets are compared "
            "privately and intentionally omitted."
        )
    alignment_controls = profile["alignment_controls"]
    if alignment_controls["present"]:
        lines.extend(
            [
                "",
                "## Cell alignment controls",
                "",
                (
                    "- **Default definition / direct-cell assignments / row assignments / "
                    "column assignments:** "
                    f"{alignment_controls['default_alignment_definition_count']} / "
                    f"{alignment_controls['cell_alignment_assignment_count']} / "
                    f"{alignment_controls['row_alignment_assignment_count']} / "
                    f"{alignment_controls['column_alignment_assignment_count']}"
                ),
            ]
        )
        if alignment_controls["unrecognized_alignment_count"]:
            lines.append(
                "- **Unrecognized or malformed alignment controls:** "
                f"{alignment_controls['unrecognized_alignment_count']}"
            )
        lines.append(
            "Alignment definitions, style indexes, and cell/row/column targets are "
            "compared privately and intentionally omitted."
        )
    border_controls = profile["border_controls"]
    if border_controls["present"]:
        lines.extend(
            [
                "",
                "## Cell border controls",
                "",
                (
                    "- **Default definition / direct-cell assignments / row assignments / "
                    "column assignments:** "
                    f"{border_controls['default_border_definition_count']} / "
                    f"{border_controls['cell_border_assignment_count']} / "
                    f"{border_controls['row_border_assignment_count']} / "
                    f"{border_controls['column_border_assignment_count']}"
                ),
            ]
        )
        if border_controls["unrecognized_border_count"]:
            lines.append(
                "- **Unrecognized or malformed border controls:** "
                f"{border_controls['unrecognized_border_count']}"
            )
        lines.append(
            "Border definitions, style indexes, and cell/row/column targets are "
            "compared privately and intentionally omitted."
        )
    worksheet_dimension_controls = profile["worksheet_dimension_controls"]
    if worksheet_dimension_controls["present"]:
        lines.extend(
            [
                "",
                "## Worksheet dimension controls",
                "",
                (
                    "- **Default row-height / default column-width / baseline-adjustment / "
                    "automatic border-adjustment sheets:** "
                    f"{worksheet_dimension_controls['default_row_height_count']} / "
                    f"{worksheet_dimension_controls['default_column_width_count']} / "
                    f"{worksheet_dimension_controls['default_baseline_adjustment_sheet_count']} / "
                    f"{worksheet_dimension_controls['default_border_adjustment_sheet_count']}"
                ),
                (
                    "- **Direct row heights / row baseline adjustments / row border "
                    "adjustments / direct column widths / best-fit columns:** "
                    f"{worksheet_dimension_controls['row_height_assignment_count']} / "
                    f"{worksheet_dimension_controls['row_baseline_adjustment_count']} / "
                    f"{worksheet_dimension_controls['row_border_adjustment_count']} / "
                    f"{worksheet_dimension_controls['column_width_assignment_count']} / "
                    f"{worksheet_dimension_controls['best_fit_column_assignment_count']}"
                ),
            ]
        )
        if worksheet_dimension_controls["unrecognized_dimension_count"]:
            lines.append(
                "- **Unrecognized or malformed dimension controls:** "
                f"{worksheet_dimension_controls['unrecognized_dimension_count']}"
            )
        lines.append(
            "Dimension values, row/column targets, and raw worksheet XML are compared "
            "privately and intentionally omitted."
        )
    worksheet_display_controls = profile["worksheet_display_controls"]
    if worksheet_display_controls["present"]:
        lines.extend(
            [
                "",
                "## Worksheet display controls",
                "",
                (
                    "- **Zero-hidden / formula / gridlines-hidden / custom-gridline-color / "
                    "headers-hidden / outline-symbols-hidden views:** "
                    f"{worksheet_display_controls['zero_hidden_view_count']} / "
                    f"{worksheet_display_controls['formula_view_count']} / "
                    f"{worksheet_display_controls['gridlines_hidden_view_count']} / "
                    f"{worksheet_display_controls['custom_gridline_color_view_count']} / "
                    f"{worksheet_display_controls['headers_hidden_view_count']} / "
                    f"{worksheet_display_controls['outline_symbols_hidden_view_count']}"
                ),
                (
                    "- **Ruler-hidden / page-whitespace-hidden / right-to-left / "
                    "non-normal views / split-or-frozen panes:** "
                    f"{worksheet_display_controls['ruler_hidden_view_count']} / "
                    f"{worksheet_display_controls['white_space_hidden_view_count']} / "
                    f"{worksheet_display_controls['right_to_left_view_count']} / "
                    f"{worksheet_display_controls['non_normal_view_count']} / "
                    f"{worksheet_display_controls['split_or_frozen_pane_count']}"
                ),
            ]
        )
        if worksheet_display_controls["unrecognized_display_control_count"]:
            lines.append(
                "- **Unrecognized or malformed display controls:** "
                f"{worksheet_display_controls['unrecognized_display_control_count']}"
            )
        lines.append(
            "Sheet names, view targets, and raw display XML are compared privately "
            "and intentionally omitted."
        )
    worksheet_print_layout_controls = profile["worksheet_print_layout_controls"]
    if worksheet_print_layout_controls["present"]:
        horizontally_centered_print_sheet_count = worksheet_print_layout_controls[
            "horizontally_centered_print_sheet_count"
        ]
        vertically_centered_print_sheet_count = worksheet_print_layout_controls[
            "vertically_centered_print_sheet_count"
        ]
        lines.extend(
            [
                "",
                "## Worksheet print-layout controls",
                "",
                (
                    "- **Print-area / print-title declarations / gridline-print sheets / "
                    "heading-print sheets:** "
                    f"{worksheet_print_layout_controls['print_area_definition_count']} / "
                    f"{worksheet_print_layout_controls['print_title_definition_count']} / "
                    f"{worksheet_print_layout_controls['print_gridlines_sheet_count']} / "
                    f"{worksheet_print_layout_controls['print_headings_sheet_count']}"
                ),
                (
                    "- **Horizontally centred / vertically centred / page-margin / "
                    "page-setup / header-footer sheets:** "
                    f"{horizontally_centered_print_sheet_count} / "
                    f"{vertically_centered_print_sheet_count} / "
                    f"{worksheet_print_layout_controls['page_margin_sheet_count']} / "
                    f"{worksheet_print_layout_controls['page_setup_sheet_count']} / "
                    f"{worksheet_print_layout_controls['header_footer_sheet_count']}"
                ),
                (
                    "- **Manual row / column page breaks:** "
                    f"{worksheet_print_layout_controls['manual_row_page_break_count']} / "
                    f"{worksheet_print_layout_controls['manual_column_page_break_count']}"
                ),
            ]
        )
        if worksheet_print_layout_controls["unrecognized_print_layout_count"]:
            lines.append(
                "- **Unrecognized or malformed print-layout controls:** "
                f"{worksheet_print_layout_controls['unrecognized_print_layout_count']}"
            )
        lines.append(
            "Print ranges, header/footer text, page values, and raw "
            "print-layout XML are compared privately and intentionally omitted."
        )
    workbook_theme = profile["workbook_theme"]
    if workbook_theme["present"]:
        lines.extend(
            [
                "",
                "## Workbook theme controls",
                "",
                (
                    "- **Theme parts / colour schemes / font schemes / format schemes:** "
                    f"{workbook_theme['theme_part_count']} / "
                    f"{workbook_theme['colour_scheme_count']} / "
                    f"{workbook_theme['font_scheme_count']} / "
                    f"{workbook_theme['format_scheme_count']}"
                ),
                (
                    "- **Workbook theme relationships / external relationships:** "
                    f"{workbook_theme['theme_relationship_count']} / "
                    f"{workbook_theme['external_theme_relationship_count']}"
                ),
                (
                    "- **Theme image parts / relationships / external relationships:** "
                    f"{workbook_theme['theme_image_part_count']} / "
                    f"{workbook_theme['theme_image_relationship_count']} / "
                    f"{workbook_theme['external_theme_image_relationship_count']}"
                ),
            ]
        )
        if workbook_theme["unrecognized_theme_count"]:
            lines.append(
                "- **Unrecognized or malformed workbook-theme metadata:** "
                f"{workbook_theme['unrecognized_theme_count']}"
            )
        lines.append(
            "Theme XML, scheme names, colour values, font names, image payloads, "
            "relationship IDs, and targets are compared privately and intentionally "
            "omitted. FormulaFence does not render a workbook, resolve effective "
            "styles, decode an image, fetch a target, or infer Excel client behavior."
        )
    formula_cached_results = profile["formula_cached_results"]
    if formula_cached_results["present"]:
        lines.extend(
            [
                "",
                "## Stored formula results",
                "",
                (
                    "- **Formula cells / cached results / missing results:** "
                    f"{formula_cached_results['formula_cell_count']} / "
                    f"{formula_cached_results['cached_result_cell_count']} / "
                    f"{formula_cached_results['missing_cached_result_cell_count']}"
                ),
                (
                    "- **Cached result types (numeric / string / Boolean / error):** "
                    f"{formula_cached_results['numeric_cached_result_count']} / "
                    f"{formula_cached_results['string_cached_result_count']} / "
                    f"{formula_cached_results['boolean_cached_result_count']} / "
                    f"{formula_cached_results['error_cached_result_count']}"
                ),
            ]
        )
        if formula_cached_results["unrecognized_cached_result_count"]:
            lines.append(
                "- **Unrecognized or malformed cached results:** "
                f"{formula_cached_results['unrecognized_cached_result_count']}"
            )
        lines.append(
            "Cached result values and formula-cell locations are compared privately and "
            "intentionally omitted."
        )
    rich_text_runs = profile["rich_text_runs"]
    if rich_text_runs["present"]:
        lines.extend(
            [
                "",
                "## Rich-text run controls",
                "",
                (
                    "- **Shared items / shared cells / shared runs:** "
                    f"{rich_text_runs['shared_rich_text_item_count']} / "
                    f"{rich_text_runs['shared_rich_text_cell_count']} / "
                    f"{rich_text_runs['shared_rich_text_run_count']}"
                ),
                (
                    "- **Inline cells / inline runs:** "
                    f"{rich_text_runs['inline_rich_text_cell_count']} / "
                    f"{rich_text_runs['inline_rich_text_run_count']}"
                ),
                (
                    "- **Phonetic runs / phonetic properties:** "
                    f"{rich_text_runs['phonetic_run_count']} / "
                    f"{rich_text_runs['phonetic_property_count']}"
                ),
            ]
        )
        if rich_text_runs["unrecognized_rich_text_count"]:
            lines.append(
                "- **Unrecognized or malformed rich-text controls:** "
                f"{rich_text_runs['unrecognized_rich_text_count']}"
            )
        lines.append(
            "Character-level text, formatting, phonetic hints, shared-string indexes, "
            "and cell locations are compared privately and intentionally omitted."
        )
    cell_hyperlinks = profile["cell_hyperlinks"]
    if cell_hyperlinks["present"]:
        lines.extend(
            [
                "",
                "## Worksheet cell hyperlinks",
                "",
                (
                    "- **Worksheets / hyperlinks:** "
                    f"{cell_hyperlinks['worksheet_hyperlink_sheet_count']} / "
                    f"{cell_hyperlinks['hyperlink_count']}"
                ),
                (
                    "- **With location / display override / ScreenTip:** "
                    f"{cell_hyperlinks['hyperlink_with_location_count']} / "
                    f"{cell_hyperlinks['hyperlink_with_display_count']} / "
                    f"{cell_hyperlinks['hyperlink_with_tooltip_count']}"
                ),
                (
                    "- **Package binding relationships:** "
                    f"{cell_hyperlinks['binding_relationship_count']} "
                    f"({cell_hyperlinks['external_relationship_count']} external)"
                ),
            ]
        )
        if cell_hyperlinks["unrecognized_cell_hyperlink_count"]:
            lines.append(
                "- **Unrecognized or malformed cell-hyperlink metadata:** "
                f"{cell_hyperlinks['unrecognized_cell_hyperlink_count']}"
            )
        lines.append(
            "Hyperlink targets, cell references, locations, display strings, "
            "ScreenTips, and relationship IDs are compared privately and intentionally "
            "omitted."
        )
    worksheet_sparklines = profile["worksheet_sparklines"]
    if worksheet_sparklines["present"]:
        lines.extend(
            [
                "",
                "## Worksheet sparklines",
                "",
                (
                    "- **Worksheets / groups / sparklines:** "
                    f"{worksheet_sparklines['worksheet_sparkline_sheet_count']} / "
                    f"{worksheet_sparklines['sparkline_group_count']} / "
                    f"{worksheet_sparklines['sparkline_count']}"
                ),
                (
                    "- **With data source / date-axis source / colour controls:** "
                    f"{worksheet_sparklines['sparkline_with_source_count']} / "
                    f"{worksheet_sparklines['group_date_axis_source_count']} / "
                    f"{worksheet_sparklines['color_control_count']}"
                ),
            ]
        )
        if worksheet_sparklines["unrecognized_worksheet_sparkline_count"]:
            lines.append(
                "- **Unrecognized or malformed worksheet-sparkline metadata:** "
                f"{worksheet_sparklines['unrecognized_worksheet_sparkline_count']}"
            )
        lines.append(
            "Sparkline source formulas, destination cells, group properties, and "
            "colour definitions are compared privately and intentionally omitted."
        )
    xml_mapping_controls = profile["xml_mapping_controls"]
    if xml_mapping_controls["present"]:
        lines.extend(
            [
                "",
                "## XML-mapped workbook controls",
                "",
                (
                    "- **Map parts / schemas / maps:** "
                    f"{xml_mapping_controls['xml_map_part_count']} / "
                    f"{xml_mapping_controls['xml_schema_count']} / "
                    f"{xml_mapping_controls['xml_map_count']}"
                ),
                (
                    "- **Map data bindings (file / connection):** "
                    f"{xml_mapping_controls['xml_map_data_binding_count']} "
                    f"({xml_mapping_controls['xml_map_file_binding_count']} / "
                    f"{xml_mapping_controls['xml_map_connection_binding_count']})"
                ),
                (
                    "- **Table mapping parts / bindings:** "
                    f"{xml_mapping_controls['table_xml_binding_part_count']} / "
                    f"{xml_mapping_controls['table_xml_binding_count']}"
                ),
                (
                    "- **Worksheets / single-cell mapping parts / bindings:** "
                    f"{xml_mapping_controls['single_cell_xml_binding_sheet_count']} / "
                    f"{xml_mapping_controls['single_cell_xml_binding_part_count']} / "
                    f"{xml_mapping_controls['single_cell_xml_binding_count']}"
                ),
                (
                    "- **Single-cell connection bindings:** "
                    f"{xml_mapping_controls['single_cell_xml_connection_binding_count']}"
                ),
            ]
        )
        if xml_mapping_controls["unrecognized_xml_mapping_count"]:
            lines.append(
                "- **Unrecognized or malformed XML-mapping metadata:** "
                f"{xml_mapping_controls['unrecognized_xml_mapping_count']}"
            )
        lines.append(
            "Schemas, map names, XPath expressions, table identities, target cells, "
            "and connection identities are compared privately and intentionally omitted."
        )
    digital_signatures = profile["digital_signatures"]
    if digital_signatures["present"]:
        lines.extend(
            [
                "",
                "## Digital-signature controls",
                "",
                (
                    "- **Package signature origins / XML signatures:** "
                    f"{digital_signatures['package_signature_origin_count']} / "
                    f"{digital_signatures['package_xml_signature_count']}"
                ),
                (
                    "- **Signed references / embedded certificate values / certificate parts:** "
                    f"{digital_signatures['package_signature_reference_count']} / "
                    f"{digital_signatures['package_signature_certificate_count']} / "
                    f"{digital_signatures['package_signature_certificate_part_count']}"
                ),
                (
                    "- **Certificate-part relationships:** "
                    f"{digital_signatures['package_signature_certificate_relationship_count']}"
                ),
                (
                    "- **VBA project signature payloads / relationships:** "
                    f"{digital_signatures['vba_project_signature_count']} / "
                    f"{digital_signatures['vba_project_signature_relationship_count']}"
                ),
            ]
        )
        if digital_signatures["unrecognized_digital_signature_count"]:
            lines.append(
                "- **Unrecognized or malformed digital-signature metadata:** "
                f"{digital_signatures['unrecognized_digital_signature_count']}"
            )
        lines.append(
            "Signature XML, signed-part references, certificate identities and "
            "contents, VBA signature payloads, and relationship targets are compared "
            "privately and intentionally omitted. FormulaFence inventories envelope "
            "changes only: it does not validate cryptography, certificate trust, "
            "expiration, revocation, timestamps, or signed contents."
        )
    rich_data = profile["rich_data"]
    if rich_data["present"]:
        lines.extend(
            [
                "",
                "## Rich data controls",
                "",
                (
                    "- **Rich value data / structures / type parts:** "
                    f"{rich_data['rich_value_data_part_count']} / "
                    f"{rich_data['rich_value_structure_part_count']} / "
                    f"{rich_data['rich_value_type_part_count']}"
                ),
                (
                    "- **Rich values / structures / linked-entity structures:** "
                    f"{rich_data['rich_value_count']} / "
                    f"{rich_data['rich_value_structure_count']} / "
                    f"{rich_data['linked_entity_structure_count']}"
                ),
                (
                    "- **Arrays / supporting property bags / styles:** "
                    f"{rich_data['rich_value_array_count']} / "
                    f"{rich_data['supporting_property_bag_count']} / "
                    f"{rich_data['rich_style_part_count']}"
                ),
                (
                    "- **Metadata bindings / bound cells:** "
                    f"{rich_data['rich_value_metadata_binding_count']} / "
                    f"{rich_data['rich_value_bound_cell_count']}"
                ),
                (
                    "- **Web images / relationship references / external references:** "
                    f"{rich_data['web_image_count']} / "
                    f"{rich_data['web_image_relationship_count']} / "
                    f"{rich_data['external_web_image_relationship_count']}"
                ),
                (
                    "- **Rich-value relationship references / external references:** "
                    f"{rich_data['rich_value_relationship_reference_count']} / "
                    f"{rich_data['external_rich_value_relationship_count']}"
                ),
            ]
        )
        if rich_data["unrecognized_rich_data_count"]:
            lines.append(
                "- **Unrecognized or malformed rich-data metadata:** "
                f"{rich_data['unrecognized_rich_data_count']}"
            )
        lines.append(
            "Entity values, provider data, field names, identifiers, URLs, image "
            "references, relationship IDs, and bound-cell locations are compared "
            "privately and intentionally omitted. FormulaFence does not contact "
            "providers, refresh data, fetch image targets, or validate their content."
        )
    custom_data_stores = profile["custom_data_stores"]
    if custom_data_stores["present"]:
        lines.extend(
            [
                "",
                "## Custom workbook data stores",
                "",
                (
                    "- **Custom XML data / property parts / schema references:** "
                    f"{custom_data_stores['custom_xml_part_count']} / "
                    f"{custom_data_stores['custom_xml_property_part_count']} / "
                    f"{custom_data_stores['custom_xml_schema_reference_count']}"
                ),
                (
                    "- **Custom XML relationships / external relationships:** "
                    f"{custom_data_stores['custom_xml_relationship_count']} / "
                    f"{custom_data_stores['external_custom_xml_relationship_count']}"
                ),
                (
                    "- **Custom binary-data property parts / payload parts:** "
                    f"{custom_data_stores['custom_data_properties_part_count']} / "
                    f"{custom_data_stores['custom_data_part_count']}"
                ),
                (
                    "- **Custom document property parts / values / linked values:** "
                    f"{custom_data_stores['document_custom_property_part_count']} / "
                    f"{custom_data_stores['document_custom_property_count']} / "
                    f"{custom_data_stores['linked_document_custom_property_count']}"
                ),
            ]
        )
        if custom_data_stores["unrecognized_custom_data_store_count"]:
            lines.append(
                "- **Unrecognized or malformed custom data-store metadata:** "
                f"{custom_data_stores['unrecognized_custom_data_store_count']}"
            )
        lines.append(
            "Custom XML, document-property names and values, storage IDs, binary "
            "payloads, relationship IDs, and targets are compared privately and "
            "intentionally omitted. FormulaFence does not execute an add-in, resolve "
            "a property, fetch a target, or interpret a payload."
        )
    legacy_comments = profile["legacy_comments"]
    if legacy_comments["present"]:
        lines.extend(
            [
                "",
                "## Legacy Excel Notes and threaded placeholders",
                "",
                (
                    "- **Worksheets / comment parts / authors / comments:** "
                    f"{legacy_comments['worksheet_comment_sheet_count']} / "
                    f"{legacy_comments['comment_part_count']} / "
                    f"{legacy_comments['comment_author_count']} / "
                    f"{legacy_comments['comment_count']}"
                ),
                (
                    "- **Comments with text / rich text / phonetic hints / properties:** "
                    f"{legacy_comments['comment_with_text_count']} / "
                    f"{legacy_comments['rich_text_comment_count']} / "
                    f"{legacy_comments['phonetic_comment_count']} / "
                    f"{legacy_comments['comment_property_count']}"
                ),
                (
                    "- **Threaded-comment placeholders:** "
                    f"{legacy_comments['threaded_placeholder_count']}"
                ),
                (
                    "- **Note VML worksheets / drawings / shapes:** "
                    f"{legacy_comments['worksheet_note_drawing_sheet_count']} / "
                    f"{legacy_comments['note_vml_drawing_part_count']} / "
                    f"{legacy_comments['note_shape_count']}"
                ),
                (
                    "- **Visible / anchored Note shapes:** "
                    f"{legacy_comments['visible_note_shape_count']} / "
                    f"{legacy_comments['anchored_note_shape_count']}"
                ),
                (
                    "- **Package binding relationships:** "
                    f"{legacy_comments['binding_relationship_count']} "
                    f"({legacy_comments['external_relationship_count']} external)"
                ),
            ]
        )
        if legacy_comments["unrecognized_legacy_comment_count"]:
            lines.append(
                "- **Unrecognized or malformed Note metadata:** "
                f"{legacy_comments['unrecognized_legacy_comment_count']}"
            )
        lines.append(
            "Note text, authors, cell references, placeholder IDs, VML markup, and "
            "layout declarations are compared privately and intentionally omitted."
        )
    threaded_comments = profile["threaded_comments"]
    if threaded_comments["present"]:
        lines.extend(
            [
                "",
                "## Modern threaded comments",
                "",
                (
                    "- **Worksheets / threaded-comment parts / threads:** "
                    f"{threaded_comments['worksheet_threaded_comment_sheet_count']} / "
                    f"{threaded_comments['threaded_comment_part_count']} / "
                    f"{threaded_comments['comment_thread_count']}"
                ),
                (
                    "- **Comments (replies / resolved / with text):** "
                    f"{threaded_comments['comment_count']} "
                    f"({threaded_comments['reply_count']} / "
                    f"{threaded_comments['resolved_comment_count']} / "
                    f"{threaded_comments['comment_with_text_count']})"
                ),
                (
                    "- **Mentions / mentioned people:** "
                    f"{threaded_comments['mention_count']} / "
                    f"{threaded_comments['mentioned_person_count']}"
                ),
                (
                    "- **Person parts / people / unreferenced people:** "
                    f"{threaded_comments['person_part_count']} / "
                    f"{threaded_comments['person_count']} / "
                    f"{threaded_comments['orphan_person_count']}"
                ),
                (
                    "- **Package binding relationships:** "
                    f"{threaded_comments['binding_relationship_count']} "
                    f"({threaded_comments['external_relationship_count']} external)"
                ),
            ]
        )
        if threaded_comments["unrecognized_threaded_comment_count"]:
            lines.append(
                "- **Unrecognized or malformed threaded-comment metadata:** "
                f"{threaded_comments['unrecognized_threaded_comment_count']}"
            )
        lines.append(
            "Comment text, cell references, timestamps, reply links, person identities, "
            "and raw IDs are compared privately and intentionally omitted."
        )
    worksheet_drawing_shapes = profile["worksheet_drawing_shapes"]
    if worksheet_drawing_shapes["present"]:
        lines.extend(
            [
                "",
                "## Worksheet DrawingML shape, connector, and graphic-frame controls",
                "",
                (
                    "- **Worksheets / drawing parts / shape anchors:** "
                    f"{worksheet_drawing_shapes['worksheet_drawing_sheet_count']} / "
                    f"{worksheet_drawing_shapes['worksheet_drawing_part_count']} / "
                    f"{worksheet_drawing_shapes['shape_anchor_count']}"
                ),
                (
                    "- **Shapes (text-bearing / connectors / grouped):** "
                    f"{worksheet_drawing_shapes['shape_count']} "
                    f"({worksheet_drawing_shapes['text_shape_count']} / "
                    f"{worksheet_drawing_shapes['connector_shape_count']} / "
                    f"{worksheet_drawing_shapes['group_shape_count']})"
                ),
                (
                    "- **Non-chart graphic frames / SmartArt diagrams:** "
                    f"{worksheet_drawing_shapes['graphic_frame_count']} / "
                    f"{worksheet_drawing_shapes['diagram_graphic_frame_count']}"
                ),
                (
                    "- **SmartArt data / layouts / quick styles / colours / renderings:** "
                    f"{worksheet_drawing_shapes['diagram_data_part_count']} / "
                    f"{worksheet_drawing_shapes['diagram_layout_part_count']} / "
                    f"{worksheet_drawing_shapes['diagram_quick_style_part_count']} / "
                    f"{worksheet_drawing_shapes['diagram_colour_part_count']} / "
                    f"{worksheet_drawing_shapes['diagram_drawing_part_count']}"
                ),
                (
                    "- **SmartArt Diagram Data images (parts / fingerprinted / "
                    "uninspected):** "
                    f"{worksheet_drawing_shapes['diagram_image_part_count']} / "
                    f"{worksheet_drawing_shapes['fingerprinted_diagram_image_part_count']} / "
                    f"{worksheet_drawing_shapes['uninspected_diagram_image_part_count']}"
                ),
                (
                    "- **Connector attachments:** "
                    f"{worksheet_drawing_shapes['connector_attachment_count']}"
                ),
                (
                    "- **Text paragraphs / runs:** "
                    f"{worksheet_drawing_shapes['text_paragraph_count']} / "
                    f"{worksheet_drawing_shapes['text_run_count']}"
                ),
                (
                    "- **Macro assignments / text links / hyperlinks:** "
                    f"{worksheet_drawing_shapes['macro_assignment_count']} / "
                    f"{worksheet_drawing_shapes['text_link_count']} / "
                    f"{worksheet_drawing_shapes['hyperlink_count']}"
                ),
                (
                    "- **Related package relationships:** "
                    f"{worksheet_drawing_shapes['related_relationship_count']} "
                    f"({worksheet_drawing_shapes['external_relationship_count']} external)"
                ),
            ]
        )
        if worksheet_drawing_shapes["unrecognized_shape_count"]:
            lines.append(
                "- **Unrecognized or malformed shape controls:** "
                f"{worksheet_drawing_shapes['unrecognized_shape_count']}"
            )
        if worksheet_drawing_shapes["unrecognized_graphic_frame_count"]:
            lines.append(
                "- **Unsupported non-chart graphic frames:** "
                f"{worksheet_drawing_shapes['unrecognized_graphic_frame_count']}"
            )
        lines.append(
            "Shape, connector, and SmartArt presentation; attachment targets; anchors; "
            "diagram content and bounded Diagram Data image payloads; macro assignments; "
            "text links; hyperlink targets; and raw XML are compared privately and "
            "intentionally omitted."
        )
    worksheet_images = profile["worksheet_images"]
    if worksheet_images["present"]:
        lines.extend(
            [
                "",
                "## Native worksheet image controls",
                "",
                (
                    "- **Worksheets / anchored pictures / picture anchors:** "
                    f"{worksheet_images['worksheet_image_sheet_count']} / "
                    f"{worksheet_images['anchored_picture_count']} / "
                    f"{worksheet_images['anchored_picture_anchor_count']}"
                ),
                (
                    "- **Sheet backgrounds / header-footer images:** "
                    f"{worksheet_images['worksheet_background_image_count']} / "
                    f"{worksheet_images['header_footer_image_count']}"
                ),
                (
                    "- **Image parts (fingerprinted / uninspected):** "
                    f"{worksheet_images['image_part_count']} "
                    f"({worksheet_images['fingerprinted_image_part_count']} / "
                    f"{worksheet_images['uninspected_image_part_count']})"
                ),
                (
                    "- **Related package relationships:** "
                    f"{worksheet_images['related_relationship_count']} "
                    f"({worksheet_images['external_relationship_count']} external)"
                ),
            ]
        )
        if worksheet_images["unrecognized_image_count"]:
            lines.append(
                "- **Unrecognized or malformed image controls:** "
                f"{worksheet_images['unrecognized_image_count']}"
            )
        lines.append(
            "Image bytes, names, descriptions, visual formatting, anchors, relationship "
            "targets, and raw XML are compared privately and intentionally omitted."
        )
    chart_definitions = profile["chart_definitions"]
    if chart_definitions["present"]:
        lines.extend(
            [
                "",
                "## Chart definitions and cached presentation data",
                "",
                (
                    "- **Chart host sheets / drawing parts / legacy refs / ChartEx refs:** "
                    f"{chart_definitions['chart_host_sheet_count']} / "
                    f"{chart_definitions['chart_drawing_part_count']} / "
                    f"{chart_definitions['chart_reference_count']} / "
                    f"{chart_definitions['chart_ex_reference_count']}"
                ),
                f"- **Chart parts:** {chart_definitions['chart_part_count']}",
                f"- **Office 2016+ ChartEx parts:** {chart_definitions['chart_ex_part_count']}",
                (
                    "- **ChartEx series / titles / data references:** "
                    f"{chart_definitions['chart_ex_series_count']} / "
                    f"{chart_definitions['chart_ex_title_count']} / "
                    f"{chart_definitions['chart_ex_data_reference_count']}"
                ),
                (
                    "- **Overlay parts / shapes:** "
                    f"{chart_definitions['chart_user_shape_part_count']} / "
                    f"{chart_definitions['chart_user_shape_count']}"
                ),
                f"- **Chart-type elements:** {chart_definitions['chart_type_count']}",
                f"- **Series:** {chart_definitions['series_count']}",
                f"- **Titles:** {chart_definitions['title_count']}",
                (
                    "- **Data references (numeric / string): "
                    f"{chart_definitions['data_reference_count']} "
                    f"({chart_definitions['numeric_data_reference_count']} / "
                    f"{chart_definitions['string_data_reference_count']})"
                ),
                (
                    "- **Literal / cached data points:** "
                    f"{chart_definitions['literal_data_point_count']} / "
                    f"{chart_definitions['cached_data_point_count']}"
                ),
                (
                    "- **Pivot / external-data / overlay references:** "
                    f"{chart_definitions['pivot_source_count']} / "
                    f"{chart_definitions['external_data_reference_count']} / "
                    f"{chart_definitions['user_shape_reference_count']}"
                ),
                (
                    "- **Related package relationships:** "
                    f"{chart_definitions['related_relationship_count']} "
                    f"({chart_definitions['external_relationship_count']} external)"
                ),
                (
                    "- **Internal direct related parts:** "
                    f"{chart_definitions['internal_related_part_count']} "
                    f"({chart_definitions['fingerprinted_related_part_count']} "
                    "fingerprinted)"
                ),
            ]
        )
        if chart_definitions["uninspected_related_part_count"]:
            lines.append(
                "- **Uninspected direct related parts:** "
                f"{chart_definitions['uninspected_related_part_count']}"
            )
        if chart_definitions["unrecognized_part_count"]:
            lines.append(
                "- **Unrecognized or uninspected chart parts/bindings:** "
                f"{chart_definitions['unrecognized_part_count']}"
            )
        lines.append(
            "Chart formulas, labels, cached values, formatting, overlay text, relationship "
            "targets, and direct related-part contents are compared privately and intentionally "
            "omitted."
        )
    worksheet_embedded_controls = profile["worksheet_embedded_controls"]
    if worksheet_embedded_controls["present"]:
        lines.extend(
            [
                "",
                "## Worksheet embedded controls, legacy VML controls, and OLE objects",
                "",
                (
                    "- **Control-bearing worksheets:** "
                    f"{worksheet_embedded_controls['control_sheet_count']}"
                ),
                (
                    "- **Worksheet controls:** "
                    f"{worksheet_embedded_controls['worksheet_control_count']}"
                ),
                (
                    "- **ActiveX persistence parts:** "
                    f"{worksheet_embedded_controls['active_x_part_count']}"
                ),
                (
                    "- **ActiveX binary references:** "
                    f"{worksheet_embedded_controls['active_x_binary_reference_count']}"
                ),
                (
                    "- **Form-control properties parts:** "
                    f"{worksheet_embedded_controls['form_control_property_part_count']}"
                ),
                (
                    "- **Legacy VML drawing parts / controls:** "
                    f"{worksheet_embedded_controls['legacy_vml_drawing_part_count']} / "
                    f"{worksheet_embedded_controls['legacy_vml_control_count']}"
                ),
                (
                    "- **Legacy VML macro assignments:** "
                    f"{worksheet_embedded_controls['legacy_vml_macro_assignment_count']}"
                ),
                (
                    "- **Legacy VML cell links / source ranges / camera ranges:** "
                    f"{worksheet_embedded_controls['legacy_vml_cell_link_count']} / "
                    f"{worksheet_embedded_controls['legacy_vml_source_range_count']} / "
                    f"{worksheet_embedded_controls['legacy_vml_camera_source_range_count']}"
                ),
                (
                    "- **Control macro assignments:** "
                    f"{worksheet_embedded_controls['control_macro_assignment_count']}"
                ),
                (
                    "- **Control cell links:** "
                    f"{worksheet_embedded_controls['control_cell_link_count']}"
                ),
                (
                    "- **Control source-range bindings:** "
                    f"{worksheet_embedded_controls['control_source_range_count']}"
                ),
                (
                    "- **Form-control formula bindings:** "
                    f"{worksheet_embedded_controls['form_control_formula_binding_count']}"
                ),
                (
                    "- **OLE objects:** "
                    f"{worksheet_embedded_controls['ole_object_count']} "
                    f"({worksheet_embedded_controls['linked_ole_object_count']} linked)"
                ),
                (
                    "- **OLE auto-load / auto-update requests:** "
                    f"{worksheet_embedded_controls['auto_load_ole_object_count']} / "
                    f"{worksheet_embedded_controls['auto_update_ole_object_count']}"
                ),
                (
                    "- **Related package relationships:** "
                    f"{worksheet_embedded_controls['related_relationship_count']} "
                    f"({worksheet_embedded_controls['external_relationship_count']} external)"
                ),
                (
                    "- **Internal direct payloads:** "
                    f"{worksheet_embedded_controls['internal_related_part_count']} "
                    f"({worksheet_embedded_controls['fingerprinted_related_part_count']} "
                    "fingerprinted)"
                ),
            ]
        )
        if worksheet_embedded_controls["uninspected_related_part_count"]:
            lines.append(
                "- **Uninspected direct payloads:** "
                f"{worksheet_embedded_controls['uninspected_related_part_count']}"
            )
        if worksheet_embedded_controls["unrecognized_part_count"]:
            lines.append(
                "- **Unrecognized or uninspected parts/bindings:** "
                f"{worksheet_embedded_controls['unrecognized_part_count']}"
            )
        lines.append(
            "Control names, class IDs, VML and modern-control macros, linked "
            "formulas/ranges, OLE identities, relationship targets, and direct payloads "
            "are compared privately and intentionally omitted."
        )
    power_query = profile["power_query"]
    if power_query["present"]:
        permission_controls = power_query["permission_controls"]
        lines.extend(
            [
                "",
                "## Power Query controls",
                "",
                (
                    "- **Data Mashup custom XML parts:** "
                    f"{power_query['mashup_count']} "
                    f"({power_query['parsed_mashup_count']} structurally parsed)"
                ),
                (
                    "- **Formula documents:** "
                    f"{power_query['formula_document_count']} across "
                    f"{power_query['package_part_count']} package part(s)"
                ),
                (
                    "- **Embedded content parts:** "
                    f"{power_query['embedded_content_part_count']}"
                ),
                (
                    "- **Query metadata records:** "
                    f"{power_query['metadata_item_count']} in "
                    f"{power_query['metadata_document_count']} document(s)"
                ),
                (
                    "- **Formula-firewall enabled:** "
                    f"{permission_controls['firewall_enabled_count']} of "
                    f"{permission_controls['payload_count']} permission payload(s)"
                ),
                (
                    "- **Future package evaluation allowed:** "
                    f"{permission_controls['future_packages_allowed_count']}"
                ),
                (
                    "- **Permission-binding payloads:** "
                    f"{power_query['permission_binding_count']}"
                ),
            ]
        )
        if permission_controls["opaque_metadata"]["present"]:
            lines.append(
                "- **Unmodelled permission items:** "
                f"{permission_controls['opaque_metadata']['count']}"
            )
        if power_query["opaque_metadata"]["present"]:
            lines.append(
                "- **Unmodelled Data Mashup items:** "
                f"{power_query['opaque_metadata']['count']}"
            )
        lines.append(
            "M formulas, query names, source locations, metadata values, embedded content, "
            "telemetry IDs, and user-bound permission bindings are intentionally omitted."
        )
    features = profile["features"]
    if features["external_reference_cells"] or features["broken_reference_cells"]:
        lines.extend(["", "## Static hazards", ""])
        for location in features["external_reference_cells"]:
            lines.append(f"- Explicit external reference: `{location}`")
        for location in features["broken_reference_cells"]:
            lines.append(f"- Broken `#REF!` formula: `{location}`")
    if features["three_d_reference_cells"]:
        lines.extend(["", "## 3-D worksheet references", ""])
        for issue in features["three_d_reference_cells"]:
            tokens = ", ".join(f"`{token}`" for token in issue["tokens"])
            lines.append(f"- Static 3-D reference at `{issue['location']}`: {tokens}")
    if features["spill_reference_cells"]:
        lines.extend(["", "## Dynamic-array spill references", ""])
        for issue in features["spill_reference_cells"]:
            tokens = ", ".join(f"`{token}`" for token in issue["tokens"])
            lines.append(
                f"- Spill reference at `{issue['location']}`: {tokens} "
                "(the anchor is traced; dynamic extent and blockers are coverage limits)"
            )
    if features["implicit_intersection_cells"]:
        lines.extend(["", "## Explicit implicit intersection", ""])
        for issue in features["implicit_intersection_cells"]:
            tokens = ", ".join(f"`{token}`" for token in issue["tokens"])
            lines.append(
                f"- Implicit intersection at `{issue['location']}`: {tokens} "
                "(direct static A1 ranges resolve to their selected cell; other "
                "expressions retain conservative input edges)"
            )
    if features["legacy_array_formula_ranges"]:
        lines.extend(
            [
                "",
                "## Legacy CSE array formulas",
                "",
                "| Anchor | Fixed output range | Output cells |",
                "| --- | --- | ---: |",
            ]
        )
        for array_range in features["legacy_array_formula_ranges"]:
            lines.append(
                "| {anchor} | {ref} | {output_cell_count} |".format(
                    anchor=_markdown_escape(array_range["anchor"]),
                    ref=_markdown_escape(array_range["ref"]),
                    output_cell_count=array_range["output_cell_count"],
                )
            )
        lines.append(
            "Fixed result members are linked back to their anchor when a static "
            "formula reads them; the range is not expanded into virtual cells."
        )
    if features["dynamic_array_formula_cells"]:
        lines.extend(
            [
                "",
                "## Dynamic-array formula anchors",
                "",
                "| Anchor | Observed output range | Observed output cells |",
                "| --- | --- | ---: |",
            ]
        )
        observed_ranges = {
            array_range["anchor"]: array_range
            for array_range in features["dynamic_array_observed_output_ranges"]
        }
        for location in features["dynamic_array_formula_cells"]:
            array_range = observed_ranges.get(location)
            if array_range is None:
                lines.append(f"| {_markdown_escape(location)} | unavailable | — |")
                continue
            lines.append(
                "| {anchor} | {ref} | {output_cell_count} |".format(
                    anchor=_markdown_escape(location),
                    ref=_markdown_escape(array_range["ref"]),
                    output_cell_count=array_range["output_cell_count"],
                )
            )
        lines.append(
            "The output range is observed from this workbook, not fixed: Excel can "
            "resize it during recalculation. FormulaFence only links a static formula "
            "back to the anchor when it currently reads a non-anchor output member."
        )
    if features["dynamic_array_output_reference_cells"]:
        lines.extend(
            [
                "",
                "## Observed dynamic-array output-member references",
                "",
                "| Formula | Dynamic anchor | Observed spill range |",
                "| --- | --- | --- |",
            ]
        )
        for issue in features["dynamic_array_output_reference_cells"]:
            for reference in issue["references"]:
                lines.append(
                    "| {location} | {anchor} | {observed_range} |".format(
                        location=_markdown_escape(issue["location"]),
                        anchor=_markdown_escape(reference["anchor"]),
                        observed_range=_markdown_escape(reference["observed_range"]),
                    )
                )
        lines.append(
            "These graph aliases explain the currently observed relationship; a future "
            "recalculation can grow, shrink, or block the spill."
        )
    if (
        features["parser_warnings"]
        or features["unresolved_reference_cells"]
        or features["dynamic_reference_cells"]
        or features["unclassified_array_formula_cells"]
        or features["tokenization_failure_cells"]
    ):
        lines.extend(["", "## Inspection coverage notes", ""])
        for warning in features["parser_warnings"]:
            lines.append(f"- {_markdown_escape(warning)}")
        for issue in features["unresolved_reference_cells"]:
            tokens = ", ".join(f"`{token}`" for token in issue["tokens"])
            lines.append(
                f"- Non-static formula reference at `{issue['location']}`: {tokens}"
            )
        for issue in features["dynamic_reference_cells"]:
            functions = ", ".join(f"`{function}`" for function in issue["functions"])
            lines.append(
                f"- Dynamic reference function at `{issue['location']}`: {functions}"
            )
        for location in features["unclassified_array_formula_cells"]:
            lines.append(
                f"- Array formula at `{location}` could not be classified as fixed CSE "
                "or dynamic; fixed-output aliases were not added"
            )
        for location in features["tokenization_failure_cells"]:
            lines.append(
                f"- Formula tokenizer could not inspect `{location}`; "
                "dependency impact may be incomplete"
            )
    return "\n".join(lines) + "\n"


def _append_report_markdown_sections(
    lines: list[str], payload: dict[str, Any], heading: str
) -> None:
    """Append reusable finding/change tables at the requested heading level."""
    if payload["findings"]:
        lines.extend(
            [
                f"{heading} Findings",
                "",
                "| Severity | Rule | Location | Finding |",
                "| --- | --- | --- | --- |",
            ]
        )
        for finding in payload["findings"]:
            lines.append(
                "| {severity} | `{rule_id}` | `{location}` | {message} |".format(
                    severity=_markdown_escape(finding["severity"]),
                    rule_id=_markdown_escape(finding["rule_id"]),
                    location=_markdown_escape(finding["location"] or "workbook"),
                    message=_markdown_escape(finding["message"]),
                )
            )
        lines.append("")
    lines.extend(
        [
            f"{heading} Semantic changes",
            "",
            "| Risk | Change | Location | Downstream formulas |",
            "| --- | --- | --- | ---: |",
        ]
    )
    if not payload["changes"]:
        lines.append("| note | no semantic changes | — | 0 |")
    else:
        for change in payload["changes"]:
            lines.append(
                "| {severity} | `{kind}` | `{location}` | {impact_count} |".format(
                    severity=_markdown_escape(change["severity"]),
                    kind=_markdown_escape(change["kind"]),
                    location=_markdown_escape(change["location"] or "workbook"),
                    impact_count=change["impact_count"],
                )
            )
    impacted = [change for change in payload["changes"] if change["impacted_cells"]]
    if impacted:
        lines.extend(["", f"{heading} Impact samples", ""])
        for change in impacted:
            sample = ", ".join(f"`{cell}`" for cell in change["impacted_cells"])
            lines.append(f"- `{change['location']}` affects: {sample}")
        path_samples = [
            (change["location"], path)
            for change in impacted
            for path in change["details"].get("impact_paths", [])
        ]
        if path_samples:
            lines.extend(["", f"{heading} Dependency paths", ""])
            for _, path_sample in path_samples:
                path = " → ".join(f"`{step}`" for step in path_sample["path"])
                lines.append(f"- {path}")


def report_to_markdown(report: DiffReport, extra_findings: Iterable[Finding] = ()) -> str:
    policy_findings = list(extra_findings)
    payload = report.to_dict(policy_findings)
    summary = payload["summary"]
    lines = [
        "# FormulaFence change report",
        "",
        f"- **Baseline:** `{payload['before']['path']}`",
        f"- **Candidate:** `{payload['after']['path']}`",
        f"- **Changes:** {summary['change_count']}",
        f"- **Findings:** {summary['finding_count']}",
        f"- **Highest severity:** `{summary['highest_severity']}`",
        "",
    ]
    _append_report_markdown_sections(lines, payload, "##")
    return "\n".join(lines) + "\n"


_SARIF_LEVELS = {
    "critical": "error",
    "high": "error",
    "medium": "warning",
    "low": "warning",
    "note": "note",
}


def report_to_sarif(report: DiffReport, extra_findings: Iterable[Finding] = ()) -> dict[str, Any]:
    findings = [*report.findings, *extra_findings]
    rule_ids = sorted({finding.rule_id for finding in findings})
    rules = [
        {
            "id": rule_id,
            "name": rule_id,
            "shortDescription": {"text": "FormulaFence spreadsheet-control finding"},
        }
        for rule_id in rule_ids
    ]
    results: list[dict[str, Any]] = []
    for finding in findings:
        result: dict[str, Any] = {
            "ruleId": finding.rule_id,
            "level": _SARIF_LEVELS.get(finding.severity, "warning"),
            "message": {"text": finding.message},
            "properties": {"severity": finding.severity, **finding.details},
        }
        if finding.location is not None:
            result["locations"] = [
                {
                    "physicalLocation": {
                        "artifactLocation": {"uri": str(report.after.path)},
                    },
                    "logicalLocations": [
                        {
                            "kind": "excel-cell",
                            "name": display_location(finding.location),
                        }
                    ],
                }
            ]
        results.append(result)
    return {
        "$schema": "https://json.schemastore.org/sarif-2.1.0.json",
        "version": "2.1.0",
        "runs": [
            {
                "tool": {
                    "driver": {
                        "name": "FormulaFence",
                        "version": __version__,
                        "informationUri": "https://github.com/SybilGambleyyu/formulafence",
                        "rules": rules,
                    }
                },
                "results": results,
            }
        ],
    }


def _markdown_code(value: object) -> str:
    """Render a short untrusted logical path safely inside a Markdown code span."""
    escaped = _markdown_escape(value).replace("`", "\\`")
    return f"`{escaped}`"


def _append_cross_workbook_impact_samples(
    lines: list[str], entry: dict[str, Any], heading: str
) -> None:
    """Render safe portfolio impact paths without exposing external link spellings."""
    findings = [
        finding for finding in entry["findings"] if finding["rule_id"] == "FF079"
    ]
    if not findings:
        return
    lines.extend([f"{heading} Static cross-workbook impacts", ""])
    for finding in findings:
        details = finding.get("details", {})
        lines.append(
            "- Source {location} reaches {formula_count} formula(s) in {workbook_count} "
            "other workbook(s).".format(
                location=_markdown_code(finding["location"] or "workbook"),
                formula_count=details["impacted_formula_count"],
                workbook_count=details["impacted_workbook_count"],
            )
        )
        for impact in details["sample_impacts"]:
            path = " → ".join(
                "{} {}".format(
                    _markdown_code(step["workbook"]),
                    _markdown_code(step["location"]),
                )
                for step in impact["path"]
            )
            lines.append(f"  - {path}")
    lines.append("")


def portfolio_to_markdown(report: PortfolioReport) -> str:
    """Render a complete multi-workbook review without absolute filesystem paths."""
    payload = report.to_dict()
    summary = payload["summary"]
    lines = [
        "# FormulaFence portfolio report",
        "",
        (
            "- **Baseline / candidate workbooks:** "
            f"{payload['before']['workbook_count']} / {payload['after']['workbook_count']}"
        ),
        f"- **Matched workbooks:** {summary['matched_workbook_count']}",
        (
            "- **Added / removed / unreadable:** "
            f"{summary['added_workbook_count']} / {summary['removed_workbook_count']} / "
            f"{summary['unreadable_workbook_count']}"
        ),
        f"- **Semantic changes:** {summary['change_count']}",
        (
            "- **Cross-workbook impact sources / impacted formulas:** "
            f"{summary['cross_workbook_impact_source_count']} / "
            f"{summary['cross_workbook_impacted_formula_count']}"
        ),
        (
            "- **Cross-workbook impact analysis:** `"
            f"{'incomplete' if summary['cross_workbook_impact_incomplete'] else 'complete'}`"
        ),
        f"- **Findings:** {summary['finding_count']}",
        f"- **Highest severity:** `{summary['highest_severity']}`",
        "",
        (
            "Relative workbook paths are the comparison identity. A move is deliberately "
            "reported as a removal plus an addition rather than inferred as a rename."
        ),
        "",
        "## Workbook summary",
        "",
        "| Workbook | Status | Changes | Findings | Highest severity |",
        "| --- | --- | ---: | ---: | --- |",
    ]
    for entry in payload["workbooks"]:
        entry_summary = entry["summary"]
        lines.append(
            "| {path} | `{status}` | {changes} | {findings} | `{severity}` |".format(
                path=_markdown_code(entry["path"]),
                status=_markdown_escape(entry["status"]),
                changes=entry_summary["change_count"],
                findings=entry_summary["finding_count"],
                severity=_markdown_escape(entry_summary["highest_severity"]),
            )
        )

    lines.extend(["", "## Workbook details", ""])
    for entry in payload["workbooks"]:
        entry_summary = entry["summary"]
        lines.extend(
            [
                f"### {_markdown_code(entry['path'])}",
                "",
                f"- **Status:** `{_markdown_escape(entry['status'])}`",
                (
                    "- **Baseline / candidate present:** "
                    f"{'yes' if entry['baseline_present'] else 'no'} / "
                    f"{'yes' if entry['candidate_present'] else 'no'}"
                ),
                f"- **Changes / findings:** {entry_summary['change_count']} / "
                f"{entry_summary['finding_count']}",
                "",
            ]
        )
        _append_report_markdown_sections(lines, entry, "####")
        _append_cross_workbook_impact_samples(lines, entry, "####")
        lines.append("")
    return "\n".join(lines) + "\n"


def portfolio_to_sarif(report: PortfolioReport) -> dict[str, Any]:
    """Render every portfolio finding in one SARIF run with relative artifacts."""
    rule_ids = sorted(
        {
            finding.rule_id
            for workbook in report.workbooks
            for finding in workbook.findings
        }
    )
    rules = [
        {
            "id": rule_id,
            "name": rule_id,
            "shortDescription": {"text": "FormulaFence spreadsheet-control finding"},
        }
        for rule_id in rule_ids
    ]
    results: list[dict[str, Any]] = []
    for workbook in report.workbooks:
        for finding in workbook.findings:
            location: dict[str, Any] = {
                "physicalLocation": {"artifactLocation": {"uri": workbook.path}},
            }
            if finding.location is not None:
                location["logicalLocations"] = [
                    {
                        "kind": "excel-cell",
                        "name": display_location(finding.location),
                    }
                ]
            results.append(
                {
                    "ruleId": finding.rule_id,
                    "level": _SARIF_LEVELS.get(finding.severity, "warning"),
                    "message": {"text": finding.message},
                    "properties": {
                        **finding.details,
                        "severity": finding.severity,
                        "portfolio_status": workbook.status,
                    },
                    "locations": [location],
                }
            )
    return {
        "$schema": "https://json.schemastore.org/sarif-2.1.0.json",
        "version": "2.1.0",
        "runs": [
            {
                "tool": {
                    "driver": {
                        "name": "FormulaFence",
                        "version": __version__,
                        "informationUri": "https://github.com/SybilGambleyyu/formulafence",
                        "rules": rules,
                    }
                },
                "results": results,
            }
        ],
    }
