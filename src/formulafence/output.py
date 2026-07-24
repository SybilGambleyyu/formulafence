"""Deterministic JSON, Markdown, and SARIF renderers for review artifacts."""

from __future__ import annotations

import json
from collections.abc import Iterable
from typing import Any

from formulafence import __version__
from formulafence.models import DiffReport, Finding, display_location


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
    if payload["findings"]:
        lines.extend(
            [
                "## Findings",
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
            "## Semantic changes",
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
        lines.extend(["", "## Impact samples", ""])
        for change in impacted:
            sample = ", ".join(f"`{cell}`" for cell in change["impacted_cells"])
            lines.append(f"- `{change['location']}` affects: {sample}")
        path_samples = [
            (change["location"], path)
            for change in impacted
            for path in change["details"].get("impact_paths", [])
        ]
        if path_samples:
            lines.extend(["", "## Dependency paths", ""])
            for _, path_sample in path_samples:
                path = " → ".join(f"`{step}`" for step in path_sample["path"])
                lines.append(f"- {path}")
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
