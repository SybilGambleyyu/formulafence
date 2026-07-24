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
