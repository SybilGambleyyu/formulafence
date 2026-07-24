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
        f"- **3-D reference formulas:** {workbook['three_d_reference_cells']}",
        f"- **Spill-reference formulas:** {workbook['spill_reference_cells']}",
        (
            "- **Implicit-intersection formulas:** "
            f"{workbook['implicit_intersection_cells']}"
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
    if (
        features["parser_warnings"]
        or features["unresolved_reference_cells"]
        or features["dynamic_reference_cells"]
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
