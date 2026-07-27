"""Command-line interface for FormulaFence."""

from __future__ import annotations

import argparse
import sys
from collections.abc import Sequence
from pathlib import Path

from formulafence import __version__
from formulafence.diff import compare_snapshots, report_severities
from formulafence.models import SEVERITY_ORDER, FormulaFenceError
from formulafence.output import (
    as_json,
    portfolio_to_markdown,
    portfolio_to_sarif,
    profile_to_markdown,
    redact_external_workbook_link_material,
    redact_formula_defined_xlm_evaluation_portfolio_payload,
    redact_formula_defined_xlm_evaluation_report_payload,
    redact_formula_defined_xlm_registration_portfolio_payload,
    redact_formula_defined_xlm_registration_report_payload,
    redact_formula_external_action_portfolio_payload,
    redact_formula_external_action_report_payload,
    redact_office_custom_function_portfolio_payload,
    redact_office_custom_function_report_payload,
    redact_python_in_excel_portfolio_payload,
    redact_python_in_excel_report_payload,
    redact_unqualified_runtime_function_portfolio_payload,
    redact_unqualified_runtime_function_report_payload,
    redact_worksheet_code_resource_registration_portfolio_payload,
    redact_worksheet_code_resource_registration_report_payload,
    report_to_markdown,
    report_to_sarif,
)
from formulafence.policy import DEFAULT_POLICY, evaluate_policy, load_policy
from formulafence.portfolio import DEFAULT_MAX_LINK_IMPACT, compare_portfolios
from formulafence.workbook import load_snapshot, profile_snapshot

_FAIL_LEVELS = ("none", "low", "medium", "high", "critical")


def _add_output_arguments(parser: argparse.ArgumentParser, formats: Sequence[str]) -> None:
    parser.add_argument("--format", choices=formats, default="markdown", help="Report format")
    parser.add_argument("--output", type=Path, help="Write report to this file instead of stdout")


def _add_external_workbook_link_redaction_argument(parser: argparse.ArgumentParser) -> None:
    parser.add_argument(
        "--redact-external-workbook-links",
        action="store_true",
        help=(
            "Replace visible literal external-workbook link material in this shared "
            "report without changing comparison or policy results"
        ),
    )


def _add_formula_external_action_redaction_argument(parser: argparse.ArgumentParser) -> None:
    parser.add_argument(
        "--redact-formula-external-actions",
        action="store_true",
        help=(
            "Replace visible formula external-action/DDE material and known static "
            "action inputs in this shared report without changing comparison or policy results"
        ),
    )


def _add_python_in_excel_redaction_argument(parser: argparse.ArgumentParser) -> None:
    parser.add_argument(
        "--redact-python-in-excel",
        action="store_true",
        help=(
            "Replace visible Python-in-Excel PY source and known static PY inputs "
            "in this shared report without changing comparison or policy results"
        ),
    )


def _add_office_custom_function_redaction_argument(
    parser: argparse.ArgumentParser,
) -> None:
    parser.add_argument(
        "--redact-office-custom-functions",
        action="store_true",
        help=(
            "Replace visible Office custom-function material and known static "
            "custom-function inputs in this shared report without changing comparison "
            "or policy results"
        ),
    )


def _add_unqualified_runtime_function_redaction_argument(
    parser: argparse.ArgumentParser,
) -> None:
    parser.add_argument(
        "--redact-unqualified-runtime-functions",
        action="store_true",
        help=(
            "Replace visible unqualified runtime-function material and known static "
            "runtime-function inputs in this shared report without changing comparison "
            "or policy results"
        ),
    )


def _add_worksheet_code_resource_registration_redaction_argument(
    parser: argparse.ArgumentParser,
) -> None:
    parser.add_argument(
        "--redact-worksheet-code-resource-registrations",
        action="store_true",
        help=(
            "Replace visible worksheet code-resource registration material and known "
            "static registration inputs in this shared report without changing comparison "
            "or policy results"
        ),
    )


def _add_formula_defined_xlm_registration_redaction_argument(
    parser: argparse.ArgumentParser,
) -> None:
    parser.add_argument(
        "--redact-formula-defined-xlm-registrations",
        action="store_true",
        help=(
            "Replace visible formula-defined XLM registration material and known "
            "static registration inputs in this shared report without changing comparison "
            "or policy results"
        ),
    )


def _add_formula_defined_xlm_evaluation_redaction_argument(
    parser: argparse.ArgumentParser,
) -> None:
    parser.add_argument(
        "--redact-formula-defined-xlm-evaluations",
        action="store_true",
        help=(
            "Replace visible formula-defined XLM evaluation material and known "
            "static evaluation inputs in this shared report without changing comparison "
            "or policy results"
        ),
    )


def _positive_integer(value: str) -> int:
    try:
        parsed = int(value)
    except ValueError as error:
        raise argparse.ArgumentTypeError("must be a positive integer") from error
    if parsed < 1:
        raise argparse.ArgumentTypeError("must be a positive integer")
    return parsed


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="formulafence",
        description="Local-first spreadsheet change assurance for CI and audit workflows.",
    )
    parser.add_argument("--version", action="version", version=f"FormulaFence {__version__}")
    commands = parser.add_subparsers(dest="command", required=True)

    profile = commands.add_parser(
        "profile", help="Inventory a workbook without exposing cell values"
    )
    profile.add_argument("workbook", type=Path)
    _add_output_arguments(profile, ("json", "markdown"))

    diff = commands.add_parser("diff", help="Compare two workbooks semantically")
    diff.add_argument("before", type=Path, help="Approved or baseline workbook")
    diff.add_argument("after", type=Path, help="Candidate workbook")
    _add_output_arguments(diff, ("json", "markdown", "sarif"))
    _add_external_workbook_link_redaction_argument(diff)
    _add_formula_external_action_redaction_argument(diff)
    _add_python_in_excel_redaction_argument(diff)
    _add_office_custom_function_redaction_argument(diff)
    _add_unqualified_runtime_function_redaction_argument(diff)
    _add_worksheet_code_resource_registration_redaction_argument(diff)
    _add_formula_defined_xlm_registration_redaction_argument(diff)
    _add_formula_defined_xlm_evaluation_redaction_argument(diff)
    diff.add_argument(
        "--fail-on",
        choices=_FAIL_LEVELS,
        default="none",
        help="Exit 1 when a semantic change or finding reaches this severity",
    )

    check = commands.add_parser("check", help="Compare workbooks and enforce a FormulaFence policy")
    check.add_argument("before", type=Path, help="Approved or baseline workbook")
    check.add_argument("after", type=Path, help="Candidate workbook")
    check.add_argument("--policy", required=True, type=Path, help="Path to formulafence.yml")
    _add_output_arguments(check, ("json", "markdown", "sarif"))
    _add_external_workbook_link_redaction_argument(check)
    _add_formula_external_action_redaction_argument(check)
    _add_python_in_excel_redaction_argument(check)
    _add_office_custom_function_redaction_argument(check)
    _add_unqualified_runtime_function_redaction_argument(check)
    _add_worksheet_code_resource_registration_redaction_argument(check)
    _add_formula_defined_xlm_registration_redaction_argument(check)
    _add_formula_defined_xlm_evaluation_redaction_argument(check)
    check.add_argument(
        "--fail-on",
        choices=_FAIL_LEVELS,
        default="none",
        help="Also exit 1 when a non-policy change or finding reaches this severity",
    )

    portfolio = commands.add_parser(
        "portfolio",
        help="Compare recursively matched .xlsx/.xlsm workbook directories",
    )
    portfolio.add_argument("before", type=Path, help="Approved baseline workbook directory")
    portfolio.add_argument("after", type=Path, help="Candidate workbook directory")
    portfolio.add_argument(
        "--policy",
        type=Path,
        help="Optional FormulaFence policy applied independently to matched workbooks",
    )
    portfolio.add_argument(
        "--max-workbooks",
        type=_positive_integer,
        default=512,
        help="Fail when either portfolio contains more than this many supported workbooks",
    )
    portfolio.add_argument(
        "--max-link-impact",
        type=_positive_integer,
        default=DEFAULT_MAX_LINK_IMPACT,
        help=(
            "Fail closed after this many static cross-workbook dependency graph states"
        ),
    )
    _add_output_arguments(portfolio, ("json", "markdown", "sarif"))
    _add_external_workbook_link_redaction_argument(portfolio)
    _add_formula_external_action_redaction_argument(portfolio)
    _add_python_in_excel_redaction_argument(portfolio)
    _add_office_custom_function_redaction_argument(portfolio)
    _add_unqualified_runtime_function_redaction_argument(portfolio)
    _add_worksheet_code_resource_registration_redaction_argument(portfolio)
    _add_formula_defined_xlm_registration_redaction_argument(portfolio)
    _add_formula_defined_xlm_evaluation_redaction_argument(portfolio)
    portfolio.add_argument(
        "--fail-on",
        choices=_FAIL_LEVELS,
        default="none",
        help="Exit 1 when a portfolio change or finding reaches this severity",
    )

    init = commands.add_parser("init", help="Write a documented starter policy")
    init.add_argument("path", nargs="?", default=Path("formulafence.yml"), type=Path)
    init.add_argument("--force", action="store_true", help="Replace an existing file")
    return parser


def _emit(content: str, output: Path | None) -> None:
    if output is None:
        sys.stdout.write(content)
        return
    output.write_text(content, encoding="utf-8")


def _ensure_output_safe(
    output: Path | None,
    *inputs: Path,
    directory_inputs: bool = False,
) -> None:
    """Keep a report write from mutating an inspected artifact or portfolio."""
    if output is None:
        return
    resolved_output = output.resolve()
    for source in inputs:
        resolved_source = source.resolve()
        if directory_inputs:
            try:
                resolved_output.relative_to(resolved_source)
            except ValueError:
                continue
            raise FormulaFenceError(
                "Refusing to write a portfolio report inside an input directory: "
                f"{output}"
            )
        if resolved_output == resolved_source:
            raise FormulaFenceError(f"Refusing to overwrite an input: {output}")


def _threshold_failed(severities: Sequence[str], fail_on: str) -> bool:
    if fail_on == "none":
        return False
    threshold = SEVERITY_ORDER[fail_on]
    return any(SEVERITY_ORDER[severity] >= threshold for severity in severities)


def _run_profile(arguments: argparse.Namespace) -> int:
    _ensure_output_safe(arguments.output, arguments.workbook)
    profile = profile_snapshot(load_snapshot(arguments.workbook))
    content = as_json(profile) if arguments.format == "json" else profile_to_markdown(profile)
    _emit(content, arguments.output)
    return 0


def _run_comparison(arguments: argparse.Namespace, enforce_policy: bool) -> int:
    inputs = [arguments.before, arguments.after]
    if enforce_policy:
        inputs.append(arguments.policy)
    _ensure_output_safe(arguments.output, *inputs)
    report = compare_snapshots(load_snapshot(arguments.before), load_snapshot(arguments.after))
    policy_findings = []
    if enforce_policy:
        policy_findings = evaluate_policy(report, load_policy(arguments.policy))

    if arguments.format == "json":
        payload = report.to_dict(policy_findings)
        if arguments.redact_external_workbook_links:
            payload = redact_external_workbook_link_material(payload)
        if arguments.redact_formula_external_actions:
            payload = redact_formula_external_action_report_payload(report, payload)
        if arguments.redact_python_in_excel:
            payload = redact_python_in_excel_report_payload(report, payload)
        if arguments.redact_office_custom_functions:
            payload = redact_office_custom_function_report_payload(report, payload)
        if arguments.redact_unqualified_runtime_functions:
            payload = redact_unqualified_runtime_function_report_payload(report, payload)
        if arguments.redact_worksheet_code_resource_registrations:
            payload = redact_worksheet_code_resource_registration_report_payload(
                report, payload
            )
        if arguments.redact_formula_defined_xlm_registrations:
            payload = redact_formula_defined_xlm_registration_report_payload(
                report, payload
            )
        if arguments.redact_formula_defined_xlm_evaluations:
            payload = redact_formula_defined_xlm_evaluation_report_payload(
                report, payload
            )
        content = as_json(payload)
    elif arguments.format == "sarif":
        content = as_json(
            report_to_sarif(
                report,
                policy_findings,
                redact_external_workbook_links=arguments.redact_external_workbook_links,
                redact_formula_external_actions=arguments.redact_formula_external_actions,
                redact_python_in_excel=arguments.redact_python_in_excel,
                redact_office_custom_functions=arguments.redact_office_custom_functions,
                redact_unqualified_runtime_functions=(
                    arguments.redact_unqualified_runtime_functions
                ),
                redact_worksheet_code_resource_registrations=(
                    arguments.redact_worksheet_code_resource_registrations
                ),
                redact_formula_defined_xlm_registrations=(
                    arguments.redact_formula_defined_xlm_registrations
                ),
                redact_formula_defined_xlm_evaluations=(
                    arguments.redact_formula_defined_xlm_evaluations
                ),
            )
        )
    else:
        content = report_to_markdown(
            report,
            policy_findings,
            redact_external_workbook_links=arguments.redact_external_workbook_links,
            redact_formula_external_actions=arguments.redact_formula_external_actions,
            redact_python_in_excel=arguments.redact_python_in_excel,
            redact_office_custom_functions=arguments.redact_office_custom_functions,
            redact_unqualified_runtime_functions=(
                arguments.redact_unqualified_runtime_functions
            ),
            redact_worksheet_code_resource_registrations=(
                arguments.redact_worksheet_code_resource_registrations
            ),
            redact_formula_defined_xlm_registrations=(
                arguments.redact_formula_defined_xlm_registrations
            ),
            redact_formula_defined_xlm_evaluations=(
                arguments.redact_formula_defined_xlm_evaluations
            ),
        )
    _emit(content, arguments.output)

    if policy_findings or _threshold_failed(
        report_severities(report, policy_findings), arguments.fail_on
    ):
        return 1
    return 0


def _run_portfolio(arguments: argparse.Namespace) -> int:
    _ensure_output_safe(
        arguments.output,
        arguments.before,
        arguments.after,
        directory_inputs=True,
    )
    if arguments.policy is not None:
        _ensure_output_safe(arguments.output, arguments.policy)
    policy = load_policy(arguments.policy) if arguments.policy is not None else None
    report = compare_portfolios(
        arguments.before,
        arguments.after,
        policy=policy,
        max_workbooks=arguments.max_workbooks,
        max_link_impact=arguments.max_link_impact,
    )
    if arguments.format == "json":
        payload = report.to_dict()
        if arguments.redact_external_workbook_links:
            payload = redact_external_workbook_link_material(payload)
        if arguments.redact_formula_external_actions:
            payload = redact_formula_external_action_portfolio_payload(report, payload)
        if arguments.redact_python_in_excel:
            payload = redact_python_in_excel_portfolio_payload(report, payload)
        if arguments.redact_office_custom_functions:
            payload = redact_office_custom_function_portfolio_payload(report, payload)
        if arguments.redact_unqualified_runtime_functions:
            payload = redact_unqualified_runtime_function_portfolio_payload(
                report, payload
            )
        if arguments.redact_worksheet_code_resource_registrations:
            payload = redact_worksheet_code_resource_registration_portfolio_payload(
                report, payload
            )
        if arguments.redact_formula_defined_xlm_registrations:
            payload = redact_formula_defined_xlm_registration_portfolio_payload(
                report, payload
            )
        if arguments.redact_formula_defined_xlm_evaluations:
            payload = redact_formula_defined_xlm_evaluation_portfolio_payload(
                report, payload
            )
        content = as_json(payload)
    elif arguments.format == "sarif":
        content = as_json(
            portfolio_to_sarif(
                report,
                redact_external_workbook_links=arguments.redact_external_workbook_links,
                redact_formula_external_actions=arguments.redact_formula_external_actions,
                redact_python_in_excel=arguments.redact_python_in_excel,
                redact_office_custom_functions=arguments.redact_office_custom_functions,
                redact_unqualified_runtime_functions=(
                    arguments.redact_unqualified_runtime_functions
                ),
                redact_worksheet_code_resource_registrations=(
                    arguments.redact_worksheet_code_resource_registrations
                ),
                redact_formula_defined_xlm_registrations=(
                    arguments.redact_formula_defined_xlm_registrations
                ),
                redact_formula_defined_xlm_evaluations=(
                    arguments.redact_formula_defined_xlm_evaluations
                ),
            )
        )
    else:
        content = portfolio_to_markdown(
            report,
            redact_external_workbook_links=arguments.redact_external_workbook_links,
            redact_formula_external_actions=arguments.redact_formula_external_actions,
            redact_python_in_excel=arguments.redact_python_in_excel,
            redact_office_custom_functions=arguments.redact_office_custom_functions,
            redact_unqualified_runtime_functions=(
                arguments.redact_unqualified_runtime_functions
            ),
            redact_worksheet_code_resource_registrations=(
                arguments.redact_worksheet_code_resource_registrations
            ),
            redact_formula_defined_xlm_registrations=(
                arguments.redact_formula_defined_xlm_registrations
            ),
            redact_formula_defined_xlm_evaluations=(
                arguments.redact_formula_defined_xlm_evaluations
            ),
        )
    _emit(content, arguments.output)

    if report.incomplete:
        return 2
    if report.policy_findings or _threshold_failed(report.severities(), arguments.fail_on):
        return 1
    return 0


def _run_init(arguments: argparse.Namespace) -> int:
    if arguments.path.exists() and not arguments.force:
        raise FormulaFenceError(
            f"Refusing to replace existing policy: {arguments.path} (use --force to replace it)"
        )
    arguments.path.write_text(DEFAULT_POLICY, encoding="utf-8")
    sys.stdout.write(f"Wrote starter FormulaFence policy to {arguments.path}\n")
    return 0


def main(argv: Sequence[str] | None = None) -> int:
    parser = build_parser()
    arguments = parser.parse_args(argv)
    try:
        if arguments.command == "profile":
            return _run_profile(arguments)
        if arguments.command == "diff":
            return _run_comparison(arguments, enforce_policy=False)
        if arguments.command == "check":
            return _run_comparison(arguments, enforce_policy=True)
        if arguments.command == "portfolio":
            return _run_portfolio(arguments)
        if arguments.command == "init":
            return _run_init(arguments)
    except (FormulaFenceError, OSError) as error:
        print(f"formulafence: error: {error}", file=sys.stderr)
        return 2
    parser.error(f"Unknown command: {arguments.command}")
    return 2  # pragma: no cover - argparse exits above


if __name__ == "__main__":  # pragma: no cover
    raise SystemExit(main())
