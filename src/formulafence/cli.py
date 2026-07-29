"""Command-line interface for FormulaFence."""

from __future__ import annotations

import argparse
import os
import sys
import tempfile
from collections.abc import Sequence
from pathlib import Path

from formulafence import __version__
from formulafence.diff import (
    DEFAULT_MAX_CHANGE_ANALYSIS_STATES,
    compare_snapshots,
    report_severities,
)
from formulafence.lint import (
    DEFAULT_MAX_AGGREGATE_OMISSION_GAP_CELLS,
    DEFAULT_MAX_FORMULA_PATTERN_FINDINGS,
    lint_snapshot,
)
from formulafence.models import SEVERITY_ORDER, FormulaFenceError
from formulafence.output import (
    DEFAULT_MAX_REPORT_BYTES,
    as_json,
    lint_to_markdown,
    lint_to_sarif,
    portfolio_to_html,
    portfolio_to_markdown,
    portfolio_to_sarif,
    profile_to_markdown,
    redact_external_workbook_link_material,
    redact_formula_defined_xlm_action_portfolio_payload,
    redact_formula_defined_xlm_action_report_payload,
    redact_formula_defined_xlm_environment_information_portfolio_payload,
    redact_formula_defined_xlm_environment_information_report_payload,
    redact_formula_defined_xlm_evaluation_portfolio_payload,
    redact_formula_defined_xlm_evaluation_report_payload,
    redact_formula_defined_xlm_get_cell_portfolio_payload,
    redact_formula_defined_xlm_get_cell_report_payload,
    redact_formula_defined_xlm_registration_portfolio_payload,
    redact_formula_defined_xlm_registration_report_payload,
    redact_formula_environment_information_portfolio_payload,
    redact_formula_environment_information_report_payload,
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
    report_to_html,
    report_to_markdown,
    report_to_sarif,
)
from formulafence.policy import DEFAULT_POLICY, evaluate_policy, load_policy
from formulafence.portfolio import (
    DEFAULT_MAX_INVENTORY_ENTRIES,
    DEFAULT_MAX_LINK_IMPACT,
    DEFAULT_MAX_PORTFOLIO_SNAPSHOT_CELLS,
    DEFAULT_MAX_PORTFOLIO_SOURCE_BYTES,
    compare_portfolios,
)
from formulafence.workbook import (
    DEFAULT_MAX_DEPENDENCY_EDGES,
    DEFAULT_MAX_FORMULA_DEFINED_NAME_STATES,
    DEFAULT_MAX_PROFILE_RECORDS,
    load_snapshot,
    profile_snapshot,
)

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


def _add_formula_defined_xlm_action_redaction_argument(
    parser: argparse.ArgumentParser,
) -> None:
    parser.add_argument(
        "--redact-formula-defined-xlm-actions",
        action="store_true",
        help=(
            "Replace visible formula-defined XLM action material and known static "
            "action inputs in this shared report without changing comparison "
            "or policy results"
        ),
    )


def _add_formula_defined_xlm_get_cell_redaction_argument(
    parser: argparse.ArgumentParser,
) -> None:
    parser.add_argument(
        "--redact-formula-defined-xlm-get-cell-calls",
        action="store_true",
        help=(
            "Replace visible formula-defined XLM GET.CELL material and known static "
            "GET.CELL inputs in this shared report without changing comparison "
            "or policy results"
        ),
    )


def _add_formula_defined_xlm_environment_information_redaction_argument(
    parser: argparse.ArgumentParser,
) -> None:
    parser.add_argument(
        "--redact-formula-defined-xlm-environment-information-calls",
        action="store_true",
        help=(
            "Replace visible formula-defined XLM environment-information material "
            "and known static inputs in this shared report without changing comparison "
            "or policy results"
        ),
    )


def _add_formula_environment_information_redaction_argument(
    parser: argparse.ArgumentParser,
) -> None:
    parser.add_argument(
        "--redact-formula-environment-information",
        action="store_true",
        help=(
            "Replace visible native CELL/INFO/SHEET/SHEETS material and known "
            "static inputs in this shared report without changing comparison "
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


def _at_least_two_integer(value: str) -> int:
    parsed = _positive_integer(value)
    if parsed < 2:
        raise argparse.ArgumentTypeError("must be an integer of at least 2")
    return parsed


def _add_change_analysis_state_limit_argument(parser: argparse.ArgumentParser) -> None:
    parser.add_argument(
        "--max-change-analysis-states",
        type=_positive_integer,
        default=DEFAULT_MAX_CHANGE_ANALYSIS_STATES,
        help=(
            "Fail closed after this many aggregate changed-source and local dependency "
            "states are analyzed"
        ),
    )


def _add_dependency_edge_limit_argument(parser: argparse.ArgumentParser) -> None:
    parser.add_argument(
        "--max-dependency-edges",
        type=_positive_integer,
        default=DEFAULT_MAX_DEPENDENCY_EDGES,
        help=(
            "Fail closed before FormulaFence retains more than this many static "
            "local dependency-graph edges"
        ),
    )


def _add_formula_defined_name_state_limit_argument(
    parser: argparse.ArgumentParser,
) -> None:
    parser.add_argument(
        "--max-formula-defined-name-states",
        type=_positive_integer,
        default=DEFAULT_MAX_FORMULA_DEFINED_NAME_STATES,
        help=(
            "Fail closed before FormulaFence retains more than this many "
            "formula-defined-name propagation states"
        ),
    )


def _add_report_byte_limit_argument(parser: argparse.ArgumentParser) -> None:
    parser.add_argument(
        "--max-report-bytes",
        type=_positive_integer,
        default=DEFAULT_MAX_REPORT_BYTES,
        help=(
            "Fail closed before a rendered artifact exceeds this many UTF-8 bytes"
        ),
    )


def _add_profile_record_limit_argument(parser: argparse.ArgumentParser) -> None:
    parser.add_argument(
        "--max-profile-records",
        type=_positive_integer,
        default=DEFAULT_MAX_PROFILE_RECORDS,
        help=(
            "Fail closed before a profile materializes more than this many "
            "aggregate inventory records"
        ),
    )


def _add_formula_pattern_finding_limit_argument(parser: argparse.ArgumentParser) -> None:
    parser.add_argument(
        "--max-formula-pattern-findings",
        type=_positive_integer,
        default=DEFAULT_MAX_FORMULA_PATTERN_FINDINGS,
        help=(
            "Fail closed before formula lint retains more than this many "
            "review findings"
        ),
    )


def _add_aggregate_omission_gap_limit_argument(parser: argparse.ArgumentParser) -> None:
    parser.add_argument(
        "--max-aggregate-omission-gap-cells",
        type=_at_least_two_integer,
        default=DEFAULT_MAX_AGGREGATE_OMISSION_GAP_CELLS,
        help=(
            "Inspect at most this many contiguous cells between a local aggregate "
            "range and its formula"
        ),
    )


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
    _add_profile_record_limit_argument(profile)
    _add_dependency_edge_limit_argument(profile)
    _add_formula_defined_name_state_limit_argument(profile)
    _add_report_byte_limit_argument(profile)
    _add_output_arguments(profile, ("json", "markdown"))

    lint = commands.add_parser(
        "lint",
        help="Find conservative formula, aggregate-range, protection, and calculation risks",
    )
    lint.add_argument("workbook", type=Path)
    _add_dependency_edge_limit_argument(lint)
    _add_formula_defined_name_state_limit_argument(lint)
    _add_formula_pattern_finding_limit_argument(lint)
    _add_aggregate_omission_gap_limit_argument(lint)
    _add_report_byte_limit_argument(lint)
    _add_output_arguments(lint, ("json", "markdown", "sarif"))
    lint.add_argument(
        "--fail-on",
        choices=_FAIL_LEVELS,
        default="none",
        help="Exit 1 when a formula-lint finding reaches this severity",
    )

    diff = commands.add_parser("diff", help="Compare two workbooks semantically")
    diff.add_argument("before", type=Path, help="Approved or baseline workbook")
    diff.add_argument("after", type=Path, help="Candidate workbook")
    _add_dependency_edge_limit_argument(diff)
    _add_formula_defined_name_state_limit_argument(diff)
    _add_change_analysis_state_limit_argument(diff)
    _add_report_byte_limit_argument(diff)
    _add_output_arguments(diff, ("json", "markdown", "html", "sarif"))
    _add_external_workbook_link_redaction_argument(diff)
    _add_formula_external_action_redaction_argument(diff)
    _add_python_in_excel_redaction_argument(diff)
    _add_office_custom_function_redaction_argument(diff)
    _add_unqualified_runtime_function_redaction_argument(diff)
    _add_worksheet_code_resource_registration_redaction_argument(diff)
    _add_formula_defined_xlm_registration_redaction_argument(diff)
    _add_formula_defined_xlm_evaluation_redaction_argument(diff)
    _add_formula_defined_xlm_action_redaction_argument(diff)
    _add_formula_defined_xlm_get_cell_redaction_argument(diff)
    _add_formula_defined_xlm_environment_information_redaction_argument(diff)
    _add_formula_environment_information_redaction_argument(diff)
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
    _add_dependency_edge_limit_argument(check)
    _add_formula_defined_name_state_limit_argument(check)
    _add_change_analysis_state_limit_argument(check)
    _add_report_byte_limit_argument(check)
    _add_output_arguments(check, ("json", "markdown", "html", "sarif"))
    _add_external_workbook_link_redaction_argument(check)
    _add_formula_external_action_redaction_argument(check)
    _add_python_in_excel_redaction_argument(check)
    _add_office_custom_function_redaction_argument(check)
    _add_unqualified_runtime_function_redaction_argument(check)
    _add_worksheet_code_resource_registration_redaction_argument(check)
    _add_formula_defined_xlm_registration_redaction_argument(check)
    _add_formula_defined_xlm_evaluation_redaction_argument(check)
    _add_formula_defined_xlm_action_redaction_argument(check)
    _add_formula_defined_xlm_get_cell_redaction_argument(check)
    _add_formula_defined_xlm_environment_information_redaction_argument(check)
    _add_formula_environment_information_redaction_argument(check)
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
        "--max-inventory-entries",
        type=_positive_integer,
        default=DEFAULT_MAX_INVENTORY_ENTRIES,
        help=(
            "Fail when either portfolio contains more than this many filesystem "
            "entries before workbook filtering"
        ),
    )
    portfolio.add_argument(
        "--max-portfolio-source-bytes",
        type=_positive_integer,
        default=DEFAULT_MAX_PORTFOLIO_SOURCE_BYTES,
        help=(
            "Fail when either portfolio's supported workbook sources exceed this "
            "aggregate byte budget before snapshot reads"
        ),
    )
    portfolio.add_argument(
        "--max-portfolio-snapshot-cells",
        type=_positive_integer,
        default=DEFAULT_MAX_PORTFOLIO_SNAPSHOT_CELLS,
        help=(
            "Fail when either portfolio retains more than this many populated "
            "workbook snapshot cells during comparison"
        ),
    )
    _add_dependency_edge_limit_argument(portfolio)
    _add_formula_defined_name_state_limit_argument(portfolio)
    _add_change_analysis_state_limit_argument(portfolio)
    _add_report_byte_limit_argument(portfolio)
    portfolio.add_argument(
        "--max-link-impact",
        type=_positive_integer,
        default=DEFAULT_MAX_LINK_IMPACT,
        help=(
            "Fail closed after this many static cross-workbook dependency graph states"
        ),
    )
    _add_output_arguments(portfolio, ("json", "markdown", "html", "sarif"))
    _add_external_workbook_link_redaction_argument(portfolio)
    _add_formula_external_action_redaction_argument(portfolio)
    _add_python_in_excel_redaction_argument(portfolio)
    _add_office_custom_function_redaction_argument(portfolio)
    _add_unqualified_runtime_function_redaction_argument(portfolio)
    _add_worksheet_code_resource_registration_redaction_argument(portfolio)
    _add_formula_defined_xlm_registration_redaction_argument(portfolio)
    _add_formula_defined_xlm_evaluation_redaction_argument(portfolio)
    _add_formula_defined_xlm_action_redaction_argument(portfolio)
    _add_formula_defined_xlm_get_cell_redaction_argument(portfolio)
    _add_formula_defined_xlm_environment_information_redaction_argument(portfolio)
    _add_formula_environment_information_redaction_argument(portfolio)
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


def _write_private_text_file(path: Path, content: str) -> Path:
    """Write complete text to a private sibling file and return its pathname."""
    descriptor: int | None = None
    temporary_path: Path | None = None
    try:
        descriptor, temporary_name = tempfile.mkstemp(
            prefix="formulafence-output-",
            suffix=path.suffix,
            dir=path.parent,
        )
        temporary_path = Path(temporary_name)
        payload = memoryview(content.encode("utf-8"))
        while payload:
            written = os.write(descriptor, payload)
            if written <= 0:
                raise OSError("Could not write FormulaFence output.")
            payload = payload[written:]
        os.close(descriptor)
        descriptor = None
        assert temporary_path is not None
        return temporary_path
    except BaseException:
        if temporary_path is not None:
            try:
                temporary_path.unlink(missing_ok=True)
            except OSError:
                pass
        raise
    finally:
        if descriptor is not None:
            try:
                os.close(descriptor)
            except OSError:
                pass


def _publish_text_atomically(path: Path, content: str, *, replace: bool) -> None:
    """Publish complete text by replacing, or atomically claiming, a final name."""
    temporary_path = _write_private_text_file(path, content)
    try:
        if replace:
            os.replace(temporary_path, path)
        else:
            try:
                # ``link`` creates the final directory entry only when it is
                # absent.  Unlike a preflight ``exists`` check, that condition
                # is tested by the filesystem at publication time.
                os.link(temporary_path, path)
            except FileExistsError as error:
                raise FormulaFenceError(
                    f"Refusing to replace existing policy: {path} "
                    "(use --force to replace it)"
                ) from error
    finally:
        if temporary_path is not None:
            try:
                temporary_path.unlink(missing_ok=True)
            except OSError:
                pass


def _replace_text_atomically(path: Path, content: str) -> None:
    """Publish text by replacing the final pathname instead of following it."""
    _publish_text_atomically(path, content, replace=True)


def _create_text_atomically(path: Path, content: str) -> None:
    """Publish text only when the final pathname remains absent."""
    _publish_text_atomically(path, content, replace=False)


def _emit(content: str, output: Path | None) -> None:
    if output is None:
        sys.stdout.write(content)
        return
    _replace_text_atomically(output, content)


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
    profile = profile_snapshot(
        load_snapshot(
            arguments.workbook,
            max_dependency_edges=arguments.max_dependency_edges,
            max_formula_defined_name_states=(
                arguments.max_formula_defined_name_states
            ),
        ),
        max_profile_records=arguments.max_profile_records,
    )
    content = (
        as_json(profile, max_bytes=arguments.max_report_bytes)
        if arguments.format == "json"
        else profile_to_markdown(profile, max_bytes=arguments.max_report_bytes)
    )
    _emit(content, arguments.output)
    return 0


def _run_lint(arguments: argparse.Namespace) -> int:
    _ensure_output_safe(arguments.output, arguments.workbook)
    report = lint_snapshot(
        load_snapshot(
            arguments.workbook,
            max_dependency_edges=arguments.max_dependency_edges,
            max_formula_defined_name_states=(
                arguments.max_formula_defined_name_states
            ),
        ),
        max_formula_pattern_findings=arguments.max_formula_pattern_findings,
        max_aggregate_omission_gap_cells=(
            arguments.max_aggregate_omission_gap_cells
        ),
    )
    if arguments.format == "json":
        content = as_json(report.to_dict(), max_bytes=arguments.max_report_bytes)
    elif arguments.format == "sarif":
        content = as_json(lint_to_sarif(report), max_bytes=arguments.max_report_bytes)
    else:
        content = lint_to_markdown(report, max_bytes=arguments.max_report_bytes)
    _emit(content, arguments.output)
    return int(
        _threshold_failed(
            [finding.severity for finding in report.findings],
            arguments.fail_on,
        )
    )


def _run_comparison(arguments: argparse.Namespace, enforce_policy: bool) -> int:
    inputs = [arguments.before, arguments.after]
    if enforce_policy:
        inputs.append(arguments.policy)
    _ensure_output_safe(arguments.output, *inputs)
    policy = load_policy(arguments.policy) if enforce_policy else None
    report = compare_snapshots(
        load_snapshot(
            arguments.before,
            max_dependency_edges=arguments.max_dependency_edges,
            max_formula_defined_name_states=(
                arguments.max_formula_defined_name_states
            ),
        ),
        load_snapshot(
            arguments.after,
            max_dependency_edges=arguments.max_dependency_edges,
            max_formula_defined_name_states=(
                arguments.max_formula_defined_name_states
            ),
        ),
        max_change_analysis_states=arguments.max_change_analysis_states,
    )
    policy_findings = []
    if policy is not None:
        policy_findings = evaluate_policy(report, policy)

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
        if arguments.redact_formula_defined_xlm_actions:
            payload = redact_formula_defined_xlm_action_report_payload(report, payload)
        if arguments.redact_formula_defined_xlm_get_cell_calls:
            payload = redact_formula_defined_xlm_get_cell_report_payload(report, payload)
        if arguments.redact_formula_defined_xlm_environment_information_calls:
            payload = redact_formula_defined_xlm_environment_information_report_payload(
                report, payload
            )
        if arguments.redact_formula_environment_information:
            payload = redact_formula_environment_information_report_payload(
                report, payload
            )
        content = as_json(payload, max_bytes=arguments.max_report_bytes)
    elif arguments.format == "html":
        content = report_to_html(
            report,
            policy_findings,
            max_bytes=arguments.max_report_bytes,
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
            redact_formula_defined_xlm_actions=(
                arguments.redact_formula_defined_xlm_actions
            ),
            redact_formula_defined_xlm_get_cell_calls=(
                arguments.redact_formula_defined_xlm_get_cell_calls
            ),
            redact_formula_defined_xlm_environment_information_calls=(
                arguments.redact_formula_defined_xlm_environment_information_calls
            ),
            redact_formula_environment_information=(
                arguments.redact_formula_environment_information
            ),
        )
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
                redact_formula_defined_xlm_actions=(
                    arguments.redact_formula_defined_xlm_actions
                ),
                redact_formula_defined_xlm_get_cell_calls=(
                    arguments.redact_formula_defined_xlm_get_cell_calls
                ),
                redact_formula_defined_xlm_environment_information_calls=(
                    arguments.redact_formula_defined_xlm_environment_information_calls
                ),
                redact_formula_environment_information=(
                    arguments.redact_formula_environment_information
                ),
            ),
            max_bytes=arguments.max_report_bytes,
        )
    else:
        content = report_to_markdown(
            report,
            policy_findings,
            max_bytes=arguments.max_report_bytes,
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
            redact_formula_defined_xlm_actions=(
                arguments.redact_formula_defined_xlm_actions
            ),
            redact_formula_defined_xlm_get_cell_calls=(
                arguments.redact_formula_defined_xlm_get_cell_calls
            ),
            redact_formula_defined_xlm_environment_information_calls=(
                arguments.redact_formula_defined_xlm_environment_information_calls
            ),
            redact_formula_environment_information=(
                arguments.redact_formula_environment_information
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
        max_inventory_entries=arguments.max_inventory_entries,
        max_portfolio_source_bytes=arguments.max_portfolio_source_bytes,
        max_portfolio_snapshot_cells=arguments.max_portfolio_snapshot_cells,
        max_dependency_edges=arguments.max_dependency_edges,
        max_formula_defined_name_states=arguments.max_formula_defined_name_states,
        max_change_analysis_states=arguments.max_change_analysis_states,
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
        if arguments.redact_formula_defined_xlm_actions:
            payload = redact_formula_defined_xlm_action_portfolio_payload(
                report, payload
            )
        if arguments.redact_formula_defined_xlm_get_cell_calls:
            payload = redact_formula_defined_xlm_get_cell_portfolio_payload(
                report, payload
            )
        if arguments.redact_formula_defined_xlm_environment_information_calls:
            payload = redact_formula_defined_xlm_environment_information_portfolio_payload(
                report, payload
            )
        if arguments.redact_formula_environment_information:
            payload = redact_formula_environment_information_portfolio_payload(
                report, payload
            )
        content = as_json(payload, max_bytes=arguments.max_report_bytes)
    elif arguments.format == "html":
        content = portfolio_to_html(
            report,
            max_bytes=arguments.max_report_bytes,
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
            redact_formula_defined_xlm_actions=(
                arguments.redact_formula_defined_xlm_actions
            ),
            redact_formula_defined_xlm_get_cell_calls=(
                arguments.redact_formula_defined_xlm_get_cell_calls
            ),
            redact_formula_defined_xlm_environment_information_calls=(
                arguments.redact_formula_defined_xlm_environment_information_calls
            ),
            redact_formula_environment_information=(
                arguments.redact_formula_environment_information
            ),
        )
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
                redact_formula_defined_xlm_actions=(
                    arguments.redact_formula_defined_xlm_actions
                ),
                redact_formula_defined_xlm_get_cell_calls=(
                    arguments.redact_formula_defined_xlm_get_cell_calls
                ),
                redact_formula_defined_xlm_environment_information_calls=(
                    arguments.redact_formula_defined_xlm_environment_information_calls
                ),
                redact_formula_environment_information=(
                    arguments.redact_formula_environment_information
                ),
            ),
            max_bytes=arguments.max_report_bytes,
        )
    else:
        content = portfolio_to_markdown(
            report,
            max_bytes=arguments.max_report_bytes,
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
            redact_formula_defined_xlm_actions=(
                arguments.redact_formula_defined_xlm_actions
            ),
            redact_formula_defined_xlm_get_cell_calls=(
                arguments.redact_formula_defined_xlm_get_cell_calls
            ),
            redact_formula_defined_xlm_environment_information_calls=(
                arguments.redact_formula_defined_xlm_environment_information_calls
            ),
            redact_formula_environment_information=(
                arguments.redact_formula_environment_information
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
    if arguments.force:
        _replace_text_atomically(arguments.path, DEFAULT_POLICY)
    else:
        _create_text_atomically(arguments.path, DEFAULT_POLICY)
    sys.stdout.write(f"Wrote starter FormulaFence policy to {arguments.path}\n")
    return 0


def main(argv: Sequence[str] | None = None) -> int:
    parser = build_parser()
    arguments = parser.parse_args(argv)
    try:
        if arguments.command == "profile":
            return _run_profile(arguments)
        if arguments.command == "lint":
            return _run_lint(arguments)
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
