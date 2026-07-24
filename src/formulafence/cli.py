"""Command-line interface for FormulaFence."""

from __future__ import annotations

import argparse
import sys
from collections.abc import Sequence
from pathlib import Path

from formulafence import __version__
from formulafence.diff import compare_snapshots, report_severities
from formulafence.models import SEVERITY_ORDER, FormulaFenceError
from formulafence.output import as_json, profile_to_markdown, report_to_markdown, report_to_sarif
from formulafence.policy import DEFAULT_POLICY, evaluate_policy, load_policy
from formulafence.workbook import load_snapshot, profile_snapshot

_FAIL_LEVELS = ("none", "low", "medium", "high", "critical")


def _add_output_arguments(parser: argparse.ArgumentParser, formats: Sequence[str]) -> None:
    parser.add_argument("--format", choices=formats, default="markdown", help="Report format")
    parser.add_argument("--output", type=Path, help="Write report to this file instead of stdout")


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
    check.add_argument(
        "--fail-on",
        choices=_FAIL_LEVELS,
        default="none",
        help="Also exit 1 when a non-policy change or finding reaches this severity",
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


def _threshold_failed(severities: Sequence[str], fail_on: str) -> bool:
    if fail_on == "none":
        return False
    threshold = SEVERITY_ORDER[fail_on]
    return any(SEVERITY_ORDER[severity] >= threshold for severity in severities)


def _run_profile(arguments: argparse.Namespace) -> int:
    profile = profile_snapshot(load_snapshot(arguments.workbook))
    content = as_json(profile) if arguments.format == "json" else profile_to_markdown(profile)
    _emit(content, arguments.output)
    return 0


def _run_comparison(arguments: argparse.Namespace, enforce_policy: bool) -> int:
    report = compare_snapshots(load_snapshot(arguments.before), load_snapshot(arguments.after))
    policy_findings = []
    if enforce_policy:
        policy_findings = evaluate_policy(report, load_policy(arguments.policy))

    if arguments.format == "json":
        content = as_json(report.to_dict(policy_findings))
    elif arguments.format == "sarif":
        content = as_json(report_to_sarif(report, policy_findings))
    else:
        content = report_to_markdown(report, policy_findings)
    _emit(content, arguments.output)

    if policy_findings or _threshold_failed(
        report_severities(report, policy_findings), arguments.fail_on
    ):
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
        if arguments.command == "init":
            return _run_init(arguments)
    except (FormulaFenceError, OSError) as error:
        print(f"formulafence: error: {error}", file=sys.stderr)
        return 2
    parser.error(f"Unknown command: {arguments.command}")
    return 2  # pragma: no cover - argparse exits above


if __name__ == "__main__":  # pragma: no cover
    raise SystemExit(main())
