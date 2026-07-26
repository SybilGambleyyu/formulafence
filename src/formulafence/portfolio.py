"""Bounded, fail-closed comparison of workbook portfolios.

Portfolio comparison deliberately uses a relative path as the identity of a
workbook.  It does not guess that two differently named files are a rename:
that would make a review less trustworthy when a file was intentionally
replaced.  Every supported workbook is inspected independently, so one
malformed archive becomes visible evidence instead of erasing the rest of a
portfolio report.
"""

from __future__ import annotations

from collections import Counter
from collections.abc import Iterable
from dataclasses import dataclass
from pathlib import Path
from typing import Any

from formulafence.diff import compare_snapshots
from formulafence.models import (
    SEVERITY_ORDER,
    Change,
    DiffReport,
    Finding,
    FormulaFenceError,
    WorkbookSnapshot,
)
from formulafence.policy import (
    Policy,
    evaluate_policy,
    evaluate_portfolio_membership_policy,
)
from formulafence.workbook import load_snapshot

_SUPPORTED_SUFFIXES = frozenset({".xlsx", ".xlsm"})
_UNSUPPORTED_EXCEL_SUFFIXES = frozenset(
    {".xls", ".xlsb", ".xlt", ".xltx", ".xltm", ".xlam", ".ods"}
)
_DEFAULT_MAX_WORKBOOKS = 512


class PortfolioError(FormulaFenceError):
    """The supplied portfolio cannot be safely compared."""


def _path_sort_key(value: str) -> tuple[str, str]:
    return value.casefold(), value


def _safe_relative_path(path: Path, root: Path) -> str:
    """Return a portable relative identity or fail before scanning a file."""
    try:
        relative = path.relative_to(root).as_posix()
    except ValueError as error:  # pragma: no cover - defensive against races
        raise PortfolioError("Workbook escaped its supplied portfolio root.") from error
    if "\\" in relative or any(
        ord(character) < 32 or ord(character) == 127 for character in relative
    ):
        raise PortfolioError(
            "Portfolio workbook paths must not contain control characters or backslashes."
        )
    return relative


def _resolve_directory(path: str | Path, label: str) -> Path:
    supplied = Path(path)
    try:
        resolved = supplied.resolve(strict=True)
    except OSError as error:
        raise PortfolioError(
            f"Could not resolve {label} portfolio directory: {supplied}"
        ) from error
    if not resolved.is_dir():
        raise PortfolioError(f"{label.capitalize()} portfolio is not a directory: {supplied}")
    return resolved


def discover_workbooks(
    root: str | Path,
    *,
    label: str,
    max_workbooks: int = _DEFAULT_MAX_WORKBOOKS,
) -> dict[str, Path]:
    """Return supported workbook files keyed by a stable relative path.

    Office lock files (``~$*.xlsx``) are transient and intentionally ignored.
    Every other known spreadsheet extension outside FormulaFence's `.xlsx` /
    `.xlsm` contract is an explicit error rather than a silent coverage hole.
    """
    if max_workbooks < 1:
        raise PortfolioError("max_workbooks must be at least 1.")

    resolved_root = _resolve_directory(root, label)
    try:
        candidates = sorted(
            resolved_root.rglob("*"), key=lambda item: _path_sort_key(item.as_posix())
        )
    except OSError as error:
        raise PortfolioError(f"Could not inventory {label} portfolio directory.") from error

    workbooks: dict[str, Path] = {}
    casefolded_paths: dict[str, str] = {}
    unsupported: list[str] = []
    for candidate in candidates:
        try:
            if candidate.is_symlink():
                relative = _safe_relative_path(candidate, resolved_root)
                raise PortfolioError(
                    f"Refusing symlinked path in {label} portfolio: {relative}"
                )
            if not candidate.is_file():
                continue
        except OSError as error:
            raise PortfolioError(f"Could not inspect a path in {label} portfolio.") from error

        relative = _safe_relative_path(candidate, resolved_root)
        if candidate.name.startswith("~$"):
            continue
        suffix = candidate.suffix.casefold()
        if suffix in _UNSUPPORTED_EXCEL_SUFFIXES:
            unsupported.append(relative)
            continue
        if suffix not in _SUPPORTED_SUFFIXES:
            continue

        try:
            candidate.resolve(strict=True).relative_to(resolved_root)
        except (OSError, ValueError) as error:
            raise PortfolioError(
                f"Workbook escaped its supplied {label} portfolio root: {relative}"
            ) from error

        portable_key = relative.casefold()
        previous = casefolded_paths.get(portable_key)
        if previous is not None:
            raise PortfolioError(
                "Portfolio contains paths that differ only by case: "
                f"{previous} and {relative}"
            )
        casefolded_paths[portable_key] = relative
        workbooks[relative] = candidate

    if unsupported:
        shown = ", ".join(unsupported[:5])
        remainder = "" if len(unsupported) <= 5 else f" (+{len(unsupported) - 5} more)"
        raise PortfolioError(
            f"{label.capitalize()} portfolio contains unsupported spreadsheet files: "
            f"{shown}{remainder}"
        )
    if len(workbooks) > max_workbooks:
        raise PortfolioError(
            f"{label.capitalize()} portfolio contains {len(workbooks)} supported workbooks, "
            f"exceeding max_workbooks={max_workbooks}."
        )
    return workbooks


def _safe_workbook_summary(snapshot: WorkbookSnapshot | None, path: str) -> dict[str, Any] | None:
    if snapshot is None:
        return None
    summary = snapshot.summary()
    # A directory report must stay portable and must not leak a worker's absolute
    # filesystem layout through the nested single-workbook summary.
    summary["path"] = path
    return summary


def _summary_for(
    changes: Iterable[Change],
    raw_findings: Iterable[Finding],
    policy_findings: Iterable[Finding],
) -> dict[str, Any]:
    changes = tuple(changes)
    raw_findings = tuple(raw_findings)
    policy_findings = tuple(policy_findings)
    findings = (*raw_findings, *policy_findings)
    counts = Counter(finding.severity for finding in findings)
    highest = (
        max(findings, key=lambda finding: SEVERITY_ORDER[finding.severity]).severity
        if findings
        else "note"
    )
    return {
        "change_count": len(changes),
        "finding_count": len(findings),
        "raw_finding_count": len(raw_findings),
        "policy_finding_count": len(policy_findings),
        "highest_severity": highest,
        "findings_by_severity": {
            severity: counts[severity] for severity in SEVERITY_ORDER if counts[severity]
        },
    }


@dataclass(frozen=True)
class PortfolioWorkbookReport:
    """One relative-path workbook result within a portfolio report."""

    path: str
    status: str
    baseline_present: bool
    candidate_present: bool
    before: WorkbookSnapshot | None = None
    after: WorkbookSnapshot | None = None
    report: DiffReport | None = None
    standalone_changes: tuple[Change, ...] = ()
    standalone_findings: tuple[Finding, ...] = ()
    policy_findings: tuple[Finding, ...] = ()

    @property
    def changes(self) -> tuple[Change, ...]:
        if self.report is not None:
            return tuple(self.report.changes)
        return self.standalone_changes

    @property
    def raw_findings(self) -> tuple[Finding, ...]:
        if self.report is not None:
            return tuple(self.report.findings)
        return self.standalone_findings

    @property
    def findings(self) -> tuple[Finding, ...]:
        return (*self.raw_findings, *self.policy_findings)

    @property
    def incomplete(self) -> bool:
        return self.status == "unreadable"

    def to_dict(self) -> dict[str, Any]:
        if self.report is not None:
            payload = self.report.to_dict(self.policy_findings)
            payload["before"]["path"] = self.path
            payload["after"]["path"] = self.path
            payload["summary"]["raw_finding_count"] = len(self.report.findings)
            payload["summary"]["policy_finding_count"] = len(self.policy_findings)
        else:
            payload = {
                "before": _safe_workbook_summary(self.before, self.path),
                "after": _safe_workbook_summary(self.after, self.path),
                "summary": _summary_for(
                    self.standalone_changes,
                    self.standalone_findings,
                    self.policy_findings,
                ),
                "changes": [change.to_dict() for change in self.standalone_changes],
                "findings": [finding.to_dict() for finding in self.findings],
            }
        return {
            "path": self.path,
            "status": self.status,
            "baseline_present": self.baseline_present,
            "candidate_present": self.candidate_present,
            **payload,
        }


@dataclass(frozen=True)
class PortfolioReport:
    """A deterministic inventory and comparison report for two directories."""

    baseline_workbook_count: int
    candidate_workbook_count: int
    workbooks: tuple[PortfolioWorkbookReport, ...]

    @property
    def incomplete(self) -> bool:
        return any(entry.incomplete for entry in self.workbooks)

    @property
    def policy_findings(self) -> tuple[Finding, ...]:
        return tuple(
            finding
            for entry in self.workbooks
            for finding in entry.policy_findings
        )

    def severities(self) -> list[str]:
        """Return every change/finding severity for ``--fail-on`` handling."""
        return [
            *(change.severity for entry in self.workbooks for change in entry.changes),
            *(finding.severity for entry in self.workbooks for finding in entry.findings),
        ]

    def to_dict(self) -> dict[str, Any]:
        entries = [entry.to_dict() for entry in self.workbooks]
        findings = [finding for entry in self.workbooks for finding in entry.findings]
        counts = Counter(finding.severity for finding in findings)
        highest = (
            max(findings, key=lambda finding: SEVERITY_ORDER[finding.severity]).severity
            if findings
            else "note"
        )
        status_counts = Counter(entry.status for entry in self.workbooks)
        return {
            "schema_version": "1.0",
            "report_type": "portfolio",
            "before": {"workbook_count": self.baseline_workbook_count},
            "after": {"workbook_count": self.candidate_workbook_count},
            "summary": {
                "matched_workbook_count": sum(
                    entry.baseline_present and entry.candidate_present
                    for entry in self.workbooks
                ),
                "unchanged_workbook_count": status_counts["unchanged"],
                "changed_workbook_count": status_counts["changed"],
                "added_workbook_count": sum(
                    not entry.baseline_present and entry.candidate_present
                    for entry in self.workbooks
                ),
                "removed_workbook_count": sum(
                    entry.baseline_present and not entry.candidate_present
                    for entry in self.workbooks
                ),
                "unreadable_workbook_count": status_counts["unreadable"],
                "change_count": sum(len(entry.changes) for entry in self.workbooks),
                "finding_count": len(findings),
                "policy_finding_count": len(self.policy_findings),
                "highest_severity": highest,
                "findings_by_severity": {
                    severity: counts[severity] for severity in SEVERITY_ORDER if counts[severity]
                },
                "incomplete": self.incomplete,
            },
            "workbooks": entries,
        }


def _membership_evidence(
    status: str, policy: Policy | None
) -> tuple[tuple[Change, ...], tuple[Finding, ...], tuple[Finding, ...]]:
    """Return the evidence for a known added or removed relative path."""
    change_kind = "workbook_added" if status == "added" else "workbook_removed"
    raw_finding = Finding(
        "FF077",
        "high",
        (
            "Workbook portfolio membership changed; the workbook was "
            f"{status} and was not matched to a peer for semantic comparison."
        ),
        details={"portfolio_change": status},
    )
    policy_findings = (
        tuple(evaluate_portfolio_membership_policy((raw_finding,), policy))
        if policy is not None
        else ()
    )
    return (
        (Change(change_kind, None, "high", details={"portfolio_change": status}),),
        (raw_finding,),
        policy_findings,
    )


def _membership_entry(
    path: str,
    *,
    before: WorkbookSnapshot | None,
    after: WorkbookSnapshot | None,
    policy: Policy | None,
) -> PortfolioWorkbookReport:
    status = "added" if after is not None else "removed"
    changes, findings, policy_findings = _membership_evidence(status, policy)
    return PortfolioWorkbookReport(
        path=path,
        status=status,
        baseline_present=before is not None,
        candidate_present=after is not None,
        before=before,
        after=after,
        standalone_changes=changes,
        standalone_findings=findings,
        policy_findings=policy_findings,
    )


def _unreadable_entry(
    path: str,
    *,
    baseline_present: bool,
    candidate_present: bool,
    before: WorkbookSnapshot | None,
    after: WorkbookSnapshot | None,
    unreadable_sides: tuple[str, ...],
    policy: Policy | None,
) -> PortfolioWorkbookReport:
    membership_changes: tuple[Change, ...] = ()
    membership_findings: tuple[Finding, ...] = ()
    membership_policy_findings: tuple[Finding, ...] = ()
    if baseline_present != candidate_present:
        membership_status = "added" if candidate_present else "removed"
        (
            membership_changes,
            membership_findings,
            membership_policy_findings,
        ) = _membership_evidence(membership_status, policy)
    return PortfolioWorkbookReport(
        path=path,
        status="unreadable",
        baseline_present=baseline_present,
        candidate_present=candidate_present,
        before=before,
        after=after,
        standalone_changes=membership_changes,
        standalone_findings=(
            *membership_findings,
            Finding(
                "FF078",
                "critical",
                "Workbook could not be inspected; portfolio comparison is incomplete.",
                details={"unreadable_sides": list(unreadable_sides)},
            ),
        ),
        policy_findings=membership_policy_findings,
    )


def _load_portfolio_workbook(path: Path | None) -> tuple[WorkbookSnapshot | None, bool]:
    if path is None:
        return None, False
    try:
        return load_snapshot(path), False
    except Exception:  # noqa: BLE001 - malformed workbooks must not erase portfolio evidence
        return None, True


def compare_portfolios(
    baseline_directory: str | Path,
    candidate_directory: str | Path,
    *,
    policy: Policy | None = None,
    max_workbooks: int = _DEFAULT_MAX_WORKBOOKS,
) -> PortfolioReport:
    """Compare every workbook at the same relative path in two directories.

    The scan is deliberately sequential and bounded.  A folder-wide run should
    be deterministic, memory-safe for CI, and explicit about unsupported or
    unreadable material rather than silently treating it as unchanged.
    """
    baseline = discover_workbooks(
        baseline_directory, label="baseline", max_workbooks=max_workbooks
    )
    candidate = discover_workbooks(
        candidate_directory, label="candidate", max_workbooks=max_workbooks
    )
    if not baseline and not candidate:
        raise PortfolioError(
            "No supported .xlsx or .xlsm workbooks were found in either portfolio."
        )

    entries: list[PortfolioWorkbookReport] = []
    for path in sorted(set(baseline) | set(candidate), key=_path_sort_key):
        baseline_path = baseline.get(path)
        candidate_path = candidate.get(path)
        before, before_unreadable = _load_portfolio_workbook(baseline_path)
        after, after_unreadable = _load_portfolio_workbook(candidate_path)
        if before_unreadable or after_unreadable:
            unreadable_sides = tuple(
                side
                for side, unreadable in (
                    ("baseline", before_unreadable),
                    ("candidate", after_unreadable),
                )
                if unreadable
            )
            entries.append(
                _unreadable_entry(
                    path,
                    baseline_present=baseline_path is not None,
                    candidate_present=candidate_path is not None,
                    before=before,
                    after=after,
                    unreadable_sides=unreadable_sides,
                    policy=policy,
                )
            )
            continue
        if before is None or after is None:
            entries.append(_membership_entry(path, before=before, after=after, policy=policy))
            continue

        report = compare_snapshots(before, after)
        policy_findings = tuple(evaluate_policy(report, policy)) if policy is not None else ()
        status = "changed" if report.changes or report.findings or policy_findings else "unchanged"
        entries.append(
            PortfolioWorkbookReport(
                path=path,
                status=status,
                baseline_present=True,
                candidate_present=True,
                before=before,
                after=after,
                report=report,
                policy_findings=policy_findings,
            )
        )

    return PortfolioReport(
        baseline_workbook_count=len(baseline),
        candidate_workbook_count=len(candidate),
        workbooks=tuple(entries),
    )
