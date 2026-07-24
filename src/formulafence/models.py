"""Data structures shared by FormulaFence's parser, diff, and policy layers."""

from __future__ import annotations

from collections.abc import Iterable
from dataclasses import asdict, dataclass, field
from datetime import date, datetime, time
from pathlib import Path
from typing import Any, TypeAlias

CellKey: TypeAlias = tuple[str, str]

SEVERITY_ORDER = {"note": 0, "low": 1, "medium": 2, "high": 3, "critical": 4}


def display_location(location: CellKey | None) -> str | None:
    """Return an Excel-friendly cell location."""
    if location is None:
        return None
    sheet, coordinate = location
    escaped_sheet = sheet.replace("'", "''")
    if any(character in sheet for character in " !'[]"):
        escaped_sheet = f"'{escaped_sheet}'"
    return f"{escaped_sheet}!{coordinate}"


def json_safe_value(value: Any) -> Any:
    """Convert an openpyxl value to a deterministic JSON-compatible representation."""
    if value is None or isinstance(value, (str, int, float, bool)):
        return value
    if isinstance(value, (datetime, date, time)):
        return value.isoformat()
    return str(value)


@dataclass(frozen=True)
class CellSnapshot:
    """The semantic content FormulaFence compares for one non-empty cell."""

    sheet: str
    coordinate: str
    cell_type: str
    value: Any
    value_type: str
    formula: str | None = None
    formula_fingerprint: str | None = None

    @property
    def location(self) -> CellKey:
        return (self.sheet, self.coordinate)

    @property
    def is_formula(self) -> bool:
        return self.cell_type == "formula"

    def to_dict(self) -> dict[str, Any]:
        return {
            "location": display_location(self.location),
            "cell_type": self.cell_type,
            "value": self.value,
            "value_type": self.value_type,
            "formula": self.formula,
            "formula_fingerprint": self.formula_fingerprint,
        }


@dataclass(frozen=True)
class SheetSnapshot:
    title: str
    state: str
    nonempty_cells: int
    formula_cells: int
    max_row: int
    max_column: int

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)


@dataclass(frozen=True)
class RangeDependency:
    """A range used by a formula, stored without expanding large Excel ranges."""

    source_sheet: str
    min_column: int
    min_row: int
    max_column: int
    max_row: int
    dependent: CellKey

    def contains(self, location: CellKey) -> bool:
        sheet, coordinate = location
        if sheet != self.source_sheet:
            return False
        from openpyxl.utils.cell import column_index_from_string, coordinate_to_tuple

        row, column = coordinate_to_tuple(coordinate)
        if isinstance(column, str):  # defensive for older openpyxl versions
            column = column_index_from_string(column)
        return (
            self.min_column <= column <= self.max_column
            and self.min_row <= row <= self.max_row
        )


@dataclass(frozen=True)
class Finding:
    rule_id: str
    severity: str
    message: str
    location: CellKey | None = None
    details: dict[str, Any] = field(default_factory=dict)

    def to_dict(self) -> dict[str, Any]:
        result = {
            "rule_id": self.rule_id,
            "severity": self.severity,
            "message": self.message,
            "location": display_location(self.location),
        }
        if self.details:
            result["details"] = self.details
        return result


@dataclass(frozen=True)
class Change:
    kind: str
    location: CellKey | None
    severity: str
    before: CellSnapshot | None = None
    after: CellSnapshot | None = None
    impact_count: int = 0
    impacted_cells: tuple[CellKey, ...] = ()
    details: dict[str, Any] = field(default_factory=dict)

    def to_dict(self) -> dict[str, Any]:
        return {
            "kind": self.kind,
            "severity": self.severity,
            "location": display_location(self.location),
            "before": self.before.to_dict() if self.before else None,
            "after": self.after.to_dict() if self.after else None,
            "impact_count": self.impact_count,
            "impacted_cells": [display_location(cell) for cell in self.impacted_cells],
            "details": self.details,
        }


@dataclass
class WorkbookSnapshot:
    """Workbook semantics plus the dependency indexes used for impact analysis."""

    path: Path
    sha256: str
    file_type: str
    sheets: dict[str, SheetSnapshot]
    cells: dict[CellKey, CellSnapshot]
    reverse_dependencies: dict[CellKey, set[CellKey]]
    range_dependencies: list[RangeDependency]
    external_references: set[CellKey]
    broken_references: set[CellKey]
    defined_names: dict[str, str]
    macro_hash: str | None
    calculation_settings: dict[str, Any]

    def direct_dependents(self, location: CellKey) -> set[CellKey]:
        dependents = set(self.reverse_dependencies.get(location, set()))
        for dependency in self.range_dependencies:
            if dependency.contains(location):
                dependents.add(dependency.dependent)
        return dependents

    def summary(self) -> dict[str, Any]:
        return {
            "path": str(self.path),
            "sha256": self.sha256,
            "file_type": self.file_type,
            "sheet_count": len(self.sheets),
            "nonempty_cells": len(self.cells),
            "formula_cells": sum(1 for cell in self.cells.values() if cell.is_formula),
            "defined_names": len(self.defined_names),
            "has_vba": self.macro_hash is not None,
            "external_reference_cells": len(self.external_references),
            "broken_reference_cells": len(self.broken_references),
        }


@dataclass
class DiffReport:
    before: WorkbookSnapshot
    after: WorkbookSnapshot
    changes: list[Change]
    findings: list[Finding]

    def severity_counts(self) -> dict[str, int]:
        counts = {severity: 0 for severity in SEVERITY_ORDER}
        for finding in self.findings:
            counts[finding.severity] = counts.get(finding.severity, 0) + 1
        return {severity: count for severity, count in counts.items() if count}

    @property
    def highest_severity(self) -> str:
        if not self.findings:
            return "note"
        return max(self.findings, key=lambda finding: SEVERITY_ORDER[finding.severity]).severity

    def to_dict(self, extra_findings: Iterable[Finding] = ()) -> dict[str, Any]:
        findings = [*self.findings, *extra_findings]
        counts = {severity: 0 for severity in SEVERITY_ORDER}
        for finding in findings:
            counts[finding.severity] = counts.get(finding.severity, 0) + 1
        highest = (
            max(findings, key=lambda finding: SEVERITY_ORDER[finding.severity]).severity
            if findings
            else "note"
        )
        return {
            "schema_version": "1.0",
            "before": self.before.summary(),
            "after": self.after.summary(),
            "summary": {
                "change_count": len(self.changes),
                "finding_count": len(findings),
                "highest_severity": highest,
                "findings_by_severity": {
                    severity: count for severity, count in counts.items() if count
                },
            },
            "changes": [change.to_dict() for change in self.changes],
            "findings": [finding.to_dict() for finding in findings],
        }


class FormulaFenceError(Exception):
    """Base class for errors that should be presented cleanly by the CLI."""


class WorkbookLoadError(FormulaFenceError):
    """The workbook could not be safely inspected."""


class PolicyError(FormulaFenceError):
    """The policy file is invalid or cannot be interpreted."""
