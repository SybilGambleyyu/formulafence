"""Deterministic JSON, Markdown, HTML, and SARIF renderers for review artifacts."""

from __future__ import annotations

import json
import re
from collections.abc import Callable, Iterable
from html import escape as _html_escape
from io import StringIO
from typing import Any

from formulafence import __version__
from formulafence.formulas import inspect_formula
from formulafence.models import (
    DiffReport,
    Finding,
    FormulaFenceError,
    FormulaLintReport,
    display_location,
)
from formulafence.portfolio import PortfolioReport

DEFAULT_MAX_REPORT_BYTES = 32 * 1024 * 1024

EXTERNAL_WORKBOOK_LINK_REDACTION = "[external-workbook link material redacted]"
FORMULA_EXTERNAL_ACTION_REDACTION = "[formula external-action material redacted]"
PYTHON_IN_EXCEL_REDACTION = "[Python-in-Excel material redacted]"
OFFICE_CUSTOM_FUNCTION_REDACTION = "[Office custom-function material redacted]"
UNQUALIFIED_RUNTIME_FUNCTION_REDACTION = (
    "[unqualified runtime-function material redacted]"
)
WORKSHEET_CODE_RESOURCE_REGISTRATION_REDACTION = (
    "[worksheet code-resource registration material redacted]"
)
FORMULA_DEFINED_XLM_REGISTRATION_REDACTION = (
    "[formula-defined XLM registration material redacted]"
)
FORMULA_DEFINED_XLM_EVALUATION_REDACTION = (
    "[formula-defined XLM evaluation material redacted]"
)
FORMULA_DEFINED_XLM_ACTION_REDACTION = (
    "[formula-defined XLM action material redacted]"
)
FORMULA_DEFINED_XLM_GET_CELL_REDACTION = (
    "[formula-defined XLM GET.CELL material redacted]"
)
FORMULA_DEFINED_XLM_ENVIRONMENT_INFORMATION_REDACTION = (
    "[formula-defined XLM environment-information material redacted]"
)
FORMULA_ENVIRONMENT_INFORMATION_REDACTION = (
    "[formula environment-information material redacted]"
)

# This fallback deliberately supplements, rather than replaces, FormulaFence's
# static formula inspection below.  A visible `INDIRECT("'[Book]Sheet'!A1")`
# string is not a static dependency, but it still exposes a workbook endpoint
# in a shared artifact.  The grammar is intentionally conservative: it only
# matches a bracketed book followed by a sheet separator, or a common Excel
# file extension in a direct-name/table-style literal.  It never evaluates or
# concatenates formula text.
_VISIBLE_EXTERNAL_WORKBOOK_LITERAL = re.compile(
    r"\[[^\[\]\r\n]{1,255}\][^!\r\n]{0,255}!"
    r"|\[[^\[\]\r\n]{1,255}\.(?:xls(?:x|m|b)?|xlt(?:x|m)?|xlam)\]"
    r"[A-Za-z_\\]"
    r"|['\"][^'\"\r\n]{1,1024}\.(?:xls(?:x|m|b)?|xlt(?:x|m)?|xlam)['\"]!",
    re.IGNORECASE,
)
_FORMULA_EXTERNAL_ACTION_HINTS = (
    "HYPERLINK",
    "WEBSERVICE",
    "IMAGE",
    "RTD",
    "STOCKHISTORY",
    "CUBE",
)


def _report_output_limit_error(max_bytes: int) -> FormulaFenceError:
    return FormulaFenceError(f"Rendered report exceeds max_report_bytes={max_bytes}.")


class _BoundedText:
    """Accumulate UTF-8 text without exceeding one rendered-artifact budget."""

    def __init__(self, max_bytes: int | None) -> None:
        if max_bytes is not None and max_bytes < 1:
            raise FormulaFenceError("max_report_bytes must be at least 1.")
        self._max_bytes = max_bytes
        self._byte_count = 0
        self._buffer = StringIO()

    def append(self, value: str) -> None:
        byte_count = len(value.encode("utf-8"))
        if self._max_bytes is not None and self._byte_count + byte_count > self._max_bytes:
            raise _report_output_limit_error(self._max_bytes)
        self._buffer.write(value)
        self._byte_count += byte_count

    def render(self) -> str:
        return self._buffer.getvalue()


class _BoundedLines:
    """Stream line-oriented renderers into one bounded UTF-8 artifact."""

    def __init__(self, values: Iterable[str], *, max_bytes: int | None) -> None:
        if max_bytes is not None and max_bytes < 1:
            raise FormulaFenceError("max_report_bytes must be at least 1.")
        self._max_bytes = max_bytes
        self._byte_count = 0
        self._buffer = StringIO()
        self.extend(values)

    def append(self, value: str) -> None:
        byte_count = len(value.encode("utf-8")) + 1
        if self._max_bytes is not None and self._byte_count + byte_count > self._max_bytes:
            raise _report_output_limit_error(self._max_bytes)
        self._buffer.write(value)
        self._buffer.write("\n")
        self._byte_count += byte_count

    def extend(self, values: Iterable[str]) -> None:
        for value in values:
            self.append(value)

    def render(self) -> str:
        return self._buffer.getvalue()


def as_json(payload: dict[str, Any], *, max_bytes: int | None = None) -> str:
    """Serialize a report payload, stopping before a bounded artifact overage."""
    if max_bytes is None:
        return json.dumps(payload, indent=2, sort_keys=True, ensure_ascii=False) + "\n"
    rendered = _BoundedText(max_bytes)
    encoder = json.JSONEncoder(indent=2, sort_keys=True, ensure_ascii=False)
    for chunk in encoder.iterencode(payload):
        rendered.append(chunk)
    rendered.append("\n")
    return rendered.render()


def _contains_external_workbook_link_material(value: str) -> bool:
    """Return whether one rendered string exposes a literal workbook endpoint.

    Most FormulaFence fields are formula strings, but some Excel controls omit
    the leading ``=``.  Normalising only that presentation detail lets the
    existing static parser recognise direct A1, 3-D, defined-name, and table
    links without resolving a source or evaluating any expression.  The small
    lexical fallback covers plainly visible dynamic literals such as
    ``INDIRECT("'[Book]Sheet'!A1")``; it intentionally cannot prove or recover
    a link assembled from separate text fragments.
    """
    # All direct grammar forms include a workbook bracket, except book-only
    # table syntax such as `'../inputs/source.xlsx'!Sales[#Data]`.  Avoid a
    # tokenizer pass over the many ordinary locations, messages, and labels in
    # a large report while retaining both forms.
    if "[" not in value and ".xls" not in value.casefold():
        return False
    formula = value if value.lstrip().startswith("=") else f"={value}"
    inspection = inspect_formula(formula)
    if (
        inspection.external_workbook_references
        or inspection.external_workbook_three_d_references
        or inspection.external_workbook_structured_references
        or inspection.external_workbook_defined_name_references
    ):
        return True
    return bool(_VISIBLE_EXTERNAL_WORKBOOK_LITERAL.search(value))


def redact_external_workbook_link_material(payload: Any) -> Any:
    """Return an output-only copy with visible external-workbook links hidden.

    FormulaFence's ordinary diff objects intentionally retain full local
    evidence.  This helper is for artifacts that leave that trusted review
    boundary: it walks JSON-compatible output values and replaces an entire
    string that exposes a literal external-workbook endpoint.  It does not
    mutate the supplied report, affect policy evaluation, or promise to redact
    a source assembled dynamically at Excel calculation time.
    """
    if isinstance(payload, dict):
        return {
            key: redact_external_workbook_link_material(value)
            for key, value in payload.items()
        }
    if isinstance(payload, list):
        return [redact_external_workbook_link_material(value) for value in payload]
    if isinstance(payload, tuple):
        return tuple(redact_external_workbook_link_material(value) for value in payload)
    if isinstance(payload, str) and _contains_external_workbook_link_material(payload):
        return EXTERNAL_WORKBOOK_LINK_REDACTION
    return payload


def _contains_formula_external_action_material(value: str) -> bool:
    """Return whether one rendered string exposes an inventoried action or DDE.

    This keeps the redaction boundary aligned with FormulaFence's existing
    ``FF064`` and ``FF074`` scanners.  A stored formula can carry a private URL,
    provider, Cube connection, DDE application/topic/item, or related argument;
    FormulaFence only recognizes the function/syntax, never evaluates it.
    """
    upper_value = value.upper()
    if "|" not in value and not any(
        hint in upper_value for hint in _FORMULA_EXTERNAL_ACTION_HINTS
    ):
        return False
    formula = value if value.lstrip().startswith("=") else f"={value}"
    inspection = inspect_formula(formula)
    return bool(
        inspection.external_action_functions or inspection.formula_dde_link_markers
    )


def redact_formula_external_action_material(payload: Any) -> Any:
    """Return an output-only copy with direct action and DDE formula text hidden.

    This handles formula-bearing strings directly.  Report-level helpers below
    also hide a changed cell value when its static impact reaches one of those
    inventoried formulas, because an endpoint can live in an ordinary input
    cell such as ``HYPERLINK(A9, ...)`` rather than in the action formula.
    """
    if isinstance(payload, dict):
        return {
            key: redact_formula_external_action_material(value)
            for key, value in payload.items()
        }
    if isinstance(payload, list):
        return [redact_formula_external_action_material(value) for value in payload]
    if isinstance(payload, tuple):
        return tuple(redact_formula_external_action_material(value) for value in payload)
    if isinstance(payload, str) and _contains_formula_external_action_material(payload):
        return FORMULA_EXTERNAL_ACTION_REDACTION
    return payload


def _formula_external_action_sensitive_change_locations(report: DiffReport) -> set[str]:
    """Return changed cells whose evidence can carry a static action input."""
    sensitive_cells = (
        report.before.formula_external_actions.action_cells
        | report.after.formula_external_actions.action_cells
        | report.before.formula_dde_links.dde_cells
        | report.after.formula_dde_links.dde_cells
        | report.formula_external_action_static_input_cells
        | report.formula_dde_link_static_input_cells
    )
    if not sensitive_cells:
        return set()

    locations: set[str] = set()
    for change in report.changes:
        if change.location is None or change.location not in sensitive_cells:
            continue
        location = display_location(change.location)
        if location is not None:
            locations.add(location)
    return locations


def _redact_formula_external_action_change_cells(
    payload: dict[str, Any], report: DiffReport
) -> dict[str, Any]:
    """Hide raw before/after cells that are an action/DDE formula or input."""
    sensitive_locations = _formula_external_action_sensitive_change_locations(report)
    if not sensitive_locations:
        return payload
    for change in payload.get("changes", []):
        if change.get("location") not in sensitive_locations:
            continue
        for side in ("before", "after"):
            snapshot = change.get(side)
            if not isinstance(snapshot, dict):
                continue
            for field in ("value", "formula", "formula_fingerprint"):
                if snapshot.get(field) is not None:
                    snapshot[field] = FORMULA_EXTERNAL_ACTION_REDACTION
    return payload


def _formula_external_action_definition_material_changed(report: DiffReport) -> bool:
    """Return whether a named action/DDE definition changed privately."""
    return (
        report.before.formula_external_actions.definition_signature
        != report.after.formula_external_actions.definition_signature
        or report.before.formula_dde_links.definition_signature
        != report.after.formula_dde_links.definition_signature
    )


def _redact_formula_external_action_defined_name_details(
    details: dict[str, Any], report: DiffReport
) -> None:
    """Hide FF008 formula evidence when a resolved action-name chain changed.

    A named definition can invoke another named LAMBDA that contains the action
    or DDE syntax.  Its own text may therefore be an ordinary-looking
    ``=FENCE.WRAPPER(\"endpoint\")`` value which a lexical formula scan cannot
    classify in isolation.  The comparison already has a private resolved
    definition signature; when it changed, replacing every changed defined-name
    before/after value is the safe output boundary.  This is deliberately
    conservative and never changes the name, finding, or policy result.
    """
    if not _formula_external_action_definition_material_changed(report):
        return
    for field in ("before", "after"):
        if isinstance(details.get(field), str):
            details[field] = FORMULA_EXTERNAL_ACTION_REDACTION


def _redact_formula_external_action_defined_name_evidence(
    payload: dict[str, Any], report: DiffReport
) -> dict[str, Any]:
    """Hide resolved action/DDE name-chain text in JSON-shaped report evidence."""
    if not _formula_external_action_definition_material_changed(report):
        return payload
    for change in payload.get("changes", []):
        if change.get("kind") == "defined_name_changed":
            details = change.get("details")
            if isinstance(details, dict):
                _redact_formula_external_action_defined_name_details(details, report)
    for finding in payload.get("findings", []):
        if finding.get("rule_id") != "FF008":
            continue
        details = finding.get("details")
        if isinstance(details, dict):
            _redact_formula_external_action_defined_name_details(details, report)
    return payload


def _safe_finding_details(
    finding: Finding,
    report: DiffReport | None,
    *,
    redact_formula_external_actions: bool = False,
    redact_office_custom_functions: bool = False,
    redact_unqualified_runtime_functions: bool = False,
    redact_worksheet_code_resource_registrations: bool = False,
    redact_formula_defined_xlm_registrations: bool = False,
    redact_formula_defined_xlm_evaluations: bool = False,
    redact_formula_defined_xlm_actions: bool = False,
    redact_formula_defined_xlm_get_cell_calls: bool = False,
    redact_formula_defined_xlm_environment_information_calls: bool = False,
    redact_formula_environment_information: bool = False,
) -> dict[str, Any]:
    """Copy one SARIF finding's properties through active name-chain bounds."""
    details = dict(finding.details)
    if (
        redact_formula_external_actions
        and report is not None
        and finding.rule_id == "FF008"
    ):
        _redact_formula_external_action_defined_name_details(details, report)
    if (
        redact_office_custom_functions
        and report is not None
        and finding.rule_id == "FF008"
    ):
        _redact_office_custom_function_defined_name_details(details, report)
    if (
        redact_unqualified_runtime_functions
        and report is not None
        and finding.rule_id == "FF008"
    ):
        _redact_unqualified_runtime_function_defined_name_details(details, report)
    if (
        redact_worksheet_code_resource_registrations
        and report is not None
        and finding.rule_id == "FF008"
    ):
        _redact_worksheet_code_resource_registration_defined_name_details(
            details, report
        )
    if (
        redact_formula_defined_xlm_registrations
        and report is not None
        and finding.rule_id == "FF008"
    ):
        _redact_formula_defined_xlm_registration_defined_name_details(details, report)
    if (
        redact_formula_defined_xlm_evaluations
        and report is not None
        and finding.rule_id == "FF008"
    ):
        _redact_formula_defined_xlm_evaluation_defined_name_details(details, report)
    if (
        redact_formula_defined_xlm_actions
        and report is not None
        and finding.rule_id == "FF008"
    ):
        _redact_formula_defined_xlm_action_defined_name_details(details, report)
    if (
        redact_formula_defined_xlm_get_cell_calls
        and report is not None
        and finding.rule_id == "FF008"
    ):
        _redact_formula_defined_xlm_get_cell_defined_name_details(details, report)
    if (
        redact_formula_defined_xlm_environment_information_calls
        and report is not None
        and finding.rule_id == "FF008"
    ):
        _redact_formula_defined_xlm_environment_information_defined_name_details(
            details, report
        )
    if (
        redact_formula_environment_information
        and report is not None
        and finding.rule_id == "FF008"
    ):
        _redact_formula_environment_information_defined_name_details(details, report)
    return details


def redact_formula_external_action_report_payload(
    report: DiffReport, payload: dict[str, Any]
) -> dict[str, Any]:
    """Redact direct action text and private static action-input cell evidence."""
    redacted = redact_formula_external_action_material(payload)
    _redact_formula_external_action_change_cells(redacted, report)
    return _redact_formula_external_action_defined_name_evidence(redacted, report)


def redact_formula_external_action_portfolio_payload(
    report: PortfolioReport, payload: dict[str, Any]
) -> dict[str, Any]:
    """Apply the same action/DDE boundary to nested portfolio reports."""
    redacted = redact_formula_external_action_material(payload)
    for workbook, entry in zip(
        report.workbooks, redacted.get("workbooks", []), strict=False
    ):
        if workbook.report is not None and isinstance(entry, dict):
            _redact_formula_external_action_change_cells(entry, workbook.report)
            _redact_formula_external_action_defined_name_evidence(entry, workbook.report)
    return redacted


def _contains_office_custom_function_material(value: str) -> bool:
    """Return whether one rendered string exposes a namespaced call candidate.

    Office Add-in custom functions are stored as namespaced formulas. A direct
    call can therefore expose an add-in namespace, callable, and arguments to
    a shared report even though FormulaFence's FF066 ledger publishes only
    aggregate counts. The parser retains its conservative exclusions for native
    dotted functions and workbook-defined names wherever that context is
    available; this output scan intentionally hides any standalone formula
    shaped like an inventoried namespaced callable.
    """
    if "." not in value or "(" not in value:
        return False
    formula = value if value.lstrip().startswith("=") else f"={value}"
    return bool(inspect_formula(formula).office_custom_function_candidates)


def redact_office_custom_function_material(payload: Any) -> Any:
    """Return an output-only copy with direct custom-function formulas hidden.

    This only covers direct stored formula material. Report-level helpers below
    also hide exact changed static inputs and changed formula-defined-name
    bodies that the private FF066 comparison identifies as custom-function
    relevant.
    """
    if isinstance(payload, dict):
        return {
            key: redact_office_custom_function_material(value)
            for key, value in payload.items()
        }
    if isinstance(payload, list):
        return [redact_office_custom_function_material(value) for value in payload]
    if isinstance(payload, tuple):
        return tuple(redact_office_custom_function_material(value) for value in payload)
    if isinstance(payload, str) and _contains_office_custom_function_material(payload):
        return OFFICE_CUSTOM_FUNCTION_REDACTION
    return payload


def _office_custom_function_sensitive_change_locations(report: DiffReport) -> set[str]:
    """Return changed cells whose evidence can carry a custom-function input."""
    sensitive_cells = (
        report.before.office_custom_functions.call_cells
        | report.after.office_custom_functions.call_cells
        | report.office_custom_function_static_input_cells
    )
    if not sensitive_cells:
        return set()

    locations: set[str] = set()
    for change in report.changes:
        if change.location is None or change.location not in sensitive_cells:
            continue
        location = display_location(change.location)
        if location is not None:
            locations.add(location)
    return locations


def _redact_office_custom_function_change_cells(
    payload: dict[str, Any], report: DiffReport
) -> dict[str, Any]:
    """Hide raw before/after cells that are a custom call or exact input."""
    sensitive_locations = _office_custom_function_sensitive_change_locations(report)
    if not sensitive_locations:
        return payload
    for change in payload.get("changes", []):
        if change.get("location") not in sensitive_locations:
            continue
        for side in ("before", "after"):
            snapshot = change.get(side)
            if not isinstance(snapshot, dict):
                continue
            for field in ("value", "formula", "formula_fingerprint"):
                if snapshot.get(field) is not None:
                    snapshot[field] = OFFICE_CUSTOM_FUNCTION_REDACTION
    return payload


def _office_custom_function_definition_material_changed(report: DiffReport) -> bool:
    """Return whether a custom-function-relevant defined-name body changed."""
    return (
        report.before.office_custom_functions.definition_signature
        != report.after.office_custom_functions.definition_signature
    )


def _redact_office_custom_function_defined_name_details(
    details: dict[str, Any], report: DiffReport
) -> None:
    """Hide FF008 evidence when a resolved custom-function name chain changed.

    A named LAMBDA can pass a private argument through an ordinary-looking
    unqualified named call whose eventual custom function lives deeper in the
    chain. Its text cannot reliably be classified without the workbook's
    private fixed-point resolution. When that private custom-function
    definition signature changed, hiding every changed defined-name body is the
    safe artifact boundary.
    """
    if not _office_custom_function_definition_material_changed(report):
        return
    for field in ("before", "after"):
        if isinstance(details.get(field), str):
            details[field] = OFFICE_CUSTOM_FUNCTION_REDACTION


def _redact_office_custom_function_defined_name_evidence(
    payload: dict[str, Any], report: DiffReport
) -> dict[str, Any]:
    """Hide custom-function-resolved name-chain text in report evidence."""
    if not _office_custom_function_definition_material_changed(report):
        return payload
    for change in payload.get("changes", []):
        if change.get("kind") == "defined_name_changed":
            details = change.get("details")
            if isinstance(details, dict):
                _redact_office_custom_function_defined_name_details(details, report)
    for finding in payload.get("findings", []):
        if finding.get("rule_id") != "FF008":
            continue
        details = finding.get("details")
        if isinstance(details, dict):
            _redact_office_custom_function_defined_name_details(details, report)
    return payload


def redact_office_custom_function_report_payload(
    report: DiffReport, payload: dict[str, Any]
) -> dict[str, Any]:
    """Redact direct custom calls and private static/name-chain evidence."""
    redacted = redact_office_custom_function_material(payload)
    _redact_office_custom_function_change_cells(redacted, report)
    return _redact_office_custom_function_defined_name_evidence(redacted, report)


def redact_office_custom_function_portfolio_payload(
    report: PortfolioReport, payload: dict[str, Any]
) -> dict[str, Any]:
    """Apply the custom-function sharing boundary to nested reports."""
    redacted = redact_office_custom_function_material(payload)
    for workbook, entry in zip(
        report.workbooks, redacted.get("workbooks", []), strict=False
    ):
        if workbook.report is not None and isinstance(entry, dict):
            _redact_office_custom_function_change_cells(entry, workbook.report)
            _redact_office_custom_function_defined_name_evidence(entry, workbook.report)
    return redacted


def _contains_unqualified_runtime_function_material(value: str) -> bool:
    """Return whether one rendered string exposes a bare runtime-call candidate.

    An unknown unqualified worksheet call can expose a private UDF name and its
    arguments even when FormulaFence's FF075 ledger publishes only counts. The
    workbook-aware classifier excludes defined names and local bindings during
    snapshot inspection; this generic output scan intentionally hides any
    standalone formula shaped like an unqualified runtime candidate because
    report strings do not retain that private workbook context.
    """
    if "(" not in value:
        return False
    formula = value if value.lstrip().startswith("=") else f"={value}"
    return bool(inspect_formula(formula).unqualified_runtime_function_candidates)


def redact_unqualified_runtime_function_material(payload: Any) -> Any:
    """Return an output-only copy with direct unknown runtime calls hidden.

    This handles direct stored formula material. The report-level helpers below
    also hide exact changed static inputs and changed formula-defined-name
    bodies that the private FF075 comparison identifies as runtime-function
    relevant.
    """
    if isinstance(payload, dict):
        return {
            key: redact_unqualified_runtime_function_material(value)
            for key, value in payload.items()
        }
    if isinstance(payload, list):
        return [
            redact_unqualified_runtime_function_material(value) for value in payload
        ]
    if isinstance(payload, tuple):
        return tuple(
            redact_unqualified_runtime_function_material(value) for value in payload
        )
    if isinstance(payload, str) and _contains_unqualified_runtime_function_material(
        payload
    ):
        return UNQUALIFIED_RUNTIME_FUNCTION_REDACTION
    return payload


def _unqualified_runtime_function_sensitive_change_locations(
    report: DiffReport,
) -> set[str]:
    """Return changed cells whose evidence can carry a runtime-call input."""
    sensitive_cells = (
        report.before.unqualified_runtime_functions.call_cells
        | report.after.unqualified_runtime_functions.call_cells
        | report.unqualified_runtime_function_static_input_cells
    )
    if not sensitive_cells:
        return set()

    locations: set[str] = set()
    for change in report.changes:
        if change.location is None or change.location not in sensitive_cells:
            continue
        location = display_location(change.location)
        if location is not None:
            locations.add(location)
    return locations


def _redact_unqualified_runtime_function_change_cells(
    payload: dict[str, Any], report: DiffReport
) -> dict[str, Any]:
    """Hide before/after cells that are a runtime call or exact static input."""
    sensitive_locations = _unqualified_runtime_function_sensitive_change_locations(
        report
    )
    if not sensitive_locations:
        return payload
    for change in payload.get("changes", []):
        if change.get("location") not in sensitive_locations:
            continue
        for side in ("before", "after"):
            snapshot = change.get(side)
            if not isinstance(snapshot, dict):
                continue
            for field in ("value", "formula", "formula_fingerprint"):
                if snapshot.get(field) is not None:
                    snapshot[field] = UNQUALIFIED_RUNTIME_FUNCTION_REDACTION
    return payload


def _unqualified_runtime_function_definition_material_changed(
    report: DiffReport,
) -> bool:
    """Return whether a runtime-function-relevant defined-name body changed."""
    return (
        report.before.unqualified_runtime_functions.definition_signature
        != report.after.unqualified_runtime_functions.definition_signature
    )


def _redact_unqualified_runtime_function_defined_name_details(
    details: dict[str, Any], report: DiffReport
) -> None:
    """Hide FF008 evidence when a resolved runtime-function name chain changed.

    A named LAMBDA can call a dotted workbook-defined wrapper whose eventual
    implementation contains the bare runtime candidate. Its own text can look
    ordinary to a standalone lexical scanner. When the private FF075 resolved
    definition signature changes, hiding every changed defined-name body is the
    safe shared-artifact boundary.
    """
    if not _unqualified_runtime_function_definition_material_changed(report):
        return
    for field in ("before", "after"):
        if isinstance(details.get(field), str):
            details[field] = UNQUALIFIED_RUNTIME_FUNCTION_REDACTION


def _redact_unqualified_runtime_function_defined_name_evidence(
    payload: dict[str, Any], report: DiffReport
) -> dict[str, Any]:
    """Hide runtime-function-resolved name-chain text in report evidence."""
    if not _unqualified_runtime_function_definition_material_changed(report):
        return payload
    for change in payload.get("changes", []):
        if change.get("kind") == "defined_name_changed":
            details = change.get("details")
            if isinstance(details, dict):
                _redact_unqualified_runtime_function_defined_name_details(details, report)
    for finding in payload.get("findings", []):
        if finding.get("rule_id") != "FF008":
            continue
        details = finding.get("details")
        if isinstance(details, dict):
            _redact_unqualified_runtime_function_defined_name_details(details, report)
    return payload


def redact_unqualified_runtime_function_report_payload(
    report: DiffReport, payload: dict[str, Any]
) -> dict[str, Any]:
    """Redact direct runtime calls and private static/name-chain evidence."""
    redacted = redact_unqualified_runtime_function_material(payload)
    _redact_unqualified_runtime_function_change_cells(redacted, report)
    return _redact_unqualified_runtime_function_defined_name_evidence(redacted, report)


def redact_unqualified_runtime_function_portfolio_payload(
    report: PortfolioReport, payload: dict[str, Any]
) -> dict[str, Any]:
    """Apply the runtime-function sharing boundary to nested reports."""
    redacted = redact_unqualified_runtime_function_material(payload)
    for workbook, entry in zip(
        report.workbooks, redacted.get("workbooks", []), strict=False
    ):
        if workbook.report is not None and isinstance(entry, dict):
            _redact_unqualified_runtime_function_change_cells(entry, workbook.report)
            _redact_unqualified_runtime_function_defined_name_evidence(
                entry, workbook.report
            )
    return redacted


def _contains_worksheet_code_resource_registration_material(value: str) -> bool:
    """Return whether one rendered string exposes a ``REGISTER.ID`` call.

    ``REGISTER.ID`` can register a DLL or code resource from a worksheet. Its
    stored module, procedure, type string, and arguments can be sensitive even
    though FormulaFence's FF067 ledger publishes only aggregate counts. A
    renderer has no workbook-local name context, so this generic scan safely
    hides any standalone formula shaped like the documented registration call.
    """
    if "register.id" not in value.casefold() or "(" not in value:
        return False
    formula = value if value.lstrip().startswith("=") else f"={value}"
    return bool(inspect_formula(formula).worksheet_code_resource_registration_functions)


def redact_worksheet_code_resource_registration_material(payload: Any) -> Any:
    """Return an output-only copy with direct ``REGISTER.ID`` text hidden.

    This handles direct stored formula material. The report-level helpers below
    also hide exact changed static inputs and changed formula-defined-name
    bodies that the private FF067 comparison identifies as registration
    relevant.
    """
    if isinstance(payload, dict):
        return {
            key: redact_worksheet_code_resource_registration_material(value)
            for key, value in payload.items()
        }
    if isinstance(payload, list):
        return [
            redact_worksheet_code_resource_registration_material(value)
            for value in payload
        ]
    if isinstance(payload, tuple):
        return tuple(
            redact_worksheet_code_resource_registration_material(value)
            for value in payload
        )
    if isinstance(payload, str) and _contains_worksheet_code_resource_registration_material(
        payload
    ):
        return WORKSHEET_CODE_RESOURCE_REGISTRATION_REDACTION
    return payload


def _worksheet_code_resource_registration_sensitive_change_locations(
    report: DiffReport,
) -> set[str]:
    """Return changed cells whose evidence can carry a registration input."""
    sensitive_cells = (
        report.before.worksheet_code_resource_registrations.registration_cells
        | report.after.worksheet_code_resource_registrations.registration_cells
        | report.worksheet_code_resource_registration_static_input_cells
    )
    if not sensitive_cells:
        return set()

    locations: set[str] = set()
    for change in report.changes:
        if change.location is None or change.location not in sensitive_cells:
            continue
        location = display_location(change.location)
        if location is not None:
            locations.add(location)
    return locations


def _redact_worksheet_code_resource_registration_change_cells(
    payload: dict[str, Any], report: DiffReport
) -> dict[str, Any]:
    """Hide before/after cells that register code or are exact static inputs."""
    sensitive_locations = _worksheet_code_resource_registration_sensitive_change_locations(
        report
    )
    if not sensitive_locations:
        return payload
    for change in payload.get("changes", []):
        if change.get("location") not in sensitive_locations:
            continue
        for side in ("before", "after"):
            snapshot = change.get(side)
            if not isinstance(snapshot, dict):
                continue
            for field in ("value", "formula", "formula_fingerprint"):
                if snapshot.get(field) is not None:
                    snapshot[field] = WORKSHEET_CODE_RESOURCE_REGISTRATION_REDACTION
    return payload


def _worksheet_code_resource_registration_definition_material_changed(
    report: DiffReport,
) -> bool:
    """Return whether a registration-relevant defined-name body changed."""
    return (
        report.before.worksheet_code_resource_registrations.definition_signature
        != report.after.worksheet_code_resource_registrations.definition_signature
    )


def _redact_worksheet_code_resource_registration_defined_name_details(
    details: dict[str, Any], report: DiffReport
) -> None:
    """Hide FF008 evidence when a resolved registration chain changed.

    A named LAMBDA can pass a private module or procedure through an
    ordinary-looking dotted workbook-defined wrapper whose eventual body calls
    ``REGISTER.ID``. Its own text cannot be classified without the private
    fixed-point resolution. When that private definition signature changes,
    hiding every changed defined-name body is the safe sharing boundary.
    """
    if not _worksheet_code_resource_registration_definition_material_changed(report):
        return
    for field in ("before", "after"):
        if isinstance(details.get(field), str):
            details[field] = WORKSHEET_CODE_RESOURCE_REGISTRATION_REDACTION


def _redact_worksheet_code_resource_registration_defined_name_evidence(
    payload: dict[str, Any], report: DiffReport
) -> dict[str, Any]:
    """Hide registration-resolved name-chain text in report evidence."""
    if not _worksheet_code_resource_registration_definition_material_changed(report):
        return payload
    for change in payload.get("changes", []):
        if change.get("kind") == "defined_name_changed":
            details = change.get("details")
            if isinstance(details, dict):
                _redact_worksheet_code_resource_registration_defined_name_details(
                    details, report
                )
    for finding in payload.get("findings", []):
        if finding.get("rule_id") != "FF008":
            continue
        details = finding.get("details")
        if isinstance(details, dict):
            _redact_worksheet_code_resource_registration_defined_name_details(
                details, report
            )
    return payload


def redact_worksheet_code_resource_registration_report_payload(
    report: DiffReport, payload: dict[str, Any]
) -> dict[str, Any]:
    """Redact direct registrations and private static/name-chain evidence."""
    redacted = redact_worksheet_code_resource_registration_material(payload)
    _redact_worksheet_code_resource_registration_change_cells(redacted, report)
    return _redact_worksheet_code_resource_registration_defined_name_evidence(
        redacted, report
    )


def redact_worksheet_code_resource_registration_portfolio_payload(
    report: PortfolioReport, payload: dict[str, Any]
) -> dict[str, Any]:
    """Apply the registration sharing boundary to nested reports."""
    redacted = redact_worksheet_code_resource_registration_material(payload)
    for workbook, entry in zip(
        report.workbooks, redacted.get("workbooks", []), strict=False
    ):
        if workbook.report is not None and isinstance(entry, dict):
            _redact_worksheet_code_resource_registration_change_cells(
                entry, workbook.report
            )
            _redact_worksheet_code_resource_registration_defined_name_evidence(
                entry, workbook.report
            )
    return redacted


def _contains_formula_defined_xlm_registration_material(value: str) -> bool:
    """Return whether one rendered string exposes a stored XLM ``REGISTER`` call.

    ``REGISTER`` is an XLM primitive whose formula-defined-name forms can
    register a DLL function, command, or XLL. The ordinary worksheet scanner
    deliberately excludes it, so this renderer explicitly enables the same
    defined-name-only classifier used by FF068. A renderer lacks name-scope
    resolution, making a standalone matching call a conservative redaction
    candidate rather than proof that it is the Excel primitive.
    """
    if "register" not in value.casefold() or "(" not in value:
        return False
    formula = value if value.lstrip().startswith("=") else f"={value}"
    return bool(
        inspect_formula(
            formula, inspect_formula_defined_xlm_registrations=True
        ).formula_defined_xlm_registration_functions
    )


def redact_formula_defined_xlm_registration_material(payload: Any) -> Any:
    """Return an output-only copy with direct stored ``REGISTER`` text hidden.

    This covers direct definition material. The report-level helpers below also
    hide exact changed static inputs and changed formula-defined-name bodies
    that the private FF068 comparison identifies as registration relevant.
    """
    if isinstance(payload, dict):
        return {
            key: redact_formula_defined_xlm_registration_material(value)
            for key, value in payload.items()
        }
    if isinstance(payload, list):
        return [
            redact_formula_defined_xlm_registration_material(value)
            for value in payload
        ]
    if isinstance(payload, tuple):
        return tuple(
            redact_formula_defined_xlm_registration_material(value)
            for value in payload
        )
    if isinstance(payload, str) and _contains_formula_defined_xlm_registration_material(
        payload
    ):
        return FORMULA_DEFINED_XLM_REGISTRATION_REDACTION
    return payload


def _formula_defined_xlm_registration_sensitive_change_locations(
    report: DiffReport,
) -> set[str]:
    """Return changed cells whose evidence can carry a registration input."""
    sensitive_cells = (
        report.before.formula_defined_xlm_registrations.registration_cells
        | report.after.formula_defined_xlm_registrations.registration_cells
        | report.formula_defined_xlm_registration_static_input_cells
    )
    if not sensitive_cells:
        return set()

    locations: set[str] = set()
    for change in report.changes:
        if change.location is None or change.location not in sensitive_cells:
            continue
        location = display_location(change.location)
        if location is not None:
            locations.add(location)
    return locations


def _redact_formula_defined_xlm_registration_change_cells(
    payload: dict[str, Any], report: DiffReport
) -> dict[str, Any]:
    """Hide before/after cells that invoke a stored registration or feed it."""
    sensitive_locations = _formula_defined_xlm_registration_sensitive_change_locations(
        report
    )
    if not sensitive_locations:
        return payload
    for change in payload.get("changes", []):
        if change.get("location") not in sensitive_locations:
            continue
        for side in ("before", "after"):
            snapshot = change.get(side)
            if not isinstance(snapshot, dict):
                continue
            for field in ("value", "formula", "formula_fingerprint"):
                if snapshot.get(field) is not None:
                    snapshot[field] = FORMULA_DEFINED_XLM_REGISTRATION_REDACTION
    return payload


def _formula_defined_xlm_registration_definition_material_changed(
    report: DiffReport,
) -> bool:
    """Return whether a stored registration-relevant definition changed."""
    return (
        report.before.formula_defined_xlm_registrations.definition_signature
        != report.after.formula_defined_xlm_registrations.definition_signature
    )


def _redact_formula_defined_xlm_registration_defined_name_details(
    details: dict[str, Any], report: DiffReport
) -> None:
    """Hide FF008 evidence when a resolved XLM registration chain changed.

    A named LAMBDA can pass module or procedure text through an ordinary-looking
    dotted workbook-defined wrapper whose eventual body calls ``REGISTER``.
    Its own text cannot be classified without the private fixed-point
    resolution. When that definition signature changes, hiding every changed
    defined-name body is the safe sharing boundary.
    """
    if not _formula_defined_xlm_registration_definition_material_changed(report):
        return
    for field in ("before", "after"):
        if isinstance(details.get(field), str):
            details[field] = FORMULA_DEFINED_XLM_REGISTRATION_REDACTION


def _redact_formula_defined_xlm_registration_defined_name_evidence(
    payload: dict[str, Any], report: DiffReport
) -> dict[str, Any]:
    """Hide registration-resolved name-chain text in report evidence."""
    if not _formula_defined_xlm_registration_definition_material_changed(report):
        return payload
    for change in payload.get("changes", []):
        if change.get("kind") == "defined_name_changed":
            details = change.get("details")
            if isinstance(details, dict):
                _redact_formula_defined_xlm_registration_defined_name_details(
                    details, report
                )
    for finding in payload.get("findings", []):
        if finding.get("rule_id") != "FF008":
            continue
        details = finding.get("details")
        if isinstance(details, dict):
            _redact_formula_defined_xlm_registration_defined_name_details(
                details, report
            )
    return payload


def redact_formula_defined_xlm_registration_report_payload(
    report: DiffReport, payload: dict[str, Any]
) -> dict[str, Any]:
    """Redact direct XLM registrations and private static/name-chain evidence."""
    redacted = redact_formula_defined_xlm_registration_material(payload)
    _redact_formula_defined_xlm_registration_change_cells(redacted, report)
    return _redact_formula_defined_xlm_registration_defined_name_evidence(
        redacted, report
    )


def redact_formula_defined_xlm_registration_portfolio_payload(
    report: PortfolioReport, payload: dict[str, Any]
) -> dict[str, Any]:
    """Apply the stored XLM registration sharing boundary to nested reports."""
    redacted = redact_formula_defined_xlm_registration_material(payload)
    for workbook, entry in zip(
        report.workbooks, redacted.get("workbooks", []), strict=False
    ):
        if workbook.report is not None and isinstance(entry, dict):
            _redact_formula_defined_xlm_registration_change_cells(
                entry, workbook.report
            )
            _redact_formula_defined_xlm_registration_defined_name_evidence(
                entry, workbook.report
            )
    return redacted


def _contains_formula_defined_xlm_evaluation_material(value: str) -> bool:
    """Return whether one rendered string exposes a stored XLM ``EVALUATE`` call.

    ``EVALUATE`` reduces supplied expression text at calculation time. The
    ordinary worksheet scanner deliberately excludes it, so this renderer
    explicitly enables the same defined-name-only classifier used by FF069. A
    renderer lacks name-scope resolution, making a standalone matching call a
    conservative redaction candidate rather than proof that it is the Excel
    primitive.
    """
    if "evaluate" not in value.casefold() or "(" not in value:
        return False
    formula = value if value.lstrip().startswith("=") else f"={value}"
    return bool(
        inspect_formula(
            formula, inspect_formula_defined_xlm_evaluations=True
        ).formula_defined_xlm_evaluation_functions
    )


def redact_formula_defined_xlm_evaluation_material(payload: Any) -> Any:
    """Return an output-only copy with direct stored ``EVALUATE`` text hidden.

    This covers direct definition material. The report-level helpers below also
    hide exact changed static inputs and changed formula-defined-name bodies
    that the private FF069 comparison identifies as evaluation relevant. It
    does not interpret or redact values referenced only inside runtime text.
    """
    if isinstance(payload, dict):
        return {
            key: redact_formula_defined_xlm_evaluation_material(value)
            for key, value in payload.items()
        }
    if isinstance(payload, list):
        return [
            redact_formula_defined_xlm_evaluation_material(value)
            for value in payload
        ]
    if isinstance(payload, tuple):
        return tuple(
            redact_formula_defined_xlm_evaluation_material(value)
            for value in payload
        )
    if isinstance(payload, str) and _contains_formula_defined_xlm_evaluation_material(
        payload
    ):
        return FORMULA_DEFINED_XLM_EVALUATION_REDACTION
    return payload


def _formula_defined_xlm_evaluation_sensitive_change_locations(
    report: DiffReport,
) -> set[str]:
    """Return changed cells whose evidence can carry a static expression input."""
    sensitive_cells = (
        report.before.formula_defined_xlm_evaluations.evaluation_cells
        | report.after.formula_defined_xlm_evaluations.evaluation_cells
        | report.formula_defined_xlm_evaluation_static_input_cells
    )
    if not sensitive_cells:
        return set()

    locations: set[str] = set()
    for change in report.changes:
        if change.location is None or change.location not in sensitive_cells:
            continue
        location = display_location(change.location)
        if location is not None:
            locations.add(location)
    return locations


def _redact_formula_defined_xlm_evaluation_change_cells(
    payload: dict[str, Any], report: DiffReport
) -> dict[str, Any]:
    """Hide before/after cells that invoke a stored evaluation or feed it."""
    sensitive_locations = _formula_defined_xlm_evaluation_sensitive_change_locations(
        report
    )
    if not sensitive_locations:
        return payload
    for change in payload.get("changes", []):
        if change.get("location") not in sensitive_locations:
            continue
        for side in ("before", "after"):
            snapshot = change.get(side)
            if not isinstance(snapshot, dict):
                continue
            for field in ("value", "formula", "formula_fingerprint"):
                if snapshot.get(field) is not None:
                    snapshot[field] = FORMULA_DEFINED_XLM_EVALUATION_REDACTION
    return payload


def _formula_defined_xlm_evaluation_definition_material_changed(
    report: DiffReport,
) -> bool:
    """Return whether a stored evaluation-relevant definition changed."""
    return (
        report.before.formula_defined_xlm_evaluations.definition_signature
        != report.after.formula_defined_xlm_evaluations.definition_signature
    )


def _redact_formula_defined_xlm_evaluation_defined_name_details(
    details: dict[str, Any], report: DiffReport
) -> None:
    """Hide FF008 evidence when a resolved XLM evaluation chain changed.

    A named LAMBDA can pass expression text through an ordinary-looking dotted
    workbook-defined wrapper whose eventual body calls ``EVALUATE``. Its own
    text cannot be classified without the private fixed-point resolution. When
    that definition signature changes, hiding every changed defined-name body
    is the safe sharing boundary.
    """
    if not _formula_defined_xlm_evaluation_definition_material_changed(report):
        return
    for field in ("before", "after"):
        if isinstance(details.get(field), str):
            details[field] = FORMULA_DEFINED_XLM_EVALUATION_REDACTION


def _redact_formula_defined_xlm_evaluation_defined_name_evidence(
    payload: dict[str, Any], report: DiffReport
) -> dict[str, Any]:
    """Hide evaluation-resolved name-chain text in report evidence."""
    if not _formula_defined_xlm_evaluation_definition_material_changed(report):
        return payload
    for change in payload.get("changes", []):
        if change.get("kind") == "defined_name_changed":
            details = change.get("details")
            if isinstance(details, dict):
                _redact_formula_defined_xlm_evaluation_defined_name_details(
                    details, report
                )
    for finding in payload.get("findings", []):
        if finding.get("rule_id") != "FF008":
            continue
        details = finding.get("details")
        if isinstance(details, dict):
            _redact_formula_defined_xlm_evaluation_defined_name_details(
                details, report
            )
    return payload


def redact_formula_defined_xlm_evaluation_report_payload(
    report: DiffReport, payload: dict[str, Any]
) -> dict[str, Any]:
    """Redact direct XLM evaluations and private static/name-chain evidence."""
    redacted = redact_formula_defined_xlm_evaluation_material(payload)
    _redact_formula_defined_xlm_evaluation_change_cells(redacted, report)
    return _redact_formula_defined_xlm_evaluation_defined_name_evidence(
        redacted, report
    )


def redact_formula_defined_xlm_evaluation_portfolio_payload(
    report: PortfolioReport, payload: dict[str, Any]
) -> dict[str, Any]:
    """Apply the stored XLM evaluation sharing boundary to nested reports."""
    redacted = redact_formula_defined_xlm_evaluation_material(payload)
    for workbook, entry in zip(
        report.workbooks, redacted.get("workbooks", []), strict=False
    ):
        if workbook.report is not None and isinstance(entry, dict):
            _redact_formula_defined_xlm_evaluation_change_cells(
                entry, workbook.report
            )
            _redact_formula_defined_xlm_evaluation_defined_name_evidence(
                entry, workbook.report
            )
    return redacted


def _contains_formula_defined_xlm_action_material(value: str) -> bool:
    """Return whether one rendered string exposes a selected stored XLM action.

    FF073 covers a finite set of legacy action and event-dispatch spellings only
    while FormulaFence is inspecting formula-defined names. A renderer lacks
    workbook name-scope resolution, so a standalone matching call is a
    conservative redaction candidate rather than proof that it is the XLM
    primitive.
    """
    folded = value.casefold()
    if "(" not in value or not any(
        token in folded for token in ("call", "exec", "run", "send.keys", "on.")
    ):
        return False
    formula = value if value.lstrip().startswith("=") else f"={value}"
    return bool(
        inspect_formula(
            formula, inspect_formula_defined_xlm_actions=True
        ).formula_defined_xlm_action_functions
    )


def redact_formula_defined_xlm_action_material(payload: Any) -> Any:
    """Return an output-only copy with direct selected XLM action text hidden.

    This covers direct definition material. The report-level helpers below also
    hide exact changed static inputs and changed formula-defined-name bodies
    that the private FF073 comparison identifies as action relevant. It does
    not resolve a target, handler, or dynamically assembled action.
    """
    if isinstance(payload, dict):
        return {
            key: redact_formula_defined_xlm_action_material(value)
            for key, value in payload.items()
        }
    if isinstance(payload, list):
        return [redact_formula_defined_xlm_action_material(value) for value in payload]
    if isinstance(payload, tuple):
        return tuple(redact_formula_defined_xlm_action_material(value) for value in payload)
    if isinstance(payload, str) and _contains_formula_defined_xlm_action_material(
        payload
    ):
        return FORMULA_DEFINED_XLM_ACTION_REDACTION
    return payload


def _formula_defined_xlm_action_sensitive_change_locations(
    report: DiffReport,
) -> set[str]:
    """Return changed cells whose evidence can carry an action input."""
    sensitive_cells = (
        report.before.formula_defined_xlm_actions.action_cells
        | report.after.formula_defined_xlm_actions.action_cells
        | report.formula_defined_xlm_action_static_input_cells
    )
    if not sensitive_cells:
        return set()

    locations: set[str] = set()
    for change in report.changes:
        if change.location is None or change.location not in sensitive_cells:
            continue
        location = display_location(change.location)
        if location is not None:
            locations.add(location)
    return locations


def _redact_formula_defined_xlm_action_change_cells(
    payload: dict[str, Any], report: DiffReport
) -> dict[str, Any]:
    """Hide before/after cells that invoke a stored action or feed it."""
    sensitive_locations = _formula_defined_xlm_action_sensitive_change_locations(report)
    if not sensitive_locations:
        return payload
    for change in payload.get("changes", []):
        if change.get("location") not in sensitive_locations:
            continue
        for side in ("before", "after"):
            snapshot = change.get(side)
            if not isinstance(snapshot, dict):
                continue
            for field in ("value", "formula", "formula_fingerprint"):
                if snapshot.get(field) is not None:
                    snapshot[field] = FORMULA_DEFINED_XLM_ACTION_REDACTION
    return payload


def _formula_defined_xlm_action_definition_material_changed(
    report: DiffReport,
) -> bool:
    """Return whether a stored action-relevant definition changed."""
    return (
        report.before.formula_defined_xlm_actions.definition_signature
        != report.after.formula_defined_xlm_actions.definition_signature
    )


def _redact_formula_defined_xlm_action_defined_name_details(
    details: dict[str, Any], report: DiffReport
) -> None:
    """Hide FF008 evidence when a resolved XLM action chain changed.

    A named LAMBDA can pass target or handler material through an
    ordinary-looking dotted workbook-defined wrapper whose eventual body calls
    a selected legacy XLM action. Its own text cannot be classified without the
    private fixed-point resolution. When that definition signature changes,
    hiding every changed defined-name body is the safe sharing boundary.
    """
    if not _formula_defined_xlm_action_definition_material_changed(report):
        return
    for field in ("before", "after"):
        if isinstance(details.get(field), str):
            details[field] = FORMULA_DEFINED_XLM_ACTION_REDACTION


def _redact_formula_defined_xlm_action_defined_name_evidence(
    payload: dict[str, Any], report: DiffReport
) -> dict[str, Any]:
    """Hide action-resolved name-chain text in report evidence."""
    if not _formula_defined_xlm_action_definition_material_changed(report):
        return payload
    for change in payload.get("changes", []):
        if change.get("kind") == "defined_name_changed":
            details = change.get("details")
            if isinstance(details, dict):
                _redact_formula_defined_xlm_action_defined_name_details(details, report)
    for finding in payload.get("findings", []):
        if finding.get("rule_id") != "FF008":
            continue
        details = finding.get("details")
        if isinstance(details, dict):
            _redact_formula_defined_xlm_action_defined_name_details(details, report)
    return payload


def redact_formula_defined_xlm_action_report_payload(
    report: DiffReport, payload: dict[str, Any]
) -> dict[str, Any]:
    """Redact direct selected XLM actions and private static/name-chain evidence."""
    redacted = redact_formula_defined_xlm_action_material(payload)
    _redact_formula_defined_xlm_action_change_cells(redacted, report)
    return _redact_formula_defined_xlm_action_defined_name_evidence(redacted, report)


def redact_formula_defined_xlm_action_portfolio_payload(
    report: PortfolioReport, payload: dict[str, Any]
) -> dict[str, Any]:
    """Apply the selected XLM action sharing boundary to nested reports."""
    redacted = redact_formula_defined_xlm_action_material(payload)
    for workbook, entry in zip(
        report.workbooks, redacted.get("workbooks", []), strict=False
    ):
        if workbook.report is not None and isinstance(entry, dict):
            _redact_formula_defined_xlm_action_change_cells(entry, workbook.report)
            _redact_formula_defined_xlm_action_defined_name_evidence(
                entry, workbook.report
            )
    return redacted


def _contains_formula_defined_xlm_get_cell_material(value: str) -> bool:
    """Return whether one string exposes a stored XLM ``GET.CELL`` call.

    FF070 only inventories this legacy information primitive while resolving
    formula-defined names. A generic renderer does not retain workbook name
    scope, so a standalone matching expression is conservatively redacted
    rather than treated as proof that Excel will invoke the primitive.
    """
    if "get.cell" not in value.casefold() or "(" not in value:
        return False
    formula = value if value.lstrip().startswith("=") else f"={value}"
    return bool(
        inspect_formula(
            formula, inspect_formula_defined_xlm_get_cell_calls=True
        ).formula_defined_xlm_get_cell_functions
    )


def redact_formula_defined_xlm_get_cell_material(payload: Any) -> Any:
    """Return an output-only copy with direct stored ``GET.CELL`` text hidden.

    The report-level helpers below also hide exact changed static inputs and
    changed formula-defined-name bodies which the private FF070 comparison
    resolves as GET.CELL relevant. This neither evaluates a call nor derives
    the information type, a dynamic reference, or display state.
    """
    if isinstance(payload, dict):
        return {
            key: redact_formula_defined_xlm_get_cell_material(value)
            for key, value in payload.items()
        }
    if isinstance(payload, list):
        return [redact_formula_defined_xlm_get_cell_material(value) for value in payload]
    if isinstance(payload, tuple):
        return tuple(
            redact_formula_defined_xlm_get_cell_material(value) for value in payload
        )
    if isinstance(payload, str) and _contains_formula_defined_xlm_get_cell_material(
        payload
    ):
        return FORMULA_DEFINED_XLM_GET_CELL_REDACTION
    return payload


def _formula_defined_xlm_get_cell_sensitive_change_locations(
    report: DiffReport,
) -> set[str]:
    """Return changed cells whose evidence can carry a static GET.CELL input."""
    sensitive_cells = (
        report.before.formula_defined_xlm_get_cell_calls.get_cell_cells
        | report.after.formula_defined_xlm_get_cell_calls.get_cell_cells
        | report.formula_defined_xlm_get_cell_static_input_cells
    )
    if not sensitive_cells:
        return set()

    locations: set[str] = set()
    for change in report.changes:
        if change.location is None or change.location not in sensitive_cells:
            continue
        location = display_location(change.location)
        if location is not None:
            locations.add(location)
    return locations


def _redact_formula_defined_xlm_get_cell_change_cells(
    payload: dict[str, Any], report: DiffReport
) -> dict[str, Any]:
    """Hide before/after cells that invoke a stored GET.CELL call or feed it."""
    sensitive_locations = _formula_defined_xlm_get_cell_sensitive_change_locations(
        report
    )
    if not sensitive_locations:
        return payload
    for change in payload.get("changes", []):
        if change.get("location") not in sensitive_locations:
            continue
        for side in ("before", "after"):
            snapshot = change.get(side)
            if not isinstance(snapshot, dict):
                continue
            for field in ("value", "formula", "formula_fingerprint"):
                if snapshot.get(field) is not None:
                    snapshot[field] = FORMULA_DEFINED_XLM_GET_CELL_REDACTION
    return payload


def _formula_defined_xlm_get_cell_definition_material_changed(
    report: DiffReport,
) -> bool:
    """Return whether a stored GET.CELL-relevant definition changed."""
    return (
        report.before.formula_defined_xlm_get_cell_calls.definition_signature
        != report.after.formula_defined_xlm_get_cell_calls.definition_signature
    )


def _redact_formula_defined_xlm_get_cell_defined_name_details(
    details: dict[str, Any], report: DiffReport
) -> None:
    """Hide FF008 evidence when a resolved GET.CELL name chain changed.

    A named LAMBDA can pass a reference or information code through an
    ordinary-looking dotted wrapper whose eventual body calls ``GET.CELL``.
    Its own text cannot be classified without FormulaFence's private
    fixed-point resolution. When that definition signature changes, replacing
    every changed defined-name body is the safe sharing boundary.
    """
    if not _formula_defined_xlm_get_cell_definition_material_changed(report):
        return
    for field in ("before", "after"):
        if isinstance(details.get(field), str):
            details[field] = FORMULA_DEFINED_XLM_GET_CELL_REDACTION


def _redact_formula_defined_xlm_get_cell_defined_name_evidence(
    payload: dict[str, Any], report: DiffReport
) -> dict[str, Any]:
    """Hide resolved GET.CELL name-chain text in report evidence."""
    if not _formula_defined_xlm_get_cell_definition_material_changed(report):
        return payload
    for change in payload.get("changes", []):
        if change.get("kind") == "defined_name_changed":
            details = change.get("details")
            if isinstance(details, dict):
                _redact_formula_defined_xlm_get_cell_defined_name_details(
                    details, report
                )
    for finding in payload.get("findings", []):
        if finding.get("rule_id") != "FF008":
            continue
        details = finding.get("details")
        if isinstance(details, dict):
            _redact_formula_defined_xlm_get_cell_defined_name_details(details, report)
    return payload


def redact_formula_defined_xlm_get_cell_report_payload(
    report: DiffReport, payload: dict[str, Any]
) -> dict[str, Any]:
    """Redact direct XLM GET.CELL and private static/name-chain evidence."""
    redacted = redact_formula_defined_xlm_get_cell_material(payload)
    _redact_formula_defined_xlm_get_cell_change_cells(redacted, report)
    return _redact_formula_defined_xlm_get_cell_defined_name_evidence(redacted, report)


def redact_formula_defined_xlm_get_cell_portfolio_payload(
    report: PortfolioReport, payload: dict[str, Any]
) -> dict[str, Any]:
    """Apply the stored XLM GET.CELL sharing boundary to nested reports."""
    redacted = redact_formula_defined_xlm_get_cell_material(payload)
    for workbook, entry in zip(
        report.workbooks, redacted.get("workbooks", []), strict=False
    ):
        if workbook.report is not None and isinstance(entry, dict):
            _redact_formula_defined_xlm_get_cell_change_cells(entry, workbook.report)
            _redact_formula_defined_xlm_get_cell_defined_name_evidence(
                entry, workbook.report
            )
    return redacted


def _contains_formula_defined_xlm_environment_information_material(value: str) -> bool:
    """Return whether a string exposes a selected stored XLM environment call.

    FF071 only inventories GET.WORKBOOK, GET.WORKSPACE, and GET.DOCUMENT while
    resolving formula-defined names. A generic renderer does not retain name
    scope, so a standalone matching expression is conservatively redacted
    rather than treated as proof that Excel will invoke the primitive.
    """
    folded = value.casefold()
    if "(" not in value or not any(
        token in folded
        for token in ("get.workbook", "get.workspace", "get.document")
    ):
        return False
    formula = value if value.lstrip().startswith("=") else f"={value}"
    return bool(
        inspect_formula(
            formula,
            inspect_formula_defined_xlm_environment_information_calls=True,
        ).formula_defined_xlm_environment_information_functions
    )


def redact_formula_defined_xlm_environment_information_material(payload: Any) -> Any:
    """Return an output-only copy with selected stored XLM calls hidden.

    The report-level helpers below also hide exact changed static inputs and
    changed formula-defined-name bodies which the private FF071 comparison
    resolves as environment-information relevant. This neither evaluates a
    call nor determines an information type, resolves a dynamic reference, or
    simulates workbook, workspace, or document state.
    """
    if isinstance(payload, dict):
        return {
            key: redact_formula_defined_xlm_environment_information_material(value)
            for key, value in payload.items()
        }
    if isinstance(payload, list):
        return [
            redact_formula_defined_xlm_environment_information_material(value)
            for value in payload
        ]
    if isinstance(payload, tuple):
        return tuple(
            redact_formula_defined_xlm_environment_information_material(value)
            for value in payload
        )
    if (
        isinstance(payload, str)
        and _contains_formula_defined_xlm_environment_information_material(payload)
    ):
        return FORMULA_DEFINED_XLM_ENVIRONMENT_INFORMATION_REDACTION
    return payload


def _formula_defined_xlm_environment_information_sensitive_change_locations(
    report: DiffReport,
) -> set[str]:
    """Return cells whose evidence can carry a static environment call input."""
    before_calls = report.before.formula_defined_xlm_environment_information_calls
    after_calls = report.after.formula_defined_xlm_environment_information_calls
    sensitive_cells = (
        before_calls.environment_information_cells
        | after_calls.environment_information_cells
        | report.formula_defined_xlm_environment_information_static_input_cells
    )
    if not sensitive_cells:
        return set()

    locations: set[str] = set()
    for change in report.changes:
        if change.location is None or change.location not in sensitive_cells:
            continue
        location = display_location(change.location)
        if location is not None:
            locations.add(location)
    return locations


def _redact_formula_defined_xlm_environment_information_change_cells(
    payload: dict[str, Any], report: DiffReport
) -> dict[str, Any]:
    """Hide cells that invoke a selected stored XLM call or feed it."""
    sensitive_locations = (
        _formula_defined_xlm_environment_information_sensitive_change_locations(
            report
        )
    )
    if not sensitive_locations:
        return payload
    for change in payload.get("changes", []):
        if change.get("location") not in sensitive_locations:
            continue
        for side in ("before", "after"):
            snapshot = change.get(side)
            if not isinstance(snapshot, dict):
                continue
            for field in ("value", "formula", "formula_fingerprint"):
                if snapshot.get(field) is not None:
                    snapshot[field] = (
                        FORMULA_DEFINED_XLM_ENVIRONMENT_INFORMATION_REDACTION
                    )
    return payload


def _formula_defined_xlm_environment_information_definition_material_changed(
    report: DiffReport,
) -> bool:
    """Return whether a selected stored XLM definition changed."""
    return (
        report.before.formula_defined_xlm_environment_information_calls.definition_signature
        != report.after.formula_defined_xlm_environment_information_calls.definition_signature
    )


def _redact_formula_defined_xlm_environment_information_defined_name_details(
    details: dict[str, Any], report: DiffReport
) -> None:
    """Hide FF008 evidence when a resolved environment name chain changed.

    A named LAMBDA can pass a private information code or reference through an
    ordinary-looking dotted wrapper whose eventual body calls one of the
    selected XLM primitives. Its own text cannot be classified without the
    private fixed-point resolution. When that definition signature changes,
    replacing every changed defined-name body is the safe sharing boundary.
    """
    if not _formula_defined_xlm_environment_information_definition_material_changed(
        report
    ):
        return
    for field in ("before", "after"):
        if isinstance(details.get(field), str):
            details[field] = FORMULA_DEFINED_XLM_ENVIRONMENT_INFORMATION_REDACTION


def _redact_formula_defined_xlm_environment_information_defined_name_evidence(
    payload: dict[str, Any], report: DiffReport
) -> dict[str, Any]:
    """Hide resolved XLM environment-information name-chain report evidence."""
    if not _formula_defined_xlm_environment_information_definition_material_changed(
        report
    ):
        return payload
    for change in payload.get("changes", []):
        if change.get("kind") == "defined_name_changed":
            details = change.get("details")
            if isinstance(details, dict):
                _redact_formula_defined_xlm_environment_information_defined_name_details(
                    details, report
                )
    for finding in payload.get("findings", []):
        if finding.get("rule_id") != "FF008":
            continue
        details = finding.get("details")
        if isinstance(details, dict):
            _redact_formula_defined_xlm_environment_information_defined_name_details(
                details, report
            )
    return payload


def redact_formula_defined_xlm_environment_information_report_payload(
    report: DiffReport, payload: dict[str, Any]
) -> dict[str, Any]:
    """Redact selected XLM calls and private static/name-chain evidence."""
    redacted = redact_formula_defined_xlm_environment_information_material(payload)
    _redact_formula_defined_xlm_environment_information_change_cells(redacted, report)
    return _redact_formula_defined_xlm_environment_information_defined_name_evidence(
        redacted, report
    )


def redact_formula_defined_xlm_environment_information_portfolio_payload(
    report: PortfolioReport, payload: dict[str, Any]
) -> dict[str, Any]:
    """Apply the selected XLM environment sharing boundary to nested reports."""
    redacted = redact_formula_defined_xlm_environment_information_material(payload)
    for workbook, entry in zip(
        report.workbooks, redacted.get("workbooks", []), strict=False
    ):
        if workbook.report is not None and isinstance(entry, dict):
            _redact_formula_defined_xlm_environment_information_change_cells(
                entry, workbook.report
            )
            _redact_formula_defined_xlm_environment_information_defined_name_evidence(
                entry, workbook.report
            )
    return redacted


def _contains_formula_environment_information_material(value: str) -> bool:
    """Return whether a string exposes a native environment-information call.

    FF072 inventories native CELL, INFO, SHEET, and SHEETS calls in formulas
    and formula-defined names. A generic renderer does not retain name scope,
    so a standalone matching expression is conservatively redacted rather than
    treated as proof that Excel will invoke the native function.
    """
    folded = value.casefold()
    if "(" not in value or not any(
        token in folded for token in ("cell", "info", "sheet", "sheets")
    ):
        return False
    formula = value if value.lstrip().startswith("=") else f"={value}"
    return bool(inspect_formula(formula).formula_environment_information_functions)


def redact_formula_environment_information_material(payload: Any) -> Any:
    """Return an output-only copy with native environment-call text hidden.

    The report-level helpers below also hide exact changed static inputs and
    changed formula-defined-name bodies which the private FF072 comparison
    resolves as environment-information relevant. This neither evaluates a
    call nor determines an information type, resolves a dynamic reference, or
    simulates file, client, workspace, selection, or workbook state.
    """
    if isinstance(payload, dict):
        return {
            key: redact_formula_environment_information_material(value)
            for key, value in payload.items()
        }
    if isinstance(payload, list):
        return [redact_formula_environment_information_material(value) for value in payload]
    if isinstance(payload, tuple):
        return tuple(
            redact_formula_environment_information_material(value) for value in payload
        )
    if isinstance(payload, str) and _contains_formula_environment_information_material(
        payload
    ):
        return FORMULA_ENVIRONMENT_INFORMATION_REDACTION
    return payload


def _formula_environment_information_sensitive_change_locations(
    report: DiffReport,
) -> set[str]:
    """Return changed cells whose evidence can carry a native call input."""
    before_calls = report.before.formula_environment_information_calls
    after_calls = report.after.formula_environment_information_calls
    sensitive_cells = (
        before_calls.environment_information_cells
        | after_calls.environment_information_cells
        | report.formula_environment_information_static_input_cells
    )
    if not sensitive_cells:
        return set()

    locations: set[str] = set()
    for change in report.changes:
        if change.location is None or change.location not in sensitive_cells:
            continue
        location = display_location(change.location)
        if location is not None:
            locations.add(location)
    return locations


def _redact_formula_environment_information_change_cells(
    payload: dict[str, Any], report: DiffReport
) -> dict[str, Any]:
    """Hide cells that invoke a native environment call or feed it."""
    sensitive_locations = _formula_environment_information_sensitive_change_locations(
        report
    )
    if not sensitive_locations:
        return payload
    for change in payload.get("changes", []):
        if change.get("location") not in sensitive_locations:
            continue
        for side in ("before", "after"):
            snapshot = change.get(side)
            if not isinstance(snapshot, dict):
                continue
            for field in ("value", "formula", "formula_fingerprint"):
                if snapshot.get(field) is not None:
                    snapshot[field] = FORMULA_ENVIRONMENT_INFORMATION_REDACTION
    return payload


def _formula_environment_information_definition_material_changed(
    report: DiffReport,
) -> bool:
    """Return whether a native environment-information definition changed."""
    return (
        report.before.formula_environment_information_calls.definition_signature
        != report.after.formula_environment_information_calls.definition_signature
    )


def _redact_formula_environment_information_defined_name_details(
    details: dict[str, Any], report: DiffReport
) -> None:
    """Hide FF008 evidence when a resolved native name chain changed.

    A named LAMBDA can pass a private information code or reference through an
    ordinary-looking dotted wrapper whose eventual body calls CELL, INFO, SHEET,
    or SHEETS. Its own text cannot be classified without FormulaFence's private
    fixed-point resolution. When that definition signature changes, replacing
    every changed defined-name body is the safe sharing boundary.
    """
    if not _formula_environment_information_definition_material_changed(report):
        return
    for field in ("before", "after"):
        if isinstance(details.get(field), str):
            details[field] = FORMULA_ENVIRONMENT_INFORMATION_REDACTION


def _redact_formula_environment_information_defined_name_evidence(
    payload: dict[str, Any], report: DiffReport
) -> dict[str, Any]:
    """Hide resolved native environment-information name-chain evidence."""
    if not _formula_environment_information_definition_material_changed(report):
        return payload
    for change in payload.get("changes", []):
        if change.get("kind") == "defined_name_changed":
            details = change.get("details")
            if isinstance(details, dict):
                _redact_formula_environment_information_defined_name_details(
                    details, report
                )
    for finding in payload.get("findings", []):
        if finding.get("rule_id") != "FF008":
            continue
        details = finding.get("details")
        if isinstance(details, dict):
            _redact_formula_environment_information_defined_name_details(
                details, report
            )
    return payload


def redact_formula_environment_information_report_payload(
    report: DiffReport, payload: dict[str, Any]
) -> dict[str, Any]:
    """Redact native environment calls and private static/name-chain evidence."""
    redacted = redact_formula_environment_information_material(payload)
    _redact_formula_environment_information_change_cells(redacted, report)
    return _redact_formula_environment_information_defined_name_evidence(
        redacted, report
    )


def redact_formula_environment_information_portfolio_payload(
    report: PortfolioReport, payload: dict[str, Any]
) -> dict[str, Any]:
    """Apply the native environment sharing boundary to nested reports."""
    redacted = redact_formula_environment_information_material(payload)
    for workbook, entry in zip(
        report.workbooks, redacted.get("workbooks", []), strict=False
    ):
        if workbook.report is not None and isinstance(entry, dict):
            _redact_formula_environment_information_change_cells(
                entry, workbook.report
            )
            _redact_formula_environment_information_defined_name_evidence(
                entry, workbook.report
            )
    return redacted


def _contains_python_in_excel_material(value: str) -> bool:
    """Return whether one rendered string exposes a direct stored ``PY`` call.

    Microsoft documents ``PY``'s first argument as static Python source text.
    The formula can therefore disclose source or an ``xl()`` reference even
    though FormulaFence's FF065 ledger deliberately publishes only safe
    aggregates. Keep this lexical prefilter small, then defer recognition to
    the same formula inspector that inventories Python-in-Excel cells.
    """
    if "PY" not in value.upper() or "(" not in value:
        return False
    formula = value if value.lstrip().startswith("=") else f"={value}"
    return bool(inspect_formula(formula).python_functions)


def redact_python_in_excel_material(payload: Any) -> Any:
    """Return an output-only copy with direct ``PY`` formula text hidden.

    This covers source text held directly in a formula. Report-level helpers
    additionally hide changed ordinary cells whose complete static impact set
    reaches an inventoried PY cell, because an ``xl()`` input can be visible as
    a normal semantic cell change rather than inside the Python formula.
    """
    if isinstance(payload, dict):
        return {
            key: redact_python_in_excel_material(value) for key, value in payload.items()
        }
    if isinstance(payload, list):
        return [redact_python_in_excel_material(value) for value in payload]
    if isinstance(payload, tuple):
        return tuple(redact_python_in_excel_material(value) for value in payload)
    if isinstance(payload, str) and _contains_python_in_excel_material(payload):
        return PYTHON_IN_EXCEL_REDACTION
    return payload


def _python_in_excel_sensitive_change_locations(report: DiffReport) -> set[str]:
    """Return changed cells whose evidence can expose PY source or input data."""
    sensitive_cells = (
        report.before.python_in_excel.python_cells
        | report.after.python_in_excel.python_cells
        | report.python_in_excel_static_input_cells
    )
    if not sensitive_cells:
        return set()

    locations: set[str] = set()
    for change in report.changes:
        if change.location is None or change.location not in sensitive_cells:
            continue
        location = display_location(change.location)
        if location is not None:
            locations.add(location)
    return locations


def _redact_python_in_excel_change_cells(
    payload: dict[str, Any], report: DiffReport
) -> dict[str, Any]:
    """Hide raw before/after cells that are a PY formula or exact static input."""
    sensitive_locations = _python_in_excel_sensitive_change_locations(report)
    if not sensitive_locations:
        return payload
    for change in payload.get("changes", []):
        if change.get("location") not in sensitive_locations:
            continue
        for side in ("before", "after"):
            snapshot = change.get(side)
            if not isinstance(snapshot, dict):
                continue
            for field in ("value", "formula", "formula_fingerprint"):
                if snapshot.get(field) is not None:
                    snapshot[field] = PYTHON_IN_EXCEL_REDACTION
    return payload


def redact_python_in_excel_report_payload(
    report: DiffReport, payload: dict[str, Any]
) -> dict[str, Any]:
    """Redact direct PY source and private static-input cell evidence."""
    redacted = redact_python_in_excel_material(payload)
    return _redact_python_in_excel_change_cells(redacted, report)


def redact_python_in_excel_portfolio_payload(
    report: PortfolioReport, payload: dict[str, Any]
) -> dict[str, Any]:
    """Apply the Python-in-Excel sharing boundary to nested reports."""
    redacted = redact_python_in_excel_material(payload)
    for workbook, entry in zip(
        report.workbooks, redacted.get("workbooks", []), strict=False
    ):
        if workbook.report is not None and isinstance(entry, dict):
            _redact_python_in_excel_change_cells(entry, workbook.report)
    return redacted


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


def profile_to_markdown(
    profile: dict[str, Any], *, max_bytes: int | None = None
) -> str:
    workbook = profile["workbook"]
    initial_lines = [
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
            "- **Namespaced custom-function formula cells / calls / namespaces / "
            "named definitions:** "
            f"{workbook['namespaced_custom_function_formula_cell_count']} / "
            f"{workbook['namespaced_custom_function_call_count']} / "
            f"{workbook['namespaced_custom_function_namespace_count']} / "
            f"{workbook['namespaced_custom_function_defined_name_count']}"
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
    lines = _BoundedLines(initial_lines, max_bytes=max_bytes)
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
                    "- **Formula cells / calls / namespaces / named definitions:** "
                    f"{office_custom_functions['namespaced_custom_function_formula_cell_count']} / "
                    f"{office_custom_functions['namespaced_custom_function_call_count']} / "
                    f"{office_custom_functions['namespaced_custom_function_namespace_count']} / "
                    f"{office_custom_functions['namespaced_custom_function_defined_name_count']}"
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
    return lines.render()


def _report_payload_for_rendering(
    report: DiffReport,
    extra_findings: Iterable[Finding] = (),
    *,
    redact_external_workbook_links: bool = False,
    redact_formula_external_actions: bool = False,
    redact_python_in_excel: bool = False,
    redact_office_custom_functions: bool = False,
    redact_unqualified_runtime_functions: bool = False,
    redact_worksheet_code_resource_registrations: bool = False,
    redact_formula_defined_xlm_registrations: bool = False,
    redact_formula_defined_xlm_evaluations: bool = False,
    redact_formula_defined_xlm_actions: bool = False,
    redact_formula_defined_xlm_get_cell_calls: bool = False,
    redact_formula_defined_xlm_environment_information_calls: bool = False,
    redact_formula_environment_information: bool = False,
) -> dict[str, Any]:
    """Return a report payload with the requested sharing boundaries applied."""
    payload = report.to_dict(extra_findings)
    if redact_external_workbook_links:
        payload = redact_external_workbook_link_material(payload)
    if redact_formula_external_actions:
        payload = redact_formula_external_action_report_payload(report, payload)
    if redact_python_in_excel:
        payload = redact_python_in_excel_report_payload(report, payload)
    if redact_office_custom_functions:
        payload = redact_office_custom_function_report_payload(report, payload)
    if redact_unqualified_runtime_functions:
        payload = redact_unqualified_runtime_function_report_payload(report, payload)
    if redact_worksheet_code_resource_registrations:
        payload = redact_worksheet_code_resource_registration_report_payload(
            report, payload
        )
    if redact_formula_defined_xlm_registrations:
        payload = redact_formula_defined_xlm_registration_report_payload(
            report, payload
        )
    if redact_formula_defined_xlm_evaluations:
        payload = redact_formula_defined_xlm_evaluation_report_payload(report, payload)
    if redact_formula_defined_xlm_actions:
        payload = redact_formula_defined_xlm_action_report_payload(report, payload)
    if redact_formula_defined_xlm_get_cell_calls:
        payload = redact_formula_defined_xlm_get_cell_report_payload(report, payload)
    if redact_formula_defined_xlm_environment_information_calls:
        payload = redact_formula_defined_xlm_environment_information_report_payload(
            report, payload
        )
    if redact_formula_environment_information:
        payload = redact_formula_environment_information_report_payload(report, payload)
    return payload


def _portfolio_payload_for_rendering(
    report: PortfolioReport,
    *,
    redact_external_workbook_links: bool = False,
    redact_formula_external_actions: bool = False,
    redact_python_in_excel: bool = False,
    redact_office_custom_functions: bool = False,
    redact_unqualified_runtime_functions: bool = False,
    redact_worksheet_code_resource_registrations: bool = False,
    redact_formula_defined_xlm_registrations: bool = False,
    redact_formula_defined_xlm_evaluations: bool = False,
    redact_formula_defined_xlm_actions: bool = False,
    redact_formula_defined_xlm_get_cell_calls: bool = False,
    redact_formula_defined_xlm_environment_information_calls: bool = False,
    redact_formula_environment_information: bool = False,
) -> dict[str, Any]:
    """Return a portfolio payload with the requested sharing boundaries applied."""
    payload = report.to_dict()
    if redact_external_workbook_links:
        payload = redact_external_workbook_link_material(payload)
    if redact_formula_external_actions:
        payload = redact_formula_external_action_portfolio_payload(report, payload)
    if redact_python_in_excel:
        payload = redact_python_in_excel_portfolio_payload(report, payload)
    if redact_office_custom_functions:
        payload = redact_office_custom_function_portfolio_payload(report, payload)
    if redact_unqualified_runtime_functions:
        payload = redact_unqualified_runtime_function_portfolio_payload(report, payload)
    if redact_worksheet_code_resource_registrations:
        payload = redact_worksheet_code_resource_registration_portfolio_payload(
            report, payload
        )
    if redact_formula_defined_xlm_registrations:
        payload = redact_formula_defined_xlm_registration_portfolio_payload(
            report, payload
        )
    if redact_formula_defined_xlm_evaluations:
        payload = redact_formula_defined_xlm_evaluation_portfolio_payload(
            report, payload
        )
    if redact_formula_defined_xlm_actions:
        payload = redact_formula_defined_xlm_action_portfolio_payload(report, payload)
    if redact_formula_defined_xlm_get_cell_calls:
        payload = redact_formula_defined_xlm_get_cell_portfolio_payload(report, payload)
    if redact_formula_defined_xlm_environment_information_calls:
        payload = redact_formula_defined_xlm_environment_information_portfolio_payload(
            report, payload
        )
    if redact_formula_environment_information:
        payload = redact_formula_environment_information_portfolio_payload(
            report, payload
        )
    return payload


def _active_redaction_labels(
    *,
    redact_external_workbook_links: bool = False,
    redact_formula_external_actions: bool = False,
    redact_python_in_excel: bool = False,
    redact_office_custom_functions: bool = False,
    redact_unqualified_runtime_functions: bool = False,
    redact_worksheet_code_resource_registrations: bool = False,
    redact_formula_defined_xlm_registrations: bool = False,
    redact_formula_defined_xlm_evaluations: bool = False,
    redact_formula_defined_xlm_actions: bool = False,
    redact_formula_defined_xlm_get_cell_calls: bool = False,
    redact_formula_defined_xlm_environment_information_calls: bool = False,
    redact_formula_environment_information: bool = False,
) -> list[str]:
    """Return stable labels for active output-only sharing boundaries."""
    labels: list[str] = []
    if redact_external_workbook_links:
        labels.append("External-workbook link material redacted for sharing")
    if redact_formula_external_actions:
        labels.append("Formula external-action / DDE material redacted for sharing")
    if redact_python_in_excel:
        labels.append("Python-in-Excel material redacted for sharing")
    if redact_office_custom_functions:
        labels.append("Office custom-function material redacted for sharing")
    if redact_unqualified_runtime_functions:
        labels.append("Unqualified runtime-function material redacted for sharing")
    if redact_worksheet_code_resource_registrations:
        labels.append("Worksheet code-resource registration material redacted for sharing")
    if redact_formula_defined_xlm_registrations:
        labels.append("Formula-defined XLM registration material redacted for sharing")
    if redact_formula_defined_xlm_evaluations:
        labels.append("Formula-defined XLM evaluation material redacted for sharing")
    if redact_formula_defined_xlm_actions:
        labels.append("Formula-defined XLM action material redacted for sharing")
    if redact_formula_defined_xlm_get_cell_calls:
        labels.append("Formula-defined XLM GET.CELL material redacted for sharing")
    if redact_formula_defined_xlm_environment_information_calls:
        labels.append(
            "Formula-defined XLM environment-information material redacted for sharing"
        )
    if redact_formula_environment_information:
        labels.append("Formula environment-information material redacted for sharing")
    return labels


_HTML_STYLE = """
:root {
  color-scheme: light dark;
  font-family: Inter, Arial, sans-serif;
  line-height: 1.5;
}
body {
  background: #f6f8fa;
  color: #1f2328;
  margin: 0;
}
main {
  margin: 0 auto;
  max-width: 1120px;
  padding: 2.5rem 1.25rem 4rem;
}
h1, h2, h3, p { margin-top: 0; }
header { margin-bottom: 1.5rem; }
.lede { color: #57606a; max-width: 70ch; }
.cards {
  display: grid;
  gap: .75rem;
  grid-template-columns: repeat(auto-fit, minmax(150px, 1fr));
  margin: 1.25rem 0;
}
.card, .review-entry, .workbook, .filters, .redactions {
  background: #fff;
  border: 1px solid #d0d7de;
  border-radius: .55rem;
  box-shadow: 0 1px 2px rgb(31 35 40 / 4%);
}
.card { padding: .85rem 1rem; }
.card dt { color: #57606a; font-size: .83rem; }
.card dd { font-size: 1.45rem; font-weight: 700; margin: .2rem 0 0; }
.filters { align-items: end; display: flex; flex-wrap: wrap; gap: .75rem; padding: 1rem; }
.filters label { display: grid; font-weight: 600; gap: .25rem; }
input, select { font: inherit; padding: .38rem .5rem; }
#filter-status { color: #57606a; margin-left: auto; }
.redactions { border-color: #bf8700; margin: 1rem 0; padding: .85rem 1rem; }
.redactions ul { margin: .35rem 0 0; padding-left: 1.25rem; }
.review-section { margin-top: 2rem; }
.review-entry { border-left-width: .45rem; margin: .8rem 0; padding: 1rem; }
.review-entry h3 {
  align-items: center;
  display: flex;
  flex-wrap: wrap;
  gap: .5rem;
  margin-bottom: .75rem;
}
.severity-critical, .severity-high { border-left-color: #cf222e; }
.severity-medium { border-left-color: #bf8700; }
.severity-low { border-left-color: #0969da; }
.severity-note { border-left-color: #57606a; }
.severity {
  border-radius: 999px;
  font-size: .75rem;
  font-weight: 700;
  padding: .12rem .5rem;
  text-transform: uppercase;
}
.severity-critical .severity, .severity-high .severity { color: #cf222e; }
.severity-medium .severity { color: #9a6700; }
.severity-low .severity { color: #0969da; }
.severity-note .severity { color: #57606a; }
.metadata { display: flex; flex-wrap: wrap; gap: .5rem 1rem; margin: 0 0 .8rem; }
.metadata span { color: #57606a; }
details { margin-top: .8rem; }
summary { cursor: pointer; font-weight: 600; }
pre {
  background: #f6f8fa;
  border-radius: .4rem;
  overflow-x: auto;
  padding: .8rem;
  white-space: pre-wrap;
  word-break: break-word;
}
code { font-family: ui-monospace, SFMono-Regular, Consolas, monospace; }
.workbook { margin: 1.5rem 0; padding: 1.25rem; }
.empty { color: #57606a; font-style: italic; }
@media (prefers-color-scheme: dark) {
  body { background: #0d1117; color: #e6edf3; }
  .lede, #filter-status, .metadata span, .empty { color: #8b949e; }
  .card, .review-entry, .workbook, .filters, .redactions {
    background: #161b22;
    border-color: #30363d;
  }
  pre { background: #0d1117; }
}
""".strip()


_HTML_FILTER_SCRIPT = """
(() => {
  const query = document.getElementById("review-filter");
  const severity = document.getElementById("severity-filter");
  const status = document.getElementById("filter-status");
  const entries = [...document.querySelectorAll(".review-entry")];
  const apply = () => {
    const text = query.value.trim().toLocaleLowerCase();
    const level = severity.value;
    let visible = 0;
    for (const entry of entries) {
      const matchesText = !text || entry.innerText.toLocaleLowerCase().includes(text);
      const matchesSeverity = level === "all" || entry.dataset.severity === level;
      entry.hidden = !(matchesText && matchesSeverity);
      if (!entry.hidden) visible += 1;
    }
    status.textContent = `${visible} of ${entries.length} entries shown`;
  };
  query.addEventListener("input", apply);
  severity.addEventListener("change", apply);
  apply();
})();
""".strip()


def _html_text(value: object) -> str:
    """Escape untrusted review material before placing it in HTML text."""
    return _html_escape(str(value), quote=True)


def _html_json(value: object) -> str:
    """Render JSON evidence as escaped text, never executable page content."""
    return _html_text(json.dumps(value, indent=2, sort_keys=True, ensure_ascii=False))


def _html_review_entry(entry: dict[str, Any], category: str) -> str:
    """Render one finding or change with escaped expandable evidence."""
    severity = str(entry.get("severity", "note"))
    identifier = str(entry.get("rule_id") if category == "finding" else entry.get("kind"))
    location = entry.get("location") or "workbook"
    title = entry.get("message") if category == "finding" else identifier
    metadata = [f"<span><strong>Location:</strong> {_html_text(location)}</span>"]
    if category == "change":
        metadata.append(
            "<span><strong>Downstream formulas:</strong> "
            f"{_html_text(entry.get('impact_count', 0))}</span>"
        )
    return "\n".join(
        [
            (
                f'<article class="review-entry severity-{_html_text(severity)}" '
                f'data-severity="{_html_text(severity)}">'
            ),
            "<h3>"
            f"<code>{_html_text(identifier)}</code> "
            f'<span class="severity">{_html_text(severity)}</span>'
            "</h3>",
            f"<p>{_html_text(title)}</p>",
            f'<div class="metadata">{"".join(metadata)}</div>',
            "<details>",
            "<summary>Review evidence</summary>",
            f"<pre>{_html_json(entry)}</pre>",
            "</details>",
            "</article>",
        ]
    )


def _html_review_section(
    heading: str, entries: list[dict[str, Any]], category: str
) -> str:
    """Render an HTML review section or an explicit empty-state note."""
    content = (
        "\n".join(_html_review_entry(entry, category) for entry in entries)
        if entries
        else '<p class="empty">No entries.</p>'
    )
    return "\n".join(
        [
            '<section class="review-section">',
            f"<h2>{_html_text(heading)}</h2>",
            content,
            "</section>",
        ]
    )


def _append_html_line(rendered: _BoundedText, value: str) -> None:
    """Append one HTML line while accounting for its final UTF-8 bytes."""
    rendered.append(value)
    rendered.append("\n")


def _append_html_json(rendered: _BoundedText, value: object) -> None:
    """Write escaped JSON evidence incrementally instead of one giant string."""
    encoder = json.JSONEncoder(indent=2, sort_keys=True, ensure_ascii=False)
    for chunk in encoder.iterencode(value):
        rendered.append(_html_escape(chunk, quote=True))


def _append_html_review_entry(
    rendered: _BoundedText, entry: dict[str, Any], category: str
) -> None:
    """Write one bounded HTML review entry with complete escaped evidence."""
    severity = str(entry.get("severity", "note"))
    identifier = str(entry.get("rule_id") if category == "finding" else entry.get("kind"))
    location = entry.get("location") or "workbook"
    title = entry.get("message") if category == "finding" else identifier
    metadata = [f"<span><strong>Location:</strong> {_html_text(location)}</span>"]
    if category == "change":
        metadata.append(
            "<span><strong>Downstream formulas:</strong> "
            f"{_html_text(entry.get('impact_count', 0))}</span>"
        )
    _append_html_line(
        rendered,
        (
            f'<article class="review-entry severity-{_html_text(severity)}" '
            f'data-severity="{_html_text(severity)}">'
        ),
    )
    _append_html_line(
        rendered,
        "<h3>"
        f"<code>{_html_text(identifier)}</code> "
        f'<span class="severity">{_html_text(severity)}</span>'
        "</h3>",
    )
    _append_html_line(rendered, f"<p>{_html_text(title)}</p>")
    _append_html_line(rendered, f'<div class="metadata">{"".join(metadata)}</div>')
    _append_html_line(rendered, "<details>")
    _append_html_line(rendered, "<summary>Review evidence</summary>")
    rendered.append("<pre>")
    _append_html_json(rendered, entry)
    _append_html_line(rendered, "</pre>")
    _append_html_line(rendered, "</details>")
    _append_html_line(rendered, "</article>")


def _append_html_review_section(
    rendered: _BoundedText,
    heading: str,
    entries: list[dict[str, Any]],
    category: str,
) -> None:
    """Write one review section without retaining all rendered entries."""
    _append_html_line(rendered, '<section class="review-section">')
    _append_html_line(rendered, f"<h2>{_html_text(heading)}</h2>")
    if entries:
        for entry in entries:
            _append_html_review_entry(rendered, entry, category)
    else:
        _append_html_line(rendered, '<p class="empty">No entries.</p>')
    _append_html_line(rendered, "</section>")


def _append_bounded_html_document_prefix(
    rendered: _BoundedText,
    title: str,
    lede: str,
    cards: Iterable[tuple[str, object]],
    context: Iterable[str],
    redactions: Iterable[str],
) -> None:
    """Write the static and small summary part of a portable HTML report."""
    _append_html_line(rendered, "<!doctype html>")
    _append_html_line(rendered, '<html lang="en">')
    _append_html_line(rendered, "<head>")
    _append_html_line(rendered, '<meta charset="utf-8">')
    _append_html_line(
        rendered,
        '<meta name="viewport" content="width=device-width, initial-scale=1">',
    )
    _append_html_line(rendered, f"<title>{_html_text(title)}</title>")
    _append_html_line(rendered, f"<style>\n{_HTML_STYLE}\n</style>")
    _append_html_line(rendered, "</head>")
    _append_html_line(rendered, "<body>")
    _append_html_line(rendered, "<main>")
    _append_html_line(rendered, "<header>")
    _append_html_line(rendered, f"<h1>{_html_text(title)}</h1>")
    _append_html_line(rendered, f'<p class="lede">{_html_text(lede)}</p>')
    for item in context:
        _append_html_line(rendered, f"<p>{_html_text(item)}</p>")
    _append_html_line(rendered, '<dl class="cards">')
    for label, value in cards:
        _append_html_line(rendered, '<div class="card">')
        _append_html_line(rendered, f"<dt>{_html_text(label)}</dt>")
        _append_html_line(rendered, f"<dd>{_html_text(value)}</dd>")
        _append_html_line(rendered, "</div>")
    _append_html_line(rendered, "</dl>")
    redaction_labels = tuple(redactions)
    if redaction_labels:
        _append_html_line(rendered, '<aside class="redactions">')
        _append_html_line(rendered, "<strong>Sharing redactions enabled</strong>")
        _append_html_line(rendered, "<ul>")
        for label in redaction_labels:
            _append_html_line(rendered, f"<li>{_html_text(label)}</li>")
        _append_html_line(rendered, "</ul>")
        _append_html_line(rendered, "</aside>")
    _append_html_line(rendered, "</header>")
    _append_html_line(rendered, '<section class="filters" aria-label="Review filters">')
    _append_html_line(
        rendered,
        '<label>Search <input id="review-filter" type="search" '
        'placeholder="Rule, location, formula, or text"></label>',
    )
    _append_html_line(rendered, '<label>Severity <select id="severity-filter">')
    _append_html_line(rendered, '<option value="all">All severities</option>')
    _append_html_line(rendered, '<option value="critical">Critical</option>')
    _append_html_line(rendered, '<option value="high">High</option>')
    _append_html_line(rendered, '<option value="medium">Medium</option>')
    _append_html_line(rendered, '<option value="low">Low</option>')
    _append_html_line(rendered, '<option value="note">Note</option>')
    _append_html_line(rendered, "</select></label>")
    _append_html_line(rendered, '<output id="filter-status" aria-live="polite"></output>')
    _append_html_line(rendered, "</section>")


def _append_bounded_html_document_suffix(rendered: _BoundedText) -> None:
    """Complete a bounded portable HTML report."""
    _append_html_line(rendered, "</main>")
    _append_html_line(rendered, f"<script>\n{_HTML_FILTER_SCRIPT}\n</script>")
    _append_html_line(rendered, "</body>")
    _append_html_line(rendered, "</html>")


def _html_document(
    title: str,
    lede: str,
    cards: Iterable[tuple[str, object]],
    context: Iterable[str],
    redactions: Iterable[str],
    content: str,
) -> str:
    """Build a portable report with no external assets or network requests."""
    card_markup = "\n".join(
        "\n".join(
            [
                '<div class="card">',
                f"<dt>{_html_text(label)}</dt>",
                f"<dd>{_html_text(value)}</dd>",
                "</div>",
            ]
        )
        for label, value in cards
    )
    context_markup = "\n".join(f"<p>{_html_text(item)}</p>" for item in context)
    redaction_items = "\n".join(
        f"<li>{_html_text(label)}</li>" for label in redactions
    )
    redaction_markup = (
        "\n".join(
            [
                '<aside class="redactions">',
                "<strong>Sharing boundaries are active</strong>",
                f"<ul>{redaction_items}</ul>",
                "</aside>",
            ]
        )
        if redaction_items
        else ""
    )
    return "\n".join(
        [
            "<!doctype html>",
            '<html lang="en">',
            "<head>",
            '<meta charset="utf-8">',
            '<meta name="viewport" content="width=device-width, initial-scale=1">',
            f"<title>{_html_text(title)}</title>",
            f"<style>\n{_HTML_STYLE}\n</style>",
            "</head>",
            "<body>",
            "<main>",
            "<header>",
            f"<h1>{_html_text(title)}</h1>",
            f'<p class="lede">{_html_text(lede)}</p>',
            context_markup,
            f'<dl class="cards">{card_markup}</dl>',
            redaction_markup,
            "</header>",
            '<section class="filters" aria-label="Review filters">',
            '<label>Search <input id="review-filter" type="search" '
            'placeholder="Rule, location, formula, or text"></label>',
            '<label>Severity <select id="severity-filter">',
            '<option value="all">All severities</option>',
            '<option value="critical">Critical</option>',
            '<option value="high">High</option>',
            '<option value="medium">Medium</option>',
            '<option value="low">Low</option>',
            '<option value="note">Note</option>',
            "</select></label>",
            '<output id="filter-status" aria-live="polite"></output>',
            "</section>",
            content,
            "</main>",
            f"<script>\n{_HTML_FILTER_SCRIPT}\n</script>",
            "</body>",
            "</html>",
            "",
        ]
    )


def _render_bounded_html_document(
    title: str,
    lede: str,
    cards: Iterable[tuple[str, object]],
    context: Iterable[str],
    redactions: Iterable[str],
    *,
    max_bytes: int,
    append_content: Callable[[_BoundedText], None],
) -> str:
    """Render a self-contained HTML review without retaining an over-limit page."""
    rendered = _BoundedText(max_bytes)
    _append_bounded_html_document_prefix(
        rendered,
        title,
        lede,
        cards,
        context,
        redactions,
    )
    append_content(rendered)
    _append_bounded_html_document_suffix(rendered)
    return rendered.render()


def _append_report_markdown_sections(
    lines: _BoundedLines, payload: dict[str, Any], heading: str
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


def report_to_markdown(
    report: DiffReport,
    extra_findings: Iterable[Finding] = (),
    *,
    max_bytes: int | None = None,
    redact_external_workbook_links: bool = False,
    redact_formula_external_actions: bool = False,
    redact_python_in_excel: bool = False,
    redact_office_custom_functions: bool = False,
    redact_unqualified_runtime_functions: bool = False,
    redact_worksheet_code_resource_registrations: bool = False,
    redact_formula_defined_xlm_registrations: bool = False,
    redact_formula_defined_xlm_evaluations: bool = False,
    redact_formula_defined_xlm_actions: bool = False,
    redact_formula_defined_xlm_get_cell_calls: bool = False,
    redact_formula_defined_xlm_environment_information_calls: bool = False,
    redact_formula_environment_information: bool = False,
) -> str:
    payload = _report_payload_for_rendering(
        report,
        extra_findings,
        redact_external_workbook_links=redact_external_workbook_links,
        redact_formula_external_actions=redact_formula_external_actions,
        redact_python_in_excel=redact_python_in_excel,
        redact_office_custom_functions=redact_office_custom_functions,
        redact_unqualified_runtime_functions=redact_unqualified_runtime_functions,
        redact_worksheet_code_resource_registrations=(
            redact_worksheet_code_resource_registrations
        ),
        redact_formula_defined_xlm_registrations=(
            redact_formula_defined_xlm_registrations
        ),
        redact_formula_defined_xlm_evaluations=redact_formula_defined_xlm_evaluations,
        redact_formula_defined_xlm_actions=redact_formula_defined_xlm_actions,
        redact_formula_defined_xlm_get_cell_calls=(
            redact_formula_defined_xlm_get_cell_calls
        ),
        redact_formula_defined_xlm_environment_information_calls=(
            redact_formula_defined_xlm_environment_information_calls
        ),
        redact_formula_environment_information=redact_formula_environment_information,
    )
    summary = payload["summary"]
    lines = _BoundedLines(
        [
        "# FormulaFence change report",
        "",
        f"- **Baseline:** `{payload['before']['path']}`",
        f"- **Candidate:** `{payload['after']['path']}`",
        *(
            ["- **External-workbook link material:** redacted for sharing"]
            if redact_external_workbook_links
            else []
        ),
        *(
            ["- **Formula external-action / DDE material:** redacted for sharing"]
            if redact_formula_external_actions
            else []
        ),
        *(
            ["- **Python-in-Excel material:** redacted for sharing"]
            if redact_python_in_excel
            else []
        ),
        *(
            ["- **Office custom-function material:** redacted for sharing"]
            if redact_office_custom_functions
            else []
        ),
        *(
            ["- **Unqualified runtime-function material:** redacted for sharing"]
            if redact_unqualified_runtime_functions
            else []
        ),
        *(
            [
                "- **Worksheet code-resource registration material:** redacted for "
                "sharing"
            ]
            if redact_worksheet_code_resource_registrations
            else []
        ),
        *(
            [
                "- **Formula-defined XLM registration material:** redacted for "
                "sharing"
            ]
            if redact_formula_defined_xlm_registrations
            else []
        ),
        *(
            [
                "- **Formula-defined XLM evaluation material:** redacted for "
                "sharing"
            ]
            if redact_formula_defined_xlm_evaluations
            else []
        ),
        *(
            ["- **Formula-defined XLM action material:** redacted for sharing"]
            if redact_formula_defined_xlm_actions
            else []
        ),
        *(
            ["- **Formula-defined XLM GET.CELL material:** redacted for sharing"]
            if redact_formula_defined_xlm_get_cell_calls
            else []
        ),
        *(
            [
                "- **Formula-defined XLM environment-information material:** "
                "redacted for sharing"
            ]
            if redact_formula_defined_xlm_environment_information_calls
            else []
        ),
        *(
            ["- **Formula environment-information material:** redacted for sharing"]
            if redact_formula_environment_information
            else []
        ),
        f"- **Changes:** {summary['change_count']}",
        f"- **Findings:** {summary['finding_count']}",
        f"- **Highest severity:** `{summary['highest_severity']}`",
        "",
        ],
        max_bytes=max_bytes,
    )
    _append_report_markdown_sections(lines, payload, "##")
    return lines.render()


def report_to_html(
    report: DiffReport,
    extra_findings: Iterable[Finding] = (),
    *,
    max_bytes: int | None = None,
    redact_external_workbook_links: bool = False,
    redact_formula_external_actions: bool = False,
    redact_python_in_excel: bool = False,
    redact_office_custom_functions: bool = False,
    redact_unqualified_runtime_functions: bool = False,
    redact_worksheet_code_resource_registrations: bool = False,
    redact_formula_defined_xlm_registrations: bool = False,
    redact_formula_defined_xlm_evaluations: bool = False,
    redact_formula_defined_xlm_actions: bool = False,
    redact_formula_defined_xlm_get_cell_calls: bool = False,
    redact_formula_defined_xlm_environment_information_calls: bool = False,
    redact_formula_environment_information: bool = False,
) -> str:
    """Render an escaped, filterable, self-contained HTML review artifact."""
    payload = _report_payload_for_rendering(
        report,
        extra_findings,
        redact_external_workbook_links=redact_external_workbook_links,
        redact_formula_external_actions=redact_formula_external_actions,
        redact_python_in_excel=redact_python_in_excel,
        redact_office_custom_functions=redact_office_custom_functions,
        redact_unqualified_runtime_functions=redact_unqualified_runtime_functions,
        redact_worksheet_code_resource_registrations=(
            redact_worksheet_code_resource_registrations
        ),
        redact_formula_defined_xlm_registrations=(
            redact_formula_defined_xlm_registrations
        ),
        redact_formula_defined_xlm_evaluations=redact_formula_defined_xlm_evaluations,
        redact_formula_defined_xlm_actions=redact_formula_defined_xlm_actions,
        redact_formula_defined_xlm_get_cell_calls=(
            redact_formula_defined_xlm_get_cell_calls
        ),
        redact_formula_defined_xlm_environment_information_calls=(
            redact_formula_defined_xlm_environment_information_calls
        ),
        redact_formula_environment_information=redact_formula_environment_information,
    )
    summary = payload["summary"]
    redactions = _active_redaction_labels(
        redact_external_workbook_links=redact_external_workbook_links,
        redact_formula_external_actions=redact_formula_external_actions,
        redact_python_in_excel=redact_python_in_excel,
        redact_office_custom_functions=redact_office_custom_functions,
        redact_unqualified_runtime_functions=redact_unqualified_runtime_functions,
        redact_worksheet_code_resource_registrations=(
            redact_worksheet_code_resource_registrations
        ),
        redact_formula_defined_xlm_registrations=(
            redact_formula_defined_xlm_registrations
        ),
        redact_formula_defined_xlm_evaluations=redact_formula_defined_xlm_evaluations,
        redact_formula_defined_xlm_actions=redact_formula_defined_xlm_actions,
        redact_formula_defined_xlm_get_cell_calls=(
            redact_formula_defined_xlm_get_cell_calls
        ),
        redact_formula_defined_xlm_environment_information_calls=(
            redact_formula_defined_xlm_environment_information_calls
        ),
        redact_formula_environment_information=redact_formula_environment_information,
    )
    title = "FormulaFence change report"
    lede = (
        "A portable local review artifact. Use the filters to narrow findings and "
        "changes; expand an entry for its complete escaped review evidence."
    )
    cards = (
        ("Changes", summary["change_count"]),
        ("Findings", summary["finding_count"]),
        ("Highest severity", summary["highest_severity"]),
    )
    context = (
        f"Baseline: {payload['before']['path']}",
        f"Candidate: {payload['after']['path']}",
    )
    if max_bytes is None:
        content = "\n".join(
            [
                _html_review_section("Findings", payload["findings"], "finding"),
                _html_review_section("Semantic changes", payload["changes"], "change"),
            ]
        )
        return _html_document(title, lede, cards, context, redactions, content)

    def append_content(rendered: _BoundedText) -> None:
        _append_html_review_section(rendered, "Findings", payload["findings"], "finding")
        _append_html_review_section(
            rendered,
            "Semantic changes",
            payload["changes"],
            "change",
        )

    return _render_bounded_html_document(
        title,
        lede,
        cards,
        context,
        redactions,
        max_bytes=max_bytes,
        append_content=append_content,
    )


_SARIF_LEVELS = {
    "critical": "error",
    "high": "error",
    "medium": "warning",
    "low": "warning",
    "note": "note",
}


def report_to_sarif(
    report: DiffReport,
    extra_findings: Iterable[Finding] = (),
    *,
    redact_external_workbook_links: bool = False,
    redact_formula_external_actions: bool = False,
    redact_python_in_excel: bool = False,
    redact_office_custom_functions: bool = False,
    redact_unqualified_runtime_functions: bool = False,
    redact_worksheet_code_resource_registrations: bool = False,
    redact_formula_defined_xlm_registrations: bool = False,
    redact_formula_defined_xlm_evaluations: bool = False,
    redact_formula_defined_xlm_actions: bool = False,
    redact_formula_defined_xlm_get_cell_calls: bool = False,
    redact_formula_defined_xlm_environment_information_calls: bool = False,
    redact_formula_environment_information: bool = False,
) -> dict[str, Any]:
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
        finding_details = _safe_finding_details(
            finding,
            report,
            redact_formula_external_actions=redact_formula_external_actions,
            redact_office_custom_functions=redact_office_custom_functions,
            redact_unqualified_runtime_functions=redact_unqualified_runtime_functions,
            redact_worksheet_code_resource_registrations=(
                redact_worksheet_code_resource_registrations
            ),
            redact_formula_defined_xlm_registrations=(
                redact_formula_defined_xlm_registrations
            ),
            redact_formula_defined_xlm_evaluations=(
                redact_formula_defined_xlm_evaluations
            ),
            redact_formula_defined_xlm_actions=redact_formula_defined_xlm_actions,
            redact_formula_defined_xlm_get_cell_calls=(
                redact_formula_defined_xlm_get_cell_calls
            ),
            redact_formula_defined_xlm_environment_information_calls=(
                redact_formula_defined_xlm_environment_information_calls
            ),
            redact_formula_environment_information=redact_formula_environment_information,
        )
        result: dict[str, Any] = {
            "ruleId": finding.rule_id,
            "level": _SARIF_LEVELS.get(finding.severity, "warning"),
            "message": {"text": finding.message},
            "properties": {"severity": finding.severity, **finding_details},
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
    payload = {
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
    if redact_external_workbook_links:
        payload = redact_external_workbook_link_material(payload)
    if redact_formula_external_actions:
        payload = redact_formula_external_action_material(payload)
    if redact_python_in_excel:
        payload = redact_python_in_excel_material(payload)
    if redact_office_custom_functions:
        payload = redact_office_custom_function_material(payload)
    if redact_unqualified_runtime_functions:
        payload = redact_unqualified_runtime_function_material(payload)
    if redact_worksheet_code_resource_registrations:
        payload = redact_worksheet_code_resource_registration_material(payload)
    if redact_formula_defined_xlm_registrations:
        payload = redact_formula_defined_xlm_registration_material(payload)
    if redact_formula_defined_xlm_evaluations:
        payload = redact_formula_defined_xlm_evaluation_material(payload)
    if redact_formula_defined_xlm_actions:
        payload = redact_formula_defined_xlm_action_material(payload)
    if redact_formula_defined_xlm_get_cell_calls:
        payload = redact_formula_defined_xlm_get_cell_material(payload)
    if redact_formula_defined_xlm_environment_information_calls:
        payload = redact_formula_defined_xlm_environment_information_material(payload)
    if redact_formula_environment_information:
        payload = redact_formula_environment_information_material(payload)
    return payload


def lint_to_markdown(
    report: FormulaLintReport,
    *,
    max_bytes: int | None = None,
) -> str:
    """Render a bounded, formula-free local formula-lint review."""
    payload = report.to_dict()
    workbook = payload["workbook"]
    summary = payload["summary"]
    lines = _BoundedLines(
        [
            "# FormulaFence formula lint",
            "",
            f"- **Workbook:** {_markdown_code(workbook['path'])}",
            f"- **Formula cells:** {workbook['formula_cells']}",
            f"- **Findings:** {summary['finding_count']}",
            f"- **Highest severity:** `{summary['highest_severity']}`",
            "",
            "The lint reports an interrupted copy pattern only with two matching immediate "
            "peers and a third contiguous supporting peer. It also reports a pure local "
            "numeric aggregate only when it stops before a short contiguous numeric run. "
            "It reports a formula as unlocked only for an explicit direct cell assignment "
            "on a protected worksheet, and reports explicitly incomplete manual calculation "
            "for a formula workbook. It also reports stored error-checking suppressions "
            "and isolated interior Excel Table calculated-column exceptions. It does not "
            "evaluate formulas or expose formula text.",
            "",
            "## Findings",
            "",
            "| Severity | Rule | Location | Finding |",
            "| --- | --- | --- | --- |",
        ],
        max_bytes=max_bytes,
    )
    if not payload["findings"]:
        lines.append("| note | — | — | No conservative formula-lint findings found. |")
    else:
        for finding in payload["findings"]:
            lines.append(
                "| {severity} | `{rule_id}` | {location} | {message} |".format(
                    severity=_markdown_escape(finding["severity"]),
                    rule_id=_markdown_escape(finding["rule_id"]),
                    location=_markdown_code(finding["location"] or "workbook"),
                    message=_markdown_escape(finding["message"]),
                )
            )
    lines.append("")

    evidence_lines = [
        (finding["location"], evidence)
        for finding in payload["findings"]
        for evidence in finding.get("details", {}).get("pattern_evidence", [])
    ]
    if evidence_lines:
        lines.extend(["## Pattern evidence", ""])
        for location, evidence in evidence_lines:
            lines.append(
                "- {location}: matching {orientation} formulas {preceding} and "
                "{following}, supported by {supporting}.".format(
                    location=_markdown_code(location or "workbook"),
                    orientation=_markdown_escape(evidence["orientation"]),
                    preceding=_markdown_code(evidence["preceding_formula"]),
                    following=_markdown_code(evidence["following_formula"]),
                    supporting=_markdown_code(evidence["supporting_formula"]),
                )
            )
    aggregate_evidence = [
        (finding["location"], finding["details"])
        for finding in payload["findings"]
        if finding["rule_id"] == "FF084"
    ]
    if aggregate_evidence:
        lines.extend(["## Aggregate range evidence", ""])
        for location, evidence in aggregate_evidence:
            lines.append(
                "- {location}: `{function}` references {referenced_range} and stops before "
                "adjacent numeric cells {omitted_range} ({count}).".format(
                    location=_markdown_code(location or "workbook"),
                    function=_markdown_escape(evidence["aggregate_function"]),
                    referenced_range=_markdown_code(evidence["referenced_range"]),
                    omitted_range=_markdown_code(evidence["omitted_range"]),
                    count=_markdown_escape(evidence["omitted_cell_count"]),
                )
            )
    protection_evidence = [
        finding["location"]
        for finding in payload["findings"]
        if finding["rule_id"] == "FF085"
    ]
    if protection_evidence:
        lines.extend(["## Formula protection evidence", ""])
        for location in protection_evidence:
            lines.append(
                "- {location}: direct cell protection marks this formula cell as unlocked."
                .format(location=_markdown_code(location or "workbook"))
            )
    calculation_evidence = [
        finding["details"]
        for finding in payload["findings"]
        if finding["rule_id"] == "FF086"
    ]
    if calculation_evidence:
        lines.extend(["## Calculation freshness evidence", ""])
        for evidence in calculation_evidence:
            lines.append(
                "- Workbook: calculation mode is `{mode}` and the file records no "
                "completed calculation before save.".format(
                    mode=_markdown_escape(evidence["calculation_mode"])
                )
            )
    error_checking_suppression_evidence = [
        finding["details"]
        for finding in payload["findings"]
        if finding["rule_id"] == "FF091"
    ]
    if error_checking_suppression_evidence:
        lines.extend(["## Excel error-checking suppression evidence", ""])
        for evidence in error_checking_suppression_evidence:
            categories = ", ".join(
                "{name} ({count})".format(
                    name=_markdown_code(name.replace("_", " ")),
                    count=_markdown_escape(count),
                )
                for name, count in evidence["suppressed_warning_counts"].items()
            )
            lines.append(
                "- Workbook: {rules} suppressed warning rules across {targets} target "
                "ranges: {categories}.".format(
                    rules=_markdown_escape(evidence["suppressed_warning_rule_count"]),
                    targets=_markdown_escape(
                        evidence["suppressed_warning_target_range_count"]
                    ),
                    categories=categories,
                )
            )
    table_calculated_column_evidence = [
        (finding["location"], finding["details"])
        for finding in payload["findings"]
        if finding["rule_id"] == "FF092"
    ]
    if table_calculated_column_evidence:
        lines.extend(["## Excel Table calculated-column evidence", ""])
        exception_labels = {
            "blank": "blank",
            "formula_mismatch": "a different formula",
            "stored_error_value": "a stored error value",
            "text_value": "a text value",
            "non_formula_value": "a non-formula value",
        }
        for location, evidence in table_calculated_column_evidence:
            lines.append(
                "- {location}: this interior Table cell is {exception} while its "
                "immediate adjacent cells match the declared calculated-column formula."
                .format(
                    location=_markdown_code(location or "workbook"),
                    exception=_markdown_escape(
                        exception_labels.get(
                            str(evidence.get("exception_kind", "")),
                            "an exception",
                        )
                    ),
                )
            )
    conditional_aggregate_range_shape_evidence = [
        (finding["location"], finding["details"])
        for finding in payload["findings"]
        if finding["rule_id"] == "FF093"
    ]
    if conditional_aggregate_range_shape_evidence:
        lines.extend(["## Conditional-aggregate range-shape evidence", ""])
        for location, evidence in conditional_aggregate_range_shape_evidence:
            lines.append(
                "- {location}: {range_count} direct static range arguments across "
                "{call_count} conditional-aggregate calls have different shapes."
                .format(
                    location=_markdown_code(location or "workbook"),
                    range_count=_markdown_escape(
                        evidence["mismatched_direct_range_argument_count"]
                    ),
                    call_count=_markdown_escape(
                        evidence["conditional_aggregate_call_count"]
                    ),
                )
            )
    sumproduct_range_shape_evidence = [
        (finding["location"], finding["details"])
        for finding in payload["findings"]
        if finding["rule_id"] == "FF094"
    ]
    if sumproduct_range_shape_evidence:
        lines.extend(["## SUMPRODUCT range-shape evidence", ""])
        for location, evidence in sumproduct_range_shape_evidence:
            lines.append(
                "- {location}: {array_count} direct static array arguments across "
                "{call_count} SUMPRODUCT calls have different shapes."
                .format(
                    location=_markdown_code(location or "workbook"),
                    array_count=_markdown_escape(
                        evidence["mismatched_direct_array_argument_count"]
                    ),
                    call_count=_markdown_escape(evidence["sumproduct_call_count"]),
                )
            )
    mmult_dimension_evidence = [
        (finding["location"], finding["details"])
        for finding in payload["findings"]
        if finding["rule_id"] == "FF095"
    ]
    if mmult_dimension_evidence:
        lines.extend(["## MMULT matrix-dimension evidence", ""])
        for location, evidence in mmult_dimension_evidence:
            lines.append(
                "- {location}: {pair_count} direct static matrix pairs across "
                "{call_count} MMULT calls have incompatible inner dimensions."
                .format(
                    location=_markdown_code(location or "workbook"),
                    pair_count=_markdown_escape(
                        evidence["incompatible_direct_matrix_pair_count"]
                    ),
                    call_count=_markdown_escape(evidence["mmult_call_count"]),
                )
            )
    lookup_return_index_evidence = [
        (finding["location"], finding["details"])
        for finding in payload["findings"]
        if finding["rule_id"] == "FF096"
    ]
    if lookup_return_index_evidence:
        lines.extend(["## Lookup return-index evidence", ""])
        for location, evidence in lookup_return_index_evidence:
            lines.append(
                "- {location}: {index_count} direct static lookup calls across "
                "{call_count} calls use out-of-range literal return indices."
                .format(
                    location=_markdown_code(location or "workbook"),
                    index_count=_markdown_escape(
                        evidence["out_of_range_literal_index_count"]
                    ),
                    call_count=_markdown_escape(evidence["lookup_call_count"]),
                )
            )
    choose_literal_index_evidence = [
        (finding["location"], finding["details"])
        for finding in payload["findings"]
        if finding["rule_id"] == "FF097"
    ]
    if choose_literal_index_evidence:
        lines.extend(["## CHOOSE literal-index evidence", ""])
        for location, evidence in choose_literal_index_evidence:
            lines.append(
                "- {location}: {index_count} direct literal-index CHOOSE calls across "
                "{call_count} calls use out-of-range value-argument indices."
                .format(
                    location=_markdown_code(location or "workbook"),
                    index_count=_markdown_escape(
                        evidence["out_of_range_literal_index_count"]
                    ),
                    call_count=_markdown_escape(evidence["choose_call_count"]),
                )
            )
    randbetween_literal_bound_evidence = [
        (finding["location"], finding["details"])
        for finding in payload["findings"]
        if finding["rule_id"] == "FF098"
    ]
    if randbetween_literal_bound_evidence:
        lines.extend(["## RANDBETWEEN literal-bound evidence", ""])
        for location, evidence in randbetween_literal_bound_evidence:
            lines.append(
                "- {location}: {bound_count} direct literal-bound RANDBETWEEN calls "
                "across {call_count} calls have a bottom above the top."
                .format(
                    location=_markdown_code(location or "workbook"),
                    bound_count=_markdown_escape(
                        evidence["inverted_literal_bound_count"]
                    ),
                    call_count=_markdown_escape(evidence["randbetween_call_count"]),
                )
            )
    subtotal_literal_function_num_evidence = [
        (finding["location"], finding["details"])
        for finding in payload["findings"]
        if finding["rule_id"] == "FF099"
    ]
    if subtotal_literal_function_num_evidence:
        lines.extend(["## SUBTOTAL function-code evidence", ""])
        for location, evidence in subtotal_literal_function_num_evidence:
            lines.append(
                "- {location}: {code_count} direct literal SUBTOTAL calls across "
                "{call_count} calls use unsupported function codes."
                .format(
                    location=_markdown_code(location or "workbook"),
                    code_count=_markdown_escape(
                        evidence["unsupported_literal_function_num_count"]
                    ),
                    call_count=_markdown_escape(evidence["subtotal_call_count"]),
                )
            )
    index_literal_position_evidence = [
        (finding["location"], finding["details"])
        for finding in payload["findings"]
        if finding["rule_id"] == "FF100"
    ]
    if index_literal_position_evidence:
        lines.extend(["## INDEX literal-position evidence", ""])
        for location, evidence in index_literal_position_evidence:
            lines.append(
                "- {location}: {index_count} direct literal INDEX calls across "
                "{call_count} calls use row or column positions outside direct "
                "static arrays."
                .format(
                    location=_markdown_code(location or "workbook"),
                    index_count=_markdown_escape(
                        evidence["out_of_range_literal_index_count"]
                    ),
                    call_count=_markdown_escape(evidence["index_call_count"]),
                )
            )
    approximate_lookup_sort_evidence = [
        (finding["location"], finding["details"])
        for finding in payload["findings"]
        if finding["rule_id"] == "FF101"
    ]
    if approximate_lookup_sort_evidence:
        lines.extend(["## Approximate lookup sort evidence", ""])
        for location, evidence in approximate_lookup_sort_evidence:
            lines.append(
                "- {location}: {vector_count} direct static numeric lookup vectors "
                "across {call_count} approximate VLOOKUP or HLOOKUP calls are not "
                "sorted ascending."
                .format(
                    location=_markdown_code(location or "workbook"),
                    vector_count=_markdown_escape(
                        evidence["unsorted_direct_numeric_lookup_vector_count"]
                    ),
                    call_count=_markdown_escape(
                        evidence["approximate_lookup_call_count"]
                    ),
                )
            )
    modern_lookup_mode_evidence = [
        (finding["location"], finding["details"])
        for finding in payload["findings"]
        if finding["rule_id"] == "FF102"
    ]
    if modern_lookup_mode_evidence:
        lines.extend(["## XLOOKUP/XMATCH mode-code evidence", ""])
        for location, evidence in modern_lookup_mode_evidence:
            lines.append(
                "- {location}: {mode_count} direct literal XLOOKUP or XMATCH modes "
                "use unsupported codes."
                .format(
                    location=_markdown_code(location or "workbook"),
                    mode_count=_markdown_escape(
                        evidence["unsupported_literal_mode_count"]
                    ),
                )
            )
    large_small_rank_evidence = [
        (finding["location"], finding["details"])
        for finding in payload["findings"]
        if finding["rule_id"] == "FF103"
    ]
    if large_small_rank_evidence:
        lines.extend(["## LARGE/SMALL literal-rank evidence", ""])
        for location, evidence in large_small_rank_evidence:
            lines.append(
                "- {location}: {rank_count} direct literal LARGE or SMALL ranks are "
                "nonpositive or exceed direct static array capacity."
                .format(
                    location=_markdown_code(location or "workbook"),
                    rank_count=_markdown_escape(
                        evidence["invalid_literal_rank_count"]
                    ),
                )
            )
    text_literal_argument_evidence = [
        (finding["location"], finding["details"])
        for finding in payload["findings"]
        if finding["rule_id"] == "FF104"
    ]
    if text_literal_argument_evidence:
        lines.extend(["## Text literal-argument evidence", ""])
        for location, evidence in text_literal_argument_evidence:
            lines.append(
                "- {location}: {argument_count} direct literal text-function "
                "positions or counts are invalid."
                .format(
                    location=_markdown_code(location or "workbook"),
                    argument_count=_markdown_escape(
                        evidence["invalid_literal_argument_count"]
                    ),
                )
            )
    direct_zero_divisor_evidence = [
        (finding["location"], finding["details"])
        for finding in payload["findings"]
        if finding["rule_id"] == "FF105"
    ]
    if direct_zero_divisor_evidence:
        lines.extend(["## Direct zero-divisor evidence", ""])
        for location, evidence in direct_zero_divisor_evidence:
            lines.append(
                "- {location}: {divisor_count} division expressions use a direct "
                "literal zero divisor."
                .format(
                    location=_markdown_code(location or "workbook"),
                    divisor_count=_markdown_escape(
                        evidence["direct_zero_divisor_count"]
                    ),
                )
            )
    direct_sum_overlap_evidence = [
        (finding["location"], finding["details"])
        for finding in payload["findings"]
        if finding["rule_id"] == "FF110"
    ]
    if direct_sum_overlap_evidence:
        lines.extend(["## Direct SUM overlap evidence", ""])
        for location, evidence in direct_sum_overlap_evidence:
            pair_count = evidence["overlapping_direct_range_pair_count"]
            call_count = evidence["direct_sum_call_count"]
            lines.append(
                "- {location}: {pair_count} overlapping direct static range {pair_noun} "
                "across {call_count} SUM {call_noun} {verb} that at least one cell is "
                "included more than once."
                .format(
                    location=_markdown_code(location or "workbook"),
                    pair_count=_markdown_escape(pair_count),
                    pair_noun="pair" if pair_count == 1 else "pairs",
                    call_count=_markdown_escape(call_count),
                    call_noun="call" if call_count == 1 else "calls",
                    verb="shows" if pair_count == 1 else "show",
                )
            )
    aggregate_literal_argument_evidence = [
        (finding["location"], finding["details"])
        for finding in payload["findings"]
        if finding["rule_id"] == "FF111"
    ]
    if aggregate_literal_argument_evidence:
        lines.extend(["## AGGREGATE literal-argument evidence", ""])
        for location, evidence in aggregate_literal_argument_evidence:
            call_count = evidence["aggregate_call_count"]
            function_num_count = evidence["unsupported_literal_function_num_count"]
            option_count = evidence["unsupported_literal_option_count"]
            missing_ref2_count = evidence["missing_required_ref2_count"]
            lines.append(
                "- {location}: {call_count} AGGREGATE {call_noun} have unsupported "
                "direct literal function numbers or options, or omit required second "
                "references ({function_num_count} function-number, {option_count} "
                "option, {missing_ref2_count} ref2 {error_noun})."
                .format(
                    location=_markdown_code(location or "workbook"),
                    call_count=_markdown_escape(call_count),
                    call_noun="call" if call_count == 1 else "calls",
                    function_num_count=_markdown_escape(function_num_count),
                    option_count=_markdown_escape(option_count),
                    missing_ref2_count=_markdown_escape(missing_ref2_count),
                    error_noun="error" if missing_ref2_count == 1 else "errors",
                )
            )
    circular_reference_evidence = [
        (finding["rule_id"], finding["location"], finding["details"])
        for finding in payload["findings"]
        if finding["rule_id"] in {"FF087", "FF090"}
    ]
    if circular_reference_evidence:
        lines.extend(["## Static circular-reference evidence", ""])
        for rule_id, location, evidence in circular_reference_evidence:
            if rule_id == "FF087":
                lines.append(
                    "- {location}: a resolved direct static formula dependency returns to "
                    "the same cell while calculation iteration is disabled."
                    .format(location=_markdown_code(location or "workbook"))
                )
            else:
                lines.append(
                    "- {location}: a static multi-cell dependency component has {count} "
                    "formula cells while calculation iteration is disabled."
                    .format(
                        location=_markdown_code(location or "workbook"),
                        count=_markdown_escape(evidence["cycle_member_count"]),
                    )
                )
    broken_reference_evidence = [
        finding["location"]
        for finding in payload["findings"]
        if finding["rule_id"] == "FF088"
    ]
    if broken_reference_evidence:
        lines.extend(["## Explicit broken-reference evidence", ""])
        for location in broken_reference_evidence:
            lines.append(
                "- {location}: formula tokenization found an explicit `#REF!` error "
                "operand.".format(location=_markdown_code(location or "workbook"))
            )
    saved_broken_reference_evidence = [
        finding["location"]
        for finding in payload["findings"]
        if finding["rule_id"] == "FF089"
    ]
    if saved_broken_reference_evidence:
        lines.extend(["## Saved broken-reference evidence", ""])
        for location in saved_broken_reference_evidence:
            lines.append(
                "- {location}: the stored formula-result cache records a broken-reference "
                "error.".format(location=_markdown_code(location or "workbook"))
            )
    saved_divide_by_zero_evidence = [
        finding["location"]
        for finding in payload["findings"]
        if finding["rule_id"] == "FF106"
    ]
    if saved_divide_by_zero_evidence:
        lines.extend(["## Saved division-by-zero evidence", ""])
        for location in saved_divide_by_zero_evidence:
            lines.append(
                "- {location}: the stored formula-result cache records a division-by-zero "
                "error.".format(location=_markdown_code(location or "workbook"))
            )
    saved_numeric_error_evidence = [
        finding["location"]
        for finding in payload["findings"]
        if finding["rule_id"] == "FF107"
    ]
    if saved_numeric_error_evidence:
        lines.extend(["## Saved numeric-error evidence", ""])
        for location in saved_numeric_error_evidence:
            lines.append(
                "- {location}: the stored formula-result cache records a numeric "
                "error.".format(location=_markdown_code(location or "workbook"))
            )
    saved_name_error_evidence = [
        finding["location"]
        for finding in payload["findings"]
        if finding["rule_id"] == "FF108"
    ]
    if saved_name_error_evidence:
        lines.extend(["## Saved name-error evidence", ""])
        for location in saved_name_error_evidence:
            lines.append(
                "- {location}: the stored formula-result cache records a name "
                "error.".format(location=_markdown_code(location or "workbook"))
            )
    saved_value_error_evidence = [
        finding["location"]
        for finding in payload["findings"]
        if finding["rule_id"] == "FF109"
    ]
    if saved_value_error_evidence:
        lines.extend(["## Saved value-error evidence", ""])
        for location in saved_value_error_evidence:
            lines.append(
                "- {location}: the stored formula-result cache records a value "
                "error.".format(location=_markdown_code(location or "workbook"))
            )
    return lines.render()


def lint_to_sarif(report: FormulaLintReport) -> dict[str, Any]:
    """Return single-workbook formula-lint findings in SARIF 2.1.0."""
    descriptions = {
        "FF082": "A blank or non-formula cell interrupts a stable copied-formula pattern.",
        "FF083": "A formula differs from a stable copied-formula pattern.",
        "FF084": "A simple numeric aggregate stops before adjacent numeric cells.",
        "FF085": "A formula cell is explicitly unlocked on a protected worksheet.",
        "FF086": "A formula workbook was saved with incomplete manual calculation.",
        "FF087": (
            "A formula directly references its own cell while calculation iteration is disabled."
        ),
        "FF088": "Formula contains an explicit broken #REF! reference.",
        "FF089": "A formula's saved result is a broken-reference error.",
        "FF090": (
            "A formula participates in a static multi-cell circular reference while "
            "calculation iteration is disabled."
        ),
        "FF091": (
            "Workbook suppresses Excel error-checking prompts; review warnings may be "
            "hidden."
        ),
        "FF092": (
            "An interior Excel Table cell differs from its declared calculated-column "
            "formula."
        ),
        "FF093": (
            "A conditional aggregate uses direct static ranges with different shapes."
        ),
        "FF094": "A SUMPRODUCT call uses direct static ranges with different shapes.",
        "FF095": (
            "An MMULT call uses direct static arrays with incompatible matrix dimensions."
        ),
        "FF096": (
            "A VLOOKUP or HLOOKUP call uses a literal return index outside its "
            "direct static table range."
        ),
        "FF097": (
            "A CHOOSE call uses a literal index outside its available value arguments."
        ),
        "FF098": (
            "A RANDBETWEEN call uses direct literal bounds with the bottom above "
            "the top."
        ),
        "FF099": (
            "A SUBTOTAL call uses a literal function number outside Excel's "
            "supported codes."
        ),
        "FF100": (
            "An INDEX call uses a literal row or column number outside its "
            "direct static array."
        ),
        "FF101": (
            "An approximate VLOOKUP or HLOOKUP call uses a direct static numeric "
            "lookup vector that is not sorted ascending."
        ),
        "FF102": (
            "An XLOOKUP or XMATCH call uses a literal mode outside Excel's "
            "supported codes."
        ),
        "FF103": (
            "A LARGE or SMALL call uses a literal rank that is nonpositive or "
            "exceeds its direct static array capacity."
        ),
        "FF104": (
            "A LEFT, RIGHT, MID, FIND, or SEARCH call uses an invalid direct "
            "literal character position or count."
        ),
        "FF105": "A division expression uses a direct literal zero divisor.",
        "FF106": "A formula's saved result is a division-by-zero error.",
        "FF107": "A formula's saved result is a numeric error.",
        "FF108": "A formula's saved result is a name error.",
        "FF109": "A formula's saved result is a value error.",
        "FF110": (
            "A SUM call uses direct static ranges that overlap, so at least one "
            "cell is included more than once."
        ),
        "FF111": (
            "An AGGREGATE call uses an unsupported direct literal function number "
            "or option, or omits a required second reference."
        ),
    }
    rule_ids = sorted({finding.rule_id for finding in report.findings})
    rules = [
        {
            "id": rule_id,
            "name": rule_id,
            "shortDescription": {
                "text": descriptions.get(rule_id, "FormulaFence formula lint finding")
            },
        }
        for rule_id in rule_ids
    ]
    results: list[dict[str, Any]] = []
    for finding in report.findings:
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
                        "artifactLocation": {"uri": str(report.workbook.path)},
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
    lines: _BoundedLines, entry: dict[str, Any], heading: str
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


def portfolio_to_markdown(
    report: PortfolioReport,
    *,
    max_bytes: int | None = None,
    redact_external_workbook_links: bool = False,
    redact_formula_external_actions: bool = False,
    redact_python_in_excel: bool = False,
    redact_office_custom_functions: bool = False,
    redact_unqualified_runtime_functions: bool = False,
    redact_worksheet_code_resource_registrations: bool = False,
    redact_formula_defined_xlm_registrations: bool = False,
    redact_formula_defined_xlm_evaluations: bool = False,
    redact_formula_defined_xlm_actions: bool = False,
    redact_formula_defined_xlm_get_cell_calls: bool = False,
    redact_formula_defined_xlm_environment_information_calls: bool = False,
    redact_formula_environment_information: bool = False,
) -> str:
    """Render a complete multi-workbook review without absolute filesystem paths."""
    payload = _portfolio_payload_for_rendering(
        report,
        redact_external_workbook_links=redact_external_workbook_links,
        redact_formula_external_actions=redact_formula_external_actions,
        redact_python_in_excel=redact_python_in_excel,
        redact_office_custom_functions=redact_office_custom_functions,
        redact_unqualified_runtime_functions=redact_unqualified_runtime_functions,
        redact_worksheet_code_resource_registrations=(
            redact_worksheet_code_resource_registrations
        ),
        redact_formula_defined_xlm_registrations=(
            redact_formula_defined_xlm_registrations
        ),
        redact_formula_defined_xlm_evaluations=redact_formula_defined_xlm_evaluations,
        redact_formula_defined_xlm_actions=redact_formula_defined_xlm_actions,
        redact_formula_defined_xlm_get_cell_calls=(
            redact_formula_defined_xlm_get_cell_calls
        ),
        redact_formula_defined_xlm_environment_information_calls=(
            redact_formula_defined_xlm_environment_information_calls
        ),
        redact_formula_environment_information=redact_formula_environment_information,
    )
    summary = payload["summary"]
    lines = _BoundedLines(
        [
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
        *(
            ["- **External-workbook link material:** redacted for sharing"]
            if redact_external_workbook_links
            else []
        ),
        *(
            ["- **Formula external-action / DDE material:** redacted for sharing"]
            if redact_formula_external_actions
            else []
        ),
        *(
            ["- **Python-in-Excel material:** redacted for sharing"]
            if redact_python_in_excel
            else []
        ),
        *(
            ["- **Office custom-function material:** redacted for sharing"]
            if redact_office_custom_functions
            else []
        ),
        *(
            ["- **Unqualified runtime-function material:** redacted for sharing"]
            if redact_unqualified_runtime_functions
            else []
        ),
        *(
            [
                "- **Worksheet code-resource registration material:** redacted for "
                "sharing"
            ]
            if redact_worksheet_code_resource_registrations
            else []
        ),
        *(
            [
                "- **Formula-defined XLM registration material:** redacted for "
                "sharing"
            ]
            if redact_formula_defined_xlm_registrations
            else []
        ),
        *(
            [
                "- **Formula-defined XLM evaluation material:** redacted for "
                "sharing"
            ]
            if redact_formula_defined_xlm_evaluations
            else []
        ),
        *(
            ["- **Formula-defined XLM action material:** redacted for sharing"]
            if redact_formula_defined_xlm_actions
            else []
        ),
        *(
            ["- **Formula-defined XLM GET.CELL material:** redacted for sharing"]
            if redact_formula_defined_xlm_get_cell_calls
            else []
        ),
        *(
            [
                "- **Formula-defined XLM environment-information material:** "
                "redacted for sharing"
            ]
            if redact_formula_defined_xlm_environment_information_calls
            else []
        ),
        *(
            ["- **Formula environment-information material:** redacted for sharing"]
            if redact_formula_environment_information
            else []
        ),
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
        ],
        max_bytes=max_bytes,
    )
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
    return lines.render()


def portfolio_to_html(
    report: PortfolioReport,
    *,
    max_bytes: int | None = None,
    redact_external_workbook_links: bool = False,
    redact_formula_external_actions: bool = False,
    redact_python_in_excel: bool = False,
    redact_office_custom_functions: bool = False,
    redact_unqualified_runtime_functions: bool = False,
    redact_worksheet_code_resource_registrations: bool = False,
    redact_formula_defined_xlm_registrations: bool = False,
    redact_formula_defined_xlm_evaluations: bool = False,
    redact_formula_defined_xlm_actions: bool = False,
    redact_formula_defined_xlm_get_cell_calls: bool = False,
    redact_formula_defined_xlm_environment_information_calls: bool = False,
    redact_formula_environment_information: bool = False,
) -> str:
    """Render an escaped, filterable HTML review for a workbook portfolio."""
    payload = _portfolio_payload_for_rendering(
        report,
        redact_external_workbook_links=redact_external_workbook_links,
        redact_formula_external_actions=redact_formula_external_actions,
        redact_python_in_excel=redact_python_in_excel,
        redact_office_custom_functions=redact_office_custom_functions,
        redact_unqualified_runtime_functions=redact_unqualified_runtime_functions,
        redact_worksheet_code_resource_registrations=(
            redact_worksheet_code_resource_registrations
        ),
        redact_formula_defined_xlm_registrations=(
            redact_formula_defined_xlm_registrations
        ),
        redact_formula_defined_xlm_evaluations=redact_formula_defined_xlm_evaluations,
        redact_formula_defined_xlm_actions=redact_formula_defined_xlm_actions,
        redact_formula_defined_xlm_get_cell_calls=(
            redact_formula_defined_xlm_get_cell_calls
        ),
        redact_formula_defined_xlm_environment_information_calls=(
            redact_formula_defined_xlm_environment_information_calls
        ),
        redact_formula_environment_information=redact_formula_environment_information,
    )
    summary = payload["summary"]
    redactions = _active_redaction_labels(
        redact_external_workbook_links=redact_external_workbook_links,
        redact_formula_external_actions=redact_formula_external_actions,
        redact_python_in_excel=redact_python_in_excel,
        redact_office_custom_functions=redact_office_custom_functions,
        redact_unqualified_runtime_functions=redact_unqualified_runtime_functions,
        redact_worksheet_code_resource_registrations=(
            redact_worksheet_code_resource_registrations
        ),
        redact_formula_defined_xlm_registrations=(
            redact_formula_defined_xlm_registrations
        ),
        redact_formula_defined_xlm_evaluations=redact_formula_defined_xlm_evaluations,
        redact_formula_defined_xlm_actions=redact_formula_defined_xlm_actions,
        redact_formula_defined_xlm_get_cell_calls=(
            redact_formula_defined_xlm_get_cell_calls
        ),
        redact_formula_defined_xlm_environment_information_calls=(
            redact_formula_defined_xlm_environment_information_calls
        ),
        redact_formula_environment_information=redact_formula_environment_information,
    )
    title = "FormulaFence portfolio report"
    lede = (
        "A portable local review artifact for recursively matched workbooks. "
        "Use the filters to narrow findings and changes across the portfolio."
    )
    cards = (
        ("Matched workbooks", summary["matched_workbook_count"]),
        ("Semantic changes", summary["change_count"]),
        ("Findings", summary["finding_count"]),
        ("Highest severity", summary["highest_severity"]),
    )
    context = (
        (
            "Baseline / candidate workbooks: "
            f"{payload['before']['workbook_count']} / "
            f"{payload['after']['workbook_count']}"
        ),
        (
            "Added / removed / unreadable: "
            f"{summary['added_workbook_count']} / "
            f"{summary['removed_workbook_count']} / "
            f"{summary['unreadable_workbook_count']}"
        ),
        (
            "Cross-workbook impact analysis: "
            f"{'incomplete' if summary['cross_workbook_impact_incomplete'] else 'complete'}"
        ),
    )
    if max_bytes is None:
        workbook_sections: list[str] = []
        for entry in payload["workbooks"]:
            entry_summary = entry["summary"]
            workbook_sections.append(
                "\n".join(
                    [
                        '<section class="workbook">',
                        f"<h2><code>{_html_text(entry['path'])}</code></h2>",
                        (
                            "<p>"
                            f"Status: <strong>{_html_text(entry['status'])}</strong> · "
                            f"Changes: {_html_text(entry_summary['change_count'])} · "
                            f"Findings: {_html_text(entry_summary['finding_count'])} · "
                            "Highest severity: "
                            f"{_html_text(entry_summary['highest_severity'])}"
                            "</p>"
                        ),
                        _html_review_section("Findings", entry["findings"], "finding"),
                        _html_review_section(
                            "Semantic changes", entry["changes"], "change"
                        ),
                        "</section>",
                    ]
                )
            )
        content = "\n".join(
            ["<h2>Workbook details</h2>", *workbook_sections]
            if workbook_sections
            else ['<p class="empty">No supported workbooks were found.</p>']
        )
        return _html_document(title, lede, cards, context, redactions, content)

    def append_content(rendered: _BoundedText) -> None:
        if not payload["workbooks"]:
            _append_html_line(rendered, '<p class="empty">No supported workbooks were found.</p>')
            return
        _append_html_line(rendered, "<h2>Workbook details</h2>")
        for entry in payload["workbooks"]:
            entry_summary = entry["summary"]
            _append_html_line(rendered, '<section class="workbook">')
            _append_html_line(rendered, f"<h2><code>{_html_text(entry['path'])}</code></h2>")
            _append_html_line(
                rendered,
                "<p>"
                f"Status: <strong>{_html_text(entry['status'])}</strong> · "
                f"Changes: {_html_text(entry_summary['change_count'])} · "
                f"Findings: {_html_text(entry_summary['finding_count'])} · "
                "Highest severity: "
                f"{_html_text(entry_summary['highest_severity'])}"
                "</p>",
            )
            _append_html_review_section(rendered, "Findings", entry["findings"], "finding")
            _append_html_review_section(
                rendered,
                "Semantic changes",
                entry["changes"],
                "change",
            )
            _append_html_line(rendered, "</section>")

    return _render_bounded_html_document(
        title,
        lede,
        cards,
        context,
        redactions,
        max_bytes=max_bytes,
        append_content=append_content,
    )


def portfolio_to_sarif(
    report: PortfolioReport,
    *,
    redact_external_workbook_links: bool = False,
    redact_formula_external_actions: bool = False,
    redact_python_in_excel: bool = False,
    redact_office_custom_functions: bool = False,
    redact_unqualified_runtime_functions: bool = False,
    redact_worksheet_code_resource_registrations: bool = False,
    redact_formula_defined_xlm_registrations: bool = False,
    redact_formula_defined_xlm_evaluations: bool = False,
    redact_formula_defined_xlm_actions: bool = False,
    redact_formula_defined_xlm_get_cell_calls: bool = False,
    redact_formula_defined_xlm_environment_information_calls: bool = False,
    redact_formula_environment_information: bool = False,
) -> dict[str, Any]:
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
            finding_details = _safe_finding_details(
                finding,
                workbook.report,
                redact_formula_external_actions=redact_formula_external_actions,
                redact_office_custom_functions=redact_office_custom_functions,
                redact_unqualified_runtime_functions=redact_unqualified_runtime_functions,
                redact_worksheet_code_resource_registrations=(
                    redact_worksheet_code_resource_registrations
                ),
                redact_formula_defined_xlm_registrations=(
                    redact_formula_defined_xlm_registrations
                ),
                redact_formula_defined_xlm_evaluations=(
                    redact_formula_defined_xlm_evaluations
                ),
                redact_formula_defined_xlm_actions=redact_formula_defined_xlm_actions,
                redact_formula_defined_xlm_get_cell_calls=(
                    redact_formula_defined_xlm_get_cell_calls
                ),
                redact_formula_defined_xlm_environment_information_calls=(
                    redact_formula_defined_xlm_environment_information_calls
                ),
                redact_formula_environment_information=(
                    redact_formula_environment_information
                ),
            )
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
                        **finding_details,
                        "severity": finding.severity,
                        "portfolio_status": workbook.status,
                    },
                    "locations": [location],
                }
            )
    payload = {
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
    if redact_external_workbook_links:
        payload = redact_external_workbook_link_material(payload)
    if redact_formula_external_actions:
        payload = redact_formula_external_action_material(payload)
    if redact_python_in_excel:
        payload = redact_python_in_excel_material(payload)
    if redact_office_custom_functions:
        payload = redact_office_custom_function_material(payload)
    if redact_unqualified_runtime_functions:
        payload = redact_unqualified_runtime_function_material(payload)
    if redact_worksheet_code_resource_registrations:
        payload = redact_worksheet_code_resource_registration_material(payload)
    if redact_formula_defined_xlm_registrations:
        payload = redact_formula_defined_xlm_registration_material(payload)
    if redact_formula_defined_xlm_evaluations:
        payload = redact_formula_defined_xlm_evaluation_material(payload)
    if redact_formula_defined_xlm_actions:
        payload = redact_formula_defined_xlm_action_material(payload)
    if redact_formula_defined_xlm_get_cell_calls:
        payload = redact_formula_defined_xlm_get_cell_material(payload)
    if redact_formula_defined_xlm_environment_information_calls:
        payload = redact_formula_defined_xlm_environment_information_material(payload)
    if redact_formula_environment_information:
        payload = redact_formula_environment_information_material(payload)
    return payload
