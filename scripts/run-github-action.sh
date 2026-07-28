#!/usr/bin/env bash
# Run FormulaFence safely from the composite GitHub Action.

set -euo pipefail

fail() {
  printf 'formulafence action: %s\n' "$1" >&2
  exit 2
}

action_path=${GITHUB_ACTION_PATH:-}
workspace=${GITHUB_WORKSPACE:-$(pwd -P)}
baseline=${INPUT_BASELINE:-}
candidate=${INPUT_CANDIDATE:-}
policy=${INPUT_POLICY:-}
format=${INPUT_FORMAT:-markdown}
output=${INPUT_OUTPUT:-formulafence-report.md}
redact_external_workbook_links=${INPUT_REDACT_EXTERNAL_WORKBOOK_LINKS:-false}
redact_formula_external_actions=${INPUT_REDACT_FORMULA_EXTERNAL_ACTIONS:-false}
redact_python_in_excel=${INPUT_REDACT_PYTHON_IN_EXCEL:-false}
redact_office_custom_functions=${INPUT_REDACT_OFFICE_CUSTOM_FUNCTIONS:-false}
redact_unqualified_runtime_functions=${INPUT_REDACT_UNQUALIFIED_RUNTIME_FUNCTIONS:-false}
redact_worksheet_code_resource_registrations=${INPUT_REDACT_WORKSHEET_CODE_RESOURCE_REGISTRATIONS:-false}
redact_formula_defined_xlm_registrations=${INPUT_REDACT_FORMULA_DEFINED_XLM_REGISTRATIONS:-false}
redact_formula_defined_xlm_evaluations=${INPUT_REDACT_FORMULA_DEFINED_XLM_EVALUATIONS:-false}
redact_formula_defined_xlm_actions=${INPUT_REDACT_FORMULA_DEFINED_XLM_ACTIONS:-false}
redact_formula_defined_xlm_get_cell_calls=${INPUT_REDACT_FORMULA_DEFINED_XLM_GET_CELL_CALLS:-false}
redact_formula_defined_xlm_environment_information_calls=${INPUT_REDACT_FORMULA_DEFINED_XLM_ENVIRONMENT_INFORMATION_CALLS:-false}
redact_formula_environment_information=${INPUT_REDACT_FORMULA_ENVIRONMENT_INFORMATION:-false}
fail_on=${INPUT_FAIL_ON:-none}
max_workbooks=${INPUT_MAX_WORKBOOKS:-512}
max_inventory_entries=${INPUT_MAX_INVENTORY_ENTRIES:-32768}
max_portfolio_source_bytes=${INPUT_MAX_PORTFOLIO_SOURCE_BYTES:-4294967296}
max_portfolio_snapshot_cells=${INPUT_MAX_PORTFOLIO_SNAPSHOT_CELLS:-2000000}
max_change_analysis_states=${INPUT_MAX_CHANGE_ANALYSIS_STATES:-100000}
max_link_impact=${INPUT_MAX_LINK_IMPACT:-100000}
install=${INPUT_INSTALL:-true}
upload_artifact=${INPUT_UPLOAD_ARTIFACT:-true}

[[ -n "$action_path" ]] || fail 'GITHUB_ACTION_PATH is required.'
[[ -d "$action_path" ]] || fail "Action path does not exist: $action_path"
[[ -d "$workspace" ]] || fail "GitHub workspace does not exist: $workspace"
[[ -n "$baseline" ]] || fail 'The baseline input is required.'
[[ -n "$candidate" ]] || fail 'The candidate input is required.'
command -v python >/dev/null 2>&1 || fail 'Python is required; add actions/setup-python first.'

case "$format" in
  markdown|html|json|sarif) ;;
  *) fail "Unsupported format: $format (expected markdown, html, json, or sarif)." ;;
esac

case "$fail_on" in
  none|low|medium|high|critical) ;;
  *) fail "Unsupported fail-on level: $fail_on." ;;
esac

case "$redact_external_workbook_links" in
  true|false) ;;
  *) fail "Unsupported redact-external-workbook-links value: $redact_external_workbook_links (expected true or false)." ;;
esac

case "$redact_formula_external_actions" in
  true|false) ;;
  *) fail "Unsupported redact-formula-external-actions value: $redact_formula_external_actions (expected true or false)." ;;
esac

case "$redact_python_in_excel" in
  true|false) ;;
  *) fail "Unsupported redact-python-in-excel value: $redact_python_in_excel (expected true or false)." ;;
esac

case "$redact_office_custom_functions" in
  true|false) ;;
  *) fail "Unsupported redact-office-custom-functions value: $redact_office_custom_functions (expected true or false)." ;;
esac

case "$redact_unqualified_runtime_functions" in
  true|false) ;;
  *) fail "Unsupported redact-unqualified-runtime-functions value: $redact_unqualified_runtime_functions (expected true or false)." ;;
esac

case "$redact_worksheet_code_resource_registrations" in
  true|false) ;;
  *) fail "Unsupported redact-worksheet-code-resource-registrations value: $redact_worksheet_code_resource_registrations (expected true or false)." ;;
esac

case "$redact_formula_defined_xlm_registrations" in
  true|false) ;;
  *) fail "Unsupported redact-formula-defined-xlm-registrations value: $redact_formula_defined_xlm_registrations (expected true or false)." ;;
esac

case "$redact_formula_defined_xlm_evaluations" in
  true|false) ;;
  *) fail "Unsupported redact-formula-defined-xlm-evaluations value: $redact_formula_defined_xlm_evaluations (expected true or false)." ;;
esac

case "$redact_formula_defined_xlm_actions" in
  true|false) ;;
  *) fail "Unsupported redact-formula-defined-xlm-actions value: $redact_formula_defined_xlm_actions (expected true or false)." ;;
esac

case "$redact_formula_defined_xlm_get_cell_calls" in
  true|false) ;;
  *) fail "Unsupported redact-formula-defined-xlm-get-cell-calls value: $redact_formula_defined_xlm_get_cell_calls (expected true or false)." ;;
esac

case "$redact_formula_defined_xlm_environment_information_calls" in
  true|false) ;;
  *) fail "Unsupported redact-formula-defined-xlm-environment-information-calls value: $redact_formula_defined_xlm_environment_information_calls (expected true or false)." ;;
esac

case "$redact_formula_environment_information" in
  true|false) ;;
  *) fail "Unsupported redact-formula-environment-information value: $redact_formula_environment_information (expected true or false)." ;;
esac

if ! [[ "$max_workbooks" =~ ^[1-9][0-9]*$ ]]; then
  fail 'max-workbooks must be a positive integer.'
fi
if ! [[ "$max_inventory_entries" =~ ^[1-9][0-9]*$ ]]; then
  fail 'max-inventory-entries must be a positive integer.'
fi
if ! [[ "$max_portfolio_source_bytes" =~ ^[1-9][0-9]*$ ]]; then
  fail 'max-portfolio-source-bytes must be a positive integer.'
fi
if ! [[ "$max_portfolio_snapshot_cells" =~ ^[1-9][0-9]*$ ]]; then
  fail 'max-portfolio-snapshot-cells must be a positive integer.'
fi
if ! [[ "$max_change_analysis_states" =~ ^[1-9][0-9]*$ ]]; then
  fail 'max-change-analysis-states must be a positive integer.'
fi
if ! [[ "$max_link_impact" =~ ^[1-9][0-9]*$ ]]; then
  fail 'max-link-impact must be a positive integer.'
fi

case "$install" in
  true|false) ;;
  *) fail "Unsupported install value: $install (expected true or false)." ;;
esac

case "$upload_artifact" in
  true|false) ;;
  *) fail "Unsupported upload-artifact value: $upload_artifact (expected true or false)." ;;
esac

case "$output" in
  *$'\n'*|*$'\r'*) fail 'The output path must not contain a newline.' ;;
esac

cd "$workspace"
if [[ -f "$baseline" && -f "$candidate" ]]; then
  comparison_mode=workbook
elif [[ -d "$baseline" && -d "$candidate" ]]; then
  comparison_mode=portfolio
else
  fail 'Baseline and candidate must both be workbook files or both be directories.'
fi
if [[ -n "$policy" ]]; then
  [[ -f "$policy" ]] || fail "Policy file does not exist: $policy"
fi

export GITHUB_WORKSPACE="$workspace"
export INPUT_OUTPUT="$output"
report_path=$(python - <<'PY'
import os
import sys
from pathlib import Path

workspace = Path(os.environ["GITHUB_WORKSPACE"]).resolve()
output = Path(os.environ["INPUT_OUTPUT"])
if not output.is_absolute():
    output = workspace / output
output = output.resolve()
try:
    output.relative_to(workspace)
except ValueError:
    sys.exit("report output must stay inside GITHUB_WORKSPACE")

for key in ("INPUT_BASELINE", "INPUT_CANDIDATE", "INPUT_POLICY"):
    raw = os.environ.get(key, "")
    if not raw:
        continue
    path = Path(raw)
    if not path.is_absolute():
        path = workspace / path
    path = path.resolve()
    try:
        path.relative_to(workspace)
    except ValueError:
        sys.exit(f"{key} must stay inside GITHUB_WORKSPACE")
    if path.is_dir():
        try:
            output.relative_to(path)
        except ValueError:
            pass
        else:
            sys.exit("report output must not be written inside an input directory")
    elif path == output:
        sys.exit("report output must not overwrite an input workbook or policy")

print(output)
PY
) || fail 'Report output and inputs must stay in the workspace, and the report must not overwrite an input.'
mkdir -p "$(dirname "$report_path")"

if [[ "$install" == true ]]; then
  python -m pip install --disable-pip-version-check --no-input "$action_path"
fi

command=(python -m formulafence.cli)
if [[ "$comparison_mode" == portfolio ]]; then
  command+=(
    portfolio "$baseline" "$candidate"
    --max-workbooks "$max_workbooks"
    --max-inventory-entries "$max_inventory_entries"
    --max-portfolio-source-bytes "$max_portfolio_source_bytes"
    --max-portfolio-snapshot-cells "$max_portfolio_snapshot_cells"
    --max-link-impact "$max_link_impact"
  )
  if [[ -n "$policy" ]]; then
    command+=(--policy "$policy")
  fi
elif [[ -n "$policy" ]]; then
  command+=(check "$baseline" "$candidate" --policy "$policy")
else
  command+=(diff "$baseline" "$candidate")
fi
command+=(--max-change-analysis-states "$max_change_analysis_states")
command+=(--format "$format" --output "$report_path" --fail-on "$fail_on")
if [[ "$redact_external_workbook_links" == true ]]; then
  command+=(--redact-external-workbook-links)
fi
if [[ "$redact_formula_external_actions" == true ]]; then
  command+=(--redact-formula-external-actions)
fi
if [[ "$redact_python_in_excel" == true ]]; then
  command+=(--redact-python-in-excel)
fi
if [[ "$redact_office_custom_functions" == true ]]; then
  command+=(--redact-office-custom-functions)
fi
if [[ "$redact_unqualified_runtime_functions" == true ]]; then
  command+=(--redact-unqualified-runtime-functions)
fi
if [[ "$redact_worksheet_code_resource_registrations" == true ]]; then
  command+=(--redact-worksheet-code-resource-registrations)
fi
if [[ "$redact_formula_defined_xlm_registrations" == true ]]; then
  command+=(--redact-formula-defined-xlm-registrations)
fi
if [[ "$redact_formula_defined_xlm_evaluations" == true ]]; then
  command+=(--redact-formula-defined-xlm-evaluations)
fi
if [[ "$redact_formula_defined_xlm_actions" == true ]]; then
  command+=(--redact-formula-defined-xlm-actions)
fi
if [[ "$redact_formula_defined_xlm_get_cell_calls" == true ]]; then
  command+=(--redact-formula-defined-xlm-get-cell-calls)
fi
if [[ "$redact_formula_defined_xlm_environment_information_calls" == true ]]; then
  command+=(--redact-formula-defined-xlm-environment-information-calls)
fi
if [[ "$redact_formula_environment_information" == true ]]; then
  command+=(--redact-formula-environment-information)
fi

if "${command[@]}"; then
  exit_code=0
else
  exit_code=$?
fi

report_written=false
if [[ -f "$report_path" ]]; then
  report_written=true
fi

if [[ -n ${GITHUB_OUTPUT:-} ]]; then
  {
    printf 'report-path=%s\n' "$report_path"
    printf 'exit-code=%s\n' "$exit_code"
    printf 'report-written=%s\n' "$report_written"
  } >> "$GITHUB_OUTPUT"
fi

if [[ -n ${GITHUB_STEP_SUMMARY:-} ]]; then
  {
    printf '## FormulaFence\n\n'
    printf '**Exit code:** `%s`  \n' "$exit_code"
    printf '**Report:** `%s`\n\n' "$report_path"
    if [[ "$format" == markdown && "$report_written" == true ]]; then
      report_size=$(wc -c < "$report_path")
      if (( report_size > 900000 )); then
        head -c 900000 "$report_path"
        printf '\n\n_Report truncated in the job summary; download the artifact for the full report._\n'
      else
        cat "$report_path"
      fi
    else
      printf 'The generated report is available at the path above and as the configured artifact.\n'
    fi
  } >> "$GITHUB_STEP_SUMMARY"
fi

# Leave the composite action alive long enough to upload the report. Its final
# step re-emits this captured exit code after the artifact step has run.
exit 0
