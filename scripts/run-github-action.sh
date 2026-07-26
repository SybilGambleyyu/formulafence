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
fail_on=${INPUT_FAIL_ON:-none}
install=${INPUT_INSTALL:-true}
upload_artifact=${INPUT_UPLOAD_ARTIFACT:-true}

[[ -n "$action_path" ]] || fail 'GITHUB_ACTION_PATH is required.'
[[ -d "$action_path" ]] || fail "Action path does not exist: $action_path"
[[ -d "$workspace" ]] || fail "GitHub workspace does not exist: $workspace"
[[ -n "$baseline" ]] || fail 'The baseline input is required.'
[[ -n "$candidate" ]] || fail 'The candidate input is required.'
command -v python >/dev/null 2>&1 || fail 'Python is required; add actions/setup-python first.'

case "$format" in
  markdown|json|sarif) ;;
  *) fail "Unsupported format: $format (expected markdown, json, or sarif)." ;;
esac

case "$fail_on" in
  none|low|medium|high|critical) ;;
  *) fail "Unsupported fail-on level: $fail_on." ;;
esac

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
[[ -f "$baseline" ]] || fail "Baseline workbook does not exist: $baseline"
[[ -f "$candidate" ]] || fail "Candidate workbook does not exist: $candidate"
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
    if path == output:
        sys.exit("report output must not overwrite an input workbook or policy")

print(output)
PY
) || fail 'Report output and inputs must stay in the workspace, and the report must not overwrite an input.'
mkdir -p "$(dirname "$report_path")"

if [[ "$install" == true ]]; then
  python -m pip install --disable-pip-version-check --no-input "$action_path"
fi

command=(python -m formulafence.cli)
if [[ -n "$policy" ]]; then
  command+=(check "$baseline" "$candidate" --policy "$policy")
else
  command+=(diff "$baseline" "$candidate")
fi
command+=(--format "$format" --output "$report_path" --fail-on "$fail_on")

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
