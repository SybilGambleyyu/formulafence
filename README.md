# FormulaFence

FormulaFence is a local-first spreadsheet change-assurance CLI. It makes `.xlsx`
changes reviewable in CI: compare workbook semantics, trace downstream formula
impact, detect high-risk edits, and enforce a small policy file before a model
is shared or merged.

It never executes formulas or macros, and it does not upload workbook contents.

> Status: early alpha. The first release supports `.xlsx` and `.xlsm` inspection
> with formula-aware diffs, dependency impact, policy checks, Markdown/HTML/JSON/
> SARIF reports, and deterministic evidence metadata.

## Why

Spreadsheets remain the operating surface for financial models, planning,
operations, and research. Git sees an Excel workbook as a binary blob; ordinary
file diffs do not answer the questions a reviewer actually has:

- Did a formula become a hard-coded number?
- Which outputs now depend on the changed cell?
- Did a formula stop following the pattern used by its peers?
- Was a hidden sheet, macro payload, external link, or calculation setting changed?
- Does the change comply with the model's review policy?

FormulaFence is intentionally a guardrail, not a spreadsheet calculation engine
or hosted document system. It produces evidence a team can inspect in its normal
Git and CI workflow.

## Design basis

Spreadsheet controls already have useful specialist tools: [Git XL](https://www.xltrail.com/git-xl)
makes workbook diffs readable in Git, while [ExceLint](https://www.microsoft.com/en-us/research/publication/excelint-automatically-finding-spreadsheet-formula-errors/)
demonstrated static detection of formula anomalies. FormulaFence concentrates on
the merge boundary between those approaches: a local semantic diff coupled to
downstream impact and a review policy that can fail CI. It is complementary to,
not a replacement for, source control, model audit, or recalculation in Excel.

## Quick start

```bash
# Install the pinned public release directly from GitHub.
python -m pip install https://github.com/SybilGambleyyu/formulafence/releases/download/v0.202.0/formulafence-0.202.0-py3-none-any.whl

# Readable review report
formulafence diff baseline.xlsx candidate.xlsx --format markdown

# Enforce a policy in CI (non-zero when a rule fails)
formulafence check baseline.xlsx candidate.xlsx --policy formulafence.yml --format sarif --output results.sarif

# Self-contained browser review artifact
formulafence check baseline.xlsx candidate.xlsx --policy formulafence.yml --format html --output review.html

# Lint one workbook for conservative formula and calculation risks.
formulafence lint candidate.xlsx --fail-on high --format sarif --output formula-lint.sarif

# Compare a recursively matched portfolio of workbooks.
formulafence portfolio approved-models build/models --policy formulafence.yml --output portfolio-report.md
```

FormulaFence is not yet published to PyPI; the direct release URL above avoids
an ambiguous package-name install. See [GitHub Releases](https://github.com/SybilGambleyyu/formulafence/releases)
for the current version.

### Formula lint

`formulafence lint WORKBOOK` is a deliberately conservative, single-workbook
check for copy-paste, aggregate-range, protection, calculation-freshness,
error-checking-suppression, Excel Table calculated-column, static-circular-reference,
conditional-aggregate and `SUMPRODUCT` range-shape, `MMULT` matrix-dimension,
legacy-lookup return-index, `RANDBETWEEN` literal-bound, `SUBTOTAL`
function-code, `INDEX` literal-position, explicit-broken-reference, and
saved-result risks that a version diff cannot see. It reports a copied-formula
interruption only
when the immediately preceding and following formulas have the same
relative-copy fingerprint and a third contiguous peer repeats that fingerprint.
It also recognizes narrow aggregate-range, formula-protection,
calculation-freshness, stored error-checking suppression, interior Table
calculated-column exception, direct and multi-cell static circular-reference,
direct conditional-aggregate and `SUMPRODUCT` range-shape, direct static
`MMULT` matrix-dimension, direct static legacy-lookup return-index, direct
literal `RANDBETWEEN` bounds, direct literal `SUBTOTAL` function codes, direct
literal `INDEX` row/column positions, explicit-broken-reference, and saved
broken-reference-result signals. Together these produce nineteen reviewable
findings:

- `FF082`: a non-formula interruption. Blanks and stored error values are high;
  numeric/manual values are medium; textual markers are low.
- `FF083` (medium): a formula differs from the otherwise stable copied block.
- `FF084` (medium): a pure local `SUM`, `AVERAGE`, `MIN`, `MAX`, or `COUNT`
  stops before at least two contiguous literal numeric cells on the same row or
  column between its one-dimensional A1 range and the aggregate formula.
- `FF085` (medium): a formula has an explicit direct-cell `locked=false`
  assignment while its worksheet is actively protected.
- `FF086` (medium): a workbook with at least one formula explicitly records
  `calcMode=manual` and `calcCompleted=false`, meaning calculation was not
  completed before it was saved.
- `FF087` (high): an ordinary formula's resolved scalar static dependency
  returns to its own cell while workbook calculation iteration is disabled.
- `FF088` (critical): a stored formula's tokenized syntax contains an explicit
  `#REF!` error operand.
- `FF089` (high): a formula's well-formed saved result is a broken-reference
  error from its last recorded calculation.
- `FF090` (high): an ordinary formula participates in a multi-cell static
  circular-reference component while workbook calculation iteration is disabled.
- `FF091` (medium): the workbook has recognized stored Excel error-checking
  suppressions, so review prompts may be hidden.
- `FF092` (medium): an interior Excel Table data cell differs from the stored
  calculated-column master while its immediate neighboring rows match it.
- `FF093` (high): a native `SUMIFS`, `COUNTIFS`, `AVERAGEIFS`, `MAXIFS`, or
  `MINIFS` call uses direct static range arguments with different dimensions.
- `FF094` (high): a native `SUMPRODUCT` call uses direct static range arguments
  with different dimensions.
- `FF095` (high): a native `MMULT` call uses direct static arrays with
  incompatible matrix dimensions.
- `FF096` (high): a native `VLOOKUP` or `HLOOKUP` call uses a literal return
  index outside its direct static table range.
- `FF097` (high): a native `CHOOSE` call uses a literal index outside its
  available value arguments.
- `FF098` (high): a native `RANDBETWEEN` call uses direct literal bounds with
  the bottom above the top.
- `FF099` (high): a native `SUBTOTAL` call uses a literal function number
  outside Excel's supported codes.
- `FF100` (high): a native `INDEX` call uses a literal row or column number
  outside its direct static array.

Use `--fail-on critical` to gate explicit broken-reference operands, or
`--fail-on high` to also gate blank/error interruptions and direct static
self-references, multi-cell static cycles, direct conditional-aggregate or
`SUMPRODUCT` range-shape mismatches, direct static `MMULT` matrix-dimension
mismatches, direct static legacy-lookup return-index mismatches, direct static
`CHOOSE` literal-index mismatches, direct literal `RANDBETWEEN` bound
mismatches, direct literal `SUBTOTAL` function-code mismatches, direct literal
`INDEX` row/column-position mismatches, and saved broken-reference results.
Use `--fail-on medium` to additionally require review of manual-value,
formula-outlier, aggregate-range, explicit formula-protection,
incomplete-manual-calculation, error-checking-suppression, and Table
calculated-column exceptions.
`FF084` intentionally ignores
named/table/external/3-D references, multi-range or computed expressions, a
formula before or beside its range, one-cell gaps, nonnumeric gaps, tokenizer
failures, and array-formula territory. `FF085` deliberately does not infer
row-, column-, default-style, or allowed-edit-range protection state. `FF086`
requires both stored calculation flags and a formula; manual mode alone,
automatic calculation, a completed save, and an omitted completion marker stay
quiet. It is a configuration prompt, not a claim that a particular cached
result is wrong. `FF091` accepts only recognized stored per-range Excel
error-checking suppressions and emits aggregate warning categories and counts,
never target ranges. It does not decide whether a suppressed prompt would
apply, or whether its suppression was intentional. `FF087` remains the direct
self-reference signal. `FF092` accepts only a non-array stored Excel Table
calculated-column master and an interior data-row cell with two immediate
eligible neighboring formulas that match that master fingerprint. It leaves
first/last data rows, array territory, explicit broken-reference formulas,
uninspectable formulas, and broader or contiguous exception runs quiet; it does
not decide whether an exception was intentional. Its evidence retains only the
affected location, exception kind, and matching-peer count, never the table
identity or master formula. `FF093` accepts only native `SUMIFS`, `COUNTIFS`,
`AVERAGEIFS`, `MAXIFS`, and `MINIFS` calls (optionally with `@`) plus the exact
OOXML `_xlfn.MAXIFS` and `_xlfn.MINIFS` serializations. It requires valid arity
and every relevant range argument to be a single bounded, internal A1
cell/range or whole-column reference, then reports differing dimensions without
evaluating the formula. Named, Table, external, 3-D, full-row, union, computed,
dynamic, spill, implicit-intersection, malformed, explicit-broken-reference,
and array-formula territory stays quiet. Its evidence retains only the affected
location, number of conditional-aggregate calls, and number of mismatched
direct range arguments—never a formula, range spelling, or Table identity.
`FF094` accepts only native `SUMPRODUCT` calls (optionally with `@`) with at
least two comma-separated arguments, every one a direct bounded internal A1
cell/range or whole-column reference. It reports different dimensions without
evaluating the formula. Names, Tables, external, 3-D, full-row, union,
computed, dynamic, spill, implicit-intersection, malformed,
explicit-broken-reference, and array-formula territory stays quiet. Its
evidence retains only the affected location, number of qualifying calls, and
number of mismatched direct array arguments—never a formula, range spelling,
or source sheet identity.
`FF095` accepts only native `MMULT` calls (optionally with `@`) with exactly
two comma-separated arguments, each one a direct bounded internal A1
cell/range or whole-column reference. It reports only when the first argument's
column count differs from the second argument's row count, without inspecting
cell values. Names, Tables, external, 3-D, full-row, union, computed, dynamic,
spill, implicit-intersection, malformed, explicit-broken-reference, and
array-formula territory stay quiet. Its evidence retains only the affected
location, number of qualifying calls, and number of incompatible direct matrix
pairs—never a formula, range spelling, or source sheet identity.
`FF096` accepts only native `VLOOKUP` and `HLOOKUP` calls (optionally with
`@`) with exactly three or four comma-separated arguments. Its table argument
must be one direct bounded internal A1 cell/range or whole-column reference and
its return index must be one direct positive integer literal. It reports only
when a `VLOOKUP` index exceeds the table width or an `HLOOKUP` index exceeds
the table height, without inspecting lookup values or table values. Names,
Tables, external, 3-D, full-row, union, computed, dynamic, spill,
implicit-intersection, malformed, nonliteral or nonpositive-index,
explicit-broken-reference, and array-formula territory stay quiet. Its evidence
retains only the affected location, number of qualifying calls, and number of
out-of-range literal return indices—never a formula, range spelling, or source
sheet identity.
`FF097` accepts only native `CHOOSE` calls (optionally with `@`) with a direct
bare nonnegative decimal index and one through 254 nonempty value arguments.
It reports only when that index is zero or exceeds the supplied value-argument
count, without inspecting selected values. Computed, signed, decimal, array,
or dynamic indices; malformed calls; explicit-broken-reference operands;
array-formula territory; and arbitrary namespaces stay quiet. Its evidence
retains only the affected location, number of qualifying calls, and number of
out-of-range literal indices—never a formula, value argument, or source sheet
identity.
`FF098` accepts only native `RANDBETWEEN` calls (optionally with `@`) with
exactly two direct decimal integer literals, each optionally preceded by one
unary `+` or `-`. It reports only when the bottom literal exceeds the top
literal, without calculating a random value. Decimal or scientific notation,
computed expressions, references, arrays, malformed calls, explicit
broken-reference operands, array-formula territory, and arbitrary namespaces
stay quiet. Its evidence retains only the affected location, number of
qualifying calls, and number of inverted literal-bound pairs—never a formula,
literal value, or source sheet identity.
`FF099` accepts only native `SUBTOTAL` calls (optionally with `@`) with a
direct bare nonnegative decimal function number plus one through 254 nonempty
reference arguments. It reports only when that literal is outside Excel's
documented 1–11 and 101–111 code families, without inspecting references or
calculating a subtotal. Computed, signed, decimal, array, malformed, explicit
broken-reference operands, array-formula territory, and arbitrary namespaces
stay quiet. Its evidence retains only the affected location, number of
qualifying calls, and number of unsupported literal function codes—never a
formula, function-code value, reference, or source sheet identity.
`FF100` accepts only native `INDEX` calls (optionally with `@`) with two or
three nonempty arguments, one direct bounded internal A1 cell/range or
whole-column array, and direct bare nonnegative decimal row and optional
column literals. It reports only when a positive row literal exceeds the array
height or a positive column literal exceeds the array width. Zero preserves
Excel's documented whole-row or whole-column array behavior and stays quiet.
Computed, signed, decimal, array, malformed, explicit-broken-reference, and
namespaced forms, plus array-formula territory, remain outside the boundary.
Its evidence retains only the affected location, number of qualifying calls,
and number of out-of-range literal positions—never a formula, position value,
array range spelling, or source sheet identity.
`FF090`
uses only a strongly connected component of
resolved scalar static dependencies with at least two eligible ordinary formula
cells; it never expands a range or evaluates a formula. Both circular-reference
signals stay quiet when `iterate=true`; `FF090` also leaves dynamic-reference,
3-D, spill, explicit-intersection, array, and tokenizer-failure territory
outside its boundary. `FF088` requires a tokenizer `#REF!` error operand: a
matching text literal, quoted worksheet name, or tokenization failure stays
quiet. JSON, Markdown, and SARIF output show only locations, static range
coordinates, calculation-status flags, direct- or multi-cell-static scope, a
multi-cell component size, saved-result facts, aggregate error-checking
suppression counts, Table exception kinds, conditional-aggregate mismatch
counts, `SUMPRODUCT` mismatch counts, and `MMULT` incompatible-matrix-pair
counts, legacy-lookup out-of-range-literal-index counts, and `RANDBETWEEN`
inverted-literal-bound counts, `SUBTOTAL` unsupported-literal-function-code
counts, and `INDEX` out-of-range-literal-index counts—never formula text,
cached values, ignored-error target ranges, direct
conditional-aggregate, `SUMPRODUCT`, `MMULT`, legacy-lookup range spellings,
`INDEX` position values or direct array ranges, or Table master formulas. `FF089` accepts
only a valid saved formula-result cache whose exact
error is `#REF!`; other saved errors, missing or malformed cache records, and
locations already covered by `FF088` stay quiet. It is a high-severity record
of the last saved display state, not proof of the formula's current result. The
lint does not calculate formulas and fails closed if array metadata is
incomplete.

It retains at most 10,000 total formula-lint findings by default; use
`--max-formula-pattern-findings` to choose another positive reviewed bound.
`FF084` inspects gaps of at most 128 cells by default; set
`--max-aggregate-omission-gap-cells` to an explicit integer of at least two
when a workbook's layout calls for another bounded window.

### GitHub Actions

The public composite Action installs FormulaFence from the selected Action
source, writes the report inside the workspace, embeds Markdown in the job
summary when requested, and uploads the report before it re-emits a policy
failure. It accepts a baseline,
candidate, optional policy, report format, and output path; matching directory
inputs invoke the portfolio mode. Its `report-path` and `exit-code` outputs are
available to later steps. Use a tagged release for readability and pin to an
immutable commit in a production workflow.

```yaml
- uses: actions/setup-python@v6
  with:
    python-version: '3.12'
- id: formulafence
  uses: SybilGambleyyu/formulafence@v0.202.0
  with:
    baseline: models/approved/model.xlsx
    candidate: build/model.xlsx
    policy: models/formulafence.yml
    output: reports/formulafence.md
```

See [the CI integration guide](docs/ci.md) for browser-review, SARIF, artifact,
and preinstalled-package options.

### Review in a browser

`--format html` creates one portable review page for `diff`, `check`, or
`portfolio`. It has local text/severity filters and expandable complete evidence
for every finding and semantic change:

```bash
formulafence check baseline.xlsx candidate.xlsx \
  --policy formulafence.yml \
  --format html \
  --output reports/formulafence-review.html
```

The page contains only inline styles and a small local filtering script; it
does not fetch remote assets or send workbook content anywhere. Workbook-derived
text is HTML-escaped before rendering. The same sharing-redaction switches work
with HTML, so apply them before uploading a report outside the trusted review
boundary.

### Sharing reports with external-workbook links

Ordinary diff reports deliberately retain full local reviewer evidence. When a
JSON, Markdown, HTML, or SARIF artifact must leave that trusted review boundary, add
`--redact-external-workbook-links` to `diff`, `check`, or `portfolio`:

```bash
formulafence check baseline.xlsx candidate.xlsx \
  --policy formulafence.yml \
  --format sarif \
  --redact-external-workbook-links \
  --output results.sarif
```

The option is output-only: comparison facts, policy evaluation, exit status,
and the in-memory report remain unchanged. It replaces a whole serialized
string that contains a FormulaFence-recognized literal static external-workbook
reference (including direct A1, 3-D, defined-name, data-validation, or
book-only table forms) with `[external-workbook link material redacted]`. It
also conservatively hides plainly visible bracketed/dynamic literals such as
`INDIRECT("'[Book]Sheet'!A1")`; it never evaluates a formula or reconstructs a
link assembled from text fragments. This is deliberately not a general
sensitive-data vault or a guarantee to redact every dynamic Excel expression.
Without the option, existing local-review output is unchanged. Markdown notes
when the option was used. The GitHub Action exposes the same behavior through
`redact-external-workbook-links: 'true'`.

### Sharing reports with formula actions or DDE links

Formula action and provider formulas can contain URLs, provider names,
connection/query text, DDE service/topic/item values, or static input cells
that route into those formulas. For an artifact that crosses the local review
boundary, add `--redact-formula-external-actions` to `diff`, `check`, or
`portfolio`:

```bash
formulafence check baseline.xlsx candidate.xlsx \
  --policy formulafence.yml \
  --format json \
  --redact-formula-external-actions \
  --output shared-report.json
```

This output-only mode covers direct stored formulas that FormulaFence already
inventories under `FF064` (`HYPERLINK`, `WEBSERVICE`, `IMAGE`, `RTD`,
`STOCKHISTORY`, and the documented Cube family) and `FF074` direct DDE syntax.
It replaces the whole serialized action/DDE formula value with
`[formula external-action material redacted]`. It also hides before/after cell
evidence for a changed action/DDE formula or for an exact changed static input
that FormulaFence's private dependency analysis recorded as reaching one. When
a relevant formula-defined-name chain changes, it conservatively hides changed
defined-name before/after values as well, since a wrapper can carry an endpoint
without spelling the action itself.

The switch does not mutate snapshots, alter findings, policy evaluation, or an
exit status; it does not calculate formulas, contact a provider or DDE server,
or reconstruct a destination assembled dynamically at Excel calculation time.
It is not a general secret scrubber and does not replace
`--redact-external-workbook-links` for workbook-link material. Use both options
when a shared report can contain both surfaces. Without the option, ordinary
local-review output remains unchanged. The GitHub Action exposes the same
boundary through `redact-formula-external-actions: 'true'`.

### Sharing reports with Python in Excel

Microsoft documents `PY(python_code, return_type)` with static Python source,
so an ordinary changed formula can reveal source text, an `xl()` reference, or
a changed ordinary input used by a stored PY binding. For an artifact that
crosses the local review boundary, add `--redact-python-in-excel` to `diff`,
`check`, or `portfolio`:

```bash
formulafence check baseline.xlsx candidate.xlsx \
  --policy formulafence.yml \
  --format json \
  --redact-python-in-excel \
  --output shared-report.json
```

This output-only mode covers direct stored `PY` formula text that FormulaFence
already inventories under `FF065`. It replaces the whole serialized formula
value with `[Python-in-Excel material redacted]`, and also hides before/after
cell evidence for a changed PY formula or for an exact changed static input
that FormulaFence's private dependency analysis recorded as reaching one.
Stored `python.xml` / `pythonScripts.xml` source remains private under the
existing Python ledger regardless of this switch.

The switch does not mutate snapshots, alter findings, policy evaluation, or an
exit status; it does not parse Python, calculate a formula, contact Microsoft
Cloud, or reconstruct a value at Excel calculation time. It is not a general
secret scrubber, does not redact arbitrary workbook cells, and does not replace
the external-workbook-link or formula-action sharing boundaries. Use the
relevant switches together when a report contains multiple sensitive surfaces.
Without the option, ordinary local-review output remains unchanged. The GitHub
Action exposes the same boundary through `redact-python-in-excel: 'true'`.

### Sharing reports with Office custom functions

Office Add-in custom functions use a namespace in a worksheet formula and can
request data from a service, so a changed call can disclose an add-in namespace,
callable, query argument, or static input in an ordinary report. For an artifact
that crosses the local review boundary, add `--redact-office-custom-functions`
to `diff`, `check`, or `portfolio`:

```bash
formulafence check baseline.xlsx candidate.xlsx \
  --policy formulafence.yml \
  --format json \
  --redact-office-custom-functions \
  --output shared-report.json
```

This output-only mode covers direct stored namespaced call material that
FormulaFence inventories under `FF066`. It replaces a whole serialized value
with `[Office custom-function material redacted]`, hides before/after evidence
for a changed custom-function formula or an exact changed static input that the
private dependency analysis recorded as reaching one, and conservatively hides
changed formula-defined-name before/after values when a private
custom-function-relevant definition chain changed. That last rule protects an
ordinary-looking wrapper whose private argument reaches a namespaced call only
through another named `LAMBDA`.

The switch does not mutate snapshots, findings, policy evaluation, or an exit
status; it does not calculate formulas, load a manifest or add-in, execute
JavaScript, contact a custom-function runtime, or reconstruct a dynamically
assembled argument. It is not a general secret scrubber and does not replace
the external-workbook-link, formula-action, or Python-in-Excel sharing
boundaries. Use the relevant switches together when a report contains multiple
surfaces. Without the option, ordinary local-review output remains unchanged.
The GitHub Action exposes the same boundary through
`redact-office-custom-functions: 'true'`. The scope follows Microsoft's
[custom-functions overview](https://learn.microsoft.com/en-us/office/dev/add-ins/excel/custom-functions-overview)
and [web-data guidance](https://learn.microsoft.com/en-us/office/dev/add-ins/excel/custom-functions-web-reqs).

### Sharing reports with unqualified runtime functions

A bare unknown worksheet call can bind to a VBA UDF, COM/Automation add-in,
XLL, or another registered runtime. Its stored formula can expose a proprietary
call name, arguments, or an ordinary static input in a generic report. For an
artifact that crosses the local review boundary, add
`--redact-unqualified-runtime-functions` to `diff`, `check`, or `portfolio`:

```bash
formulafence check baseline.xlsx candidate.xlsx \
  --policy formulafence.yml \
  --format json \
  --redact-unqualified-runtime-functions \
  --output shared-report.json
```

This output-only mode covers direct stored bare-call material that FormulaFence
inventories under `FF075`. It replaces a whole serialized value with
`[unqualified runtime-function material redacted]`, hides before/after evidence
for a changed candidate formula or exact changed static input that the private
dependency analysis recorded as reaching one, and conservatively hides changed
formula-defined-name before/after values when a private resolved runtime-name
chain changed. That last rule protects an ordinary-looking dotted wrapper whose
private argument eventually reaches a bare UDF call. Because a shared renderer
does not retain workbook-local name bindings, it may conservatively replace a
standalone unknown bare-call-shaped string too; default local-review output is
unchanged.

The switch does not mutate snapshots, findings, policy evaluation, or an exit
status; it does not calculate formulas, resolve or load VBA, COM/Automation,
XLL, or another provider, execute code, contact a runtime, or reconstruct a
dynamically assembled argument. It is not a general secret scrubber and does
not replace the external-workbook-link, formula-action, Python-in-Excel, or
Office custom-function sharing boundaries. Use the relevant switches together
when a report contains multiple sensitive surfaces. The GitHub Action exposes
the same boundary through `redact-unqualified-runtime-functions: 'true'`. The
scope follows Microsoft's [installed UDF guidance](https://support.microsoft.com/en-us/excel/user-defined-functions-that-are-installed-with-add-ins-reference),
[VBA custom-function guidance](https://support.microsoft.com/en-us/excel/create-custom-functions-in-excel),
and [XLL registration/call guidance](https://learn.microsoft.com/en-us/office/client-developer/excel/accessing-xll-code-in-excel).

### Sharing reports with worksheet code-resource registrations

`REGISTER.ID(module_text, procedure, [type_text])` can register a DLL or code
resource from a worksheet, so its stored module, procedure, type string, or an
ordinary static input can appear in a generic report. For an artifact that
crosses the local review boundary, add
`--redact-worksheet-code-resource-registrations` to `diff`, `check`, or
`portfolio`:

```bash
formulafence check baseline.xlsx candidate.xlsx \
  --policy formulafence.yml \
  --format json \
  --redact-worksheet-code-resource-registrations \
  --output shared-report.json
```

This output-only mode covers direct stored `REGISTER.ID` material FormulaFence
inventories under `FF067`. It replaces a whole serialized value with
`[worksheet code-resource registration material redacted]`, hides before/after
evidence for a changed registration formula or exact changed static input that
the private dependency analysis recorded as reaching one, and conservatively
hides changed formula-defined-name before/after values when a private resolved
registration chain changed. That last rule protects an ordinary-looking dotted
wrapper whose private argument eventually reaches `REGISTER.ID`. Because a
shared renderer does not retain workbook-local name bindings, it may
conservatively replace a standalone `REGISTER.ID`-shaped string too; default
local-review output is unchanged.

The switch does not mutate snapshots, findings, policy evaluation, or an exit
status; it does not calculate formulas, resolve a module path, load a DLL/XLL,
inspect host trust settings, execute code, contact a provider, or reconstruct a
dynamically assembled argument. It is not a general secret scrubber and does
not replace the external-workbook-link, formula-action, Python-in-Excel, Office
custom-function, or unqualified-runtime-function sharing boundaries. Use the
relevant switches together when a report contains multiple sensitive surfaces.
The GitHub Action exposes the same boundary through
`redact-worksheet-code-resource-registrations: 'true'`. The scope follows
Microsoft's [`REGISTER.ID` reference](https://support.microsoft.com/en-us/office/register-id-function-f8f0af0f-fd66-4704-a0f2-87b27b175b50),
which documents the worksheet-capable DLL/code-resource registration function.

### Sharing reports with formula-defined XLM registrations

Legacy XLM `REGISTER` can be stored in a formula-defined name or named
`LAMBDA`, where its module, procedure, type string, and static arguments can
surface through ordinary changed-name or cell evidence even though the `FF068`
ledger itself publishes only counts. For an artifact that crosses the local
review boundary, add `--redact-formula-defined-xlm-registrations` to `diff`,
`check`, or `portfolio`:

```bash
formulafence check baseline.xlsx candidate.xlsx \
  --policy formulafence.yml \
  --format json \
  --redact-formula-defined-xlm-registrations \
  --output shared-report.json
```

This output-only mode covers direct stored `REGISTER` material FormulaFence
inventories under `FF068`. It replaces a whole serialized value with
`[formula-defined XLM registration material redacted]`, hides before/after
evidence for a changed invoking formula or exact changed static input that the
private dependency analysis recorded as reaching one, and conservatively hides
changed formula-defined-name before/after values when a private resolved
registration chain changed. That last rule protects an ordinary-looking dotted
wrapper whose private argument eventually reaches `REGISTER`. A shared renderer
does not retain workbook-local name bindings, so it may conservatively replace a
standalone `REGISTER`-shaped string too; default local-review output is
unchanged.

The switch does not mutate snapshots, findings, policy evaluation, or an exit
status; it does not calculate formulas, execute a macro, resolve a module path,
load a DLL/XLL, inspect host trust settings, contact a provider, or reconstruct
a dynamically assembled argument. It is not a general secret scrubber and does
not replace the external-workbook-link, formula-action, Python-in-Excel, Office
custom-function, unqualified-runtime-function, or worksheet-code-resource
registration sharing boundaries. The GitHub Action exposes the same boundary
through `redact-formula-defined-xlm-registrations: 'true'`. The scope follows
Microsoft's [`xlfRegister` Form 1 reference](https://learn.microsoft.com/en-au/office/client-developer/excel/xlfregister-form-1)
and [`Form 2` reference](https://learn.microsoft.com/en-au/office/client-developer/excel/xlfregister-form-2),
which document XLM registration for DLL functions/commands and XLL activation.

### Sharing reports with formula-defined XLM evaluation

Legacy XLM `EVALUATE` can be stored in a formula-defined name or named
`LAMBDA`, where its expression text and static arguments can surface through
ordinary changed-name or cell evidence even though the `FF069` ledger itself
publishes only counts. For an artifact that crosses the local review boundary,
add `--redact-formula-defined-xlm-evaluations` to `diff`, `check`, or
`portfolio`:

```bash
formulafence check baseline.xlsx candidate.xlsx \
  --policy formulafence.yml \
  --format json \
  --redact-formula-defined-xlm-evaluations \
  --output shared-report.json
```

This output-only mode covers direct stored `EVALUATE` material FormulaFence
inventories under `FF069`. It replaces a whole serialized value with
`[formula-defined XLM evaluation material redacted]`, hides before/after
evidence for a changed invoking formula or exact changed static input that the
private dependency analysis recorded as reaching one, and conservatively hides
changed formula-defined-name before/after values when a private resolved
evaluation chain changed. That last rule protects an ordinary-looking dotted
wrapper whose private expression eventually reaches `EVALUATE`. A shared
renderer does not retain workbook-local name bindings, so it may conservatively
replace a standalone `EVALUATE`-shaped string too; default local-review output
is unchanged.

The switch does not mutate snapshots, findings, policy evaluation, or an exit
status; it does not calculate formulas, evaluate the expression text, parse a
runtime-generated expression, execute a macro, or reconstruct dynamically
assembled text. It is not a general secret scrubber and does not replace the
external-workbook-link, formula-action, Python-in-Excel, Office
custom-function, unqualified-runtime-function, or registration sharing
boundaries. Runtime text remains a static-coverage limit: this mode protects
only direct stored material and the existing proven argument graph. The GitHub
Action exposes the same boundary through
`redact-formula-defined-xlm-evaluations: 'true'`. The scope follows Microsoft's
[Excel expression-evaluation reference](https://learn.microsoft.com/en-us/office/client-developer/excel/excel-worksheet-and-expression-evaluation),
which documents `EVALUATE` reducing a valid character string to an Excel value.

### Sharing reports with formula-defined XLM actions

Selected legacy XLM actions and event-dispatch calls can be stored in a
formula-defined name or named `LAMBDA`, where their targets, handlers, and
static arguments can surface through ordinary changed-name or cell evidence
even though the `FF073` ledger itself publishes only counts. For an artifact
that crosses the local review boundary, add
`--redact-formula-defined-xlm-actions` to `diff`, `check`, or `portfolio`:

```bash
formulafence check baseline.xlsx candidate.xlsx \
  --policy formulafence.yml \
  --format json \
  --redact-formula-defined-xlm-actions \
  --output shared-report.json
```

This output-only mode covers direct stored selected XLM action material
FormulaFence inventories under `FF073`. It replaces a whole serialized value
with `[formula-defined XLM action material redacted]`, hides before/after
evidence for a changed invoking formula or exact changed static input that the
private dependency analysis recorded as reaching one, and conservatively hides
changed formula-defined-name before/after values when a private resolved action
chain changed. That last rule protects an ordinary-looking dotted wrapper whose
private target or handler eventually reaches a selected action. A shared
renderer does not retain workbook-local name bindings, so it may conservatively
replace a standalone selected-action-shaped string too; default local-review
output is unchanged.

The switch does not mutate snapshots, findings, policy evaluation, or an exit
status; it does not calculate formulas, resolve an action target or event
handler, load a DLL, send DDE, execute a macro or program, or reconstruct a
dynamically assembled action. It is not a general secret scrubber and does not
replace the external-workbook-link, formula-action, Python-in-Excel, Office
custom-function, unqualified-runtime-function, or registration/evaluation
sharing boundaries. The GitHub Action exposes the same boundary through
`redact-formula-defined-xlm-actions: 'true'`. The scope follows Microsoft's
[Excel C API reference](https://learn.microsoft.com/en-us/office/client-developer/excel/programming-with-the-c-api-in-excel)
and [DLL-access guidance](https://learn.microsoft.com/en-us/office/client-developer/excel/how-to-access-dlls-in-excel),
which document XLM command-equivalent functions, event traps, and `CALL` as a
DLL-access route.

### Portfolio gates

`formulafence portfolio BASELINE_DIRECTORY CANDIDATE_DIRECTORY` recursively
matches supported `.xlsx` and `.xlsm` files by their relative paths, then
produces one deterministic JSON, Markdown, HTML, or SARIF review artifact. A path
present on only one side emits high-severity `FF077`; enable
`no_portfolio_membership_changes` to fail it as `FFP077`. FormulaFence never
guesses renames, so a move remains a visible removal plus addition.

When a changed candidate cell is read through a direct static external A1
reference, a static external 3-D A1 span such as
`=[Inputs.xlsx]Jan:Mar!$B$2`, the documented workbook-scoped name form such as
`=[Inputs.xlsx]InputRange`, the sheet-local name form such as
`=[Inputs.xlsx]Data!LocalInput`, or Excel's package-indexed cell/range or name
forms such as `=[1]Data!$B$2:$B$4`, `=[1]!InputRange`, and
`=[1]Data!LocalInput` (including the indexed 3-D form
`=[1]Jan:Mar!$B$2`), or an external structured-table selector such as
`='..\\inputs\\source.xlsx'!Sales[Amount]` or `=[1]!Sales[#Data]`, by a
formula in another candidate workbook,
FormulaFence emits high-severity `FF079` with relative workbook/cell paths and
deterministic shortest-path samples. Any supported form may be retained through
a finite, acyclic chain of workbook-scoped consumer aliases. Its terminal may
be a direct external A1, workbook-scoped-name, or selector-bearing table alias
such as a defined name storing `'..\\inputs\\[Inputs.xlsx]Data'!$B$2:$B$4`,
`'..\\inputs\\[Inputs.xlsx]InputRange'`, or
`'..\\inputs\\source.xlsx'!Sales[Amount]`; every intermediate definition
must be exactly one unqualified, non-A1 name identity, with or without its
leading `=`. It may also be reached through an eligible workbook-scoped
formula-defined name, for example `=SUM(ExternalInput)`,
`=SUM('..\\inputs\\[Inputs.xlsx]Data'!$B$2:$B$4)`, or a named
`LAMBDA` such as `=LAMBDA(value,SUM(value,Inputs!$B$2,ExternalInput))`.
FormulaFence does not calculate either expression: it retains every static
endpoint and fixed internal input only when every external token is one already
parsed static direct or package-validated endpoint, every remaining reference
is static, and the definition has no broken reference, unresolved token,
tokenizer failure, dynamic-reference function, relative internal A1 reference,
local 3-D form, spill reference, or explicit implicit intersection. A named
LAMBDA carries those edges only at an actual function call, never through a
bare name. Eligible global formula names and named LAMBDAs can call each other;
local or locally shadowed definitions never enter this bridge.
An indexed form is eligible only when its one-based
`externalReferences` declaration identifies exactly one `externalBook` and
external `externalLinkPath` relationship whose target resolves to one exact
relative candidate. An indexed A1 form must be one static
cell/range/whole-row/whole-column destination. A 3-D span is expanded only when
that source candidate has a complete raw OOXML tab catalog consistent with its
inspected worksheet order; its endpoints must be exact, forward, ordinary
worksheets. A source name must fully expand to static internal A1 destinations
without evaluation; a sheet-local spelling uses only the matching source sheet's
local scope, never a same-named global fallback. An external table spelling is
book-only (not source-sheet-qualified), must include an explicit selector, and
is resolved only when its source candidate has exactly one case-insensitive
matching table. FormulaFence then uses only its static `#All`, `#Data`,
`#Headers`, `#Totals`, column, or contiguous-column bounds. Enable
`no_cross_workbook_impacts` to make that evidence `FFP079`.

This is deliberately a narrow local graph: it never opens or fetches a link
target, guesses by filename, evaluates a formula, or trusts an external-link
cache. Raw external source paths, package targets, and indexed source-name
spellings stay private in portfolio evidence; ordinary source and consumer
defined-name declarations remain normal defined-name review context. Absolute,
URI, escaping,
malformed/ambiguous package declarations, non-workbook package links,
non-static package-A1 forms, sheet-scoped consumer aliases, and consumer
formula definitions that do not meet the static bridge above (including
dynamic, relative, local-3-D, spilled, explicitly intersected, unresolved, or
tokenizer-failed forms), or cyclic/missing exact alias chains (and a sheet-local
consumer name shadows a same-named workbook alias),
missing/unknown/wrong-scope source locals, bare or sheet-qualified table names,
`@`/`#This Row`, unsupported/missing/colliding source tables, dynamic or
otherwise non-statically-expanded name forms, and 3-D spans with missing,
reversed, non-worksheet, or inconsistent-tab-catalog endpoints remain ordinary
external-link coverage rather than being approximated. A defined-name
declaration change remains its ordinary defined-name review event; `FF079`
roots are changed candidate cells.

The command applies an optional policy independently to every matched workbook,
keeps paths relative in portfolio output, skips transient Office `~$` lock
files, and fails closed for unsupported Excel formats, case-colliding paths,
symlinked paths, over-limit inventories, or an unreadable workbook.
The default bounds are 512 supported workbooks, 32,768 total filesystem entries,
4 GiB of supported-workbook source bytes, 2,000,000 populated retained snapshot
cells per directory, 2,000,000 static local dependency-graph edges per input
(one shared pool for each portfolio side), 1,000,000 formula-defined-name
propagation states per input (again one shared pool for each portfolio side),
100,000 aggregate local change-analysis states, and 32 MiB of rendered artifact text. The
entry budget is enforced before FormulaFence
retains or sorts paths, so non-workbook files, directories, lock files, and
symlinks cannot make a broad CI directory consume an unbounded inventory. After
supported regular files are inventoried, FormulaFence sums their observed source
sizes before opening any snapshot; the source-byte budget bounds the aggregate
private-copy and reader workload for each side independently. During comparison,
FormulaFence then counts the actual cells in retained immutable snapshots and
stops as soon as that side exceeds its snapshot-cell budget, before opening a
later workbook. Snapshot construction also caps retained static local
dependency-graph records: direct and range dependencies plus additional
legacy-CSE or observed dynamic-array output aliases. This prevents a compact
formula-defined name from fanning out across many callers without changing how
FormulaFence treats a range as one compact record. Tune the separate limits with
`--max-workbooks`,
`--max-inventory-entries`, `--max-portfolio-source-bytes`,
`--max-portfolio-snapshot-cells`, `--max-dependency-edges`,
`--max-formula-defined-name-states`,
`--max-change-analysis-states`, and
`--max-report-bytes`.
The recursive walk uses direct directory enumeration and treats any unreadable
subdirectory as a fail-closed portfolio error; it never silently omits a branch
from the reviewed directory coverage.
For every retained workbook, FormulaFence also records the observed regular-file
identity and state before comparison. Its later private snapshot read verifies
that observation and refuses a late in-place rewrite, regular-file replacement,
or symlink replacement as redacted `FF078` incomplete evidence rather than
following a newly redirected path.
Local impact analysis has its own 100,000-state budget, configurable with
`--max-change-analysis-states`. It applies to `diff` and `check`, and one shared
pool spans every matched workbook in `portfolio`: each changed source and each
static local dependent it reaches consumes a state. This prevents a broad edit
set from multiplying the per-source traversal limit into unbounded CI work or
report data; an exhausted pool returns exit code `2` before a partial impact
report can imply complete evidence. FormulaFence reconstructs only the bounded
set of serialized shortest-path samples, not every reachable path prefix.
`profile` has a separate 100,000-record in-memory inventory ceiling by default,
configurable with `--max-profile-records`. It counts every public profile-list
record—including nested table columns, control ranges, token/function entries,
and dynamic-array references—before FormulaFence builds the new profile object.
An overage returns exit code `2` before output rendering or publication; raise
the positive limit only when a reviewer intentionally needs the complete known
inventory. This profile-state boundary is independent from source-reader and
rendered-artifact limits.
Rendered artifact size has a separate 32 MiB UTF-8 ceiling, configurable with
`--max-report-bytes`. It applies to `profile` JSON/Markdown and to JSON,
Markdown, HTML, and SARIF from `diff`, `check`, and `portfolio`; an overage
returns exit code `2` before FormulaFence writes or replaces the requested
output path. JSON/SARIF count incremental encoder chunks, Markdown streams each
line, and HTML counts each escaped review entry, so a compact but repetitive
workbook cannot inflate into an unbounded CI artifact. Set a larger positive
value only when a reviewer intentionally needs the corresponding complete
artifact.
Cross-workbook traversal has a separate global bound of 100,000 source-to-node
graph states, configurable with `--max-link-impact`; an exhausted bound emits
critical `FF080` and returns exit code `2` rather than claiming complete impact
evidence. An unreadable workbook still receives a redacted
`FF078` report entry, then returns exit code `2` because the comparison is
incomplete. A newly added or removed unreadable workbook also retains its
known `FF077` / `FFP077` membership evidence. New or removed workbook contents
are represented only by safe
aggregate profiles; no rename matching or whole-file semantic equivalence is
claimed. Place the report outside both input directories: FormulaFence refuses
an output path that could overwrite an inspected input or join a portfolio
while it is being reviewed.

Create `formulafence.yml`:

```yaml
version: 1
rules:
  no_formula_to_value: true
  no_new_external_links: true
  no_external_workbook_link_surface_changes: true
  no_new_broken_references: true
  no_macro_changes: true
  no_xlm_macro_sheet_changes: true
  no_xlm_automatic_macro_binding_changes: true
  no_ribbon_customization_changes: true
  no_office_web_addin_changes: true
  no_chart_definition_changes: true
  no_pivot_table_definition_changes: true
  no_slicer_timeline_cache_changes: true
  no_power_pivot_data_model_changes: true
  no_what_if_data_table_changes: true
  no_scenario_manager_changes: true
  no_filter_visibility_changes: true
  no_ignored_error_changes: true
  no_named_sheet_view_changes: true
  no_custom_workbook_view_changes: true
  no_table_style_control_changes: true
  no_shared_workbook_revision_changes: true
  no_number_format_changes: true
  no_cell_font_changes: true
  no_cell_fill_changes: true
  no_workbook_theme_changes: true
  no_cell_alignment_changes: true
  no_cell_border_changes: true
  no_worksheet_dimension_changes: true
  no_worksheet_display_control_changes: true
  no_worksheet_print_layout_changes: true
  no_formula_cached_result_changes: true
  no_rich_text_run_changes: true
  no_cell_hyperlink_changes: true
  no_worksheet_sparkline_changes: true
  no_xml_mapping_changes: true
  no_digital_signature_changes: true
  no_rich_data_changes: true
  no_custom_data_store_changes: true
  no_legacy_comment_changes: true
  no_threaded_comment_changes: true
  no_worksheet_drawing_shape_changes: true
  no_worksheet_image_changes: true
  no_worksheet_embedded_control_changes: true
  no_new_parser_warnings: true
  no_new_unresolved_references: true
  no_new_dynamic_references: true
  no_new_spill_references: true
  no_new_dynamic_array_output_references: true
  no_new_implicit_intersections: true
  no_array_formula_semantics_changes: true
  no_new_tokenization_failures: true
  no_table_definition_changes: true
  no_data_validation_changes: true
  no_conditional_formatting_changes: true
  no_protection_changes: true
  no_external_data_connection_changes: true
  no_external_link_package_changes: true
  no_external_relationship_changes: true
  no_formula_external_action_changes: true
  no_formula_dde_link_changes: true
  no_python_in_excel_changes: true
  no_office_custom_function_changes: true
  no_unqualified_runtime_function_changes: true
  no_worksheet_code_resource_registration_changes: true
  no_formula_defined_xlm_registration_changes: true
  no_formula_defined_xlm_evaluation_changes: true
  no_formula_defined_xlm_action_changes: true
  no_formula_defined_xlm_get_cell_changes: true
  no_formula_defined_xlm_environment_information_changes: true
  no_formula_environment_information_changes: true
  no_power_query_changes: true
  no_portfolio_membership_changes: true
  no_cross_workbook_impacts: true
  no_3d_reference_scope_changes: true
  max_changed_formulas: 20
  max_downstream_impact: 100

# Cells whose contents must not change without an explicit policy edit.
protected_cells:
  - Dashboard!B12
  - Dashboard!B18

# Optional: limit ordinary cell edits to designated input areas.
allowed_changes:
  - Inputs!B2:B100
```

## What the first release checks

| Capability | What it catches |
| --- | --- |
| Semantic cell diff | Formula/value additions, removals, and changes—not ZIP/XML noise |
| Impact trace | Downstream formula cells and deterministic shortest dependency-path samples, including cross-sheet, static named ranges, formula-defined names, static named `LAMBDA` calls, `LET`/inline-`LAMBDA`, Excel-table, 3-D worksheet references, fixed legacy CSE result members, and currently observed dynamic-array result members |
| Formula-pattern break | An edited formula that no longer matches equal neighboring formulas |
| Formula lint | Conservative blank/error, manual-value, text-marker, and formula-outlier candidates inside a single workbook's copied blocks; narrowly scoped simple aggregate ranges that stop before a contiguous numeric gap; direct static conditional-aggregate range-shape mismatches for `SUMIFS`, `COUNTIFS`, `AVERAGEIFS`, `MAXIFS`, and `MINIFS`, direct static `SUMPRODUCT` array-shape mismatches, direct static `MMULT` inner-dimension mismatches, and direct static `VLOOKUP`/`HLOOKUP` out-of-range literal return indices; explicit direct unlocks on protected formula cells; and formula workbooks explicitly saved with incomplete manual calculation; copied-pattern findings require three local matching peers |
| Portfolio control | Recursive, relative-path workbook matching with per-file semantic reports, explicit additions/removals, bounded static cross-workbook impact evidence, unreadable-file evidence, bounded inventory/traversal, and consolidated JSON/Markdown/HTML/SARIF for CI |
| Workbook controls | Sheet visibility, defined names, Excel-table definitions, AutoFilter/sort/row-and-column visibility including zero-sized dimensions, material worksheet-dimension controls, ignored-error, modern Named Sheet View and legacy Excel Custom View controls, Excel Table Style controls, legacy shared-workbook revision headers/logs, cell-number-format, cell-font, cell-fill, effective cell-alignment, material worksheet-display and worksheet print-layout controls, workbook DrawingML Theme parts/direct image relationships, native worksheet pictures/backgrounds/header-footer watermarks, character-level rich-text runs/phonetic hints, ordinary worksheet-cell hyperlinks, Office 2010 worksheet sparklines, SpreadsheetML XML Maps, OPC package XML-signature envelopes/certificate parts, VBA project signature payloads (classic, Agile, and V3), unexplained stored-formula-result controls, legacy Excel Note/VML Note-shape/threaded-placeholder controls, modern threaded-comment/reply/mention/person controls, and non-chart Worksheet DrawingML regular/connector/group shapes plus bounded SmartArt `xdr:graphicFrame` diagrams and direct Diagram Data image payloads; Excel What-If Data Tables and Scenario Manager definitions, data-validation, conditional-formatting, operational protection, external-data refresh, external-link-package, package-wide external OPC relationships, Python-in-Excel code, namespaced Office custom-function candidates, worksheet and formula-defined code-resource registration calls, formula-defined XLM `REGISTER`/`EVALUATE` calls, XLM macro-sheet programs and automatic-macro bindings, Office RibbonX, Office Web Add-in task-pane/worksheet/in-content bindings, PivotTable views/cache schema/shared items/cached records, Slicer and Timeline cache filter state, embedded Power Pivot/Data Model packages, DrawingML chart definitions/cached series/overlay shapes, modern and legacy-VML worksheet controls/OLE, and Power Query controls; array-formula mode/fixed-output range, static 3-D-reference scope, calculation settings, and VBA payload changes |
| Formula hazards | New external-workbook references and `#REF!` formulas |
| External-workbook link surfaces | Private static ledger across worksheet formulas, defined names, data-validation criteria, and standard/ChartEx chart formulas; catches same-location source or target swaps without evaluating, resolving, or exposing endpoint material |
| Formula external-action and data-provider surfaces | Material stored `HYPERLINK`, `WEBSERVICE`, `IMAGE`, `RTD`, `STOCKHISTORY`, or documented `CUBE*` call changes in cells, formula-defined names, or named `LAMBDA`s, including same-count destination, market-provider, connection, or query swaps, without evaluating formulas or exposing their arguments in the private ledger |
| Direct DDE formula links | Material lexical `application|topic!item` DDE-link changes in worksheet formulas, formula-defined names, or named `LAMBDA`s, including same-count endpoint swaps and static inputs to an invoking named `LAMBDA`, without evaluating a formula, looking up/launching a DDE server, or exposing endpoint material |
| Native workbook/environment-information boundary | Stored native `CELL`, `INFO`, `SHEET`, and `SHEETS` calls and statically visible inputs, including a private all-tab catalog comparison when `SHEET` or `SHEETS()` can observe tab position/count, without evaluating formulas or exposing arguments |
| Python in Excel boundary | Stored Python code/environment, `PY` formula bindings, and statically visible inputs, without loading Python code, executing it, or contacting its cloud runtime |
| Namespaced Office custom-function boundary | Material namespaced formula-call candidates and statically visible inputs, without loading an add-in, executing a formula, or exposing names and arguments in the private ledger |
| Formula-defined XLM registration boundary | Stored legacy XLM `REGISTER` calls in formula-defined names/named `LAMBDA`s and their statically visible inputs, without executing a macro, evaluating a formula, or loading a DLL/XLL |
| Formula-defined XLM cell-information boundary | Stored legacy XLM GET.CELL calls in formula-defined names/named LAMBDAs and their statically visible inputs, without evaluating a call, resolving dynamic references, or simulating Excel state |
| Formula-defined XLM environment-information boundary | Stored legacy XLM GET.WORKBOOK, GET.WORKSPACE, and GET.DOCUMENT calls in formula-defined names/named LAMBDAs and their statically visible inputs, without evaluating a call, resolving dynamic references, or simulating Excel state |
| Coverage changes | New parser warnings, unresolved formula references (including unsupported table syntax), dynamic-reference functions (`INDIRECT`/`OFFSET`), dynamic-array spill references, explicit implicit intersection, and formula-tokenization failures |
| Policy as code | Protected cells, allowed edit areas, bans, and change/impact limits |
| CI output | Deterministic JSON, reviewer-friendly Markdown, and SARIF |

See [the policy reference](docs/policy.md) for the configuration contract and
[the threat model](docs/threat-model.md) for important limits. The
[external validation notes](docs/validation.md) record an independently
maintained financial-model compatibility check.

For Excel tables, FormulaFence statically resolves a table name, a single
column or contiguous column range, and the `#All`, `#Data`, `#Headers`, and
`#Totals` regions. It also resolves row-scoped `@` / `#This Row` references
when their formula location proves the row: unqualified forms such as
`[@[Sales Amount]]` and `[Sales Amount]` only inside a table data cell, plus
qualified forms such as `Sales[@Amount]` and
`Sales[[#This Row],[Amount]:[Rate]]` anywhere on that table's data row. Header,
total, cross-sheet, ambiguous, and complex bracket-escape cases remain explicit
coverage notes. The supported subset follows Excel's documented
[structured-reference semantics](https://support.microsoft.com/en-us/excel/using-structured-references-with-excel-tables).

FormulaFence expands a formula-defined name when its complete definition can be
statically resolved to internal dependencies or a constant. It follows nested
workbook and sheet-local definitions, so a formula such as
`=DiscountedValue`, where `DiscountedValue` uses `TaxRate`, receives edges to
the underlying inputs. Relative references, cycles, external links,
`INDIRECT`/`OFFSET`, 3-D spans inside a name definition, and syntax the formula
tokenizer cannot inspect remain unresolved at the use site rather than being
treated as a safe dependency. This builds on Excel's documented support for
[names that represent formulas and constants](https://support.microsoft.com/en-us/excel/names-in-formulas).

FormulaFence recognizes the lexical names introduced by inline `LET` and
`LAMBDA` expressions, including nested lambdas supplied to functions such as
`REDUCE`. Those local variables no longer masquerade as unresolved workbook
names; the static A1, named, and table references around them remain graph
edges. The implementation follows Excel's documented
[LET scope](https://support.microsoft.com/en-us/excel/functions/let-function)
and [LAMBDA parameter syntax](https://support.microsoft.com/en-us/excel/functions/lambda-function).

FormulaFence traces the static anchor behind a direct internal spilled-array
reference such as `=SUM(Inputs!B2#)`, including Excel-compatible OOXML spelling
such as `=SUM(_xlfn.ANCHORARRAY(Inputs!B2))`. That creates an edge from the
anchor cell to the consumer, so changes to the anchor formula and its statically
visible inputs reach the consumer. Profiles retain every spill token, and a
diff emits `FF015`; `no_new_spill_references` can make new instances a CI
failure. FormulaFence does not invent the spill extent or a dependency on every
possible blocking cell. External, 3-D, range, named, implicit-intersection, and
malformed spill forms remain explicit coverage limits. A formula-defined name
that contains a spill reference also remains unresolved at its call site rather
than hiding that dynamic boundary behind a static name edge.

FormulaFence also inventories Excel's explicit implicit-intersection operator:
the display form `@A1:A3`, function forms such as `@INDEX(...)`, and persisted
OOXML `_xlfn.SINGLE(...)`. When a direct static A1 cell or range has one
unambiguous documented row/column intersection, it adds that single-cell
edge—for example, `=_xlfn.SINGLE(Inputs!B2:B4)` in row 3
depends on `Inputs!B3`, not every input in the range. Other expressions retain
their visible static input edges without evaluating Excel. Profiles record each
recognized explicit use; a new use emits `FF017`, and
`no_new_implicit_intersections` can require review in CI. This is distinct from
Excel-table `[@Column]` syntax, which follows the table-specific resolver.
Formula-defined names containing explicit implicit intersection remain
unresolved at call sites because their selected cell depends on the caller's
position. The scope follows Microsoft's
[implicit-intersection guidance](https://support.microsoft.com/en-us/excel/implicit-intersection-operator)
and [XlsxWriter's documentation](https://xlsxwriter.readthedocs.io/working_with_formulas.html)
that persisted `@` behavior uses `SINGLE()`.

FormulaFence also recognizes a multi-cell **legacy CSE** array formula as a
fixed output range when its OOXML array anchor has no dynamic-array cell
metadata. It does not inflate that range into virtual cells. Instead, if an
ordinary formula reads a member such as `Model!B2`, FormulaFence adds a compact
alias from the CSE anchor (for example `Model!B1`) to that consumer. Thus a
change to an input of `{=LEN(Inputs!A1:A3)}` can reach a downstream `=B2*10`
calculation even though Excel stores the formula only at `B1`. The profile
lists each declared fixed range. An OOXML dynamic-array marker (`cm` resolving
to `XLDAPR`/`fDynamic`) records the current serialized output range as an
**observed** spill surface. If a static formula reads a non-anchor member such
as `Model!B2`, FormulaFence adds the same compact anchor-to-consumer graph edge
and records the relationship in the profile. It never calls that range fixed:
Excel can grow, shrink, or block a spill during recalculation. A newly observed
output-member relationship emits `FF019`; enable
`no_new_dynamic_array_output_references` for `FFP019`. Before it reads the
canonical `xl/metadata.xml` mapping that identifies `XLDAPR`/`fDynamic`,
FormulaFence streams it under a 16 MiB and 32,768-element boundary. A
successfully streamed overage stays out of the tree parser, leaves array
formulas unclassified with no aliases, and emits a visible parser-coverage
warning. The direct worksheet `c`/`f` binding pass is streamed too, so it does
not retain a second complete worksheet tree. Private opaque fallback evidence
keeps a material unavailable-metadata change reviewable as `FF018` without
disclosing XML or metadata material. FormulaFence emits `FF018` when a
legacy-CSE or dynamic-array formula is added, removed, or changes mode, when a
legacy CSE fixed output range changes, or when raw array-formula metadata
coverage materially changes; enable
`no_array_formula_semantics_changes` for `FFP018`. This follows Microsoft's
[dynamic-versus-legacy array guidance](https://support.microsoft.com/en-us/excel/dynamic-array-formulas-and-spilled-array-behavior)
and [XlsxWriter's documented CSE/dynamic serialization behavior](https://xlsxwriter.readthedocs.io/working_with_formulas.html).

FormulaFence inventories **data-validation controls** as reviewable workbook
semantics. A validation can restrict a whole column, select a list from another
sheet, enforce a custom formula, hide a list arrow, or change whether Excel
shows a prompt or blocks an invalid entry. The profile records each target range
and its effective control settings without exposing validation formulas or
prompt/error text; a local change report includes the full before/after rule.
It normalizes OOXML defaults such as `operator=between` and `errorStyle=stop`,
the optional leading `=` in a criterion, and equivalent grouping of identical
target rules, so equivalent writers do not create noise. Any changed validation
control emits `FF020`; enable
`no_data_validation_changes` for `FFP020`. FormulaFence does not evaluate a
validation formula or predict whether a user entry will be accepted. The scope
follows Microsoft's [data-validation guidance](https://support.microsoft.com/en-US/Excel/get-started/apply-data-validation-to-cells)
and [openpyxl's documented range model](https://openpyxl.readthedocs.io/en/stable/validation.html).

FormulaFence also inventories **conditional-formatting controls** directly from
worksheet OOXML, because a priority move, `Stop If True` toggle, exception
formula, or changed red/green threshold can alter what reviewers see without
altering a normal formula cell. It records compact target ranges, worksheet-wide
precedence, criteria, rule flags, differential styles, color scales, data bars,
and icon sets. Differential styles are resolved from their actual OOXML rather
than their unstable `dxfId`, while omitted boolean defaults, optional leading
`=` formula spelling, priority-number gaps, and extension GUID links are
normalized to avoid writer-only noise. Excel-2010 extension fragments are retained as
opaque local evidence even when the parser cannot model them. Profiles redact
criteria, text rules, and raw style/extension XML; local change reports retain
the full before/after control. Any change emits `FF021`; enable
`no_conditional_formatting_changes` for `FFP021`. FormulaFence does not
calculate a conditional-formatting formula, resolve a relative rule for every
target cell, or predict the final visual display. The scope follows Microsoft's
[conditional-formatting precedence guidance](https://support.microsoft.com/en-us/excel/use-conditional-formatting-to-highlight-information-in-excel)
and [OOXML conditional-formatting model](https://learn.microsoft.com/en-us/office/open-xml/spreadsheet/working-with-conditional-formatting).

FormulaFence also inventories **operational protection controls** directly from
OOXML: workbook structure/windows/revision locks; worksheet and dialog-sheet
permissions; chart-sheet content/object locks; protected-range target areas;
and explicit cell, row, and column `locked`/`hidden` assignments when a normal
sheet is protected. This catches a newly editable input or newly hidden formula
without expanding styled rows or columns into cells. It normalizes effective
OOXML defaults for sheet actions, so writers that omit or explicitly serialize
the same defaults do not create a diff. The profile and every report redact
legacy verifiers, password hashes, salts, protected-range names, and security
descriptors; FormulaFence retains only private comparison fingerprints plus
safe presence metadata. Any change emits `FF022`; enable
`no_protection_changes` for `FFP022`. These controls are not file encryption,
identity enforcement, or a security guarantee: workbook and worksheet
protection are an operational review surface, and FormulaFence does not decide
who can successfully edit a file or fully emulate Excel's style cascade. The
scope follows Microsoft's [workbook protection guidance](https://support.microsoft.com/en-US/Excel/protect-a-workbook),
[worksheet protection guidance](https://support.microsoft.com/en-us/excel/protect-a-worksheet),
and [protection-and-security overview](https://support.microsoft.com/en-US/Excel/protection-and-security-in-excel).

FormulaFence also inventories **external-data refresh controls** directly from
OOXML before a reader library can omit them: workbook-wide external-link and
refresh-on-open flags; connection refresh schedules, background behavior,
credential/cache flags, source-kind metadata, connection-file reload policy,
and parameter-triggered refreshes; linked query-table refresh/growth behavior;
and pivot-cache source and refresh controls. This catches a workbook that can
change its inputs on open or during use even when its ordinary formulas did not
change. Profiles and reports intentionally omit connection names and
descriptions, paths, URLs, connection strings, commands, parameter values, SSO
identifiers, cached records, and opaque extension XML; private fingerprints
still expose a material source or identity change. Any change emits `FF023`;
enable `no_external_data_connection_changes` for `FFP023`. FormulaFence never
opens a connection or refreshes data, does not assess source trust or actual
returned values, or calculate a PivotTable report. The scope follows Microsoft's
[external-data refresh guidance](https://support.microsoft.com/en-us/excel/refresh-an-external-data-connection-in-excel)
and the [SpreadsheetML Connections part](https://c-rex.net/samples/ooxml/e1/Part1/OOXML_P1_Fundamentals_Connections_topic_ID0EQLGK.html).

Before FormulaFence privately materializes a raw `xl/connections*.xml` part, it
streams its XML structure with a 32,768-element per-part and 65,536-element
Connections-scan limit, alongside 16 MiB per-part, 64 MiB total, and 512-part
read limits. A successfully streamed structural overage is retained as
private opaque coverage evidence (`FF010`) and remains diff-visible through
`FF023`; malformed XML keeps its ordinary parser diagnostic. This bound applies
to raw Connections XML, not to a claim that every external-data package reader
has the same structural limit.

FormulaFence applies the same bounded streaming pattern to selected raw
query-table XML targets reached from worksheet or table `queryTable`
relationships (normally `xl/queryTables/queryTable*.xml`): 32,768 elements per
part, 65,536 across the shared query-table scan, 16 MiB per part, 64 MiB total,
and 512 parts. A private template cache means one reused part is neither
reparsed nor recursively canonicalized once per worksheet binding. A
successfully streamed structural overage produces private opaque coverage
evidence (`FF010`) that remains diff-visible through `FF023`; malformed XML
keeps its ordinary parser diagnostic. This allocation boundary is specific to
the raw query-table reader, not a claim that every table or external-data
reader has the same structural limit. The part shape follows the
[SpreadsheetML `queryTable` definition](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.querytable?view=openxml-3.0.1).

FormulaFence separately inventories raw **external-link packages**
(`xl/externalLinks/externalLink*.xml`). It recognizes external-workbook,
DDE, and OLE definitions; binds `externalReferences` declarations to their
package parts privately; and compares their source relationship, definition,
cached-data, item-behavior, and unmodelled-extension material. Profiles expose
only safe counts—never workbook targets, sheet/defined names, DDE services or
items, OLE program/item names, or cached values. A material package change
emits `FF025`; enable `no_external_link_package_changes` for `FFP025`.
FormulaFence never follows or executes these links, establishes source trust,
or infers returned data. The package shape follows the
[SpreadsheetML `externalLink` definition](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.externallink?view=openxml-3.0.1).

Before FormulaFence privately parses a selected raw external-link XML part, it
streams `xl/externalLinks/externalLink*.xml` and the direct `.rels` parts used
by the external-link inventory or package-indexed resolver through a
32,768-element per-part and 65,536-element shared-scan limit. The same boundary
also allows 16 MiB per part, 64 MiB in aggregate, and 512 parts. A successfully
streamed structural overage produces private opaque coverage evidence (`FF010`)
that remains diff-visible through `FF025`; malformed XML keeps its ordinary
parser diagnostic. This allocation boundary is specific to these raw
external-link readers, not a claim that every package relationship reader has
the same structural limit.

FormulaFence also keeps a separate private **static external-workbook
link-surface ledger**. It recognizes literal external endpoints persisted in
worksheet formulas, workbook or sheet-local defined names, data-validation
criteria, and standard DrawingML/ChartEx chart formula elements. This catches a
source or target swap at the same cell or object—an intentional gap in
`no_new_external_links`, which only guards a newly external worksheet formula
location. A material ledger change emits `FF081`; enable
`no_external_workbook_link_surface_changes` for `FFP081`.

Profiles and the `FF081` finding expose only surface and endpoint counts. Source
paths, workbook/sheet/name identities, formulas, validation ranges, and chart
part identities remain inside this ledger's private signature. Chart parts with
unavailable formula coverage make that guard fail closed. FormulaFence does not
evaluate a formula, open or resolve a source workbook, trust a cache, or infer
text-built references. The boundary follows the SpreadsheetML
[data-validation formula rules](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/20ed0abd-113f-4b8a-8de3-c68e733a300a)
and DrawingML chart [`c:f` formula element](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.charts.numberreference?view=openxml-3.0.1).

FormulaFence also keeps a bounded **package-wide external-relationship ledger**.
It inspects the root and every canonical OPC `.rels` part for
`TargetMode="External"`, including relationship sources that no specialized
workbook feature reader recognizes. That closes the gap where an otherwise
ordinary `.xlsx` can gain a remote hyperlink, image, or opaque relationship
outside the known `externalLink`, drawing, add-in, or worksheet-markup paths.
The ledger follows the [Open Packaging Conventions relationship
model](https://learn.microsoft.com/en-us/previous-versions/windows/desktop/opc/open-packaging-conventions-overview)
and its [`TargetMode`](https://learn.microsoft.com/en-us/dotnet/api/system.io.packaging.packagerelationship.targetmode)
semantics, but never resolves, opens, fetches, or rates a target.

Profiles and reports expose only counts for relationship parts, sources,
targets, and hyperlink/image/other target types. Source part names, relationship
types and IDs, targets, unknown attributes, and raw XML stay inside private
comparison signatures. Writer-chosen relationship-ID rewrites normalize away;
a material endpoint, type, source, or malformed-coverage change emits `FF063`.
Enable `no_external_relationship_changes` for `FFP063`. Relationship XML reads
are bounded to 16 MiB per part, 64 MiB per workbook, and 512 parts. Duplicate,
orphaned, malformed, unsafe, oversized, unreadable, or over-budget relationship
metadata remains explicit coverage evidence rather than silently disappearing.

FormulaFence also keeps a private **formula external-action and data-provider
ledger** for stored `HYPERLINK`, `WEBSERVICE`, `IMAGE`, `RTD`, `STOCKHISTORY`,
and the documented Cube family (`CUBEKPIMEMBER`, `CUBEMEMBER`,
`CUBEMEMBERPROPERTY`, `CUBERANKEDMEMBER`, `CUBESET`, `CUBESETCOUNT`, and
`CUBEVALUE`) calls, including their `_xlfn.` compatibility spelling and calls
held in formula-defined names or named `LAMBDA` bodies. Microsoft documents
that `HYPERLINK` can use a text link location or a cell reference, that
[`WEBSERVICE`](https://support.microsoft.com/en-US/Excel/functions/webservice-function)
calls a web-service URL, that
[`IMAGE`](https://support.microsoft.com/en-us/excel/functions/image-function)
uses an HTTPS image source, and that
[`RTD`](https://support.microsoft.com/en-us/excel/functions/rtd-function)
retrieves data through a COM-automation provider. [`STOCKHISTORY`](https://support.microsoft.com/en-us/office/stockhistory-function-1ac8b5b3-5f62-4d94-8ab8-7504ec7239a8)
retrieves financial history, while [`CUBEVALUE`](https://support.microsoft.com/en-us/excel/functions/cubevalue-function)
and [`CUBESET`](https://support.microsoft.com/en-us/excel/functions/cubeset-function)
bind stored expressions to a workbook Cube connection and can retrieve data
from a server or offline cube. `HYPERLINK` calls are all inventoried—including
known in-workbook links—because their destination can be computed dynamically
and a later formula edit can retarget a reviewer.

Profiles and `FF064`/`FFP064` details expose only formula-cell,
formula-defined-name, and per-function counts. Private signatures retain the
stored cell and relevant name-definition material so a destination, market
symbol, provider, connection, query, argument, call-location, or
name-definition change stays visible even when public counts do not move.
FormulaFence also raises `FF064` when an ordinary cell edit can reach an
invoking action/provider formula through its static dependency graph, catching
`=HYPERLINK(A1, ...)`- or `=STOCKHISTORY(A1, ...)`-style input retargeting
without evaluating `A1`. Dynamic or unresolvable arguments remain explicit
parser-coverage limits. The ordinary semantic cell diff intentionally continues
to show changed formulas to reviewers; its general payload is not a redacted
formula vault. `--redact-formula-external-actions` is an opt-in, output-only
sharing boundary for the action/DDE formulas and exact static input evidence
described above; `--redact-external-workbook-links` independently covers
literal external-workbook link material. Neither option turns the general
payload into a redacted formula vault. FormulaFence does not
calculate, resolve, fetch, open, click, follow, authenticate to, query, or
execute any formula action/provider, and it
does not decide whether a dynamic `HYPERLINK` destination is local or remote.
Enable
`no_formula_external_action_changes` to block this boundary in CI.

FormulaFence also keeps a separate private **direct DDE formula-link ledger**.
Windows documents Excel's DDE formula form as
`='Quote'|'NYSE'!ZAXX`: application and topic are separated by a pipe outside
quoted text, followed by an item boundary after `!`. FormulaFence recognizes only that
conservative lexical shape (including quoted components and command-style
missing-item forms), so an ordinary quoted sheet name such as
`='cmd|/C calc'!A0` is not misclassified. It follows the signal through
formula-defined names and named `LAMBDA`s without evaluating either one.

The public profile and `FF074`/`FFP074` details expose only formula-cell,
DDE-link, and relevant formula-defined-name counts. Private signatures retain
the stored formula and relevant name chain, so a service/topic/item or
same-count definition change remains visible without emitting it. An ordinary
cell change that reaches an invoking named `LAMBDA` through the static graph is
also reviewable. Direct-DDE raw formula syntax can be outside the underlying
tokenizer grammar; FormulaFence records its dedicated ledger before that parser
boundary and preserves the normal tokenization-coverage signal separately.
It never evaluates a formula, resolves a service/topic/item, looks up or starts
a DDE server, sends a DDE command, or determines what local Trust Center
settings allow. Microsoft's [DDE overview](https://learn.microsoft.com/en-us/windows/win32/dataxchg/about-dynamic-data-exchange)
documents the formula and the application/topic/item conversation model; Excel's
[DDE security settings](https://learn.microsoft.com/en-us/troubleshoot/microsoft-365-apps/excel/security-settings)
describe separate server lookup and server-launch controls. Raw OOXML
`externalLink` DDE/OLE package definitions remain the separate `FF025` boundary
because package metadata can exist without a direct formula. Enable
`no_formula_dde_link_changes` to block this surface in CI.

FormulaFence separately inventories **Python in Excel** workbooks. Microsoft
documents that `PY` has static Python source and runs through a Microsoft Cloud
runtime; related workbook package material can also retain script state.
FormulaFence recognizes stored `PY` formula spellings, the documented 2023
`python.xml` package contract, and the separately stored 2022
`pythonScripts.xml` compatibility contract before the ordinary reader can
discard that material. It privately fingerprints bounded raw Python XML,
including code, environment definitions, script ordering, and extensions, then
compares the stored PY formula binding separately. If both package contracts
are present, FormulaFence inventories each stored part independently rather
than assuming that they agree or choosing a runtime implementation.

Profiles and `FF065`/`FFP065` details expose only package, formula-cell,
function-call, script, environment, initialization, and coverage counts. Python
source, environment IDs, script indexes, formula arguments, formula locations,
and raw XML remain private. A source-code/environment/package change, a changed
PY formula binding, or an ordinary cell change that reaches a PY formula through
the static dependency graph emits `FF065`; the latter catches a static source
such as `=_xlfn._xlws.PY(0,0,A1)` without interpreting `A1` or Python code.
Relationship-ID-only rewrites normalize. Missing, malformed, unbound,
oversized, unreadable, or over-budget package material remains visible as a
coverage gap. The scan is bounded to 16 MiB per XML part, 64 MiB per workbook,
and 512 parts. Before it materializes package XML, FormulaFence streams each
part under a 32,768-element limit and the complete Python-in-Excel scan under a
65,536-element limit. A successfully parsed structural overage remains visible
as `FF010`/`FF065` coverage evidence rather than allocating its private tree.

FormulaFence does not parse Python source as Python, run code, evaluate `PY`,
resolve its result, contact Microsoft Cloud, or verify runtime package support.
Dynamic or unresolved formula inputs stay explicit parser-coverage limits. The
ordinary semantic diff deliberately retains changed PY formulas and values for
local review; its general payload is not a redacted code vault.
`--redact-python-in-excel` is an opt-in, output-only sharing boundary for
direct PY source and exact static input evidence, while
`no_python_in_excel_changes` blocks the boundary in CI. This scope follows
Microsoft's [PY function reference](https://support.microsoft.com/en-us/excel/functions/py-function),
[Python in Excel introduction](https://support.microsoft.com/en-US/Excel/python/introduction-to-python-in-excel),
and the [Python part](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/151e4bcd-90a0-4d82-8b98-f16bf273e4ff)
definition.

FormulaFence also inventories **namespaced Office custom-function call
candidates**. Microsoft documents Office Add-in custom functions as JavaScript
or TypeScript functions surfaced in Excel with a manifest namespace—for example
`=CONTOSO.ADD(10,200)`—and documents that they can request or stream data from
the web. A formula alone does not embed that add-in's manifest, code, identity,
or runtime, so FormulaFence intentionally treats the formula shape as a review
candidate rather than proof that an add-in is installed or runnable.

The direct-call classifier accepts only a namespaced callable that is not a
known native dotted Excel function or a workbook-defined name. It excludes
OOXML `_xlfn.` / `_xlws.` compatibility forms. Unqualified VBA, COM, and XLL
UDF-shaped calls are intentionally handled by the separate `FF075` boundary.
FormulaFence separately propagates a
candidate stored inside a formula-defined name or named `LAMBDA` body to the
worksheet formulas that invoke that definition. Profiles and `FF066`/`FFP066`
details show only formula-cell, call, namespace, and relevant formula-defined
name counts. Candidate names,
namespaces, cells, formulas, and arguments remain private, so a same-count
call or argument change stays reviewable without publishing the add-in surface.

FormulaFence also emits `FF066` when an ordinary cell edit reaches a candidate
call through its static dependency graph. That catches a stored source such as
`=CONTOSO.GETDATA(A1)` without interpreting `A1`, evaluating the call, resolving
the candidate to an add-in, loading a manifest, or contacting a runtime.
Dynamic or unresolved inputs remain explicit coverage limits. The ordinary
semantic diff deliberately keeps reviewer context unless the opt-in
`--redact-office-custom-functions` sharing boundary is enabled; that output
mode hides direct custom-function material, exact static input evidence, and
conservatively changed relevant defined-name bodies without changing the
comparison or policy result. Enable `no_office_custom_function_changes` to
block this boundary in CI. This scope
follows Microsoft's [custom-functions overview](https://learn.microsoft.com/en-us/office/dev/add-ins/excel/custom-functions-overview),
[tutorial](https://learn.microsoft.com/en-us/office/dev/add-ins/tutorials/excel-tutorial-create-custom-functions),
and [external-data guidance](https://learn.microsoft.com/en-us/office/dev/add-ins/excel/custom-functions-web-reqs).

FormulaFence also inventories **unqualified runtime-function candidates**.
Microsoft documents that installed add-ins can expose user-defined and
Automation functions, that VBA custom functions can be made available through
an add-in, and that registered XLL worksheet functions can be called anywhere
a built-in function can be called. A normal workbook formula does not identify
which provider—if any—will resolve a bare unknown call, so FormulaFence treats
it as a review candidate rather than proof that code exists, is trusted, or can
run.

The classifier accepts only a bare identifier such as `=MYUDF(A1)`, excludes
workbook-defined names and local `LET`/`LAMBDA` bindings, and uses a pinned
catalogue of native Excel function spellings so ordinary calls such as `SUM`,
`XLOOKUP`, `VSTACK`, `FIELDVALUE`, and `PY` are not candidates. Qualified
namespaced calls remain under `FF066`; direct `REGISTER.ID` has its dedicated
`FF067` ledger. Candidate calls stored inside formula-defined names and named
`LAMBDA` bodies are propagated to their invoking formulas. The native catalogue
is intentionally versioned with FormulaFence rather than inherited from a
third-party parser, so a newly introduced Excel function may be conservatively
reported until its spelling is added in a FormulaFence release.

A stored candidate definition is also compared even if no worksheet formula
currently invokes it. This catches dormant runtime bindings without pretending
that they are active.

Profiles and `FF075`/`FFP075` details expose only formula-cell, call, and
relevant formula-defined-name counts. Candidate names, formula text, arguments,
provider identities, and locations stay in private signatures; same-count call
or definition changes remain reviewable. `FF075` also covers an ordinary cell
edit that statically reaches a candidate. FormulaFence does not evaluate a
formula, resolve or load VBA, COM/Automation, XLL, or any registered provider,
inspect host trust settings, or execute code. Dynamic and unresolved inputs
remain explicit coverage limits. Enable
`no_unqualified_runtime_function_changes` to block this boundary in CI. This
scope follows Microsoft's [native function catalogue](https://support.microsoft.com/en-us/office/excel-functions-alphabetical-b3944572-255d-4efb-bb96-c6d90033e188),
[installed UDF guidance](https://support.microsoft.com/en-us/excel/user-defined-functions-that-are-installed-with-add-ins-reference),
[VBA custom-function guidance](https://support.microsoft.com/en-us/excel/create-custom-functions-in-excel),
and [XLL registration/call guidance](https://learn.microsoft.com/en-us/office/client-developer/excel/accessing-xll-code-in-excel).

FormulaFence separately inventories **worksheet code-resource registrations**
through `REGISTER.ID`. Microsoft's [function reference](https://support.microsoft.com/en-us/office/register-id-function-f8f0af0f-fd66-4704-a0f2-87b27b175b50)
documents that it returns a DLL or code-resource registration ID and registers
the resource when needed; unlike `REGISTER`, it can be used from a worksheet.
That makes a stored `REGISTER.ID` expression a distinct boundary from an
ordinary namespaced add-in candidate.

The `REGISTER.ID` ledger propagates calls inside formula-defined names and
named `LAMBDA` bodies to their invoking worksheet formulas. Profiles and
`FF067`/`FFP067` details expose only formula-cell, call, and relevant
formula-defined-name counts. Module paths, procedure names, type strings,
cells, formulas, arguments, and name identities remain private, including for
same-count changes. `FF067` also covers ordinary cell edits that statically
reach a registration expression. FormulaFence does not evaluate a formula,
resolve a path, load a DLL/XLL, inspect the host's trust settings, or determine
whether registration succeeds. Dynamic or unresolved inputs remain explicit
coverage limits. Enable `no_worksheet_code_resource_registration_changes` to
block this boundary in CI.

This does not reinterpret raw XLM macro-sheet programs: Microsoft's
[`CALL` reference](https://support.microsoft.com/en-us/office/call-function-32d58445-e646-4ffd-8d5e-b45077a5e995)
places `CALL` on macro sheets, and FormulaFence continues to guard complete raw
XLM macro-sheet material through its separate `FF026` boundary below.

FormulaFence also keeps a distinct **formula-defined XLM registration ledger**
for legacy `REGISTER` calls stored in formula-defined names and named `LAMBDA`
bodies. Microsoft's [`xlfRegister` Form 1 reference](https://learn.microsoft.com/en-au/office/client-developer/excel/xlfregister-form-1)
documents the XLM `REGISTER` equivalent for DLL functions or commands and
documents macro types callable from a defined-name definition; its
[`Form 2` reference](https://learn.microsoft.com/en-au/office/client-developer/excel/xlfregister-form-2)
documents the XLL load-and-activate form. FormulaFence therefore does not
broaden this into a guess about arbitrary formulas: direct worksheet
`REGISTER` calls and raw XLM macro-sheet program XML remain outside this narrow
stored-definition boundary.

The ledger propagates stored registrations through nested and sheet-local
formula names to their invoking cells. Profiles and `FF068`/`FFP068` details
expose only invocation-cell, call, and relevant formula-defined-name counts.
Module paths, procedure names, type strings, formulas, arguments, cells, and
name identities remain private; same-count definition or invocation changes,
uninvoked stored names, and ordinary static input edits remain reviewable using
private signatures. FormulaFence does not evaluate a formula, execute an XLM
macro, resolve a path, load a DLL/XLL, or inspect host trust settings. Dynamic
or unresolved inputs remain explicit coverage limits. Enable
`no_formula_defined_xlm_registration_changes` to block this boundary in CI.
For a shared artifact, the separate output-only
`--redact-formula-defined-xlm-registrations` mode hides direct stored
`REGISTER` material, exact changed static inputs, and changed private
name-chain evidence without changing comparison, policy, or exit status.

FormulaFence also keeps a separate **formula-defined XLM expression-evaluation
ledger** for `EVALUATE` calls stored in formula-defined names and named
`LAMBDA` bodies. Microsoft's [Excel expression-evaluation
reference](https://learn.microsoft.com/en-us/office/client-developer/excel/excel-worksheet-and-expression-evaluation)
identifies `EVALUATE` as an XLM function that reduces a valid character string
to a worksheet value. A stored call can therefore make the expression Excel
calculates differ from the static formula text surrounding it. FormulaFence
does not claim to calculate that expression or turn its runtime-generated text
into static dependencies.

The ledger propagates stored calls through nested and sheet-local formula names
to invoking cells. Profiles and `FF069`/`FFP069` details expose only
invocation-cell, call, and relevant formula-defined-name counts; expression
text, formulas, arguments, cells, and name identities remain private.
Same-count definition or invocation changes, uninvoked stored names, and
ordinary edits that reach a stored call through a statically visible argument
edge remain reviewable through private signatures. Formula text parsed by
`EVALUATE` is not re-tokenized, so dependencies inside it remain an explicit
coverage limit. Direct worksheet `EVALUATE` calls and raw XLM macro-sheet parts
are deliberately outside this narrow stored-definition boundary. Enable
`no_formula_defined_xlm_evaluation_changes` to block this boundary in CI.
For a shared artifact, the separate output-only
`--redact-formula-defined-xlm-evaluations` mode hides direct stored `EVALUATE`
material, exact changed static inputs, and changed private name-chain evidence
without evaluating or parsing runtime-generated expression text, or changing
comparison, policy, or exit status.

FormulaFence also keeps a separate **formula-defined XLM action and
event-dispatch ledger** for the selected legacy calls `CALL`, `EXEC`,
`EXECUTE`, `RUN`, `SEND.KEYS`, `ON.DATA`, `ON.DOUBLECLICK`, `ON.ENTRY`,
`ON.KEY`, `ON.RECALC`, `ON.SHEET`, `ON.TIME`, and `ON.WINDOW` when they are
stored in formula-defined names and named `LAMBDA` bodies. Microsoft's [Excel C
API reference](https://learn.microsoft.com/en-us/office/client-developer/excel/programming-with-the-c-api-in-excel)
describes XLM command-equivalent functions and event traps including `ON.ENTRY`
and `ON.TIME`; its [DLL-access
guidance](https://learn.microsoft.com/en-us/office/client-developer/excel/how-to-access-dlls-in-excel)
documents `CALL` and `REGISTER` as XLM macro-sheet routes to DLL functions or
commands. Microsoft also documents `EXEC` as a runtime-risky XLM trigger in
its [XLM AMSI analysis](https://www.microsoft.com/en-us/security/blog/2021/03/03/xlm-amsi-new-runtime-defense-against-excel-4-0-macro-malware/).

This is intentionally a finite, stored-definition inventory rather than an
attempt to interpret all XLM commands. The ledger propagates selected calls
through nested and sheet-local formula names to invoking cells. Profiles and
`FF073`/`FFP073` details expose only invocation-cell, call, and relevant
formula-defined-name counts; function targets, handler names, formulas,
arguments, cells, and name identities remain private. Same-count definition or
invocation changes, uninvoked stored names, and ordinary static input edits are
reviewable through private signatures. FormulaFence does not evaluate a
formula, resolve an action target or event handler, load a DLL, send DDE, run a
macro or program, or infer whether an action succeeds. Direct worksheet action
calls and raw XLM macro-sheet parts are deliberately outside this narrow
stored-definition boundary. Enable `no_formula_defined_xlm_action_changes` to
block this boundary in CI.
For a shared artifact, the separate output-only
`--redact-formula-defined-xlm-actions` mode hides direct stored selected-action
material, exact changed static inputs, and changed private name-chain evidence
without resolving a target or handler, executing an action, or changing
comparison, policy, or exit status.

FormulaFence also keeps a separate **formula-defined XLM cell-information
ledger** for GET.CELL calls stored in formula-defined names and named LAMBDA
bodies. Microsoft identifies
[GET.CELL / xlfGetCell](https://learn.microsoft.com/en-us/office/client-developer/excel/programming-with-the-c-api-in-excel)
as an XLM information function. A stored call can inspect cell information
outside an ordinary formula's visible value dependencies, including display and
format-related state; FormulaFence does not claim to evaluate that call or
determine the requested information type.

The ledger propagates stored calls through nested and sheet-local formula names
to invoking cells. Profiles and FF070/FFP070 details expose only
invocation-cell, call, and relevant formula-defined-name counts; information
types, references, formulas, arguments, cells, and name identities remain
private. Same-count definition or invocation changes, uninvoked stored names,
and ordinary edits that reach a stored call through a statically visible
argument edge remain reviewable through private signatures. FormulaFence does
not resolve dynamic references, render formatting or display text, inspect
comments/protection, or simulate Excel state. Direct worksheet GET.CELL calls
and raw XLM macro-sheet parts are deliberately outside this narrow
stored-definition boundary. Enable no_formula_defined_xlm_get_cell_changes to
block this boundary in CI.

The ordinary semantic diff deliberately retains full local reviewer evidence.
For an artifact that leaves that boundary, add
`--redact-formula-defined-xlm-get-cell-calls` to `diff`, `check`, or
`portfolio`. The output-only mode replaces direct stored GET.CELL material,
changed invoking-formula evidence, exact changed static inputs, and changed
resolved formula-defined-name-chain evidence with
`[formula-defined XLM GET.CELL material redacted]`; it preserves comparison
facts, policy results, and exit status. It does not evaluate GET.CELL,
determine an information type, resolve a dynamic reference, or reconstruct a
value from Excel state, and is not a general secret scrubber.

FormulaFence also keeps a separate **formula-defined XLM environment-information
ledger** for selected GET.WORKBOOK, GET.WORKSPACE, and GET.DOCUMENT calls
stored in formula-defined names and named LAMBDA bodies. Microsoft's [C API
reference](https://learn.microsoft.com/en-us/office/client-developer/excel/programming-with-the-c-api-in-excel)
identifies workspace information functions such as GET.CELL and GET.WORKBOOK;
the [GET.WORKSPACE example](https://learn.microsoft.com/en-us/office/client-developer/excel/xlfree)
shows that it can return platform information, and the [expression-evaluation
reference](https://learn.microsoft.com/en-us/office/client-developer/excel/excel-worksheet-and-expression-evaluation)
documents GET.DOCUMENT as an XLM information function. A stored call can
therefore depend on workbook, workspace/client, or document state outside an
ordinary formula's visible cell dependencies.

The ledger propagates selected stored calls through nested and sheet-local
formula names to invoking cells. Profiles and FF071/FFP071 details expose only
invocation-cell, call, and relevant formula-defined-name counts; information
types, references, formulas, arguments, cells, and name identities remain
private. Same-count definition or invocation changes, uninvoked stored names,
and ordinary edits that reach a stored call through a statically visible
argument edge remain reviewable through private signatures.

FormulaFence does not determine which information type is requested, resolve a
dynamic reference, simulate workbook/workspace/document/client/add-in/printer
state, or infer dependencies from that state. A state-only workbook change is
not asserted to change a stored call. Direct worksheet calls and raw XLM
macro-sheet parts are deliberately outside this narrow stored-definition
boundary. Enable no_formula_defined_xlm_environment_information_changes to
block this boundary in CI.

The ordinary semantic diff deliberately retains full local reviewer evidence.
For an artifact that leaves that boundary, add
`--redact-formula-defined-xlm-environment-information-calls` to `diff`,
`check`, or `portfolio`. The output-only mode replaces direct stored selected
environment-information material, changed invoking-formula evidence, exact
changed static inputs, and changed resolved formula-defined-name-chain evidence
with `[formula-defined XLM environment-information material redacted]`; it
preserves comparison facts, policy results, and exit status. It does not
evaluate a call, determine an information type, resolve a dynamic reference,
simulate workbook/workspace/document state, or reconstruct a runtime value,
and is not a general secret scrubber.

FormulaFence also keeps a separate **native workbook and environment-information
ledger** for `CELL`, `INFO`, `SHEET`, and `SHEETS` calls in ordinary worksheet
formulas, formula-defined names, and named LAMBDA bodies. Microsoft's [CELL function
documentation](https://support.microsoft.com/en-us/office/cell-function-51bd39a5-f338-4dbe-a33f-955d67c2b2cf)
explains that CELL can return file, location, formatting, or content
information and can use the selected cell when its optional reference is
omitted. Microsoft's [INFO function
documentation](https://support.microsoft.com/en-au/office/info-function-725f259a-0e4b-49b3-8b52-58815c69acae)
lists operating-environment values such as the current folder, operating-system
version, calculation mode, and workbook-count information. Those values can
change even when visible precedents do not.

Microsoft's [SHEET function
documentation](https://support.microsoft.com/en-us/excel/functions/sheet-function)
states that SHEET returns a sheet number and, without its optional value,
returns the number of the sheet containing the function. Its [SHEETS function
documentation](https://support.microsoft.com/en-us/excel/functions/sheets-function)
states that an omitted reference returns the number of sheets in the containing
workbook. Both document that hidden, very-hidden, macro, chart, and dialog
sheets are included. A formula can therefore retain identical text while a tab
is inserted, removed, or moved and its result changes.

The ledger propagates calls through nested and sheet-local formula names to
invoking cells. Profiles and FF072/FFP072 details expose only formula-cell,
call, relevant formula-defined-name, omitted-CELL-reference, `SHEET`, `SHEETS`,
and omitted-SHEETS-reference counts; information types, references, formulas,
arguments, cells, name identities, and raw tab-catalog comparison material
remain private. The ordinary per-sheet inventory remains normal reviewer
context. Same-count definition or invocation changes, uninvoked stored names,
and ordinary edits that reach a call through a statically visible argument edge
remain reviewable through private signatures.

FormulaFence does not evaluate a formula or information call, determine an
information type, resolve a dynamic reference, infer the selected cell,
inspect a file/folder/client/workspace state, or simulate any of those states.
For stored `SHEET` calls and `SHEETS()` calls with an omitted reference, it
privately compares the raw OOXML workbook tab catalog—not only ordinary
worksheets—when it can read that catalog completely. A membership, order, or
tab-name change then emits FF072 because Excel may calculate from different
workbook-structure information; visibility-only changes do not trigger that
condition because Excel includes hidden tabs. FormulaFence does not determine
which explicit `SHEET`/`SHEETS` argument Excel resolves, calculate a particular
result, or infer whether a non-omitted SHEETS reference is a one-sheet or 3-D
reference. A malformed or unavailable tab catalog remains a parser coverage
warning. The no_formula_environment_information_changes rule blocks material
call, definition, invocation, statically visible input, and applicable tab-
catalog changes in CI.

The ordinary semantic diff deliberately retains full local reviewer evidence.
For an artifact that leaves that boundary, add
`--redact-formula-environment-information` to `diff`, `check`, or `portfolio`.
The output-only mode replaces direct stored `CELL`, `INFO`, `SHEET`, and
`SHEETS` material, exact changed static inputs, and changed resolved
formula-defined-name-chain evidence with
`[formula environment-information material redacted]`; it preserves comparison
facts, policy results, and exit status. It does not evaluate a formula or
information call, determine an information type, resolve a dynamic reference,
infer a selected cell, simulate workbook/client/workspace state, or reconstruct
a runtime value, and is not a general secret scrubber.

FormulaFence separately inventories **Excel 4.0 / XLM macro sheets**. Unlike
VBA, this executable automation is stored in raw macro-sheet XML parts (usually
`xl/macrosheets/*.xml`), not `xl/vbaProject.bin`. FormulaFence binds the
documented `xlMacrosheet` and `xlIntlMacrosheet` workbook relationships to
their parts, privately fingerprints accepted macro XML and related package
relationships, then streams direct safe internal relationship targets into
private payload fingerprints. Profiles expose only safe counts for sheets,
formula cells, visibility, related OLE/package parts, and fingerprinted versus
uninspected internal targets. A material change emits `FF026`; enable
`no_xlm_macro_sheet_changes` for `FFP026`. FormulaFence never executes,
emulates, resolves, or parses XLM commands, relationship targets, or embedded
objects. It never follows external targets; direct internal payload scanning is
bounded to 32 MiB per part, 64 MiB per workbook, and 256 parts, after which a
coverage warning remains visible. The package shape follows Microsoft's [Macro Sheet
part](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-offmacro/b8bee527-ef5a-4734-bb8c-6eae4166b6c9)
and [International Macro Sheet
part](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-offmacro/450634cb-ca5a-4350-9edb-940a90707f49).

Before private tree parsing, the macro-sheet scanner streams raw macro XML and
allows 32,768 XML elements per part and 65,536 across its scan, alongside 16
MiB per part, 64 MiB in aggregate, and 512 parts. A successfully streamed
structural overage remains opaque program evidence with a private streamed
content fingerprint and explicit parser-coverage warning, so a comparison can
surface `FF010` and `FF026` without retaining raw macro XML. After this raw
scan, FormulaFence gives its temporary ordinary-workbook reader an empty
worksheet replacement for the selected XLM targets, preventing that reader
from materializing the macro program; Custom View sanitization also excludes
those targets. An invalid ordinary-sheet relationship alias to the same target
is treated as XLM too and leaves visible coverage evidence. Malformed macro XML
reached before a structural overage retains its ordinary parser diagnostic.
These are bounds for the named raw macro-sheet XML readers, not a claim that
all legacy workbook XML is semantically interpreted.

FormulaFence also guards **XLM automatic-macro bindings**, the legacy
workbook-name dispatch path for `Auto_Open`, `Auto_Close`, `Auto_Activate`, and
`Auto_Deactivate`. A macro sheet's XML can remain byte-for-byte identical while
one of those special names is added, removed, or retargeted to a different
macro-sheet cell. The scanner reads raw workbook defined names, accepts only a
workbook-scoped name (including its optional `_xlnm.` built-in prefix) with a
direct internal single-cell A1 reference to a declared XLM macro sheet, and
keeps the spelling, target, and stored formula in a private signature. Public profiles
and `FF076` expose only aggregate event counts; enable
`no_xlm_automatic_macro_binding_changes` for `FFP076` in CI. Sheet-local names,
ordinary-sheet targets, external references, and dynamic/non-direct name
formulas are deliberately not asserted to be XLM automatic bindings. The
ordinary defined-name diff remains readable separately. FormulaFence never
evaluates or resolves a name, parses or executes XLM code, reads the reserved
`definedName@xlm` attribute as evidence, or infers whether Excel security
settings will run anything. Microsoft's [automatic-macro API](https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.runautomacros)
and [enumeration](https://learn.microsoft.com/en-us/office/vba/api/excel.xlrunautomacro)
define the four event names.

FormulaFence separately inventories **Office RibbonX custom UI** package parts.
RibbonX can bind controls to workbook callback names even when no ordinary cell
or `vbaProject.bin` payload changes. FormulaFence reads the documented root
package relationships and `customUI` XML for the 2006 and Office 2010-era
schemas, privately fingerprints full control XML plus direct image
relationships, and reports only safe part, control, callback-attribute, and
relationship counts. A material change emits `FF027`; enable
`no_ribbon_customization_changes` for `FFP027`. Control IDs, labels, callback
names, XML, and image targets never enter profiles or reports. FormulaFence
does not execute callbacks, follow external relationships, or parse image
payloads. Custom-UI XML reads are bounded to 16 MiB per part, 32 MiB per
workbook, and eight parts; a bound or malformed part becomes a coverage
warning. The package forms follow Microsoft's [Ribbon Extensibility
Part](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-customui/52faf7b6-fecc-48d9-96db-ee80a631a5ac)
and [Ribbon and Backstage Customizations
part](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-customui2/452a58ae-cb0a-4926-83f8-fb1cbaa6114c)
specifications; its `onLoad`, `loadImage`, and `onAction` callback surface is
documented in the [Custom UI schema](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-customui2/a232628d-f6fb-4630-a463-459989a68e7a).

FormulaFence separately inventories **Office Web Add-ins**. A workbook can
declare a task pane, bind a documented worksheet `x15:webExtensions` entry to
a definition `appref`, or host a `we:webextensionref` frame in worksheet
DrawingML even though no ordinary cell, VBA payload, or RibbonX control
changes. FormulaFence follows the bounded workbook-to-`taskpanes.xml` chain,
validates worksheet `appRef` bindings against direct `webextension*.xml`
definitions, and follows direct in-content frame relationships in the active
`mc:Choice` branch. It compares task-pane configuration, visible/locked state,
add-in references, auto-show properties, bindings, snapshots, active frame
placement/XML, and direct relationship semantics privately; profiles expose
only safe counts. The native-picture fallback is retained under the separate
worksheet-image boundary. A material change emits `FF028`; enable
`no_office_web_addin_changes` for `FFP028`. Add-in IDs, store references,
property values, binding values, worksheet formulas, snapshot data, frame XML,
and relationship targets never enter profiles or reports. FormulaFence does
not install, load, execute, or fetch an add-in or manifest, and it never
follows an external relationship. Task-pane and web-extension XML reads are
bounded to 16 MiB per part, 32 MiB per workbook, and 64 parts; worksheet
binding and in-content DrawingML XML reads are each bounded to 16 MiB per part,
64 MiB per workbook, and 512 parts. Malformed, unbound, oversized, or
over-budget parts remain explicit coverage warnings. The package surface
follows Microsoft's [Taskpane Web Extension File](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-owexml/3d04f8ce-65f2-4dc3-bafa-636d0a7e41a1),
[Web Extension](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-owexml/56fe5a64-dd6d-422c-beac-19d72dd10ade),
[Worksheet](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/07d607af-5618-4ca2-b683-6a78dc0d9627),
and [CT_WebExtension](https://learn.microsoft.com/en-nz/openspecs/office_standards/ms-xlsx/386851b6-b7b6-42b8-8cf1-d94bab7b0731)
specifications.

FormulaFence also inventories **DrawingML chart definitions and cached
presentation data**. A worksheet or chartsheet can point to a drawing part,
which binds a legacy `c:chart` part holding the chart type, series, titles,
axes, formatting, formulas, and last cached values outside ordinary cells. It
also recognizes Office 2016+ `cx:chart` ChartEx bindings, including the
`mc:AlternateContent` graphic-frame form Excel writes with an older-client
fallback. A legacy chart can point to a `c:userShapes` overlay part whose text
or image relationship changes what a reader sees; a ChartEx part can carry
fixed direct style, colour-style, drawing, image, theme-override, and embedded
package targets.

FormulaFence follows the bounded worksheet/chartsheet → drawing →
legacy-chart/ChartEx chain, compares private legacy chart definition and cache
material separately, fingerprints private ChartEx XML, and hashes bounded
direct related payloads without parsing their formats. It emits `FF030` for a
material change; enable `no_chart_definition_changes` for `FFP030`.

Profiles expose only safe counts for host sheets, legacy and ChartEx parts and
references, series, titles, chart-type elements, cached/literal point counts,
pivot/external/overlay references, relationships, and inspected versus
uninspected direct targets. Series formulas, labels, cached values, formatting,
overlay text, relationship targets, XML, and payload bytes never enter profiles
or reports. Writer-chosen relationship IDs and equivalent internal target
spellings are normalized away. Malformed, missing, orphaned, unbound,
unsupported, oversized, or over-budget chart material becomes a visible
coverage warning. XML reads are bounded to 16 MiB per part, 64 MiB per
workbook, and 512 parts; direct related payload hashes are bounded to 32 MiB
per part, 64 MiB per workbook, and 512 parts. FormulaFence does not calculate
a series formula, map chart inputs into the cell-impact graph, render a chart,
assess its visual truthfulness, follow external targets, parse direct media or
embedded-package formats, resolve ChartEx second-hop relationships, or
interpret ChartEx-specific visualization semantics. The boundary follows the
OOXML [Chart Part](https://c-rex.net/samples/ooxml/e1/Part1/OOXML_P1_Fundamentals_Chart_topic_ID0ELZLM.html), Microsoft's [ChartEx part definition](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-odrawxml/5d0d453e-adac-43be-a797-59b9916593dd), its [ChartEx relationship-ID type](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-odrawxml/d8ede39e-a36c-48ad-8a17-0086a2d0889b), the documented
[number-reference cache](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.charts.numberreference?view=openxml-3.0.1),
and the [chart-overlay relationship](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.charts.usershapesreference?view=openxml-2.20.0).

FormulaFence also inventories **PivotTable views, cache definitions, and cached
report material**. PivotTable packages can regroup, filter, aggregate, or
present a report without changing an ordinary formula cell. FormulaFence
follows the bounded workbook-cache and worksheet-PivotTable relationship graph,
then compares private PivotTable layout XML, cache-schema XML, shared cache
items, normalized relationships, and bounded raw cache-record hashes. A
material change emits `FF031`; enable `no_pivot_table_definition_changes` for
`FFP031`.

Source definitions and refresh behavior remain intentionally owned by the
existing `FF023` external-data control. This keeps a refresh-only edit distinct
from a report-layout or cached-report-data edit. Profiles expose only safe
counts for PivotTable parts, fields, items, cache records, relationships, and
coverage. PivotTable names, source ranges, field names, item values, formulas,
cache records, relationship targets, XML, and payload bytes never enter
profiles or reports. Writer-chosen relationship IDs, equivalent internal target
spellings, and cache-ID renumbering are normalized away.

Missing, malformed, orphaned, unbound, oversized, or over-budget PivotTable
material becomes a visible coverage warning. XML reads are bounded to 16 MiB
per part, 64 MiB per workbook, and 512 parts; raw cache-record hashes are
bounded to 32 MiB per part, 64 MiB per workbook, and 512 parts. FormulaFence
also detaches cache-record relationships in a temporary reader copy so the
underlying workbook library does not eagerly materialize unbounded record data;
the original workbook is never changed. It does not refresh a cache, calculate
a PivotTable, render a report, infer PivotTable-to-cell impact, fetch an
external target, or interpret OLAP or extension-list semantics. Slicer and
Timeline cache definitions are compared separately. The
package boundary follows the OOXML
[Pivot Table Part](https://c-rex.net/samples/ooxml/e1/Part1/OOXML_P1_Fundamentals_Pivot_topic_ID0ELLAM.html),
[Pivot Cache Definition Part](https://c-rex.net/samples/ooxml/e1/Part1/OOXML_P1_Fundamentals_Pivot_topic_ID0E1TAM.html),
and [Pivot Cache Records Part](https://c-rex.net/samples/ooxml/e1/Part1/OOXML_P1_Fundamentals_Pivot_topic_ID0EV2AM.html).

FormulaFence also inventories **Slicer and Timeline cache filter state**. A
Slicer can apply a cached item selection to a PivotTable or Excel table, and a
Timeline can apply a cached date filter to a PivotTable, without altering an
ordinary cell. FormulaFence follows documented workbook extension declarations
through their explicit workbook relationships to bounded `slicerCache` and
`TimelineCache` XML. It privately compares cache definitions, item selections,
Timeline state and filter material, PivotTable/table source bindings, filtered
PivotTable bindings, and any unexpected direct cache-part relationships. A
material change emits `FF032`; enable `no_slicer_timeline_cache_changes` for
`FFP032`.

Profiles expose only safe structural counts for cache parts, source and
PivotTable bindings, item/selection counts, Timeline states/filters, and
relationship coverage. Cache names, source fields, selected item values, date
ranges, PivotTable names, relationship targets, and XML never enter profiles
or reports. Writer-chosen relationship IDs, equivalent internal target
spellings, coordinated Slicer/Timeline PivotCache extension-ID renumbering, known optional Slicer
defaults, Boolean spellings, and Timeline GUIDs are normalized away. Missing,
malformed, orphaned, unbound, externally targeted, oversized, or over-budget
material becomes a visible coverage warning. Cache XML reads are bounded to 16
MiB per part, 64 MiB per workbook, and 512 parts. Before FormulaFence
materializes a cache tree, a streamed structural preflight permits 16,384 XML
elements per part and 32,768 across the complete cache scan. That capacity is
deliberately above Excel's documented 10,000 displayed filter drop-down items;
it is a CI allocation boundary, not a claim that a larger cache is invalid. A
well-formed structural overage stays visible as filter-cache coverage evidence
instead of being recursively canonicalized. FormulaFence does not apply a
filter, calculate or render a PivotTable or table, infer downstream cell
impact, fetch an external target, or model worksheet/drawing Slicer or Timeline
view geometry and styles. The boundary follows Microsoft's [Slicer Cache
part](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/7dbb4481-b021-45cc-8bd4-6094b566a5ff),
[Timeline Cache part](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/29a0f58c-d942-4641-8ed0-4f02010326f2),
[Slicer-to-PivotCache relationship](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/2a393f85-21f9-4a27-a2b7-4867223f4b9a),
and [Slicer view boundary](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/69c0e0f9-d014-4bd5-9f2d-2f554c715083),
alongside Excel's [published limits](https://support.microsoft.com/en-US/Excel/excel-specifications-and-limits).

FormulaFence also inventories **embedded Power Pivot/Data Model packages**. An
Excel Data Model can hold tables, relationships, calculations, and stored data
outside ordinary worksheet cells. FormulaFence follows the workbook's explicit
`powerPivotData` relationship and its `x15:dataModel` declaration, then
privately fingerprints the complete declaration and bounded raw
`xl/model/*.data` payload. A material change emits `FF033`; enable
`no_power_pivot_data_model_changes` for `FFP033`.

Profiles expose only safe counts for model parts, workbook bindings,
declarations, tables, relationships, fingerprinted payloads, and coverage.
Table names, column and relationship details, connection details, DAX,
stored values, relationship targets, XML, and raw payload bytes never enter a
profile or report. Writer-chosen relationship IDs, equivalent internal target
spellings, and GUIDs in Data Model metadata are normalized away. Missing,
malformed, orphaned, unbound, externally targeted, unexpected directly related,
oversized, or over-budget material becomes a visible coverage warning. Raw
payload reads are bounded to 512 MiB per part, 512 MiB per workbook, and 16
parts. FormulaFence does not deserialize the Analysis Services payload,
evaluate DAX, refresh a model, calculate or render a report, infer model-to-cell
impact, or fetch an external target. The boundary follows Microsoft's
[PowerPivot Model guidance](https://learn.microsoft.com/en-us/office/vba/excel/concepts/about-the-powerpivot-model-object-in-excel)
and the Open XML [`x15:dataModel` declaration](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.linq.x15.datamodel?view=openxml-3.0.1).

FormulaFence also inventories Excel **What-If Data Tables**—the sensitivity
analysis feature, not an Excel table. It reads each OOXML `f` element with
`t="dataTable"` directly from the worksheet and privately compares the master
output range, one- or two-variable mode, row/column orientation, input-cell
references, deleted-input flags, and recalculation request. A material change
emits `FF034`; enable `no_what_if_data_table_changes` for `FFP034`.

Profiles expose only safe counts for masters, table dimensions, one-/two-variable
forms, orientation, recalculation requests, deleted inputs, and malformed
definitions. Output ranges, input references, and raw formula metadata never
enter the profile, Markdown control section, `FF034` details, or SARIF.
Equivalent A1 case/absolute-reference spellings and Boolean spellings are
normalized. A missing, malformed, overlapping, or unsupported definition is a
visible coverage warning. FormulaFence does not calculate scenarios, infer the
scenario output formula, predict recalc results, or add Data Table inputs to the
ordinary dependency graph. Cached output cells remain ordinary cell values under
the regular cell-diff boundary. The declaration boundary follows the Open XML
[`f` (Formula) specification](https://c-rex.net/samples/ooxml/e1/part4/OOXML_P4_DOCX_f_topic_ID0E6TY4.html)
and Excel's [What-If Data Table guidance](https://support.microsoft.com/en-us/office/calculate-multiple-results-by-using-a-data-table-e95e2487-6ca6-4413-ad12-77542a5ea50b).

FormulaFence also inventories Excel **Scenario Manager** controls. Unlike a
Data Table, a Scenario Manager worksheet stores named, alternate input-value
sets that Excel can apply to the sheet. FormulaFence reads the worksheet's raw
`<scenarios>` declaration before the workbook reader can omit it, then privately
compares scenario selection state, result-summary references, names, protection
flags, comments/users, changing-cell references, stored values, deleted/undone
state, and display number formats. A material change emits `FF035`; enable
`no_scenario_manager_changes` for `FFP035`.

Profiles expose only safe structural counts: Scenario Manager worksheets,
scenarios, stored inputs, locked/hidden controls, comment/user presence,
summary references, selections, deleted/undone inputs, number-format presence,
and malformed declarations. Scenario names, comments, users, input values,
input references, summary references, and raw XML never enter a profile,
Markdown control section, `FF035` details, or SARIF. Equivalent local A1
case/absolute-reference, Boolean, and unsigned-integer spellings are normalized;
schema-default `locked`, `hidden`, `deleted`, and `undone` flags are normalized
too. Missing, malformed, duplicate-within-worksheet, or unsupported
declarations become visible coverage warnings. FormulaFence does not show or
apply a scenario, calculate its results, infer a scenario-to-formula dependency,
or fetch an external target. Cached worksheet cells remain under the regular
cell-diff boundary. The declaration boundary follows Open XML
[`scenarios`](https://c-rex.net/samples/ooxml/e1/part4/OOXML_P4_DOCX_scenarios_topic_ID0EVDF5.html),
[`scenario`](https://c-rex.net/samples/ooxml/e1/part4/OOXML_P4_DOCX_scenario_topic_ID0E5WE5.html),
and [`inputCells`](https://c-rex.net/samples/ooxml/e1/Part4/OOXML_P4_DOCX_inputCells_topic_ID0EE624.html)
definitions, plus Excel's [Scenario Manager guidance](https://support.microsoft.com/en-us/excel/switch-between-various-sets-of-values-by-using-scenarios).

FormulaFence also inventories **filter, sort, and row/column-visibility controls**.
An Excel AutoFilter can hide rows by private criteria without changing a formula
or a cell value, and `SUBTOTAL` always excludes filtered-out rows; its 101–111
forms also exclude manually hidden rows. A hidden or collapsed column can also
remove a report field from a reviewer’s view without changing any cell. Excel
also documents a row or column dimension of zero as hidden. FormulaFence reads
worksheet `<autoFilter>` and `<sortState>` declarations, Table Definition-part
filters, explicit row `hidden`/`outlineLevel`/`collapsed` state and zero `ht`,
the worksheet-level `sheetFormatPr@zeroHeight` hidden-by-default optimization,
and zero `defaultRowHeight` / `defaultColWidth` defaults. It also reads raw
worksheet `<cols>/<col>` declarations for `hidden`, `outlineLevel`, `collapsed`,
and zero `width`, applying overlapping records in file order so a later
*present* attribute overrides only that control. A material change emits
`FF036`; enable `no_filter_visibility_changes` for `FFP036`.

Profiles expose only counts for worksheet/table filters, filter columns and
criterion groups, sort states/conditions, default-hidden/default-zero-dimension
sheets, explicitly hidden/zero-height/outlined/collapsed rows, visible-row
overrides, hidden/zero-width/outlined/collapsed columns, and malformed controls.
Filter criteria, selected values, table names, custom sort lists, sort keys,
raw dimension values, and row/column ranges never enter profiles, Markdown control
sections, `FF036` details, or SARIF. Equivalent local A1 case/absolute-reference,
Boolean/default, unsigned-integer, equivalent zero-dimension spellings, and
equivalent column-range segmentation spellings are normalized. Unsupported
extensions, malformed declarations, exhausted control-update limits, or
unsafe/missing table relationships are visible coverage warnings rather than
silent omissions. FormulaFence does not apply a filter, calculate a result,
infer which formulas are visibility sensitive, render a report, track ordinary
positive width/height outside the dedicated worksheet-dimension boundary, or
model outline-display settings. The boundary follows
Microsoft's [SUBTOTAL documentation](https://support.microsoft.com/en-us/excel/functions/subtotal-function),
its [row-height and column-width guidance](https://support.microsoft.com/en-us/excel/change-the-column-width-and-row-height),
the Open XML [`autoFilter`](https://c-rex.net/samples/ooxml/e1/part4/OOXML_P4_DOCX_autoFilter_topic_ID0EIDM4.html),
[`filterColumn`](https://c-rex.net/samples/ooxml/e1/Part4/OOXML_P4_DOCX_filterColumn_topic_ID0ELVP5.html),
[`sheetFormatPr`](https://c-rex.net/samples/ooxml/e1/part4/OOXML_P4_DOCX_sheetFormatPr_topic_ID0EVAG5.html),
[`row`](https://c-rex.net/samples/ooxml/e1/part4/OOXML_P4_DOCX_row_topic_ID0EIKD5.html),
[`cols`](https://c-rex.net/samples/ooxml/e1/Part4/OOXML_P4_DOCX_cols_topic_ID0E5XR4.html),
and [`col`](https://c-rex.net/samples/ooxml/e1/part4/OOXML_P4_DOCX_col_topic_ID0ELFQ4.html)
definitions.

FormulaFence also inventories **ignored Excel error-checking controls**. Excel
persists a reviewed decision to suppress particular warnings for a range: a
formula evaluation error, an inconsistent formula, an omitted range, an
unlocked formula, an empty reference, a list-validation issue, a calculated
column mismatch, a number stored as text, or a two-digit text year. Those
decisions can alter the warning surface a reviewer sees without changing an
ordinary cell or formula. FormulaFence reads both standard `<ignoredErrors>`
and Office 2010 `x14:ignoredErrors` declarations. A material change emits
`FF037`; enable `no_ignored_error_changes` for `FFP037`.

Profiles expose only structural counts for worksheets, standard/extension
containers, suppressed warning rules, target ranges, and warning kinds. Target
ranges and exact suppressions never enter profiles, Markdown control sections,
`FF037` details, or SARIF. Equivalent local A1 case/absolute-reference,
Boolean, and target-order spellings are normalized. Malformed or unsupported
containers, extensions, attributes, flags, targets, or child markup become
visible coverage warnings. FormulaFence does not determine whether a warning
applies, calculate a formula, repair an error, change application-level
error-checking options, or infer a suppressed warning's downstream impact. The
boundary follows the OOXML [`ignoredError`](https://c-rex.net/samples/ooxml/e1/Part4/OOXML_P4_DOCX_ignoredError_topic_ID0EVK24.html)
definition and Microsoft's Office 2010 [`ignoredErrors`](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/0d164d85-23bf-4d43-87c5-9fcde148aabe)
documentation.

FormulaFence also inventories modern Excel **Named Sheet Views**. A saved view
is stored in a relationship-backed part and can retain an alternate AutoFilter
criterion or sort order that changes what a reviewer sees without changing a
cell, formula, or the worksheet's active filter. FormulaFence follows the
worksheet relationship to each Named Sheet View part, privately compares the
view definition, and reconciles each stored filter to its base AutoFilter using
Excel's documented AutoFilter UID, table-ID, then worksheet-filter sequence. A
material change emits `FF038`; enable `no_named_sheet_view_changes` for
`FFP038`.

Profiles expose only structural counts for worksheets, relationship-backed
parts, named views, alternate filters, column filters, criterion groups, sort
rules/conditions, and unrecognized controls. View names, IDs, criteria, target
ranges, table bindings, table-column IDs, and sort keys never enter profiles,
Markdown control sections, `FF038` details, or SARIF. Equivalent GUID, local
A1 case/absolute-reference, Boolean/default, and unsigned-integer spellings are
normalized. Missing, ambiguous, mismatched, malformed, unsupported, oversized,
or unsafe parts and filter bindings become visible coverage warnings. FormulaFence
does not activate or render a saved view, calculate a filtered result, infer
formula visibility sensitivity, repair a mismatched declaration, or support
full differential-format, future extension, or rich-sort semantics. The boundary follows Microsoft's
[`Named Sheet Views` part](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/4b4d6448-d997-4ebe-9153-5c2c67d16972),
[`CT_NsvFilter`](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/e132d9cc-c711-4fb3-aa28-e7356a791b1c),
and [reconciliation](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/dd6b2cb8-b5b3-43b1-a5bd-dccdd9c0864a)
definitions.

FormulaFence also inventories legacy Excel **Custom Views**. A workbook-level
`<customWorkbookView>` can name an alternate display/print mode and bind it by
GUID to a `<customSheetView>` on every workbook sheet. Those saved states can
change hidden rows or columns, filters, print settings, formula/gridline and
header display, panes, comments, and object visibility without changing an
ordinary cell or the workbook's active view. FormulaFence reads those raw OOXML
records before the ordinary workbook reader, privately reconciles each GUID to
its per-sheet declarations, and emits `FF060` for a material alternate-view
change. Enable `no_custom_workbook_view_changes` for `FFP060` in CI.

Profiles expose only structural counts for workbook views, per-sheet views,
sheets with alternate state, and per-sheet views carrying
hidden/filter/print/display settings or unrecognized metadata. View names,
GUIDs, sheet bindings, ranges,
filter criteria, print settings, pane locations, and raw XML never enter
profiles, Markdown control sections, `FF060` details, or SARIF. Coordinated
GUID rewrites, sheet-ID/active-sheet-ID remapping, Boolean/default spelling,
and unsigned-integer spelling normalize. Transitional and Strict SpreadsheetML
worksheets/dialog sheets plus chart sheets are supported; legacy Custom View
containers are isolated from the ordinary reader only after FormulaFence has
captured their evidence. Missing, duplicate, malformed, unsupported,
unbound, unsafe, oversized, or over-budget declarations remain explicit
coverage warnings.

FormulaFence does not activate or render a Custom View, calculate a filtered
result, determine what an application will print, interpret future extensions,
or support Custom Views on unsupported sheet types. The boundary follows
Microsoft's [`customWorkbookView`](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.customworkbookview?view=openxml-3.0.1)
and [`customSheetView`](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.customsheetview?view=openxml-3.0.1)
definitions.

FormulaFence also inventories **Excel Table Style controls**. A table can change
its visible headers, totals, banding, emphasized columns, data-area format, or
borders without changing a cell value, formula, table reference, or structured
formula. FormulaFence reads raw table `tableStyleInfo` bindings and toggles,
custom workbook `tableStyle` / `tableStyleElement` definitions, their resolved
`dxf` formatting material, and direct Table/TableColumn Dxf and named-cell-style
references. A material presentation declaration or coverage change emits
`FF061`; enable `no_table_style_control_changes` for `FFP061` in CI.

Profiles expose only structural counts for style declarations, styled tables,
custom styles/elements, direct Dxf and named-cell-style assignments, banding,
emphasized columns, and unrecognized controls. Table names, custom style names,
cell-style names, Dxf formatting, colours, IDs, and raw XML never enter
profiles, Markdown control sections, `FF061` details, or SARIF. Equivalent
Boolean/default spelling, case-only style names, `xr9:uid` revision provenance,
and coordinated Dxf reordering/ID rewriting normalize. Transitional and Strict
SpreadsheetML parts are read directly; presentation-only Table Style XML is
isolated from the ordinary reader after raw evidence is captured. Missing,
duplicate, malformed, unresolved, unsupported, oversized, or over-budget
records become explicit coverage warnings.

FormulaFence compares stored declarations rather than rendering Excel's final
appearance: it does not resolve themes, calculate values, evaluate conditional
formatting, apply a style to cells, interpret future extensions, or cover
PivotTable-only Table Style regions. `defaultTableStyle` is a preference for
newly created tables rather than a binding on an existing table, so it is not a
review-surface control. Named cell-style references are compared privately, but
their same-name underlying style definitions remain outside this Table Style
boundary. The boundary follows Microsoft's [`TableStyles`](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.tablestyles?view=openxml-3.0.1),
[`TableStyleInfo`](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.tablestyleinfo?view=openxml-3.0.1),
and [`TableColumn`](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.tablecolumn?view=openxml-3.0.1)
definitions.

FormulaFence also inventories **legacy shared-workbook revision history**. An
older shared workbook can retain headers and log files outside the ordinary
worksheet grid. They can preserve prior cell values, author identities,
timestamps, comments, formatting edits, conflict-resolution records, and
tracking/retention/protection controls even when cells and formulas are
unchanged. FormulaFence follows the workbook-to-header and header-to-log
relationships, streams each revision XML part before private parsing, and
fingerprints complete bounded revision declarations privately. Each revision
part permits 32,768 XML elements and the complete revision scan permits 65,536,
alongside its existing 16 MiB per-part, 64 MiB aggregate, and 512-part
byte/count limits. A successfully streamed overage becomes visible
`FF010`/`FF062` coverage evidence rather than a materialized private tree.
Enable `no_shared_workbook_revision_changes` for `FFP062` in CI.

Profiles expose only revision-header/log-part and record counts plus aggregate
shared/tracked/history-retention/protection state and unrecognized-metadata
counts. Prior/new values, locations, author names, timestamps, comments, GUIDs,
relationship IDs, and raw XML never enter JSON, Markdown, `FF062` details, or
SARIF. Equivalent Boolean/integer spelling, coordinated relationship-ID
rewrites, and transitional versus Strict SpreadsheetML relationship spelling
normalize. Missing, duplicate, malformed, unsafe, unsupported, oversized, or
over-budget declarations become explicit coverage evidence rather than a silent
blind spot. Structural-overage evidence keeps a private streamed content
fingerprint, so same-size hostile history changes remain diff-visible.

FormulaFence compares stored audit declarations; it does not apply revisions,
reconstruct a historical workbook state, resolve conflicts, validate an author
or timestamp, render Excel, or interpret arbitrary future-extension content.
The boundary follows Microsoft's
[`headers`](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.headers?view=openxml-3.0.1),
[`header`](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.header?view=openxml-3.0.1),
and
[`revisions`](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.revisions?view=openxml-3.0.1)
definitions.

FormulaFence also inventories **cell number-format controls**. A number format
can make an unchanged value appear blank (for example, Excel's custom `;;;`
format), scale it with commas, or render it as a percentage, date, text, or a
private literal. FormulaFence reads raw `styles.xml` custom `<numFmt>` records,
base `<cellStyleXfs>`, effective `<cellXfs>` records and their
`xfId`/`applyNumberFormat` inheritance, plus direct cell `s`, row `s` with
`customFormat=1`, and worksheet `<cols>/<col style>` assignments. A material
change emits `FF039`; enable `no_number_format_changes` for `FFP039`.

Profiles expose only counts for default overrides, direct-cell/row/effective
column assignments, built-in/custom assignments, and malformed controls.
Format codes, style IDs, and cell/row/column targets never enter profiles,
Markdown control sections, `FF039` details, or SARIF. Equivalent custom-format
ID remapping, `applyNumberFormat` Boolean spelling, base-XF inheritance, and
equivalent effective column-range splitting are normalized. Missing custom
definitions, invalid style references, invalid/out-of-range targets, conflicting
definitions, and bounded-parser failures become visible coverage warnings rather
than silent omissions. FormulaFence compares declarations only: it does not
render locale-specific output, validate a format code, calculate a value, model
width/overflow, compose number formats with font/fill/alignment/border or
separately inventoried Table Style controls, or cover quote prefixes or
arbitrary visual formatting. A
column `style` is recorded as
the OOXML column default for unallocated/new cells; FormulaFence does not claim
to apply that default retroactively to existing allocated cells. The boundary
follows Microsoft's [custom number-format guidance](https://support.microsoft.com/en-us/excel/review-guidelines-for-customizing-a-number-format),
its [hide/display guidance](https://support.microsoft.com/en-us/excel/hide-or-display-cell-values),
and Open XML's [`numFmt`](https://c-rex.net/samples/ooxml/e1/Part4/OOXML_P4_DOCX_numFmt_topic_ID0EHDH6.html),
[`cellStyleXfs`](https://c-rex.net/samples/ooxml/e1/part4/OOXML_P4_DOCX_cellStyleXfs_topic_ID0EXX65.html),
[`xf`](https://c-rex.net/samples/ooxml/e1/part4/OOXML_P4_DOCX_xf_topic_ID0E13S6.html),
and [`col`](https://c-rex.net/samples/ooxml/e1/part4/OOXML_P4_DOCX_col_topic_ID0ELFQ4.html)
definitions.

FormulaFence also inventories **cell-font controls**. A font can make an
unchanged value or warning less visible—for example, by using a white font
against a matching background—or change its face, size, emphasis, underline, or
other display effects. FormulaFence reads raw `styles.xml` `<fonts>` records,
base `<cellStyleXfs>`, effective `<cellXfs>` records and their
`xfId`/`applyFont` inheritance, plus direct cell `s`, row `s` with
`customFormat=1`, and worksheet `<cols>/<col style>` assignments. A material
change emits `FF040`; enable `no_cell_font_changes` for `FFP040`.

Profiles expose only counts for the default definition, direct-cell/row/effective
column assignments, and malformed controls. Font names, colour values, effects,
style IDs, and cell/row/column targets never enter profiles, Markdown control
sections, `FF040` details, or SARIF. Equivalent font-ID remapping,
`applyFont` Boolean spelling, base-XF inheritance, common font-child ordering,
and equivalent effective column-range splitting are normalized. Missing or
malformed font/style definitions, invalid IDs/indexes/targets, and bounded-parser
failures become visible coverage warnings rather than silent omissions.
FormulaFence compares declarations only: it does not render or resolve theme
colours, decide whether a font is visible against a fill, calculate
text/background contrast, compose font rendering with fill/border/alignment or
other display controls, rich-text run rendering, separately inventoried Table
Style controls, or arbitrary visual formatting.
A column `style` is recorded as the
OOXML default for unallocated/new cells;
FormulaFence does not claim to apply that default retroactively to existing
allocated cells. The boundary follows the OOXML [`xf`](https://c-rex.net/samples/ooxml/e1/Part4/OOXML_P4_DOCX_xf_topic_ID0E13S6.html)
form and the ICAEW's [spreadsheet-review guidance](https://www.icaew.com/-/media/corporate/files/technical/technology/excel/how-to-review-a-spreadsheet-report.ashx),
which explicitly flags white-on-white font use as a hidden-calculation risk.

FormulaFence also inventories **cell-fill controls**. A fill can change whether
unchanged text, a warning, or a visual input/output distinction is legible—for
example, by making a matching background solid, patterned, or gradient. It
reads raw `styles.xml` `<fills>` definitions, including `patternFill` and
`gradientFill` stops, base `<cellStyleXfs>`, effective `<cellXfs>` records and
their `xfId`/`applyFill` inheritance, plus direct cell `s`, row `s` with
`customFormat=1`, and worksheet `<cols>/<col style>` assignments. A material
change emits `FF041`; enable `no_cell_fill_changes` for `FFP041`.

Profiles expose only counts for the default definition, direct-cell/row/effective
column assignments, and malformed controls. Fill colours, pattern types,
gradient geometry/stops, style IDs, and cell/row/column targets never enter
profiles, Markdown control sections, `FF041` details, or SARIF. Equivalent
fill-ID remapping, `applyFill` Boolean spelling, base-XF inheritance, valid
pattern-colour child ordering, semantically inert no-fill/solid-background
declarations, and equivalent effective column-range splitting are normalized.
Missing or malformed fill/style definitions, invalid IDs/indexes/targets, and
bounded-parser failures become visible coverage warnings rather than silent
omissions. FormulaFence compares declarations only: it does not resolve theme
colours, render patterns or gradients, calculate text/background contrast,
evaluate conditional-format differential styles, apply separately inventoried
Table Style controls, calculate values, compose fill rendering with
border/alignment or other display controls, rich-text run rendering,
width/overflow, or arbitrary visual formatting. A
column `style` is recorded as the OOXML default for
unallocated/new cells; FormulaFence does not claim to apply that default
retroactively to existing allocated cells. The boundary follows OOXML's
[`xf`](https://c-rex.net/samples/ooxml/e1/Part4/OOXML_P4_DOCX_xf_topic_ID0E13S6.html),
[`patternFill`](https://c-rex.net/samples/ooxml/e1/Part4/OOXML_P4_DOCX_patternFill_topic_ID0E6KM6.html),
and [`gradientFill`](https://c-rex.net/samples/ooxml/e1/Part4/OOXML_P4_DOCX_gradientFill_topic_ID0ENWD6.html)
forms, alongside the ICAEW's [spreadsheet-review guidance](https://www.icaew.com/-/media/corporate/files/technical/technology/excel/how-to-review-a-spreadsheet-report.ashx)
to reveal text hidden by formatting.

FormulaFence also inventories **effective cell-alignment controls**. An
unchanged value, warning, or visual classification can be pushed out of a
reviewer's normal view, rotated, wrapped, or shrunk without a formula or value
edit. FormulaFence reads raw `styles.xml` `alignment` children from base
`cellStyleXfs` and effective `cellXfs` records, follows `xfId` and
`applyAlignment`, and compares direct cell `s`, row `s` with
`customFormat=1`, and worksheet `<cols>/<col style>` assignments. It covers
horizontal/vertical placement, rotation, wrapping, shrinking, indentation,
relative indentation, justification, and reading order. A material effective
change emits `FF054`; enable `no_cell_alignment_changes` for `FFP054`.

Profiles expose only counts for default definitions, direct-cell, row, effective
column, and malformed controls. Alignment values, style IDs, and cell/row/column
targets never enter profiles, Markdown control sections, `FF054` details, or
SARIF. Equivalent explicit defaults, Boolean/integer spelling, semantically
inert `mergeCell` compatibility material, base-XF inheritance,
`applyAlignment`, and equivalent effective column-range splitting are
normalized. Missing, duplicate, malformed, or unsupported alignment metadata
becomes a visible coverage warning rather than a silent omission.

This is a stored-declaration boundary, not a layout engine. FormulaFence does
not calculate column width, row height, merged-cell layout, overflow, final
visibility, font/fill/conditional-format composition, or Excel client
rendering. A column `style` remains an OOXML default for unallocated/new
cells, not a claim to restyle allocated cells. The scope follows Microsoft's
[SpreadsheetML alignment definition](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oi29500/e4ad6e3e-7702-4dbe-8c44-f5a4c686c440)
and [CellFormat alignment semantics](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oi29500/68362a4b-5589-4504-b566-e8154dce1de3).

FormulaFence also inventories **effective cell-border controls**. An unchanged
value can be reframed as a report edge, total, exception box, or warning by a
border edit alone. FormulaFence reads raw `styles.xml` `<borders>/<border>`
definitions, base `cellStyleXfs`, effective `cellXfs` records and their
`borderId`, `xfId`, and `applyBorder` inheritance, plus direct cell `s`, row
`s` with `customFormat=1`, and worksheet `<cols>/<col style>` assignments. It
covers left/right/top/bottom edges, Office 2010 logical start/end edges,
diagonals and their direction, outline, stored line styles, and stored colours.
A material effective change emits `FF057`; enable `no_cell_border_changes` for
`FFP057`.

Profiles expose only counts for default definitions, direct-cell, row,
effective-column, and unrecognized controls. Border definitions, colours,
style IDs, and cell/row/column targets never enter profiles, Markdown control
sections, `FF057` details, or SARIF. Equivalent omitted/`none` sides,
Boolean/colour spellings, unused diagonal payload, ineffective empty
`outline="false"`, base-XF inheritance, `applyBorder`, and equivalent effective
column-range splitting are normalized. Missing, duplicate, malformed, or
unsupported border metadata becomes a visible coverage warning rather than a
silent omission. Material `vertical` or `horizontal` inner sides under ordinary
cell styles are also surfaced as a coverage warning: those sides have
differential-format semantics which this ordinary-cell boundary does not model.

This is a stored-declaration boundary, not a renderer. FormulaFence does not
resolve theme or palette colours, choose adjacent-cell border precedence,
render a final visual style, apply conditional-format/table/differential-style
borders, calculate print output, or infer Excel client behavior. A column
`style` remains an OOXML default for unallocated/new cells, not a claim to
restyle allocated cells. The scope follows OOXML's
[`border`](https://c-rex.net/samples/ooxml/e1/part4/OOXML_P4_DOCX_border_topic_ID0EVV35.html)
and [`xf`](https://c-rex.net/samples/ooxml/e1/part4/OOXML_P4_DOCX_xf_topic_ID0E13S6.html)
forms, the Open XML SDK's
[`Border` schema surface](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.border?view=openxml-3.0.1),
and Microsoft's [cell-border guidance](https://support.microsoft.com/en-us/Excel/apply-or-remove-cell-borders-on-a-worksheet).

FormulaFence also inventories **material worksheet-dimension controls**. A
positive row-height or column-width edit can clip wrapped text, suppress visual
context, or reframe a report without changing a cell value or formula. It can
also move Excel's automatic page breaks. FormulaFence reads raw transitional
and strict SpreadsheetML `sheetFormatPr` defaults (`defaultRowHeight`,
`defaultColWidth`, and `baseColWidth`), explicit/default `customHeight`, Office
2010 `x14ac:dyDescent` baseline adjustments, and active automatic thick-border
row adjustments. It also compares direct row
`ht`/`customHeight`/`x14ac:dyDescent`/automatic `thickTop` or `thickBot` controls and raw
`<cols>/<col>` positive `width` and `bestFit` declarations. Overlapping column
records are applied in XML order so only a later *present* width or AutoFit
attribute overrides the earlier effective state. A material change emits
`FF058`; enable `no_worksheet_dimension_changes` for `FFP058`.

Profiles expose only counts for default row/column sizing, baseline-adjustment
and automatic border-adjustment sheets, direct row heights, row baseline/border
adjustments, effective positive-width columns, AutoFit columns, and malformed
controls. Dimension values, sheet names, row/column targets,
`customHeight`/`customWidth` flags, and raw XML never enter profiles, Markdown
control sections, `FF058` details, or SARIF. Decimal and Boolean spelling,
baseline defaults, inert thick-border
flags under a fixed custom height, the `customWidth` writer hint, and equivalent
effective column-range splitting are normalized. Zero/hidden dimensions remain
under `FF036` because they conceal content; a later positive width can therefore
both reveal a column and change an ordinary dimension.

Missing, duplicate, malformed, unsupported, or budget-exhausted metadata is a
coverage warning rather than a silent omission. FormulaFence compares stored
declarations, not a rendered workbook: it does not calculate final AutoFit
sizes, font/merged-cell/overflow layout, exact automatic page breaks, print
geometry, or client-specific visibility. The boundary follows Microsoft's
[row-height and column-width guidance](https://support.microsoft.com/en-us/excel/change-the-column-width-and-row-height),
its [wrapped-text guidance](https://support.microsoft.com/en-us/excel/wrap-text-in-a-cell-in-excel),
its [automatic page-break guidance](https://support.microsoft.com/en-US/Excel/insert-move-or-delete-page-breaks-in-a-worksheet),
and OOXML's [`sheetFormatPr`](https://c-rex.net/samples/ooxml/e1/part4/OOXML_P4_DOCX_sheetFormatPr_topic_ID0EVAG5.html),
[`row`](https://c-rex.net/samples/ooxml/e1/part4/OOXML_P4_DOCX_row_topic_ID0EIKD5.html),
and [`col`](https://c-rex.net/samples/ooxml/e1/part4/OOXML_P4_DOCX_col_topic_ID0ELFQ4.html)
forms, plus Microsoft's [`dyDescent` extension documentation](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/f11dfda4-46de-4035-8418-d76b0d3898f1).

FormulaFence also inventories **material worksheet-display controls**. A saved
worksheet view can make an unchanged zero appear blank, replace displayed
results with formulas, hide or recolour gridlines, remove row/column headers,
hide outline symbols, rulers, or page margins, switch direction or view mode,
or split/freeze the review surface.
FormulaFence reads raw transitional and strict SpreadsheetML
`sheetViews/sheetView` declarations before a workbook reader can normalize
them. It compares non-default
`showZeros`, `showFormulas`, `showGridLines`, custom gridline-colour
(`defaultGridColor`/`colorId`), `showRowColHeaders`, `showOutlineSymbols`,
`showRuler`, `showWhiteSpace`, and `rightToLeft` flags, non-normal view
modes, and material split/frozen pane state. A material change emits `FF055`; enable
`no_worksheet_display_control_changes` for `FFP055`.

Profiles expose only structural display-control counts; sheet names, target
cells, pane positions, and raw XML remain private. Omitted/default controls,
Boolean and active custom-gridline-colour spelling, and finite pane-split
decimal spelling are normalized.
Active-cell, selection, top-left navigation, and zoom state are deliberately
out of scope to avoid ordinary writer churn. Missing, malformed, duplicate, or
unsupported material becomes a coverage warning rather than a silent omission.
FormulaFence compares stored declarations, not Excel rendering, the effective
palette colour, viewport geometry, final visibility, extension-specific view
behavior, or client state. The boundary follows the Open XML SDK
[`SheetView` schema surface](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.sheetview?view=openxml-3.0.1)
and Microsoft’s [worksheet display guidance](https://learn.microsoft.com/en-us/office/dev/add-ins/excel/excel-add-ins-worksheet-display),
which describes persisted gridline, heading, and page-layout controls.

FormulaFence also inventories **material worksheet print-layout controls**. A
workbook can print a materially different report without changing a cell: a
print area can omit rows or columns, titles can repeat different context, and
margins, fit/scaling, headers, footers, or manual page breaks can reframe the
saved output. FormulaFence reads raw transitional and strict SpreadsheetML
before ordinary readers normalize it. It compares workbook `definedName`
declarations for `_xlnm.Print_Area` and `_xlnm.Print_Titles`, plus direct
worksheet `printOptions`, `pageMargins`, `pageSetup`,
`sheetPr/pageSetUpPr`, `headerFooter`, `rowBreaks`, and `colBreaks` controls.
A material declaration change emits `FF056`; enable
`no_worksheet_print_layout_changes` for `FFP056`.

Profiles and reports expose only structural counts. Print ranges, header/footer
text, page values, printer settings references, and raw XML remain private.
Equivalent omitted/default, Boolean, integer, and decimal spellings normalize
away. The scanner also keeps known no-ops quiet: printing gridlines requires
both saved flags, disabled first/even header/footer sections are ignored,
`firstPageNumber` is ignored until enabled, fit-to-page selects fit dimensions
instead of percentage scale, and automatic-break display state is not treated
as a manual pagination change. Missing, duplicate, malformed, or unsupported
metadata creates a coverage warning rather than a silent omission.

This is a stored-declaration boundary, not a print engine. FormulaFence does
not render or preview Excel, calculate page geometry/counts, determine
automatic page breaks, resolve client/printer defaults or printer-specific
`devMode` settings, or claim coverage for custom/legacy sheet-view print
layouts and extension-specific behavior. The scope follows Microsoft's
[print-area guidance](https://support.microsoft.com/en-us/excel/set-or-clear-a-print-area-on-a-worksheet)
and [`PageLayout` control surface](https://learn.microsoft.com/en-us/javascript/api/excel/excel.pagelayout?view=excel-js-preview).

FormulaFence also inventories the workbook-level **DrawingML Theme**. A Theme
can change the colour, font, or effect schemes used by themed cells, charts,
and drawing objects without changing their local style references. FormulaFence
reads the raw workbook Theme binding, Theme XML, and direct Theme-image
relationships/payloads in both transitional and strict OOXML namespaces. A
material stored control change emits `FF053`; enable
`no_workbook_theme_changes` to make that boundary `FFP053` in CI.

Profiles and `FF053` details expose only aggregate Theme-part, colour-scheme,
font-scheme, format-scheme, relationship, external-relationship, image, and
malformed-metadata counts. Theme XML, scheme names, colour values, font names,
image payloads, relationship IDs, and targets never enter profiles, Markdown,
JSON, or SARIF. Writer-selected relationship IDs/order and equivalent internal
target spelling stay quiet. Missing, duplicate, malformed, unsafe, unbound,
unreadable, oversized, or over-budget metadata becomes a visible coverage
warning; raw reads are bounded to 16 MiB per part, 64 MiB per workbook, and
512 parts. FormulaFence also streams Theme and Theme-relationship XML before
tree materialization: 32,768 elements per XML part and 65,536 across the
complete Theme scan. These are reader-allocation and coverage limits, not
workbook-validity limits; a well-formed structural overage produces visible
`FF053` coverage evidence. Direct Theme-image payloads stay byte-bounded rather
than being treated as XML.

This is a stored-declaration boundary, not a rendering engine. FormulaFence
does not resolve effective cell/chart/drawing styles, render a workbook,
calculate contrast, decode an image, fetch a relationship target, calculate
formulas, or infer Excel client behavior. The scope follows the Open XML SDK
[WorkbookPart](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.packaging.workbookpart?view=openxml-2.20.0)
Theme-part surface and Microsoft's
[conditional-formatting guidance](https://learn.microsoft.com/en-us/office/open-xml/spreadsheet/working-with-conditional-formatting),
which illustrates Theme-indexed colours in spreadsheet formatting.

FormulaFence also compares **stored formula results**. SpreadsheetML can retain
the last calculated result beside formula text in the same `<c>` cell. Most
formula readers expose one representation or the other, so FormulaFence reads
the raw `<f>` and `<v>` elements together, privately fingerprints each result,
and compares the saved display state without exposing a result value or
formula-cell location. A cache change emits `FF042` only when FormulaFence
cannot explain it by a changed formula at that cell or by an ordinary changed
cell that reaches it through the static dependency graph. Enable
`no_formula_cached_result_changes` to make that boundary `FFP042` in CI.

Profiles expose only formula-cell, cached-result, missing-result, result-type,
and malformed-metadata counts. Result values, error text, result digests, and
formula-cell locations never enter profiles, Markdown, JSON, `FF042` details,
or SARIF. Equivalent finite numeric spellings and Boolean spellings are
normalized. Supported raw result kinds are numeric, string, Boolean, and error;
an absent or blank cache stays visible as a missing cache rather than being
invented. Unsupported or malformed formula-cache metadata creates an explicit
coverage warning instead of being silently ignored.

FormulaFence does not calculate or validate a formula result, decide whether a
result is stale or tampered, or model volatile, dynamic, external, or
calculation-engine dependencies. A legitimate recalculation can change a cache
without a statically visible input edit; `FF042` is therefore review evidence,
not a mathematical-correctness verdict. The boundary follows Microsoft's
[formula OOXML guidance](https://learn.microsoft.com/da-dk/office/open-xml/spreadsheet/working-with-formulas),
the Open XML `c` [cell](https://c-rex.net/samples/ooxml/e1/part4/OOXML_P4_DOCX_c_topic_ID0E1XM4.html)
and `ST_CellType` [definition](https://c-rex.net/samples/ooxml/e1/Part4/OOXML_P4_DOCX_ST_CellType_topic_ID0E6NEFB.html),
and Excel's [calculation-mode guidance](https://support.microsoft.com/en-US/Excel/change-formula-recalculation-iteration-or-precision-in-excel).

FormulaFence also compares **rich-text run controls**. A cell can keep the same
concatenated string while its character-level SpreadsheetML `<rPr>` formatting
changes: for example, a reviewer-facing “DO NOT APPROVE” phrase can be made
white inside an otherwise unchanged cell. FormulaFence follows both
relationship-backed shared strings and direct inline strings, privately
fingerprints run-property sequences, styled character boundaries, and phonetic
hints, and emits `FF043` when a material presentation control changes. Enable
`no_rich_text_run_changes` to make that boundary `FFP043` in CI.

Profiles expose only counts for referenced shared rich-text items/cells/runs,
inline rich-text cells/runs, phonetic runs/properties, and malformed controls.
Text, character formatting, colours, fonts, shared-string indexes, and cell
locations never enter profiles, Markdown, JSON, `FF043` details, or SARIF.
Equivalent rich-run property ordering, colour case, and explicit false Boolean
properties are normalized. A plain text edit inside an otherwise unchanged
run-property sequence remains the normal cell diff; moving a styled boundary
while the displayed text stays unchanged is guarded. Malformed, unsupported, or
unreadable rich-text metadata creates a visible coverage warning rather than a
silent omission.

FormulaFence does not render a cell, resolve theme colours, calculate
foreground/background contrast, infer whether a phrase is visible, preserve or
edit rich text, or guarantee equivalence across Excel versions. It compares
stored presentation declarations only. The boundary follows Microsoft's
[shared-string-table guidance](https://learn.microsoft.com/en-us/office/open-xml/spreadsheet/working-with-the-shared-string-table),
the Open XML `r` [rich-text-run definition](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.run?view=openxml-3.0.1),
and the `is` [inline-string definition](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.inlinestring?view=openxml-3.0.1).

FormulaFence also inventories ordinary **worksheet cell hyperlinks**. A cell's
friendly value can stay unchanged while its hyperlink redirects a reviewer to a
different external URL or file, jumps to a different in-workbook location,
changes its separate display override, or changes its ScreenTip. Standard
SpreadsheetML stores those declarations in worksheet `hyperlink` elements,
with an external target normally resolved through the worksheet relationship
part; Excel can additionally retain Office 2016 `xr:hyperlink` declarations.

FormulaFence reads the raw worksheet and relationship declarations before the
ordinary workbook reader can normalize them. It privately compares cell/range
binding, standard and revision declaration material, location, display and
ScreenTip values, and selected relationship type/target/mode semantics. A
material change emits `FF047`; enable `no_cell_hyperlink_changes` to make that
boundary `FFP047` in CI.

Profiles and `FF047` details expose only aggregate worksheet/hyperlink,
location/display/ScreenTip, relationship/external-relationship, and
malformed-metadata counts. Targets, cell references, locations, display
strings, ScreenTips, relationship IDs, and revision UIDs never enter profiles,
Markdown, JSON, or SARIF. Writer-chosen relationship IDs and revision UIDs,
relationship ordering, and equivalent internal-part target spelling are
normalized. Missing, duplicate, unbound, malformed, unsafe, unreadable,
oversized, or over-budget metadata becomes a visible coverage warning; XML
reads are bounded to 16 MiB per worksheet, 64 MiB per workbook, and 512 parts.
After raw inspection, the ordinary reader receives a hyperlink-removed
temporary copy so a malformed declaration cannot erase the review evidence.

This is a stored-declaration boundary, not link execution or reputation
analysis. FormulaFence does not render a hyperlink, resolve or fetch a target,
test availability, follow redirects, inspect linked content, infer trust-zone
or client behavior. Stored `HYPERLINK()` calls are separately covered by the
formula external-action ledger (`FF064`), but FormulaFence still never evaluates
their arguments or follows a resulting link. The raw worksheet-hyperlink scope
follows the Open XML
[Hyperlink](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.hyperlink?view=openxml-3.0.1)
and Office 2016
[Hyperlink](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.office2016.excel.hyperlink?view=openxml-3.0.1)
definitions, plus Microsoft's guidance for
[working with links in Excel](https://support.microsoft.com/en-US/Excel/work-with-links-in-excel).

FormulaFence also inventories Office 2010 **worksheet sparklines**. A compact
trend in a cell can be retargeted to a different source range, moved to a
different output cell, converted between line/column/win-loss presentation, or
have its axis, empty-cell, marker, hidden-data, and colour controls changed
without changing any ordinary cell value. SpreadsheetML stores these records in
an `x14:sparklineGroups` worksheet extension: each group can hold visual
controls and an optional date-axis source, while each nested sparkline holds a
source formula and destination `xm:sqref`.

FormulaFence reads that raw extension before the ordinary workbook reader drops
it. It privately compares group membership, source and date-axis formulas,
destination cells, type/axis/display/marker controls, line weight, and colour
definitions. A material change emits `FF048`; enable
`no_worksheet_sparkline_changes` to make that boundary `FFP048` in CI.

Profiles and `FF048` details expose only aggregate worksheet/group/sparkline,
source/date-axis-source, colour-control, and malformed-metadata counts. Source
formulas, destination cells, group properties, and colour definitions never
enter profiles, Markdown, JSON, or SARIF. Equivalent direct local-range
spelling, Boolean/numeric spelling, colour case, and declaration order are
normalized. Missing, duplicate, malformed, unreadable, oversized, or
over-budget metadata produces a visible coverage warning; raw reads are bounded
to 16 MiB per worksheet, 64 MiB per workbook, and 512 parts. After raw
inspection, FormulaFence removes only Sparkline Group extensions from its
temporary reader copy so lossy reader support cannot suppress evidence.

This is a stored-declaration boundary, not spreadsheet rendering or
calculation. FormulaFence does not calculate sparkline values, resolve names or
external sources, render a line/column/win-loss graphic, assess visual
accessibility, or guarantee cross-version Excel rendering equivalence. The
scope follows Microsoft's Open XML
[SparklineGroup](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.office2010.excel.sparklinegroup?view=openxml-3.0.1)
and [CT_Sparkline](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/6b28a993-e0fd-451d-860e-35097c6baa77)
definitions.

FormulaFence also inventories SpreadsheetML **XML Maps**. An XML map can pair
an embedded schema with refresh and export behavior, then route operational data
into XML table columns or individual cells. That configuration can change the
schema, XPath, file/connection binding, target cell, or refresh behavior while
ordinary cell values and formulas remain unchanged.

FormulaFence reads raw XML Maps, table XML-column-property, and single-cell
table parts and their workbook/worksheet relationships before ordinary workbook
readers discard or normalize the mapping surface. It privately compares map
schemas, map and data-binding behavior, mapped table columns, single-cell
bindings, and relationship targets. A material change emits FF049; enable
no_xml_mapping_changes to make that boundary FFP049 in CI.

Profiles and FF049 details expose only aggregate map/schema/binding,
file/connection, table, single-cell, and malformed-metadata counts. Schemas,
map names, XPath expressions, table identities, target cells, connection
identities, and relationship targets never enter profiles, Markdown, JSON, or
SARIF. Equivalent Boolean and unsigned-integer spellings, writer-selected
relationship IDs, relationship ordering, and equivalent internal target
spelling stay quiet. Missing, duplicate, malformed, unsafe, unbound,
unreadable, oversized, or over-budget metadata becomes a visible coverage
warning; raw reads are bounded to 16 MiB per part, 64 MiB per workbook, and 512
parts. Before any XML Map, mapped-table, or single-cell table tree is
materialized, FormulaFence streams a 32,768-element per-part and
65,536-element aggregate boundary. A successfully parsed structural overage
becomes visible `FF010`/`FF049` coverage evidence rather than a private tree.

This is a stored-declaration boundary, not XML data execution. FormulaFence
does not import or export XML, validate an XML instance against a schema, open a
file or connection, fetch remote data, calculate a refresh result, or infer
Excel client behavior. The scope follows the Open XML
[Map](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.map?view=openxml-3.0.1),
[XmlProperties](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.xmlproperties.xpath?view=openxml-3.0.1),
and [SingleXmlCells](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.singlexmlcells?view=openxml-3.0.1)
definitions.

FormulaFence also inventories Excel **rich data controls**. Rich data types can
keep linked entity values, provider-backed fields, web-image associations, and
worksheet value-metadata bindings outside ordinary cell values. Those
declarations can change a workbook's operational data surface while the usual
cell and formula diff stays quiet.

FormulaFence reads the raw Rich Value Data, structure, type, array, supporting
property-bag, style, web-image, and rich-value-relationship parts, along with
their workbook/package relationships and `XLRICHVALUE` worksheet metadata
bindings. It privately compares values, structures, provider-associated
metadata, web-image and rich-value relationship endpoints, and bound-cell
metadata. A material change emits `FF051`; enable
`no_rich_data_changes` to make that boundary `FFP051` in CI.

Profiles and `FF051` details expose only aggregate part, value, structure,
array, property-bag, metadata-binding, bound-cell, web-image, relationship,
external-reference, and malformed-metadata counts. Entity values, provider
data, field names, identifiers, URLs, image references, relationship IDs, and
bound-cell locations never enter profiles, Markdown, JSON, or SARIF.
Writer-selected relationship IDs/order and equivalent internal-target spelling
stay quiet. Missing, duplicate, malformed, unsafe, unreadable, oversized, or
over-budget metadata becomes a visible coverage warning; raw reads are bounded
to 16 MiB per XML part, 64 MiB per workbook, and 512 parts. Before raw
rich-data package XML is materialized, FormulaFence streams 32,768 elements per
part and 65,536 across the inventory; a successfully parsed structural overage
becomes visible `FF010`/`FF051` coverage evidence. The separate worksheet
binding pass streams only required cell attributes after the shared
semantic-reader preflight, so it does not retain a second worksheet XML tree.

This is a stored-data-control boundary, not provider execution or validation.
FormulaFence does not contact providers, refresh entity values, calculate
formulas, fetch web-image or other relationship targets, validate their
content, or infer Excel client behavior. The scope follows Microsoft's
[Rich Value Data](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/896934fd-8df7-43f4-b154-2d39371c270d),
[Rich Value Structure](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/d90f6d91-d868-4b94-9d26-ec3b1492cec6),
[Rich Value Types](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/5d213b66-3196-4516-b63c-eef80d926f4a),
and [Rich Value Web Image](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/4f3a80fd-1776-407f-8807-2497a4692dea)
definitions.

FormulaFence also inventories **custom workbook data stores**. Generic Custom
XML, add-in-owned opaque Custom Data, and custom document properties can retain
workbook-specific state while ordinary cells and formulas stay unchanged. That
state may contain an approval gate, a workflow decision, an integration
identifier, or another add-in setting that materially changes how a workbook
is used.

FormulaFence reads generic `customXml/item*.xml` data and their property/schema
parts and relationships, workbook-bound `xl/customData` property and binary
parts, and `docProps/custom.xml`. Power Query `DataMashup` Custom XML remains
under the existing Power Query guard, so it is not double counted. A material
change emits `FF052`; enable `no_custom_data_store_changes` to make that
boundary `FFP052` in CI.

Profiles and `FF052` details expose only aggregate custom-XML,
relationship/external-relationship, binary-data, document-property, and
malformed-metadata counts. Custom XML, schema URIs, property names and values,
storage IDs, binary payloads, relationship IDs, and targets never enter
profiles, Markdown, JSON, or SARIF. Writer-selected relationship IDs/order and
document-property `pid` stay quiet. Custom XML `itemID` and Custom Data `id`
storage identities are compared privately because an add-in can bind state to
them. Missing, duplicate, malformed, unsafe, unbound, unreadable, oversized,
or over-budget metadata becomes a visible coverage warning; raw reads are
bounded to 16 MiB per part, 64 MiB per workbook, and 512 parts. Before any
custom-state XML tree is materialized, FormulaFence also streams its complete
structure: 32,768 elements per XML part and 65,536 across the custom-state
scan. This is a FormulaFence CI allocation and coverage boundary, not an Excel
file-validity limit: a well-formed overage produces visible `FF052` coverage
evidence. Opaque binary Custom Data remains byte-bounded rather than being
treated as XML. The Power Query scanner receives only `DataMashup` members that
the same bounded pass classified safely, so rejected generic Custom XML is not
materialized a second time.

This is a stored-state boundary, not add-in execution. FormulaFence does not
execute an add-in, resolve a property, follow or fetch a relationship target,
interpret a binary payload, calculate formulas, or infer Excel client
behavior. The scope follows Microsoft's guidance on
[persisting add-in state](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/persisting-add-in-state-and-settings),
[Custom Data](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/7c53f6f4-fea8-43f7-a4b0-ba6e14d0eb78),
[Custom Data Properties](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/1f4aa666-c966-4ecf-8399-28390399c891),
and Excel's
[CustomDocumentProperties](https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.customdocumentproperties).

FormulaFence also inventories **digital-signature controls** that can change
outside ordinary cells and outside the `xl/vbaProject.bin` macro payload.
Excel distinguishes a package/content signature from a VBA project code
signature: either one can be added, removed, replaced, or rendered materially
different while formulas and visible values stay fixed.

FormulaFence reads the raw OPC package-signature graph before ordinary workbook
readers can discard or normalize it. It privately compares the package-root
signature-origin relationship, origin-to-XML-signature relationships, XMLDSIG
envelopes and signed references, embedded certificate values, certificate-part
relationships and payloads, and conventional VBA signature payloads and
relationships (`vbaProjectSignature.bin`, `vbaProjectSignatureAgile.bin`, and
`vbaProjectSignatureV3.bin`). A material change emits `FF050`; enable
`no_digital_signature_changes` to make that boundary `FFP050` in CI.

Profiles and `FF050` details expose only aggregate origin/XML-signature,
signed-reference, embedded-certificate/certificate-part, VBA-signature, and
malformed-metadata counts. Signature XML, reference URIs, certificate
identities and contents, binary signature payloads, relationship IDs, and
relationship targets never enter profiles, Markdown, JSON, or SARIF. Equivalent
relationship IDs/order and internal-target spelling, plus whitespace in XMLDSIG
base64 values, stay quiet. Missing, duplicate, malformed, unsafe, unbound,
unreadable, oversized, or over-budget metadata becomes a visible coverage
warning; raw signature reads are bounded to 16 MiB per part, 64 MiB per
workbook, and 512 parts. Before an XMLDSIG envelope is materialized,
FormulaFence streams a 32,768-element per-part and 65,536-element aggregate
boundary; a successfully parsed structural overage becomes visible
`FF010`/`FF050` coverage evidence. Certificate and VBA-signature binary
payloads remain byte-bounded rather than being interpreted as XML.

This is an envelope-integrity boundary, **not cryptographic validation**.
FormulaFence does not verify signature or digest values, XML transforms,
reference coverage, certificate chains, identity, trust, expiry, revocation,
timestamps, or the validity of signed VBA code. It never fetches certificates
or contacts a trust service. The scope follows Microsoft's
[OPC digital-signature overview](https://learn.microsoft.com/en-us/previous-versions/windows/desktop/opc/open-packaging-conventions-overview),
which places signer/trust validation with the package consumer, and Excel's
[separate workbook and VBA signing guidance](https://learn.microsoft.com/en-us/troubleshoot/microsoft-365-apps/excel/digital-signatures-code-signing).

FormulaFence also inventories legacy **Excel Notes** and the legacy placeholders
that can accompany modern threaded comments. Conventional Notes keep author
records, text, cell association, comment properties, and optional rich-text
material in a SpreadsheetML comments part rather than ordinary worksheet cells.
Their visibility and layout live separately in a VML drawing reached through a
worksheet `legacyDrawing` relationship. Excel can use a Note whose author is
marked `tc={GUID}` as a reconciliation placeholder for a threaded comment.

FormulaFence follows the worksheet-to-comments and worksheet-to-VML bindings,
then privately compares author association, text/presentation, cell binding,
comment properties, placeholder reconciliation state, Note visibility/layout,
and related relationship semantics. A material change emits `FF046`; enable
`no_legacy_comment_changes` for `FFP046` in CI. This boundary follows the Open
XML [Comment](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.comment?view=openxml-3.0.1),
[Authors](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.authors?view=openxml-3.0.1),
and [LegacyDrawing](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.legacydrawing?view=openxml-3.0.1)
definitions, plus Microsoft's [threaded-comment placeholder
rule](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/6383f002-c90b-401c-a1d7-66b97b14cb3e).

Profiles and `FF046` details expose only aggregate worksheet/part, author/
comment/text/rich-text/property/placeholder, VML Note-shape/visibility/anchor,
relationship, and malformed-metadata counts. Note text, authors, cell
references, raw VML, targets, relationship IDs, and GUIDs remain private.
Writer-chosen VML shape IDs, comment shape IDs, placeholder GUIDs (when
consistently rekeyed), and package relationship IDs are normalized. Missing,
duplicate, malformed, unsafe, unbound, unreadable, oversized, or over-budget
metadata becomes a visible coverage warning; XML reads are bounded to 16 MiB
per part, 64 MiB per workbook, and 512 parts. Before materializing a comments
or Note-VML tree, FormulaFence streams its complete structure: 32,768 elements
per part and 65,536 across the legacy-Note scan. This is a CI allocation and
coverage boundary, not a workbook-validity limit; a well-formed structural
overage becomes visible `FF010`/`FF046` evidence.

The scope is deliberately static: FormulaFence does not render Note text or
VML, resolve an author, fetch a relationship target, execute linked content,
calculate a client display position, or infer notification, account,
permission, cloud-sync, or client-visibility behavior. It only compares stored
package declarations.

FormulaFence also inventories modern **threaded comments**. Excel stores a
thread's top-level comment and replies outside the cell grid, together with
resolution state, mention bindings, and separate person records. Those records
can hold assumptions, instructions, review feedback, and approval context even
when every ordinary cell is unchanged. FormulaFence follows the documented
worksheet → threaded-comments and workbook → persons package relationships,
then privately compares the comment/reply graph, text, cell binding, stored
timestamp, `done` state, mention range/person association, extension material,
and person definitions. A material change emits `FF045`; enable
`no_threaded_comment_changes` for `FFP045` in CI. The package boundary follows
Microsoft's [threaded-comment overview](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/e0fb917a-1107-409a-852f-13b47aea70dc),
[Threaded Comments part](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/66e1875d-c60a-48eb-bf88-41066d45fea8),
and [Persons part](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/1a170d26-42a2-46f0-b2b6-0ff1dec1c344).

Profiles expose only worksheet/part/thread, comment/reply/resolved/text,
mention, person/unreferenced-person, binding-relationship, and malformed-
metadata counts. Comment bodies, cell references, timestamps, parent links,
person names/user IDs/provider IDs, relationship IDs, and GUIDs never enter
profiles, Markdown, JSON, `FF045` details, or SARIF. The scanner rebuilds the
thread structure and author/mention association before hashing, so consistent
writer-generated comment, parent, person, mention, and package relationship-ID
rewrites stay quiet. It also accepts equivalent Boolean spellings for resolution
state. Missing, duplicate, unsafe, unbound, malformed, unreadable, oversized,
or over-budget metadata becomes an explicit coverage warning; XML reads are
bounded to 16 MiB per part, 64 MiB per workbook, and 512 parts. Before any
comment or person tree is materialized, FormulaFence streams the complete XML
structure: 32,768 elements per part and 65,536 across the complete threaded-
comment scan. These are CI allocation and coverage limits, not workbook-
validity limits: a well-formed structural overage becomes visible `FF010` and
`FF045` evidence. After raw inspection, the ordinary workbook reader receives
a temporary copy with threaded-comment and person relationship bindings
removed, so it cannot re-materialize a rejected raw tree; the original package
and private raw evidence remain intact.

The scope is deliberately static and narrow: FormulaFence does not render
comments, validate that a mention's character offsets match text, send a
notification, resolve an account, fetch a relationship target, or determine
whether a thread's legacy placeholder is rendered. It does not infer
collaboration permissions, cloud-sync state, or whether a comment is visible in
a particular Excel client.
The documented [threaded-comment schema](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/adb84732-9fc8-48b6-bddc-6b0bcdaad940)
defines the stored `personId`, parent/reply, `done`, and mention structures;
FormulaFence compares those declarations without executing any collaboration
behavior.

FormulaFence also inventories non-chart **Worksheet DrawingML regular shapes,
connectors, groups, and SmartArt graphic frames**. A worksheet can point to a
DrawingML part whose `xdr:sp` text boxes carry visible text, formatting,
geometry, anchors, macro assignments, text links, descriptions, and click/hover
hyperlink bindings independently of a cell. Its `xdr:cxnSp` connectors can
visually attach one process or review shape to another, so an endpoint
reattachment can materially change a workbook's meaning without changing a
cell.

For a non-chart `xdr:graphicFrame`, FormulaFence recognizes the DrawingML
Diagram `a:graphicData` URI and requires one `dgm:relIds` declaration. It
privately compares the anchored frame plus the four explicitly bound SmartArt
components—data (`r:dm`), layout (`r:lo`), quick style (`r:qs`), and colours
(`r:cs`)—and direct worksheet-drawing `diagramDrawing` rendering parts. For a
Diagram Data part only, it also accepts a direct internal Image relationship
and fingerprints the raw target bytes without decoding them. The OOXML Diagram
Data Part model permits an explicit relationship only to an
[Image part](https://ooxml.info/docs/14/14.2/14.2.4/), and Microsoft's Open XML
SDK exposes those targets as
[`DiagramDataPart.ImageParts`](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.packaging.diagramdatapart.imageparts?view=openxml-2.8.1).
The standard worksheet → drawing relationship and transitional or Strict
DrawingML are supported. FormulaFence retains `xdr:sp`, `xdr:cxnSp`, nested
`xdr:grpSp`, and recognized SmartArt declarations under the supported anchors,
and compares private signatures for anchors, frame/shape/group/connector XML,
diagram material, bounded Diagram Data image bytes, connector endpoint
attachments, referenced relationship semantics, and structural text/macro/link
counts. A material change emits `FF044`; enable
`no_worksheet_drawing_shape_changes` for `FFP044`.

Profiles expose only worksheet/drawing/anchor, shape/text/connector/group,
graphic-frame/SmartArt-component, Diagram Data image part/fingerprinted/
uninspected, connector-attachment, text paragraph/run, macro/text-link/
hyperlink, relationship, and malformed-control counts. Shape and SmartArt
text, colours, geometry, anchors, diagram node IDs, connector target IDs,
macro names, formulas, descriptions, relationship IDs, image names/bytes,
targets, and raw XML never enter profiles or reports. Consistent non-visual and
connector-endpoint ID rewrites, worksheet-DrawingML relationship-ID rewrites,
and colour case normalize when semantics stay unchanged. Missing, duplicate,
malformed, unsafe, unreadable, oversized, over-budget, unsupported, or
external Diagram Data image material becomes explicit coverage evidence. XML
reads are bounded to 16 MiB per part, 64 MiB per workbook, and 512 parts;
direct Diagram Data image hashing is separately bounded to 32 MiB per image,
64 MiB per workbook, and 512 images.

The boundary is deliberately narrow: FormulaFence does not render DrawingML,
resolve themes or contrast, decide whether text is actually visible, calculate
a text-link formula, execute a macro assignment, fetch an external target,
calculate a final SmartArt layout, or decode/render media. It hashes only
bounded direct internal Image targets from a Diagram Data part; it does not
follow hyperlinks, second-hop targets, or any component-side relationship from
layout, quick-style, colours, or diagram-drawing parts. Those edges remain
coverage gaps. Chart graphic frames remain in `FF030`, and native `xdr:pic`
objects remain in the separate `FF059` image boundary below. Other non-chart
graphic-frame URI types are coverage gaps rather than inferred SmartArt. The
boundary follows Open XML's
[`xdr:sp` Shape definition](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.spreadsheet.shape?view=openxml-3.0.1),
[`xdr:cxnSp` ConnectionShape definition](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.spreadsheet.connectionshape?view=openxml-3.0.1),
and Microsoft's [Graphic Object Data](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oe376/f58e82a5-5590-4e36-b178-e12989960415)
and [Diagram relationship IDs](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.diagrams.relationshipids?view=openxml-3.0.1)
references. Microsoft's [Excel shape API overview](https://learn.microsoft.com/en-us/office/dev/add-ins/excel/excel-add-ins-shapes)
also documents connectors that attach to shapes and move with them; XlsxWriter's
documented [worksheet text boxes](https://xlsxwriter.readthedocs.io/working_with_textboxes.html)
remain an independent regular-shape writer surface.

FormulaFence separately inventories native **worksheet image controls**. A
floating DrawingML `xdr:pic`, a worksheet `<picture>` background, or a
header/footer VML image can change what a reviewer sees or prints while every
cell stays fixed. It follows worksheet `drawing`, `picture`, and
`legacyDrawingHF` relationships, recognizes transitional and strict DrawingML,
and compares private anchor/picture/VML declarations, relationship semantics,
and bounded direct image payload hashes. A material change emits `FF059`;
enable `no_worksheet_image_changes` for `FFP059`.

Profiles expose only worksheet, picture/anchor, background, header/footer,
image-payload, relationship, and malformed-control counts. Image bytes, image
names/descriptions, visual formatting, anchors, relationship IDs/targets, and
raw XML never enter profiles or reports. Non-visual DrawingML and VML IDs plus
consistent relationship-ID rewrites normalize when semantics stay unchanged.
Missing, duplicate, malformed, unsafe, unreadable, oversized, or over-budget
metadata becomes an explicit coverage warning. XML reads are bounded to 16 MiB
per part, 64 MiB per workbook, and 512 parts; direct image payload hashes are
bounded to 32 MiB per part, 64 MiB per workbook, and 512 parts.

The image boundary does not render or decode media, fetch an external target,
resolve themes, calculate visibility/cropping/z-order, compose controls, or
calculate final pagination. It intentionally leaves charts to `FF030`, rich
data/in-cell images to `FF051`, Theme images to `FF053`, ActiveX/OLE image
controls to `FF029`, regular/group/connector/SmartArt drawing controls to
`FF044`, and
header/footer text to `FF056`. The package model follows Open XML's
[`xdr:pic` Picture definition](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.spreadsheet.picture?view=openxml-3.0.1),
Microsoft's [worksheet background guidance](https://support.microsoft.com/en-us/excel/add-or-remove-a-sheet-background),
and [header/footer watermark guidance](https://support.microsoft.com/en-us/excel/get-started/add-a-watermark-in-excel).

FormulaFence also inventories relationship-backed **worksheet controls, legacy
VML form controls, and OLE objects**. A worksheet can bind modern `<control>`
markup to persisted ActiveX state, nested form-control properties, or raw
OLE/package data without changing an ordinary cell or the VBA payload. It can
also point to a legacy VML drawing whose non-comment `ClientData` records carry
macro assignments, linked cells, source ranges, or camera source ranges.
FormulaFence follows these package relationships, compares the relevant modern
or VML control definitions privately, and hashes only bounded direct ActiveX
binary and OLE/package payloads. Ordinary VML comment notes are intentionally
excluded.

A material change emits `FF029`; enable
`no_worksheet_embedded_control_changes` for `FFP029`. Profiles expose only safe
counts. Control names, class IDs, licenses, captions, macros, formulas/ranges,
OLE identities, relationship targets, XML, and payload bytes never enter
profiles or reports. FormulaFence never initializes a control, deserializes or
opens an OLE object/package, renders a VML drawing, follows an external target,
or infers control event dispatch. Relevant XML reads are bounded to 16 MiB per
part, 64 MiB per workbook, and 512 parts. Before private XML canonicalization,
FormulaFence streams complete XML structure: 32,768 elements per part and
65,536 across the embedded-control scan. A well-formed structural overage is
visible `FF010`/`FF029` coverage evidence. Direct raw payload hashes are
bounded to 32 MiB per part, 64 MiB per workbook, and 512 parts. Malformed,
orphaned, unbound, oversized, or over-budget material remains an explicit
coverage warning. This scope follows Microsoft's guidance on [sheet ActiveX
controls](https://learn.microsoft.com/en-us/office/vba/excel/concepts/controls-dialogboxes-forms/using-activex-controls-on-sheets),
the [`ocx` ActiveX persistence
schema](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oe376/b30a660a-95eb-4716-b201-a46aae788610),
[form-control properties](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/3d054a6d-4f94-4082-837a-f939fd8d4a45),
and VML [`ClientData`](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.vml.spreadsheet.clientdata?view=openxml-3.0.1)
with [macro](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oi29500/fdd83507-7a57-4bf1-b844-66f551ee55b9)
and [list-range](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oi29500/3d0c9716-88c5-4af3-b63a-feef60b8ebd8)
bindings.

FormulaFence also inventories **Power Query Data Mashup** custom XML parts.
It reads the documented length-prefixed container, fingerprints the embedded
`Section1.m` formula document and logical package material privately, and
compares structural query metadata plus formula-firewall permissions. Its
profile exposes only safe counts and controls—not M text, query names, source
locations, metadata values, embedded content, telemetry IDs, or user-bound
permission bindings. `sqmid` telemetry and result-only refresh metadata are
ignored to avoid writer and refresh noise. A material query-definition or
semantic-control change emits `FF024`; enable `no_power_query_changes` for
`FFP024`. FormulaFence does not execute M, refresh a connection, or infer the
returned data. The implementation follows Microsoft's
[Query Definition File Format](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-qdeff/27b1dd1e-7de8-45d9-9c84-dfcc7a802e37),
which stores Power Query definitions in a Custom XML part.

The outer Custom XML safety gate does not itself bound XML decoded from the
Data Mashup stream. FormulaFence therefore streams each metadata and
formula-firewall permission document before private parsing, allowing 32,768
elements per document and 65,536 across the Power Query scan. This is a CI
allocation and coverage boundary, not a Power Query file-validity rule: a
successfully parsed overage becomes visible `FF010`/`FF024` evidence while
malformed input keeps its established diagnostic.

Every nested Data Mashup ZIP is first scanned as bounded central-directory
metadata before Python can materialize a ZIP entry catalog. This covers both
the logical package and metadata embedded-content ZIPs, limits each source to
768 KiB and raw member names to 1 KiB, and shares a 512-part budget across the
complete Power Query scan. ZIP64, multi-disk, malformed, and filename-rewriting
metadata (Unicode-path aliases, NULs, or platform separators) stays visible as a
coverage gap rather than bypassing that boundary.
The logical package alone is eligible for content reads; its members must be
stored or deflated, can expand to at most 16 MiB each, must stay at or below a
1,000:1 ratio, and share 64 MiB of declared expanded data. If either nested ZIP
exceeds its safety boundary, FormulaFence retains only a private opaque
fingerprint and raises an explicit coverage warning, so a changed candidate
remains visible through `FF010`/`FF024` rather than being treated as safely
equivalent.

FormulaFence also follows a call to a workbook- or worksheet-local defined name
when its complete definition is one statically resolvable `LAMBDA` expression.
For `=ToCelsius(A2)`, the caller keeps its explicit `A2` edge and gains the
function's static internal dependencies; nested named-LAMBDA calls and
formula-defined names that call a named LAMBDA are resolved the same way. It
recognizes both human-authored formulas and the `_xlfn.LAMBDA` / `_xlpm.` /
`_xlop.` OOXML spelling produced by Excel-compatible writers. Definition scope
and worksheet-local precedence are preserved. Relative, cyclic, dynamic, 3-D,
tokenizer-unsupported, or otherwise non-static LAMBDAs remain a visible
unresolved reference at each call site. In portfolio mode, the narrow global
named-LAMBDA external-endpoint bridge described above is the exception: it
retains only fully static endpoint and internal-input edges at a real function
call, never through a bare name. Spill extents and blockers, plus arbitrary
VBA, add-in, or other custom functions, remain coverage limits rather than
inferred dependencies.

If the underlying formula tokenizer cannot inspect a formula at all,
FormulaFence records the affected cell in the profile and reports `FF016` when
it is newly introduced. `no_new_tokenization_failures` turns that condition
into a CI failure instead of silently omitting its dependency graph.

FormulaFence also expands internal static 3-D A1 references such as
`Jan:Mar!B2:B10` over every worksheet tab between the named endpoints. Profiles
identify those formula cells. If an unchanged 3-D formula's tab span changes
because a sheet is inserted, removed, or moved, FormulaFence emits `FF014`; the
`no_3d_reference_scope_changes` policy rule can make that a hard boundary.
Unknown endpoints and non-A1 constructs remain explicit coverage cases rather
than invented dependencies. Exact static external 3-D A1 spans are separately
eligible for the candidate-only portfolio graph when their source candidate has
a complete consistent tab catalog; other external 3-D forms remain external-link
hazards. Exact external table selectors are separately eligible only after an
exact source candidate yields one matching table and static source-cell bounds;
row-relative and source-sheet-qualified table forms remain coverage gaps. This
follows Excel's
[3-D-reference behavior](https://support.microsoft.com/en-us/excel/create-a-3-d-reference-to-the-same-cell-range-on-multiple-worksheets).

## Development

```bash
python -m venv .venv
.venv/bin/python -m pip install -e '.[dev]'
.venv/bin/python -m pytest
.venv/bin/python -m ruff check .
```

## Safety and scope

FormulaFence reads workbook structure only. It does not recalculate formulas,
run VBA, follow external links, or claim that a workbook's numbers are correct.
It is a review and control layer; human review remains essential for material
models.

Policy input is bounded before FormulaFence opens either workbook: it accepts
one UTF-8 YAML document with ordinary mappings, lists, and scalars, while
rejecting duplicate keys, anchors, aliases, and merge keys. The policy limits
are 1 MiB of source, 4,096 YAML nodes, 64 nesting levels, 4,096 characters per
scalar, and 512 selectors in each selector list. FormulaFence reads that source
through one descriptor, requests nonblocking mode where the host supports it,
and verifies that the opened object is a regular file before parsing; a policy
pathname replaced with a FIFO or device after the initial check fails closed on
hosts that provide nonblocking descriptor opens rather than stalling CI. See
the [policy reference](docs/policy.md#policy-file-safety-and-syntax) for the supported
syntax and the [threat model](docs/threat-model.md) for the CI boundary.

When a CLI command writes a report, or `init --force` writes a starter policy,
FormulaFence first creates a private temporary file alongside the requested
destination and then atomically replaces the final directory entry. This keeps
a post-check symlink or hard-link swap from redirecting the write into an
inspected input: the link itself is replaced. Ordinary `init` instead uses an
atomic no-replace link publication: if any final directory entry appears after
its initial absence check, the command fails and leaves that entry untouched.
That keeps the default “use `--force` to replace it” promise true across the
publication race. Parent-directory access remains part of the caller's CI
workspace integrity boundary.

For each snapshot, FormulaFence first materializes one bounded private copy of
the regular workbook source. The archive preflight, semantic-reader gate, raw
OOXML scanners, ordinary workbook reader, and snapshot `sha256` all operate on
that one inspected copy; the snapshot's visible path remains the path supplied
by the caller. A later replacement of that pathname therefore cannot mix
preflight evidence from one workbook with report evidence from another. This
direct-snapshot process is not a lock on a producer that edits a file in place:
use atomic artifact handoff and isolated workspace permissions when
source-producer integrity is a requirement.

Portfolio scans additionally bind each later source opening to the regular file
and state observed while they built the bounded directory inventory. A source
that changes in place or is replaced after inventory is reported as unreadable,
so a report cannot silently switch to a newly named workbook between membership
review and semantic inspection.

Every `.xlsx` and `.xlsm` source passes a fail-closed OOXML archive preflight
before FormulaFence reads an OOXML part or opens the workbook reader. The
preflight does not extract the archive. It accepts only one canonical,
single-disk ZIP container with stored or deflated members, then rejects
duplicate or case-colliding paths, ZIP Unicode-path aliases, unsafe paths,
encrypted or special-file members, inconsistent local headers, and overlapping
member payloads. The fixed resource ceilings are 1 GiB for the source package,
32 MiB for its central
directory, 4,096 members, 512 MiB per expanded member, 768 MiB in aggregate,
and a 1,000:1 maximum member compression ratio. These are input-safety limits,
not a malware classification or a substitute for isolated CI runners. See the
[threat model](docs/threat-model.md) for the complete boundary.

After that header-only ZIP check, a semantic-reader preflight caps every
XML/relationship part at 64 MiB, aggregate XML material at 256 MiB, and streams
the reader-visible manifest, workbook, styles, shared strings, and bounded
workbook-selected sheets before FormulaFence starts a complete in-memory reader
or downstream raw OOXML scanning. Before XML parser construction, it bounds
each physical opening tag and every non-character-data lexical token to 128
KiB: comments, processing instructions, declarations, end tags, and entity
references cannot force an unbounded parser-side value. Document-type
declarations are explicitly forbidden in the shared defused parser. Decoded
ordinary text, tails, and CDATA remain separately limited to 1 MiB per node.
It allows at most 4,000,000 XML elements per streamed part and 256 nesting
levels, 500,000 populated SpreadsheetML cells,
16,384 reader-materialized row-dimension declarations, 16,384 column-dimension
declarations, 4,096 direct column-dimension containers, 500,000 shared-string
entries, 65,490 effective `cellXfs` styles, 32,767
characters of cell text, and 8,192 characters of stored formula/defined-name
text. It also
caps the bootstrap catalog at 4,096 content-type declarations, 4,096 workbook
relationships, 512 workbook sheet declarations (including repeated
declarations of one target part), and 100,000 direct workbook defined-name
declarations. It also caps direct workbook external-reference and pivot-cache
declarations at 4,096 each, matching the bounded relationship catalog they
select; direct workbook book-view, custom-workbook-view, function-group,
smart-tag-type, and web-publish-object catalogs are likewise capped at 4,096.
Direct legacy custom sheet-view declarations are capped at 4,096 across the
workbook-selected sheet parts before FormulaFence builds raw Custom View
records. A row declaration counts toward the row-dimension budget only when it
has an unqualified attribute other than `r` or `spans`, matching the condition
that makes `openpyxl` retain a `RowDimension`; namespace-qualified extension
attributes do not consume that budget. Every reader-visible SpreadsheetML
`col` declaration consumes the separate 16,384 column-dimension budget because
the reader dispatches it before resolving whether its attributes matter; direct
`cols` containers are also capped at 4,096 because raw dimension scanners
retain them. Direct merged-cell declarations across the reader-selected ordinary
worksheet parts are capped at 4,096; each merge range and their aggregate
expanded coordinate area are capped at 100,000 cells, and a merge reference at
256 characters, before `openpyxl` expands every coordinate into a `MergedCell`.
Data-validation declarations, conditional-formatting declarations and rules,
and Scenario Manager containers, scenarios, and input cells are likewise capped
at 4,096 each across reader-selected ordinary worksheet parts. Their `sqref`
fields are capped at 128 KiB and 4,096 target ranges per field, with 8,192
targets in aggregate for each catalog, before `openpyxl` creates a `CellRange`
object for every target; data-validation and conditional-formatting formulas
also follow the 8,192-character stored-formula ceiling. The catalog counters
follow the reader's local-name behavior, so
alternate-namespace entries cannot bypass a limit. The cell-text, formula, and
effective-cell-style ceilings match [Excel's published limits](https://support.microsoft.com/en-US/Excel/excel-specifications-and-limits).
Shared strings retain their broad 500,000-entry allowance for ordinary simple
values, but each complete `si` item is limited to 32,768 XML elements; complex
items share a 65,536-element budget, and ignored opaque direct `sst` children
are limited to 32,768 elements per selected table and 65,536 in aggregate.
The raw rich-text scanner streams one direct shared-string item at a time and
releases unrelated root children as it goes, so a compact extension tree cannot
first force either scanner to retain the full table. These are CI allocation
limits, not SpreadsheetML validity rules.
The reader ceilings count input structure, but a static formula-derived graph can
still fan out after that reader phase: one compact formula-defined name may
resolve to many local references at each caller. FormulaFence therefore also
caps the retained local dependency records with `--max-dependency-edges`
(2,000,000 by default). Each direct local dependency, compact local range
dependency, and additional fixed-CSE or observed dynamic-array output alias
consumes one record; it neither expands a range into cells nor constrains the
separate candidate-only cross-workbook graph. An overage returns status `2`
before a CLI report is rendered or published.
Formula-defined names use a separate 1,000,000-state temporary propagation
budget by default (`--max-formula-defined-name-states`). It reserves direct
sensitive-call ledger entries, direct name-to-name marker dependencies, and
each component's inherited ledger before FormulaFence retains it. That keeps a
compact acyclic chain of action, DDE, custom-function, registration, XLM, or
environment-information calls from repeatedly materializing every prefix. It
does not deduplicate genuine separate runtime calls, alter normal dependency
records, or replace the candidate-only cross-workbook graph limit. An overage
returns status `2` before a CLI report is rendered or published.
Formula-defined-name lookup itself reuses compact, scope-aware live overlays
instead of rebuilding the complete visible name catalog for each definition.
This preserves worksheet-local shadowing, qualified-name visibility, and the
explicit unresolved named-`LAMBDA` coverage signal while avoiding catalog
materialization work that otherwise grows quadratically with a large valid
defined-name list. Its eleven private safety-marker kinds are generated only
when a formula inspection resolves a name and are recovered only from bounded,
canonical markers, so an otherwise inactive catalog does not retain an
eleven-fold marker ledger. It does not relax the actual formula-token or
dependency work required by a definition.
For reader-selected transitional or Strict worksheets, the preflight also
matches the published base Worksheet root-child grammar. A direct root subtree
outside that grammar is limited to 32,768 XML elements per worksheet and
65,536 in aggregate before raw worksheet scanners or the ordinary reader can
retain it. A SpreadsheetML `extLst` is a named extension container rather than
ordinary base content, so every extension-list subtree in a selected worksheet
is separately limited to 32,768 XML elements per worksheet and 65,536 in
aggregate. Ordinary `sheetData` and other named base controls keep their
existing specialized budgets. These are CI allocation limits for opaque root
and extension content, not SpreadsheetML validity rules.
Every relationship-selected Chartsheet and Dialogsheet is separately bounded
as a non-grid sheet: its complete XML tree allows 32,768 elements per part and
65,536 across selected parts before raw control readers or the workbook reader
can materialize it. This covers their documented `extLst` containers and opaque
content while chart DrawingML remains under its dedicated structural boundary.
The bootstrap `xl/workbook.xml` part is also read into a complete tree. Its
documented [Workbook extension-list location](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.workbook?view=openxml-3.0.1)
and every nested `extLst` subtree permit 32,768 XML elements before a raw or
ordinary workbook reader starts; a foreign direct workbook-root subtree has the
same separate budget. The check follows the reader's `workbook`/`extLst` local-
name dispatch so alternate namespaces cannot bypass it. Named Workbook controls
such as `sheets`, `definedNames`, and views remain on their existing
format-aware catalog limits. These are CI allocation limits, not SpreadsheetML
validity rules.
The stylesheet reader also constructs a complete `xl/styles.xml` tree. Its
documented [Stylesheet](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.stylesheet?view=openxml-3.0.1)
and [ExtensionList](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.extensionlist?view=openxml-3.0.1)
controls, including every nested local-name `extLst` subtree, allow 32,768
elements. A foreign direct root subtree, a non-`styleSheet` root local name,
and an ignored direct child inside a named catalog each have the same separate
budget.
Every materialized direct style record permits 32,768 non-extension descendants
and those records share a 262,144-element budget. The existing format-aware
number-format, font, fill, border, base-XF, named-style, differential-style,
palette, table-style, table-style-element, and extension catalog limits remain;
effective `cellXfs` retains its separate 65,490-style ceiling. These checks
follow the reader's local-name and nested-sequence behavior, so a namespace
variation or repeated ordinary-looking child cannot evade the boundary. Before
the raw shape, native-image, in-content Office Web Add-in, worksheet-chart, or
ordinary workbook readers can materialize a shared DrawingML tree. The
preflight follows direct internal `drawing` relationships from selected
transitional or Strict worksheet parts. It streams every unique XML target under a 32,768-element
per-part and 65,536-element aggregate ceiling. A successfully parsed overage
returns the stable safety-preflight error rather than a partial metadata report;
missing, malformed, or non-XML optional targets keep their existing coverage
diagnostics, and orphan DrawingML parts remain outside this relationship-based
boundary. These are CI allocation limits, not Excel workbook-validity limits.
Before raw filter, Named Sheet View, external-data, XML Mapping, Table Style, or
ordinary workbook readers can materialize a Table Definition tree, the
preflight streams every canonical `xl/tables/*.xml` part (including an orphan
part that the Table Style scanner inventories) and every safe direct internal
worksheet `table` relationship target, including Strict relationships and
noncanonical targets. Those table-definition targets allow 32,768 XML elements
per part and 65,536 in aggregate. A successfully parsed overage returns the
stable safety-preflight error rather than a partial report; malformed, missing,
or non-XML optional targets retain their downstream coverage diagnostics.
These are CI allocation limits, not Excel workbook-validity limits.
Valid workbooks above FormulaFence's separate CI-oriented cardinality bounds
are deliberately rejected rather than partially inspected. FormulaFence
requires `defusedxml` for its XML parser, which also enables `openpyxl`'s
defused XML path in the supported installation. Partition unusually large data
workbooks before reviewing them rather than relying on an unbounded worker
allocation.

## License

[MIT](LICENSE)
