# FormulaFence

FormulaFence is a local-first spreadsheet change-assurance CLI. It makes `.xlsx`
changes reviewable in CI: compare workbook semantics, trace downstream formula
impact, detect high-risk edits, and enforce a small policy file before a model
is shared or merged.

It never executes formulas or macros, and it does not upload workbook contents.

> Status: early alpha. The first release supports `.xlsx` and `.xlsm` inspection
> with formula-aware diffs, dependency impact, policy checks, Markdown/JSON/SARIF
> reports, and deterministic evidence metadata.

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
python -m pip install https://github.com/SybilGambleyyu/formulafence/releases/download/v0.11.0/formulafence-0.11.0-py3-none-any.whl

# Readable review report
formulafence diff baseline.xlsx candidate.xlsx --format markdown

# Enforce a policy in CI (non-zero when a rule fails)
formulafence check baseline.xlsx candidate.xlsx --policy formulafence.yml --format sarif --output results.sarif
```

FormulaFence is not yet published to PyPI; the direct release URL above avoids
an ambiguous package-name install. See [GitHub Releases](https://github.com/SybilGambleyyu/formulafence/releases)
for the current version.

Create `formulafence.yml`:

```yaml
version: 1
rules:
  no_formula_to_value: true
  no_new_external_links: true
  no_new_broken_references: true
  no_macro_changes: true
  no_new_parser_warnings: true
  no_new_unresolved_references: true
  no_new_dynamic_references: true
  no_new_spill_references: true
  no_new_implicit_intersections: true
  no_new_tokenization_failures: true
  no_table_definition_changes: true
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
| Impact trace | Downstream formula cells and deterministic shortest dependency-path samples, including cross-sheet, static named ranges, formula-defined names, static named `LAMBDA` calls, `LET`/inline-`LAMBDA`, Excel-table, and 3-D worksheet references |
| Formula-pattern break | An edited formula that no longer matches equal neighboring formulas |
| Workbook controls | Sheet visibility, defined names, Excel-table definitions, static 3-D-reference scope, calculation settings, and VBA payload changes |
| Formula hazards | New external-workbook references and `#REF!` formulas |
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

FormulaFence also follows a call to a workbook- or worksheet-local defined name
when its complete definition is one statically resolvable `LAMBDA` expression.
For `=ToCelsius(A2)`, the caller keeps its explicit `A2` edge and gains the
function's static internal dependencies; nested named-LAMBDA calls and
formula-defined names that call a named LAMBDA are resolved the same way. It
recognizes both human-authored formulas and the `_xlfn.LAMBDA` / `_xlpm.` /
`_xlop.` OOXML spelling produced by Excel-compatible writers. Definition scope
and worksheet-local precedence are preserved. Relative, cyclic, external,
dynamic, 3-D, tokenizer-unsupported, or otherwise non-static LAMBDAs remain a
visible unresolved reference at each call site. Spill extents and blockers,
plus arbitrary VBA, add-in, or other custom functions, remain coverage limits
rather than inferred dependencies.

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
than invented dependencies; external 3-D references remain external-link
hazards. This follows Excel's
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

## License

[MIT](LICENSE)
