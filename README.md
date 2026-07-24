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
python -m pip install https://github.com/SybilGambleyyu/formulafence/releases/download/v0.5.0/formulafence-0.5.0-py3-none-any.whl

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
  no_table_definition_changes: true
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
| Impact trace | Downstream formula cells and deterministic shortest dependency-path samples, including cross-sheet, static named-range, and basic Excel-table references |
| Formula-pattern break | An edited formula that no longer matches equal neighboring formulas |
| Workbook controls | Sheet visibility, defined names, Excel-table definitions, calculation settings, and VBA payload changes |
| Formula hazards | New external-workbook references and `#REF!` formulas |
| Coverage changes | New parser warnings, unresolved formula references (including unsupported table syntax), and dynamic-reference functions (`INDIRECT`/`OFFSET`) |
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
