# Policy reference

FormulaFence keeps controls in a small YAML file so that the rule itself is
reviewable alongside the model. A policy is evaluated by `formulafence check`;
each violation is emitted as a `FFP…` finding and makes the command exit with
status `1`.

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
  no_sheet_visibility_changes: true
  max_changed_formulas: 20
  max_downstream_impact: 100

protected_cells:
  - Dashboard!B12
  - 'Debt Schedule'!F10:F24

allowed_changes:
  - Inputs!B2:B100
```

FormulaFence rejects unknown fields and rules. This is deliberate: a misspelled
control should fail closed rather than silently weaken review.

## Root fields

| Field | Type | Meaning |
| --- | --- | --- |
| `version` | integer | Required. The current schema is `1`. |
| `rules` | mapping | Optional control switches and limits. |
| `protected_cells` | list of selectors | Optional cells/ranges that may not change. |
| `allowed_changes` | list of selectors | Optional edit allow-list. When present and non-empty, every semantic cell change must match one selector. |

Selectors are sheet-qualified A1 cells or ranges: `Inputs!B2`,
`Inputs!B2:B100`, or `'Debt Schedule'!F10:F24`. The sheet name is
case-insensitive; quotes are required when it contains spaces or punctuation.

## Rules

| Rule | Type | Fails when |
| --- | --- | --- |
| `no_formula_to_value` | boolean | A formula is replaced with a non-formula value. |
| `no_new_external_links` | boolean | A formula adds a statically visible external-workbook reference. |
| `no_new_broken_references` | boolean | A formula adds `#REF!`. |
| `no_macro_changes` | boolean | The `xl/vbaProject.bin` payload is added, removed, or has a different SHA-256. |
| `no_new_parser_warnings` | boolean | The candidate introduces an unsupported-workbook coverage warning. |
| `no_new_unresolved_references` | boolean | A formula adds a name, table reference, or other range token that cannot be resolved statically. |
| `no_new_dynamic_references` | boolean | A formula adds a dynamic reference function such as `INDIRECT` or `OFFSET`. |
| `no_table_definition_changes` | boolean | An Excel table is added, removed, moved, renamed, or has its columns/header/total-row configuration changed. |
| `no_sheet_visibility_changes` | boolean | A sheet becomes visible, hidden, or very hidden. |
| `max_changed_formulas` | non-negative integer | More formula-bearing cells change than allowed. |
| `max_downstream_impact` | non-negative integer | A changed cell reaches more downstream formula cells than allowed. |

All booleans default to `false`; limits default to unset. Start narrowly for a
single material workbook, then expand policy only after reviewing the model's
actual change patterns.

Ordinary workbook and sheet-local names with static A1 destinations are resolved
into the dependency graph. FormulaFence also resolves a conservative Excel-table
subset: a table name, a column or contiguous column range,
`#All`/`#Data`/`#Headers`/`#Totals`, and provably row-scoped references. An
unqualified `[Column]` or `[@Column]` is resolved only when its formula cell is
inside one table's data body. A qualified `Table[@Column]` or
`Table[[#This Row],[Column]:[Other Column]]` is resolved when the formula is on
the named table's data row, including an adjacent cell on that worksheet. The
coverage controls are for remaining cases—such as named formulas, header/total
row current-row syntax, exotic bracket escapes, and dynamic address
construction—where FormulaFence intentionally does not guess at dependencies.

## Exit status

| Status | Meaning |
| --- | --- |
| `0` | No policy violation (and no selected `--fail-on` threshold reached). |
| `1` | One or more policy violations, or `--fail-on` was reached. |
| `2` | Invalid policy, unreadable workbook, unsupported format, or output error. |

## Suggested rollout

1. Start with `formulafence diff approved.xlsx candidate.xlsx --format markdown`.
2. Commit a policy that only protects headline outputs and bans new broken/external links.
3. Run `check` in non-blocking CI for a few review cycles and tune `max_*` limits.
4. Make the check required once the report matches the team's real review process.

Do not use an allow-list as a substitute for reviewing a material change. It is
most useful for separating designated input blocks from calculation and output
areas.
