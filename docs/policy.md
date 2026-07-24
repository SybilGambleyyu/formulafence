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
  no_new_spill_references: true
  no_new_dynamic_array_output_references: true
  no_new_implicit_intersections: true
  no_array_formula_semantics_changes: true
  no_new_tokenization_failures: true
  no_table_definition_changes: true
  no_3d_reference_scope_changes: true
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
| `no_new_unresolved_references` | boolean | A formula adds a name, named-LAMBDA call, table reference, or other token that cannot be resolved statically. |
| `no_new_dynamic_references` | boolean | A formula adds a dynamic reference function such as `INDIRECT` or `OFFSET`. |
| `no_new_spill_references` | boolean | A formula adds a dynamic-array spill reference; FormulaFence traces its anchor but not its variable extent or blockers. |
| `no_new_dynamic_array_output_references` | boolean | A formula newly intersects a non-anchor member of an OOXML-observed dynamic-array output range. |
| `no_new_implicit_intersections` | boolean | A formula adds explicit `@` / `SINGLE()` implicit intersection, which can change which value a range or array contributes. |
| `no_array_formula_semantics_changes` | boolean | A legacy-CSE or dynamic-array formula is added, removed, or changes mode, or a legacy CSE formula's fixed output range changes. |
| `no_new_tokenization_failures` | boolean | A formula is newly introduced that the underlying formula tokenizer cannot inspect. |
| `no_table_definition_changes` | boolean | An Excel table is added, removed, moved, renamed, or has its columns/header/total-row configuration changed. |
| `no_3d_reference_scope_changes` | boolean | The worksheet span of an unchanged static 3-D formula changes because tab order or membership changed. |
| `no_sheet_visibility_changes` | boolean | A sheet becomes visible, hidden, or very hidden. |
| `max_changed_formulas` | non-negative integer | More formula-bearing cells change than allowed. |
| `max_downstream_impact` | non-negative integer | A changed cell reaches more downstream formula cells than allowed. |

All booleans default to `false`; limits default to unset. Start narrowly for a
single material workbook, then expand policy only after reviewing the model's
actual change patterns.

Ordinary workbook and sheet-local names with static A1 destinations are resolved
into the dependency graph. FormulaFence also expands a formula-defined name
when every dependency in its definition is statically visible and internal (or
when the definition is a constant), including nested workbook and sheet-local
names. FormulaFence also resolves a conservative Excel-table subset: a table
name, a column or contiguous column range,
`#All`/`#Data`/`#Headers`/`#Totals`, and provably row-scoped references. An
unqualified `[Column]` or `[@Column]` is resolved only when its formula cell is
inside one table's data body. A qualified `Table[@Column]` or
`Table[[#This Row],[Column]:[Other Column]]` is resolved when the formula is on
the named table's data row, including an adjacent cell on that worksheet. The
coverage controls are for remaining cases—such as relative, cyclic, external,
dynamic, 3-D, or tokenizer-unsupported formula-defined names; header/total-row
current-row syntax; exotic bracket escapes; and dynamic address construction—
where FormulaFence intentionally does not guess at dependencies.

Within a formula, ordinary `LET` bindings and inline `LAMBDA` parameters are
treated as lexical local names rather than unresolved workbook names. This
preserves the static dependencies in their value expressions and bodies,
including nested lambdas in higher-order Excel functions.

FormulaFence recognizes a direct internal static spill anchor written as
`A1#` or its OOXML `ANCHORARRAY(A1)` representation. It adds the anchor cell
to the dependency graph and records the spill token in the profile; `FF015` and
`no_new_spill_references` let CI reject a newly introduced partial-coverage
case. It deliberately does not infer a variable spill extent or every cell
that could block it. External, 3-D, range, named, implicit-intersection, and
malformed spill forms remain coverage limits. Formula-defined names containing
a spill reference are left unresolved at their call sites so this boundary
cannot be hidden behind a name.

FormulaFence separately inventories explicit implicit intersection: literal
`@A1:A3`, `@` applied to a function result, and persisted
`_xlfn.SINGLE(...)`. When `SINGLE()` wraps one direct static A1 cell or range
with an unambiguous row/column intersection, the formula location selects a
precise cell edge; otherwise FormulaFence keeps the visible static inputs
conservatively and never evaluates the expression.
New instances emit `FF017`; enable `no_new_implicit_intersections` for
`FFP017`. This does not change the table-specific `[@Column]` / `#This Row`
resolver. Formula-defined names containing explicit implicit intersection stay
unresolved at a caller because that caller position is part of the semantics.

For an OOXML array formula with a fixed legacy CSE output range, FormulaFence
keeps the range compact and links its anchor to any statically known formula
that reads a member of that range. For example, a formula stored at `Model!B1`
with fixed `B1:B3` output makes both `=Model!B2*10` and
`=SUM(Model!B2:B3)` downstream consumers of the anchor. It does not expand the
range into one graph node per result cell. For a dynamic array identified by
OOXML `XLDAPR`/`fDynamic` metadata, FormulaFence records the currently observed
output range and similarly links formulas that read its non-anchor members.
That is an observed graph relationship, not a fixed-size guarantee: Excel can
grow, shrink, or block the spill during recalculation. A new observed
output-member relationship emits `FF019`; enable
`no_new_dynamic_array_output_references` to make it `FFP019` in CI. An array
formula with unrecognized metadata remains an explicit coverage warning without
aliases. A change between ordinary, fixed CSE, and dynamic modes, or a fixed CSE
output-range change, emits `FF018`; adding or removing a legacy-CSE or
dynamic-array formula also emits `FF018`. Enable
`no_array_formula_semantics_changes` to make it `FFP018` in CI.

When a formula cannot be tokenized at all, FormulaFence records its location in
the profile and emits `FF016` for a new instance. Enable
`no_new_tokenization_failures` to turn that into `FFP016`.

A call to a workbook- or worksheet-local defined name is also resolved when
the complete definition is one statically resolvable `LAMBDA` expression. The
caller retains its explicit argument references and receives the static
dependencies of the function body, including through nested named-LAMBDA calls
and formula-defined names. FormulaFence recognizes normal formula text plus the
`_xlfn.LAMBDA`, `_xlpm.`, and `_xlop.` OOXML serialization used by
Excel-compatible writers. Dynamic, relative, cyclic, external, 3-D, or
tokenizer-unsupported definitions remain unresolved at their call site, so
`no_new_unresolved_references` can make newly introduced instances a hard CI
failure.

Static internal 3-D A1 references such as `Jan:Mar!B2:B10` are expanded across
every tab between their endpoints in workbook order. FormulaFence records the
cells that use them in a profile. If sheet insertion, removal, or movement
changes the resolved span while the formula text remains the same, it emits
`FF014`; `no_3d_reference_scope_changes` turns that condition into `FFP014`.
External 3-D forms remain external-link hazards; malformed, endpoint-missing,
and non-A1 3-D forms stay visible as unresolved coverage rather than being
assigned to a synthetic sheet.

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
