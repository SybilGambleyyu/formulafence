# Changelog

## 0.14.0 — 2026-07-24

- Inventory worksheet data-validation controls as compact target ranges and
  compare their effective type, operator, criteria, blank/dropdown behavior,
  prompts, error alert, IME mode, and worksheet-level prompt-disable setting.
  This does not expand full-column validations into cells.
- Normalize omitted OOXML defaults (`none`, `between`, `stop`, and `noControl`)
  and an optional leading `=` in criteria, so equivalent Excel-compatible
  writers do not create a control-only diff. Identical controls are also joined
  when a writer splits their target groups. Profiles redact criteria and
  prompt/error text; local reports retain full before/after evidence.
- Emit `FF020` for changed validation controls and add the fail-closed
  `no_data_validation_changes` policy rule (`FFP020`). Validation expressions
  remain inspectable controls, not calculations FormulaFence evaluates.

## 0.13.0 — 2026-07-24

- Trace formulas that read a non-anchor member of a dynamic array's current
  OOXML output range. The compact anchor-to-consumer edge makes an input of a
  dynamic array reach its current direct and range consumers without expanding
  the spill into virtual cells.
- Preserve the safety boundary: dynamic output ranges are profiled as observed,
  not fixed, because recalculation can grow, shrink, or block a spill. Profiles
  list the observed range and every linked output-member consumer.
- Emit `FF019` when a formula newly intersects an observed non-anchor dynamic
  output member, including when a changed observed extent reaches an unchanged
  formula. Add `no_new_dynamic_array_output_references` (`FFP019`) as a
  fail-closed policy control.

## 0.12.0 — 2026-07-24

- Trace fixed legacy CSE array output members without expanding their declared
  ranges: an input of an array anchor now reaches ordinary formulas that read
  non-anchor result cells, including cross-sheet and range consumers.
- Inspect raw OOXML dynamic-array metadata to keep the boundary safe. Dynamic
  anchors are inventoried but never receive aliases for a current spill extent;
  unrecognized array metadata becomes a visible coverage warning instead of a
  guessed fixed CSE graph.
- Compare array-formula execution mode independently of formula text and emit
  `FF018` when a legacy-CSE or dynamic formula is added, removed, or changes
  mode, or when a legacy CSE fixed output range changes. Add the fail-closed
  `no_array_formula_semantics_changes` policy rule (`FFP018`).

## 0.11.0 — 2026-07-24

- Trace the exact selected cell for direct static A1 implicit intersection,
  including literal `@A1:A3` and persisted OOXML `_xlfn.SINGLE(A1:A3)` forms.
  Other explicit intersection expressions retain conservative static input
  edges instead of being evaluated.
- Inventory explicit implicit intersection in profiles, emit `FF017` for new
  uses, and add the fail-closed `no_new_implicit_intersections` policy rule
  (`FFP017`). Formula-defined names containing this context-dependent behavior
  remain unresolved at call sites.
- Normalize direct display and OOXML spellings of `#`/`ANCHORARRAY` and
  `@`/`SINGLE` in formula fingerprints to avoid a serialization-only formula
  diff.

## 0.10.0 — 2026-07-24

- Trace the static anchor behind direct internal `A1#` spilled-array references
  and OOXML-style `ANCHORARRAY(A1)` calls without evaluating Excel.
- Inventory spill-reference consumers in profiles and report new instances as
  `FF015`; add `no_new_spill_references` (`FFP015`) for a fail-closed CI
  boundary. Dynamic spill extent and blocking cells remain explicit limits.
- Surface formula-tokenization failures at workbook level instead of silently
  omitting their graph. New failures emit `FF016` and can be blocked with
  `no_new_tokenization_failures` (`FFP016`).
- Keep formula-defined names containing a spill reference unexpanded, so a
  named formula cannot hide the dynamic boundary behind inferred dependencies.

## 0.9.0 — 2026-07-24

- Expand calls to workbook- and worksheet-local defined names whose complete
  definitions are statically resolvable `LAMBDA` expressions, including nested
  named-LAMBDA calls and formula-defined names that call them.
- Preserve Excel name scope and local-name precedence for callable definitions.
  Recognize OOXML formula definitions without a leading `=` and serialized
  `_xlfn.LAMBDA`, `_xlpm.`, and `_xlop.` local-name forms.
- Leave dynamic, relative, cyclic, external, 3-D, tokenizer-unsupported, and
  otherwise non-static named LAMBDAs visible as unresolved references at their
  call sites rather than creating guessed graph edges.

## 0.4.0 — 2026-07-24

- Resolve the conservative, fully qualified Excel-table subset: table names,
  single or contiguous column ranges, and `#All`/`#Data`/`#Headers`/`#Totals`
  regions.
- Inventory table metadata in profiles and flag table additions, removals, and
  definition changes as `FF013`; add the `no_table_definition_changes` policy
  control (`FFP013`).
- Keep this-row (`@`) and complex table syntax as explicit unresolved coverage
  rather than inferring a dependency that static inspection cannot justify.

## 0.3.0 — 2026-07-24

- Resolve ordinary workbook and sheet-local defined names with static A1
  destinations into the dependency graph, including explicit references to
  sheet-local names and names that resolve to external workbooks.
- Make static-analysis coverage visible: profiles identify unresolved range
  tokens and dynamic `INDIRECT`/`OFFSET` formulas; diffs report new instances
  as `FF011` and `FF012`.
- Add opt-in `no_new_unresolved_references` and `no_new_dynamic_references`
  policy controls (`FFP011`, `FFP012`) for teams that need to fail closed on
  new dependency-coverage gaps.

## 0.2.0 — 2026-07-24

- Add deterministic shortest dependency-path samples to every changed cell's
  impact record, FormulaFence hazard finding metadata, Markdown reports, and
  SARIF properties.
- Include the same path evidence when an impact-limit policy fails.

## 0.1.1 — 2026-07-24

- Capture workbook-parser warnings as structured profile coverage notes instead
  of writing raw dependency warnings to the console.
- Flag newly introduced parser coverage warnings in diffs (`FF010`) and support
  the `no_new_parser_warnings` policy control.
- Validate the profile path against a public 18-sheet financial cap-table model;
  see [validation notes](docs/validation.md).

## 0.1.0 — 2026-07-24

- Initial public release: formula-aware semantic diffing, explicit dependency
  impact, workbook-control checks, policy-as-code, and Markdown/JSON/SARIF output.
