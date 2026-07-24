# Changelog

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
