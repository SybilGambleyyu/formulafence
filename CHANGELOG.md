# Changelog

## 0.23.0 — 2026-07-24

- Inspect document-linked Office Web Add-in task-pane packages directly from
  their workbook declarations, through task-pane bindings and direct
  web-extension definitions. Private fingerprints retain task-pane
  configuration, add-in references, auto-show properties, bindings, snapshots,
  and relationship material without serializing add-in IDs, store references,
  property/binding values, XML, snapshot data, or relationship targets.
- Emit `FF028` for Office Web Add-in task-pane workbook bindings,
  configuration, definitions, or relationships and add the fail-closed
  `no_office_web_addin_changes` policy rule (`FFP028`). Writer-chosen
  relationship IDs and equivalent internal target spellings are normalized
  away, while malformed, oversized, unbound, or otherwise unrecognized parts
  remain visible coverage warnings. Task-pane and web-extension XML reads are
  bounded to 16 MiB per part, 32 MiB per workbook, and 64 parts. FormulaFence
  never installs, loads, executes, or fetches an add-in or manifest, follows an
  external target, or models worksheet-scoped Web Add-in markup outside this
  task-pane chain.

## 0.22.0 — 2026-07-24

- Inspect Office RibbonX custom-UI package parts directly from their root
  declarations, including the documented 2006 and Office 2010-era package
  forms. Private fingerprints retain complete custom-UI XML and direct
  relationship material without serializing control IDs, labels, callback
  names, image targets, or XML content.
- Emit `FF027` for RibbonX package, callback/control, or relationship changes
  and add the fail-closed `no_ribbon_customization_changes` policy rule
  (`FFP027`). Writer-chosen relationship IDs and equivalent internal target
  spellings are normalized away, while malformed, oversized, unbound,
  version-mismatched, or otherwise unrecognized parts remain visible coverage
  warnings. Reads are bounded to 16 MiB per part, 32 MiB per workbook, and
  eight parts. FormulaFence never invokes RibbonX callbacks, follows external
  targets, or parses image payloads.

## 0.21.0 — 2026-07-24

- Extend the XLM macro-sheet control boundary to direct internal related parts,
  including embedded OLE objects and packages. FormulaFence streams those raw
  bytes into private fingerprints without parsing or serializing payload
  contents; a payload-only change now remains a critical `FF026` finding.
- Bound that work to 32 MiB per related part, 64 MiB across a workbook, and
  256 parts. Missing, unreadable, oversized, or over-budget targets surface as
  safe inventory counts and parser-coverage warnings rather than silent gaps.

## 0.20.0 — 2026-07-24

- Inspect Excel 4.0 / XLM Macro Sheet and International Macro Sheet package
  parts directly, before the workbook reader can omit their executable cells.
  Private fingerprints retain complete macro XML, workbook bindings, and
  related package relationships without serializing commands, cell values,
  targets, identifiers, or embedded-object payloads.
- Emit `FF026` for XLM macro-sheet changes and add the fail-closed
  `no_xlm_macro_sheet_changes` policy rule (`FFP026`). Relationship-id-only
  rewrites are normalized away, while malformed, unbound, oversized, or
  unrecognized parts remain visible coverage warnings. FormulaFence does not
  execute or emulate XLM commands, resolve targets, or load embedded objects.

## 0.19.0 — 2026-07-24

- Inspect raw `xl/externalLinks/externalLink*.xml` parts for external-workbook,
  DDE, and OLE link definitions. Private fingerprints retain declaration-to-part
  bindings, endpoint relationships, definition material, cached values,
  item behavior, and unmodelled XML without serializing targets, names, source
  data, or extension payloads.
- Emit `FF025` for external-link package changes and add the fail-closed
  `no_external_link_package_changes` policy rule (`FFP025`). FormulaFence does
  not follow or execute external-workbook, DDE, or OLE links, establish source
  trust, or infer returned data.

## 0.18.0 — 2026-07-24

- Inspect Power Query Data Mashup custom XML parts directly: private fingerprints
  cover the embedded `Section1.m` formula document, logical package material,
  stable query metadata, and formula-firewall permissions without serializing
  M text, query/source names, metadata values, embedded content, telemetry IDs,
  or user-bound permission bindings.
- Ignore documented refresh-result metadata and `sqmid` telemetry noise while
  preserving high-signal query-definition and execution-control changes. Query
  tables linked through normal Excel table relationships are now inventoried in
  addition to directly worksheet-linked tables.
- Emit `FF024` for changed Power Query formulas or semantic controls and add
  the fail-closed `no_power_query_changes` policy rule (`FFP024`). FormulaFence
  does not execute M, refresh connections, assess sources, or inspect DDE/OLE
  links or full PivotTable layout semantics.

## 0.17.0 — 2026-07-24

- Inventory external-data refresh controls directly from OOXML: workbook-wide
  external-link/refresh flags, data connections, linked query tables, and
  pivot-cache sources and refresh behavior.
- Normalize schema defaults and preserve source paths, connection strings,
  query material, identifiers, names, descriptions, parameter values, cached
  records, and opaque extension XML as private comparison fingerprints rather
  than report content.
- Emit `FF023` for changed external-data connection or refresh controls and add
  the fail-closed `no_external_data_connection_changes` policy rule (`FFP023`).
  FormulaFence does not execute connections, refresh data, or parse Power Query
  M, DDE/OLE links, or full PivotTable layout semantics.

## 0.16.0 — 2026-07-24

- Inventory operational protection controls directly from OOXML: workbook
  structure/windows/revision locks; worksheet, dialog-sheet, and chart-sheet
  permissions; protected range targets; and compact direct cell/row/column
  locked/hidden assignments on active protected sheets.
- Normalize worksheet action defaults so omitted and explicit OOXML spellings
  compare equal. Preserve unmodelled protection metadata through private
  fingerprints, and compare legacy/modern verifier material, protected-range
  names, and security descriptors without serializing any of their values.
- Emit `FF022` for changed protection controls and add the fail-closed
  `no_protection_changes` policy rule (`FFP022`). Protection remains an
  operational review boundary, not file encryption, authentication, or a claim
  to reproduce Excel's complete style cascade.

## 0.15.0 — 2026-07-24

- Inventory worksheet conditional-formatting controls directly from OOXML:
  compact target ranges, global precedence, criteria, rule flags, differential
  styles, color scales, data bars, icon sets, and retained extension fragments.
- Resolve differential styles rather than comparing unstable `dxfId` values;
  normalize schema boolean defaults, leading `=` criteria, priority-number
  gaps, and extension GUID links to avoid writer-only control diffs. Profiles
  redact criteria, text rules, and raw style/extension XML.
- Emit `FF021` for changed conditional-formatting controls and add the
  fail-closed `no_conditional_formatting_changes` policy rule (`FFP021`). The
  tool records display-control semantics but does not calculate Excel's final
  conditional formatting result.

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
