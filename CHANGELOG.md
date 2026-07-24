# Changelog

## 0.34.0 — 2026-07-25

- Inspect modern Excel Named Sheet Views from the documented relationship-backed
  worksheet parts, retaining view names, IDs, alternate filter criteria, target
  ranges, table bindings, table-column IDs, and sort keys only in private
  signatures—not profiles, `FF038`, or SARIF.
- Reconcile each stored filter to its base AutoFilter using Excel's documented
  UID, table-ID, then worksheet-owned fallback sequence. Normalize equivalent
  GUID, local A1 case/absolute-reference, Boolean/default, and unsigned-integer
  spellings while making a resolved target rebinding material to the diff.
- Emit `FF038` for a Named Sheet View definition, alternate filter/sort rule, or
  binding change; add the fail-closed `no_named_sheet_view_changes` policy rule
  (`FFP038`).
- Make missing, ambiguous, mismatched, malformed, unsupported, oversized, and
  unsafe relationship parts or filter bindings visible parser-coverage warnings
  rather than silently dropping them. FormulaFence does not activate/render a
  saved view, calculate a filtered result, infer formula visibility sensitivity,
  repair metadata, or interpret full differential-format, future extension, or
  rich-sort semantics.

## 0.33.0 — 2026-07-25

- Inspect standard worksheet `ignoredErrors` declarations and Office 2010
  `x14:ignoredErrors` extension declarations directly from raw OOXML. Private
  signatures retain target ranges and enabled warning types without serializing
  them into profiles, `FF037`, or SARIF.
- Emit `FF037` when suppressed Excel evaluation, inconsistent-formula,
  omitted-range, unlocked-formula, empty-reference, list-validation,
  calculated-column, text-number, or two-digit-year warning controls change;
  add the fail-closed `no_ignored_error_changes` policy rule (`FFP037`).
  Equivalent local A1 case/absolute-reference, Boolean, and target-order
  spellings are normalized.
- Make malformed or unsupported containers, extension material, attributes,
  flags, targets, and child markup visible parser-coverage warnings rather than
  silently dropping them. FormulaFence does not determine whether Excel would
  show a warning, calculate a formula, repair an error, or change application-
  level error-checking options.

## 0.32.0 — 2026-07-25

- Inspect worksheet and Table Definition-part AutoFilters directly from raw
  OOXML, including private filter criteria, selected values, filter-button
  state, AutoFilter sort state, and sort conditions. Also inspect explicit row
  `hidden` / outline state and the `sheetFormatPr@zeroHeight` hidden-by-default
  optimization without serializing criteria, sort keys/lists, table names, or
  row/range references into profiles, `FF036`, or SARIF.
- Emit `FF036` for a material filter, sort, or row-visibility control change,
  and add the fail-closed `no_filter_visibility_changes` policy rule (`FFP036`).
  Equivalent local A1 case/absolute-reference, Boolean/default, and unsigned
  integer spellings are normalized; filter-member ordering is canonicalized.
- Make malformed or unsupported declarations, extensions, and unsafe/missing
  table relationships visible parser-coverage warnings rather than silently
  dropping them. FormulaFence does not apply filters, calculate results,
  determine formula visibility sensitivity, render a report, or track hidden
  columns.

## 0.31.0 — 2026-07-25

- Inspect Excel Scenario Manager declarations directly from worksheet OOXML
  (`scenarios` / `scenario` / `inputCells`). Private signatures retain selected
  and shown state, summary references, names, protection flags, comments/users,
  changing-cell references, stored input values, deleted/undone state, and
  display number formats without serializing that material into profiles,
  `FF035`, or SARIF.
- Emit `FF035` for a material Scenario Manager definition or stored-input
  change, and add the fail-closed `no_scenario_manager_changes` policy rule
  (`FFP035`). Equivalent local A1 case/absolute-reference, Boolean, and
  unsigned-integer spellings plus schema-default false flags are normalized.
  Missing, malformed, duplicate-within-worksheet, or unsupported declarations
  are visible parser-coverage warnings rather than silently ignored.
- Treat Scenario Manager as worksheet-scoped: duplicate scenario names on
  different worksheets remain valid. Do not show/apply scenarios, calculate
  results, infer scenario-to-formula dependencies, or expose scenario names,
  comments, users, stored values, references, or raw XML.

## 0.30.0 — 2026-07-25

- Inspect Excel What-If Data Table masters directly from worksheet OOXML
  (`f t="dataTable"`). Private signatures retain the declared output range,
  one-/two-variable mode, orientation, input references, deleted-input flags,
  recalculation request, and supported generic formula metadata without
  serializing those controls into profiles, `FF034`, or SARIF.
- Stabilize the workbook reader's `DataTableFormula` representation, eliminating
  false self-diffs caused by process-local object addresses while preserving
  ordinary formula add/remove guards with a safe `=TABLE()` placeholder.
- Emit `FF034` for a material What-If Data Table definition or control change,
  and add the fail-closed `no_what_if_data_table_changes` policy rule
  (`FFP034`). Equivalent A1 case/absolute-reference and Boolean spellings are
  normalized. Missing, malformed, overlapping, or unsupported declarations are
  visible parser-coverage warnings rather than silently ignored.
- Do not calculate scenarios, infer a Data Table's output formula, predict
  recalculation results, or add Data Table inputs to the ordinary dependency
  graph. Cached scenario-output cells remain under the normal cell-diff
  boundary.

## 0.29.0 — 2026-07-24

- Inspect embedded Power Pivot/Data Model packages from the workbook's explicit
  `powerPivotData` relationship and `x15:dataModel` declaration. Private
  fingerprints retain complete declaration material, normalized relationship
  semantics, and bounded raw model payload hashes without serializing table,
  column, relationship, connection, DAX, stored-value, target, XML, or payload
  content.
- Emit `FF033` for a Data Model binding, declaration, direct model-part
  relationship, or bounded raw payload change, and add the fail-closed
  `no_power_pivot_data_model_changes` policy rule (`FFP033`). Relationship IDs,
  equivalent internal target spellings, and writer-generated Data Model GUIDs
  are normalized away. Missing, malformed, orphaned, unbound, externally
  targeted, unexpected directly related, oversized, over-budget, or
  unrecognized material remains a visible coverage warning. Raw model payload
  reads are bounded to 512 MiB per part, 512 MiB per workbook, and 16 parts.
- Never deserialize the embedded Analysis Services payload, evaluate DAX,
  refresh a model, calculate/render a report, infer model-to-cell impact, or
  fetch an external target.

## 0.28.0 — 2026-07-24

- Inspect Slicer and Timeline cache definitions directly from documented
  workbook extension declarations and explicit workbook relationships. Private
  fingerprints retain Slicer item selections, Timeline state/filter material,
  PivotTable/table source bindings, filtered-PivotTable bindings, normalized
  relationships, and complete cache definitions without serializing cache
  names, source fields, selected values, date ranges, PivotTable names,
  relationship targets, or XML.
- Emit `FF032` for a Slicer/Timeline workbook binding, filter state, cache
  definition, source binding, filtered-PivotTable binding, or unexpected direct
  cache-part relationship change, and add the fail-closed
  `no_slicer_timeline_cache_changes` policy rule (`FFP032`). Relationship IDs,
  equivalent internal target spellings, coordinated Slicer/Timeline PivotCache
  extension-ID renumbering,
  known optional Slicer defaults, Boolean spellings, and Timeline GUIDs are
  normalized away. Malformed, orphaned, unbound, externally targeted,
  oversized, over-budget, or unrecognized material remains a visible coverage
  warning. Cache XML reads are bounded to 16 MiB per part, 64 MiB per workbook,
  and 512 parts.
- Treat the documented 2010 Timeline-cache relationship and the widely emitted
  2011 compatibility relationship as one equivalent workbook binding.
- Never apply a Slicer or Timeline filter, calculate/render a PivotTable or
  table, infer downstream cell impact, fetch an external target, or model
  worksheet/drawing view geometry and styles.

## 0.27.0 — 2026-07-24

- Inspect PivotTable view definitions, cache schemas, shared cache items, and
  bounded raw cache-record payloads directly from the documented workbook-cache
  and worksheet-PivotTable OOXML relationships. Private fingerprints retain
  layouts, cache material, normalized relationships, and record hashes without
  serializing names, source ranges, field/item values, formulas, cache records,
  relationship targets, XML, or payload bytes.
- Emit `FF031` for PivotTable bindings/layouts, cache definitions, shared items,
  cache-record relationships, or cache-record payload changes, and add the
  fail-closed `no_pivot_table_definition_changes` policy rule (`FFP031`). Cache
  source and refresh settings deliberately remain with `FF023`. Relationship
  IDs, equivalent internal target spellings, and cache-ID renumbering are
  normalized away; malformed, orphaned, unbound, oversized, over-budget, or
  unrecognized material remains a visible coverage warning. PivotTable and
  cache-definition XML reads are bounded to 16 MiB per part, 64 MiB per
  workbook, and 512 parts; raw cache-record hashes are bounded to 32 MiB per
  part, 64 MiB per workbook, and 512 parts.
- Detach cache-record relationships only in a temporary reader copy before the
  underlying workbook library loads cells, so it does not eagerly materialize
  unbounded record streams. The original workbook is never modified.
- Never refresh or calculate a PivotTable, render a report, infer
  PivotTable-to-cell impact, fetch an external target, parse cached record
  values, or interpret OLAP, extension-list, or slicer semantics.

## 0.26.0 — 2026-07-24

- Inspect DrawingML chart definitions, cached series presentation data, and
  chart `userShapes` overlays directly through standard worksheet/chartsheet
  drawing relationships. Private fingerprints retain chart definition and cache
  material separately, normalized relationships, overlays, and bounded direct
  related-part payload hashes without serializing formulas, cached values,
  titles, shape text, relationship targets, XML, or payload bytes.
- Emit `FF030` for chart bindings, definitions, cached data, overlays,
  relationships, or direct related-payload changes, and add the fail-closed
  `no_chart_definition_changes` policy rule (`FFP030`). Writer-chosen
  relationship IDs and equivalent internal target spellings are normalized
  away; malformed, orphaned, unbound, oversized, over-budget, or unrecognized
  material remains a visible coverage warning. Chart and overlay XML reads are
  bounded to 16 MiB per part, 64 MiB per workbook, and 512 parts; direct related
  payload hashes are bounded to 32 MiB per part, 64 MiB per workbook, and 512
  parts.
- Never calculate a series formula, render a chart, infer chart-to-cell impact,
  follow an external target, parse media or embedded-package formats, or
  interpret modern `chartEx`/nested-chart semantics.

## 0.25.0 — 2026-07-24

- Extend the existing worksheet-control guardrail to legacy VML form controls.
  FormulaFence now follows standard worksheet `vmlDrawing` relationships and
  privately fingerprints non-`Note` VML `ClientData` control material, including
  macro assignments, cell/range bindings, camera source ranges, and directly
  referenced VML-part relationship semantics. Ordinary VML comment notes are
  deliberately excluded from the control inventory.
- Keep the `FF029` / `FFP029` contract while making it cover modern worksheet
  controls, legacy VML controls, and OLE objects together. Relationship IDs and
  equivalent internal target spellings remain normalized; malformed, orphaned,
  missing, oversized, and over-budget VML material remains a visible coverage
  warning. VML XML shares the existing 16 MiB-per-part, 64 MiB-per-workbook,
  512-part control XML budget.
- Never render a VML drawing, read comment text into the control profile, execute
  a macro, evaluate a binding, or open a relationship target. Macro names,
  formulas/ranges, captions, relationship targets, and VML XML remain private.

## 0.24.0 — 2026-07-24

- Inspect relationship-backed worksheet ActiveX, form-control, and OLE-object
  chains directly from raw OOXML before the workbook reader can omit them.
  Private fingerprints retain worksheet declarations, control configuration,
  ActiveX persistence XML, form-control properties, relationships, and bounded
  direct ActiveX/OLE/package payload hashes without serializing control names,
  class IDs, licenses, macro assignments, formulas/ranges, OLE identities,
  relationship targets, XML, or payload bytes.
- Emit `FF029` for worksheet-control bindings, definitions, ActiveX/form-control
  material, OLE configuration, related relationships, or direct payload changes,
  and add the fail-closed `no_worksheet_embedded_control_changes` policy rule
  (`FFP029`). Writer-chosen relationship IDs and equivalent internal target
  spellings are normalized away. Malformed, orphaned, unbound, oversized, or
  over-budget material remains a visible coverage warning. XML reads are bounded
  to 16 MiB per part, 64 MiB per workbook, and 512 parts; raw direct payload
  hashes are bounded to 32 MiB per part, 64 MiB per workbook, and 512 parts.
  FormulaFence never initializes an ActiveX control, deserializes or opens an OLE
  object/package, renders its drawing surface, follows an external target, or
  infers event dispatch.

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
