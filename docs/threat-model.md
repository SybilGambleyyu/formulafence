# Scope and threat model

FormulaFence is a static change-assurance layer. It answers whether an Excel
workbook's inspectable structure changed in a risky way; it does not certify
financial correctness or replace model review.

## Safety properties

- Workbook content stays on the machine running FormulaFence. The CLI makes no
  network requests.
- It loads formulas as text with `data_only=False`; it does not calculate them.
- It never executes VBA, XLM macro sheets, RibbonX callbacks, DDE, external
  links, Power Query, Power Pivot/DAX, Office Web Add-in code, or worksheet
  ActiveX/OLE code.
- VBA payloads, XLM macro-sheet source material, RibbonX control/callback
  material, Office Web Add-in task-pane material, and worksheet control/OLE
  material are compared through private fingerprints only.
- Embedded Power Pivot/Data Model declarations and bounded raw model payloads
  are compared through private fingerprints only; table names, relationships,
  DAX, stored values, and connection details are never emitted.
- What-If Data Table output ranges, input-cell references, and raw formula
  metadata are compared through a private signature only. Cached scenario-output
  cells remain under the normal cell-diff boundary.
- Scenario Manager names, comments, user metadata, stored input values,
  input/result references, and raw declarations are compared through a private
  signature only. Cached worksheet cells remain under the normal cell-diff
  boundary.
- Worksheet DrawingML regular-shape text, presentation, geometry, anchors,
  macro assignments, text links, descriptions, relationship identifiers, and
  targets are compared through private fingerprints only. Profiles and reports
  retain structural counts, never the underlying shape content.
- Filter criteria, selected values, custom sort lists, table names, sort keys,
  and row/range references are compared through a private signature only.
- Ignored-error target ranges and exact per-range warning suppressions are
  compared through a private signature only.
- Protection credential material is never emitted: legacy verifiers, modern
  hashes/salts, protected-range names, and security descriptors are compared
  through private fingerprints and reported only as safe presence/change metadata.
- External-data source material is never emitted: connection names/descriptions,
  paths, URLs, connection strings, commands, parameter values, SSO identifiers,
  cached records, and opaque extension XML remain private comparison evidence.
- It uses sparse cell storage rather than walking every coordinate in a workbook's
  declared used rectangle.
- Parser warnings from unsupported OOXML extensions are captured in the profile
  as coverage notes; FormulaFence does not silently discard them from its report.

## What a finding means

An impact count traces explicit A1-style cell dependencies available in the
baseline and candidate. It is an aid to review, not a claim that the cells will
recalculate correctly in Excel. FormulaFence also emits deterministic shortest
path samples from the changed cell to sampled downstream formulas. These paths
are explicit static dependencies, not proof of runtime evaluation. A
formula-pattern finding means both immediate peers have the same relative
formula fingerprint while the changed middle cell does not; it is a focused
review prompt, not proof of an error.

## Deliberate limits

- Supported files are `.xlsx` and `.xlsm`; legacy `.xls` and file-encrypted or
  password-to-open workbooks are outside scope. Workbook and worksheet
  protection flags inside an otherwise readable OOXML workbook are inspected as
  operational controls, not treated as encryption.
- Ordinary workbook and sheet-local names with static A1 destinations are
  resolved into the dependency graph. It also expands formula-defined names
  whose whole definition is statically visible and internal, including nested
  workbook and sheet-local names and constants. FormulaFence also resolves
  table names, static columns/contiguous column ranges, and
  `#All`/`#Data`/`#Headers`/`#Totals` regions; it inventories and diffs the
  table definitions that give those references meaning. It resolves `@` and
  `#This Row` only when the formula location statically identifies a named
  table's data row; unqualified current-row forms additionally require the
  formula cell itself to be in that table. Header/total-row, cross-sheet,
  ambiguous, and complex bracket-escape table syntax, `INDIRECT`, `OFFSET`,
  relative/cyclic/external/3-D/tokenizer-unsupported formula-defined names,
  cube functions, add-ins, and custom functions cannot always be statically
  resolved. FormulaFence flags newly introduced unresolved tokens and
  `INDIRECT`/`OFFSET` use, but does not fabricate dependencies for them.
- A direct internal A1 spilled-array anchor such as `A1#`, or its OOXML
  `ANCHORARRAY(A1)` representation, adds a dependency edge from the anchor
  cell to its consumer and is inventoried in the profile. FormulaFence cannot
  safely enumerate the dynamic spill extent or every possible blocking cell;
  it emits `FF015` for newly added spill references. External, 3-D, range,
  named, implicit-intersection, and malformed spill forms stay outside this
  subset. A formula-defined name containing a spill reference is not expanded,
  so callers retain a visible coverage gap.
- A multi-cell legacy CSE array formula has a fixed OOXML output range. When
  FormulaFence can verify that the array anchor has no dynamic-array cell
  metadata, it links the anchor to statically known formulas that read any
  result member of that range. The range remains compact rather than becoming
  one graph node per output cell. Dynamic-array anchors identified by OOXML
  `XLDAPR`/`fDynamic` metadata expose a current serialized output range.
  FormulaFence links the anchor to static formulas that currently read a
  non-anchor member of that observed range, and records those relationships in
  the profile. This is not a fixed-output assertion: the spill can resize or be
  blocked at recalculation time, and FormulaFence does not predict future spill
  dimensions or blockers. A newly observed output-member relationship emits
  `FF019` and can be blocked by `no_new_dynamic_array_output_references`.
  Array formulas with absent, malformed, or unrecognized metadata mappings are
  reported as coverage notes and receive no aliases. FormulaFence reports
  adding, removing, or changing mode, plus a fixed CSE output-range change, as
  `FF018`; it does not calculate either array form.
- Excel What-If Data Tables are distinct from Excel tables. FormulaFence reads
  each worksheet `f t="dataTable"` master directly from OOXML and privately
  compares its declared output range, one-/two-variable form, one-variable
  orientation, input references, deleted-input flags, recalculation request,
  and supported generic formula metadata. A material change emits `FF034` and
  can be blocked with `no_what_if_data_table_changes`. Profiles and `FF034`
  details expose only structural counts, never those references or ranges.
  Equivalent A1 case/absolute-reference and Boolean spellings are normalized.
  Missing, malformed, overlapping, or unsupported declarations remain visible
  coverage warnings. FormulaFence does not calculate scenarios, infer their
  output formula, predict recalculation results, or add Data Table inputs to
  the ordinary dependency graph; cached scenario-output cells remain ordinary
  cell values under the normal diff boundary.
- Excel Scenario Manager controls are distinct from Data Tables. FormulaFence
  reads each worksheet's raw `<scenarios>` declaration and privately compares
  current/shown selection state, summary references, scenario names,
  locked/hidden flags, declared input counts, comments/users, input references,
  stored values, deleted/undone flags, and display number formats. A material
  change emits `FF035` and can be blocked with
  `no_scenario_manager_changes`. Profiles and `FF035` details expose only
  structural counts, never names, comments, users, values, or references.
  Equivalent local A1 case/absolute-reference, Boolean, and unsigned-integer
  spellings plus schema-default false flags are normalized. Missing, malformed,
  duplicate-within-worksheet, or unsupported declarations remain visible
  coverage warnings. FormulaFence does not show/apply a scenario, calculate its
  result, infer a scenario-to-formula dependency, or fetch an external target;
  cached worksheet cells remain ordinary cell values under the normal diff
  boundary.
- Excel AutoFilters and row/column visibility can change which records or
  fields are shown, and which values vertical `SUBTOTAL` formulas include,
  without changing a formula or ordinary cell value. FormulaFence reads
  worksheet `<autoFilter>` and `<sortState>` elements, the same controls in
  relationship-backed Table Definition parts, explicit row `hidden`,
  `outlineLevel`, `collapsed`, and zero `ht` attributes,
  `sheetFormatPr@zeroHeight`, zero worksheet-default row/column dimensions, and
  raw `<cols>/<col>` `hidden`, `outlineLevel`, `collapsed`, and zero `width`
  declarations. Column declarations are applied in file order and only
  attributes present in a later declaration override earlier effective state,
  including a positive-width override of an inherited zero width. Criteria,
  selected values, custom sort lists, sort keys, raw dimension values, and row/column
  ranges remain private; profiles and `FF036` expose structural counts only.
  Local A1 case/absolute-reference, Boolean/default, unsigned-integer,
  equivalent zero-dimension, and equivalent column-range spellings are
  normalized.
  Unsupported extensions, malformed controls, exhausted column-update limits,
  and unsafe/missing table relationships remain visible coverage warnings, and
  `no_filter_visibility_changes` can block the change as `FFP036`.
  FormulaFence does not apply filters, calculate `SUBTOTAL`/`AGGREGATE`, infer
  formula sensitivity, render views, track ordinary positive widths/heights or
  styles, or model outline-display settings.
- Excel ignored-error declarations can suppress evaluation, inconsistent-formula,
  omitted-range, unlocked-formula, empty-reference, list-validation,
  calculated-column, text-number, and two-digit-year warnings without changing
  a cell or formula. FormulaFence reads standard `<ignoredErrors>` and Office
  2010 `x14:ignoredErrors` declarations, privately compares local target ranges
  and enabled warning flags, and emits `FF037`; `no_ignored_error_changes` can
  block the change as `FFP037`. Profiles and `FF037` details expose only
  structural counts, never target ranges or individual suppressions. Equivalent
  local A1 case/absolute-reference, Boolean, and target-order spellings are
  normalized. Malformed or unsupported containers, extension material,
  attributes, flags, targets, and child markup remain visible coverage warnings.
  FormulaFence does not decide whether Excel would display a warning, calculate
  a formula, repair an error, alter application-level error checking, or infer
  a suppressed warning's downstream impact.
- Modern Excel Named Sheet Views retain alternate filter and sort settings in
  relationship-backed worksheet parts, potentially changing a saved report view
  without changing ordinary cells or the active AutoFilter. FormulaFence follows
  those parts, privately compares view definitions, and resolves each filter by
  AutoFilter UID, table ID, then worksheet-owned AutoFilter. It emits `FF038`;
  `no_named_sheet_view_changes` can block the change as `FFP038`. Profiles and
  `FF038` details expose only counts for parts, views, alternate filters,
  columns, criterion groups, sort rules/conditions, and unrecognized controls;
  names, IDs, criteria, ranges, table bindings, table-column IDs, and sort keys
  remain private. Equivalent GUID, local A1 case/absolute-reference,
  Boolean/default, and unsigned-integer spellings are normalized. Missing,
  ambiguous, mismatched, malformed, unsupported, oversized, or unsafe
  parts/bindings remain visible coverage warnings. FormulaFence does not
  activate/render a saved view, calculate a filtered result, infer formula
  visibility sensitivity, repair metadata, or interpret future extension/rich-
  sort or full differential-format semantics.
- Excel number formats can hide or materially reinterpret an unchanged stored
  value: `;;;` can display it as blank, while custom sections, scaling commas,
  dates, percentages, literals, and text placeholders can change the review
  surface. FormulaFence privately resolves custom `<numFmt>` codes, base
  `<cellStyleXfs>`, effective `<cellXfs>` with `xfId` and
  `applyNumberFormat`, direct cell `s`, `customFormat=1` row `s`, and raw
  `<cols>/<col style>` assignments. It emits `FF039`; the
  `no_number_format_changes` policy rule can block it as `FFP039`. Profiles and
  `FF039` details expose only counts for default/direct/row/effective-column
  assignments, built-in/custom classes, and malformed controls—never codes,
  style IDs, or targets. Equivalent custom-ID remapping, Boolean spelling,
  base-XF inheritance, and effective column-range splitting are normalized.
  Missing custom definitions, invalid IDs/indexes/targets, conflicting
  definitions, and bounded parser failures remain coverage warnings. FormulaFence
  does not render locale-specific output, validate format syntax, calculate
  values, model width/overflow, or track fonts/fills/borders/alignment, quote
  prefixes, table styles, or arbitrary visual formatting. Column styles
  are compared only as OOXML defaults for unallocated/new cells, not as a claim
  to restyle allocated cells.
- Excel cell fonts can make an unchanged value or warning less visible, such as
  a white font against a matching background, without changing a formula or
  value. FormulaFence privately resolves raw `<fonts>` records, base
  `<cellStyleXfs>`, effective `<cellXfs>` with `xfId` and `applyFont`, direct
  cell `s`, `customFormat=1` row `s`, and raw `<cols>/<col style>` assignments.
  It emits `FF040`; `no_cell_font_changes` can block it as `FFP040`. Profiles
  and `FF040` details expose only default-definition, direct/row/effective-column,
  and malformed-control counts—never font names, colour values, effects, style
  IDs, or targets. Equivalent font-ID remapping, common font-child ordering,
  Boolean spelling, base-XF inheritance, and effective column-range splitting
  are normalized. Missing or malformed definitions, invalid IDs/indexes/targets,
  and bounded parser failures remain visible coverage warnings. FormulaFence
  does not render or resolve theme colours, decide whether a font is visible
  against a fill, calculate text/background contrast or values, track
  borders/alignment, rich-text run rendering, table styles, width/overflow, or arbitrary
  visual formatting. Column
  styles are compared only as OOXML defaults for unallocated/new cells, not as a
  claim to restyle allocated cells.
- Excel cell fills can make unchanged text, warnings, or input/output cues less
  visible without changing a formula or value. FormulaFence privately resolves
  raw `<fills>` definitions, including patterned and gradient fills, base
  `<cellStyleXfs>`, effective `<cellXfs>` with `xfId` and `applyFill`, direct
  cell `s`, `customFormat=1` row `s`, and raw `<cols>/<col style>` assignments.
  It emits `FF041`; `no_cell_fill_changes` can block it as `FFP041`. Profiles
  and `FF041` details expose only default-definition, direct/row/effective-column,
  and malformed-control counts—never fill colours, pattern types, gradient
  geometry/stops, style IDs, or targets. Equivalent fill-ID remapping, valid
  pattern-colour child ordering, Boolean spelling, base-XF inheritance,
  semantically inert no-fill/solid-background declarations, and effective
  column-range splitting are normalized. Missing or malformed definitions,
  invalid IDs/indexes/targets, and bounded parser failures remain visible
  coverage warnings. FormulaFence does not resolve theme colours, render fills,
  calculate text/background contrast, evaluate conditional-format differential
  styles, apply table styles, or claim arbitrary visual-style coverage. Column
  styles are compared only as OOXML defaults for unallocated/new cells, not as a
  claim to restyle allocated cells.
- SpreadsheetML can retain a formula's last calculated result beside its
  formula text in the same `<c>` cell. That lets a workbook save a different
  displayed result without changing the ordinary formula text. FormulaFence
  reads raw `<f>` and `<v>` elements together and privately fingerprints
  numeric, string, Boolean, and error result material. It emits `FF042` only
  when a result cache changes without a changed formula at that cell and without
  an ordinary changed cell reaching it through the static dependency graph;
  `no_formula_cached_result_changes` can block it as `FFP042`.
  Profiles and reports expose only formula/cached/missing/type/malformed counts,
  never result values, error text, digests, or locations. Equivalent finite
  numeric and Boolean spellings are normalized; blank results remain visible as
  missing caches. Unsupported or malformed metadata becomes an explicit coverage
  warning. FormulaFence does not calculate or validate results, determine
  whether they are stale or tampered, or model volatile, dynamic, external, or
  calculation-engine dependencies. A legitimate recalculation without a
  statically visible input edit can therefore still require review.
- SpreadsheetML shared strings and inline strings can split one displayed cell
  value into character-level `<r>` runs. Their `<rPr>` formatting can hide or
  alter the emphasis of a phrase while the normal cell reader still returns the
  same concatenated text. FormulaFence follows referenced shared-string items
  and direct inline strings, privately compares run-property sequences, styled
  character boundaries, and phonetic runs/properties, and emits `FF043`.
  `no_rich_text_run_changes` blocks it as `FFP043`. Profiles and report
  details expose only aggregate shared-item/cell/run, inline-cell/run, phonetic,
  and malformed-control counts; text, font/colour material, shared-string
  indexes, and locations remain private. Equivalent property ordering, colour
  case, and explicit false Boolean properties are normalized. A normal text
  edit inside an unchanged run-property sequence remains a normal cell diff,
  while a moved styled boundary with unchanged text is guarded. Malformed,
  unsupported, or unreadable metadata becomes a coverage warning. FormulaFence
  does not render a cell, resolve theme colours, calculate contrast, determine
  whether text is visible, preserve rich text, or guarantee cross-version Excel
  rendering equivalence. This boundary follows Microsoft's
  [shared-string-table guidance](https://learn.microsoft.com/en-us/office/open-xml/spreadsheet/working-with-the-shared-string-table),
  the Open XML `r` [rich-text-run definition](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.run?view=openxml-3.0.1),
  and `is` [inline-string definition](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.inlinestring?view=openxml-3.0.1).
- Non-chart Worksheet DrawingML regular shapes (`xdr:sp`) and nested group
  shapes (`xdr:grpSp`) are followed from standard worksheet `drawing`
  relationships before the workbook reader can discard their text-box
  declarations. FormulaFence privately fingerprints supported anchor/layout,
  shape/group XML, macro assignments, text links, click/hover relationship
  semantics, and visible text/presentation declarations. Profiles expose only
  safe worksheet/drawing/anchor, shape/text/group, text paragraph/run,
  macro/text-link/hyperlink, relationship, and malformed-control counts. A
  material change emits `FF044` and can be blocked with
  `no_worksheet_drawing_shape_changes`. Non-visual shape IDs, relationship-ID
  rewrites, and colour-case spelling are normalized. FormulaFence does **not**
  render DrawingML, resolve themes or contrast, calculate text links, execute
  macro assignments, retrieve external targets, parse or hash media, or cover
  pictures, connectors, graphic frames, SmartArt, or other non-`xdr:sp`
  objects. Missing, malformed, unsafe, oversized, over-budget, or unsupported
  regular-shape metadata becomes a visible parser-coverage warning. XML reads
  are bounded to 16 MiB per part, 64 MiB per workbook, and 512 parts. This
  scope follows the Open XML [`xdr:sp` Shape definition](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.spreadsheet.shape?view=openxml-3.0.1).
- Worksheet data-validation controls are inventoried and diffed as compact
  ranges rather than by expanding their target cells. FormulaFence compares the
  validation type, operator, two criteria expressions, blank/dropdown behavior,
  input prompts, error alerts, IME mode, and worksheet-level `disablePrompts`.
  It normalizes schema defaults (`none`, `between`, `stop`, and `noControl`) and
  an optional leading `=` in a criterion, and writer grouping of identical
  targets to avoid writer-only noise. Profiles omit criteria and prompt/error
  text, while local diff evidence retains them.
  A change emits `FF020` and can be blocked with
  `no_data_validation_changes`. FormulaFence does not evaluate a validation
  formula, infer list contents, or predict whether Excel will accept an entry.
- Worksheet conditional-formatting controls are read directly from OOXML so
  library support gaps cannot silently erase them before comparison. FormulaFence
  inventories compact target ranges, globally ordered priority, rule criteria
  and flags, differential styles, color scales, data bars, icon sets, and both
  rule- and worksheet-level extension fragments. It resolves a `dxfId` to its
  actual differential style and normalizes schema defaults, leading `=` formula
  spelling, priority-number gaps, and extension GUID links. An extension that
  cannot be interpreted remains opaque but is retained as full local diff
  evidence; profiles redact criteria, text rules, and raw XML. Any change emits
  `FF021` and can be blocked with `no_conditional_formatting_changes`.
  FormulaFence does not calculate a condition, expand relative criteria across
  every target, reconcile it with manual formatting, or predict the rendered
  result when overlapping rules conflict.
- Operational protection controls are read directly from OOXML: workbook
  structure/windows/revision locks; worksheet and dialog-sheet action locks;
  chart-sheet content/object locks; protected ranges; and direct cell, row, and
  column `locked`/`hidden` assignments on active protected sheets. It normalizes
  sheet-action defaults and keeps target spans compact. Raw credential and
  identity material is never serialized; private fingerprints expose a material
  change without exposing its values. A change emits `FF022` and can be blocked
  with `no_protection_changes`. This does **not** establish confidentiality,
  authentication, authorization, file encryption, rights-management behavior,
  or whether Excel's full style cascade makes an individual cell editable.
- External-data refresh controls are read directly from OOXML: workbook-wide
  external-link and refresh-on-open flags; connection refresh schedules,
  background/cache/credential controls, source-kind metadata, connection-file
  behavior, and parameter-triggered refreshes; linked query-table refresh and
  growth behavior; and pivot-cache source/refresh settings. Omitted schema
  defaults are normalized. Names, paths, URLs, connection strings, commands,
  parameter values, SSO IDs, cached records, and opaque extension XML remain
  private fingerprints; a material change emits `FF023` and can be blocked with
  `no_external_data_connection_changes`. FormulaFence does **not** connect,
  refresh data, establish source trust, or calculate/render a PivotTable report.
- Raw `xl/externalLinks/externalLink*.xml` packages are separately inspected
  for external-workbook, DDE, and OLE definitions. FormulaFence privately binds
  workbook declarations to package parts and fingerprints endpoint
  relationships, names, definitions, caches, item behavior, and opaque
  extensions. Reports expose only structural counts: targets, sheet and defined
  names, DDE services/topics/items, OLE program/item names, cached values, and
  extension payloads remain private. A material package change emits `FF025`
  and can be blocked with `no_external_link_package_changes`. FormulaFence does
  **not** follow or execute these links, establish source trust, or infer
  returned data.
- Excel 4.0 / XLM macro sheets are read directly from their raw Macro Sheet XML
  package parts before a workbook library can omit their executable cells.
  FormulaFence binds the documented workbook relationships to those parts,
  privately fingerprints complete XML plus related-part relationships, and
  streams direct safe internal targets into private payload fingerprints. It
  reports only structural counts. Commands, cell values, relationship targets,
  and embedded-object payloads remain private. A material change emits `FF026`
  and can be blocked with `no_xlm_macro_sheet_changes`. FormulaFence does
  **not** execute, emulate, resolve, or parse any XLM command, related target,
  or embedded object, and it never follows an external target. Direct internal
  payload streams are bounded to 32 MiB per part, 64 MiB per workbook, and 256
  parts. Oversized, missing, unreadable, over-budget, malformed, unbound, or
  unrecognized parts remain visible parser-coverage warnings rather than being
  silently ignored.
- Office RibbonX custom UI is read directly from its root-package declarations
  and `customUI` XML parts before the workbook reader can omit it. FormulaFence
  recognizes the documented 2006 and Office 2010-era package forms, privately
  fingerprints complete control XML and direct relationships, and reports only
  structural counts for parts, controls, callback attributes, image
  relationships, and external relationships. Control IDs, labels, callback
  names, XML, and targets remain private. A material change emits `FF027` and
  can be blocked with `no_ribbon_customization_changes`. FormulaFence does
  **not** invoke callbacks, follow an external relationship, or parse image
  payloads. Missing, oversized, malformed, unbound, version-mismatched, or
  otherwise unrecognized parts remain visible parser-coverage warnings.
  Custom-UI XML reads are bounded to 16 MiB per part, 32 MiB per workbook, and
  eight parts.
- Office Web Add-in task panes are read directly from the documented workbook
  task-pane relationship, `taskpanes.xml` parts, their direct
  task-pane-to-extension bindings, and direct `webextension*.xml` definitions
  before the workbook reader can omit them. FormulaFence privately fingerprints
  task-pane configuration, add-in references, auto-show properties, bindings,
  snapshots, and direct relationship semantics while reporting only safe
  structural counts. Add-in identities, store references, property/binding
  values, XML, snapshots, and relationship targets remain private. A material
  change emits `FF028` and can be blocked with
  `no_office_web_addin_changes`. FormulaFence does **not** install, load,
  execute, or fetch an add-in or manifest, and it never follows external
  relationships. Missing, oversized, malformed, unbound, or over-budget parts
  remain visible parser-coverage warnings. Task-pane and web-extension XML
  reads are bounded to 16 MiB per part, 32 MiB per workbook, and 64 parts.
  Worksheet-scoped web-extension markup outside this task-pane chain is not
  yet modeled.
- PivotTable packages are followed through the bounded workbook cache and
  worksheet PivotTable relationship graph, then privately fingerprinted as
  view layout, cache-schema, shared-item, normalized relationship, and bounded
  raw cache-record material. Profiles expose only structural counts; names,
  source ranges, fields, item values, formulas, cache records, targets, XML,
  and payload bytes remain private. A material change emits `FF031` and can be
  blocked with `no_pivot_table_definition_changes`. Cache source and refresh
  settings remain under `FF023` / `no_external_data_connection_changes` so a
  refresh-only edit remains distinct. Relationship IDs, equivalent internal
  target spellings, and cache-ID renumbering are normalized. FormulaFence does
  **not** refresh a cache, calculate or render a PivotTable, infer
  PivotTable-to-cell impact, fetch an external target, or interpret OLAP,
  or extension-list semantics. Slicer and Timeline cache definitions are
  compared separately. Missing, malformed, orphaned, unbound,
  oversized, or over-budget material remains a visible parser-coverage warning.
  PivotTable/cache-definition XML reads are bounded to 16 MiB per part, 64 MiB
  per workbook, and 512 parts; raw cache-record hashes are bounded to 32 MiB per
  part, 64 MiB per workbook, and 512 parts. FormulaFence detaches cache-record
  relationships in a temporary reader copy before the underlying workbook
  library loads cells, so raw records are not eagerly materialized; the original
  workbook is never changed.
- Slicer and Timeline cache definitions are followed from their documented
  workbook extension declarations through explicit workbook relationships to
  bounded cache XML. FormulaFence privately fingerprints Slicer item selection,
  Timeline state/filter material, cache definitions, PivotTable/table source
  bindings, filtered-PivotTable bindings, and unexpected direct cache-part
  relationships while exposing only structural counts. Cache names, source
  fields, selected values, date ranges, PivotTable names, relationship targets,
  and XML remain private. A material change emits `FF032` and can be blocked
  with `no_slicer_timeline_cache_changes`. Relationship IDs, equivalent
  internal target spellings, coordinated Slicer/Timeline PivotCache extension-ID renumbering, known
  optional Slicer defaults, Boolean spellings, and Timeline GUIDs are
  normalized. FormulaFence does **not** apply a filter, calculate/render a
  PivotTable or table, infer downstream impact, fetch an external target, or
  model worksheet/drawing Slicer or Timeline view geometry/styles. Missing,
  malformed, orphaned, unbound, externally targeted, oversized, or over-budget
  material remains a visible parser-coverage warning. Cache XML reads are
  bounded to 16 MiB per part, 64 MiB per workbook, and 512 parts.
- Embedded Power Pivot/Data Model packages are followed from the workbook's
  explicit `powerPivotData` relationship and `x15:dataModel` declaration.
  FormulaFence privately fingerprints declaration material, normalized workbook
  relationship semantics, and bounded raw `xl/model/*.data` payloads while
  exposing only model-part, binding, declaration, table, relationship, payload,
  and coverage counts. Table/column/relationship names, connection details,
  DAX, stored values, targets, XML, and raw bytes remain private. A material
  change emits `FF033` and can be blocked with
  `no_power_pivot_data_model_changes`. Relationship IDs, equivalent internal
  target spellings, and GUIDs in Data Model metadata are normalized. FormulaFence
  does **not** deserialize the Analysis Services payload, evaluate DAX, refresh
  the model, calculate/render a report, infer model-to-cell impact, or fetch an
  external target. Missing, malformed, orphaned, unbound, externally targeted,
  unexpected directly related, oversized, or over-budget material remains a
  visible parser-coverage warning. Raw payload reads are bounded to 512 MiB per
  part, 512 MiB per workbook, and 16 parts.
- DrawingML chart definitions and cached presentation data are followed from
  standard worksheet or chartsheet `drawing` relationships through chart parts
  and direct `userShapes` overlays. FormulaFence privately fingerprints
  non-cache chart definition material separately from `numCache`, `strCache`,
  and `multiLvlStrCache` material, plus overlay XML, normalized relationship
  semantics, and bounded direct internal related-part payloads. Profiles expose
  only structural counts; chart formulas, cached values, titles, shape text,
  relationship targets, XML, and payload bytes remain private. A material
  change emits `FF030` and can be blocked with
  `no_chart_definition_changes`. FormulaFence does **not** calculate a series
  formula, render a chart, infer chart-to-cell impact, follow an external
  target, parse media or embedded-package formats, or interpret modern
  `chartEx`/nested-chart semantics. Missing, malformed, orphaned, unbound,
  oversized, over-budget, or unrecognized material remains a visible
  parser-coverage warning. Chart and overlay XML reads are bounded to 16 MiB
  per part, 64 MiB per workbook, and 512 parts; direct related payload hashes
  are bounded to 32 MiB per part, 64 MiB per workbook, and 512 parts.
- Relationship-backed worksheet controls and OLE objects are read from raw
  worksheet control/OLE markup and direct control relationships before the
  workbook reader can omit them. FormulaFence also follows `vmlDrawing`
  relationships and privately fingerprints non-`Note` legacy VML `ClientData`
  controls, including macro, linked-cell, source-range, camera-range, and
  directly referenced relationship material. It privately fingerprints
  worksheet declarations, ActiveX `ocx` persistence XML, form-control-property
  XML, relationship semantics, and bounded direct ActiveX binary and
  OLE/package payload hashes. Profiles report only structural counts; control
  names, class IDs, licenses, captions, macros, formulas/ranges, OLE identities,
  targets, XML, and payload bytes remain private. A material change emits
  `FF029` and can be blocked with `no_worksheet_embedded_control_changes`.
  FormulaFence does **not** initialize an ActiveX control, deserialize/open an
  OLE object or package, render a VML drawing, include ordinary comment notes in
  its control inventory, follow an external relationship, or infer event
  dispatch. Relevant XML reads are bounded to 16 MiB per part, 64 MiB per
  workbook, and 512 parts; direct payload hashing is bounded to 32 MiB per
  part, 64 MiB per workbook, and 512 parts. Missing, malformed, orphaned,
  unbound, oversized, or over-budget material remains a visible parser-coverage
  warning. VML/drawing layout, embedded payload formats, and behavior outside
  this relationship-backed chain are not modeled.
- Power Query Data Mashup custom XML is inspected without serializing its M
  formulas or data/source material. FormulaFence privately compares the
  `Section1.m` formula document, logical package content, stable query metadata,
  and formula-firewall permissions. Profiles expose only structural counts and
  safe controls; query names, locations, metadata values, embedded content,
  telemetry IDs, and user-bound permission bindings remain private. `sqmid`
  telemetry and result-only refresh metadata are intentionally ignored. A
  material change emits `FF024` and can be blocked with
  `no_power_query_changes`. FormulaFence does **not** execute M, refresh a
  query, establish source trust, or infer returned values.
- Explicit implicit intersection is inventoried for literal `@` display syntax,
  `@` applied to a function, and persisted `SINGLE()` OOXML. When `SINGLE()`
  has one direct static A1 cell or range argument with an unambiguous
  row/column intersection, FormulaFence selects that single-cell edge.
  Function results, names, table syntax, external/3-D forms, and ambiguous
  placements retain conservative visible inputs or remain
  unresolved; FormulaFence never evaluates an expression to discover a value.
  New explicit uses emit `FF017`. This is separate from supported table
  current-row `[@Column]` syntax. Formula-defined names containing explicit
  implicit intersection are not expanded because the caller location matters.
- Ordinary lexical names inside inline `LET` expressions and `LAMBDA` bodies
  are not workbook references and are excluded from unresolved-token reporting;
  FormulaFence still traces the static dependencies around them. A defined name
  whose whole definition is a static `LAMBDA` can also be expanded at a call
  site, preserving name scope and explicit argument edges. FormulaFence accepts
  standard and `_xlfn.LAMBDA`/`_xlpm.`/`_xlop.` OOXML spellings. Recursive or
  non-static named LAMBDAs and arbitrary custom functions remain outside this
  subset and stay visible as coverage gaps.
- Static internal 3-D A1 references such as `Jan:Mar!B2:B10` are expanded over
  every worksheet in the endpoint tab span. FormulaFence compares the resolved
  span when the same 3-D formula survives a workbook change, because moving,
  adding, or removing tabs can change its semantics. External 3-D references
  remain external-link hazards; malformed, endpoint-missing, and non-A1 3-D
  forms remain explicit coverage gaps.
- Explicit external-workbook references are detected. References assembled from
  text or macro code are not.
- A formula that the underlying tokenizer cannot inspect is recorded by cell
  location in the profile, and a newly introduced one emits `FF016`; its graph
  is deliberately omitted rather than partially guessed.
- It inventories sheet visibility, defined names, calculation settings, the VBA
  payload, XLM macro-sheet packages, RibbonX custom UI packages, Office Web
  Add-in task-pane packages, PivotTable view/cache-schema/shared-item/cached-
  record chains, Slicer and Timeline cache filter-definition chains, embedded
  Power Pivot/Data Model declaration/raw-payload chains, What-If Data Table and
  Scenario Manager declarations, worksheet/Table AutoFilter and row/column-
  visibility controls, ignored-error warning suppressions, relationship-backed Named Sheet
  View controls, non-chart Worksheet DrawingML regular/group shape controls, DrawingML chart
  definition/cached-presentation/overlay chains,
  relationship-backed worksheet ActiveX/form-control/legacy-VML/OLE chains, the
  protection controls above, external-data refresh controls, external-link
  packages, and private Power Query definition material. It does not yet
  interpret PivotTable OLAP or other extension-list semantics; deserialize or execute
  Power Pivot/Data Model content; apply Slicer/Timeline filters or model their
  worksheet/drawing view geometry/styles; modern
  `chartEx` or nested-chart semantics; future Named Sheet View extension/rich-
  sort or full differential-format semantics; worksheet pictures, connectors,
  graphic frames, SmartArt, and other non-`xdr:sp` drawing objects or
  chart-to-cell impact; Ribbon image payloads; VML/drawing control layout or
  comment content; embedded OLE/package formats; worksheet-scoped Web Add-in
  markup; Power Query runtime behavior or returned data; ordinary styles beyond
  direct protection assignments; complete Excel style-cascade results; or every
  OOXML part.
- The tool preserves Excel formula text and uses a limited A1-reference
  normalizer for peer-pattern detection; it is not an Excel-compatible parser
  or calculation engine.

For high-stakes use, treat FormulaFence as one control among independent review,
recalculation in the approved spreadsheet engine, input reconciliation, and an
appropriately qualified model owner.
