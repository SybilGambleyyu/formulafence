# Scope and threat model

FormulaFence is a static change-assurance layer. It answers whether an Excel
workbook's inspectable structure changed in a risky way; it does not certify
financial correctness or replace model review.

## Safety properties

- Workbook content stays on the machine running FormulaFence. The CLI makes no
  network requests.
- It loads formulas as text with `data_only=False`; it does not calculate them.
- It never executes VBA, XLM macro sheets, Python-in-Excel scripts, RibbonX
  callbacks, DDE, external links, Power Query, Power Pivot/DAX, Office Web
  Add-in or custom-function code, or worksheet ActiveX/OLE code; it does not
  contact a Python-in-Excel Microsoft Cloud or custom-function runtime.
- VBA payloads, XLM macro-sheet source material, RibbonX control/callback
  material, Office Web Add-in task-pane/worksheet/in-content material, and worksheet control/OLE
  material are compared through private fingerprints only.
- Embedded Power Pivot/Data Model declarations and bounded raw model payloads
  are compared through private fingerprints only; table names, relationships,
  DAX, stored values, and connection details are never emitted.
- Python-in-Excel source, environment definitions/identifiers, script indexes,
  formula arguments and locations, and raw XML are compared through private
  fingerprints only. Public output retains aggregate package, formula-call,
  script, environment, initialization, and coverage counts.
- The dedicated namespaced custom-function ledger compares candidate names,
  namespaces, cells, formulas, and arguments through private signatures only.
  Its public inventory retains aggregate formula-cell, call, and namespace
  counts; ordinary defined-name and semantic-diff output retains its normal
  reviewer context and is not a redacted ledger. A matching formula is not
  evidence that an Office Add-in is installed, trusted, or runnable.
- What-If Data Table output ranges, input-cell references, and raw formula
  metadata are compared through a private signature only. Cached scenario-output
  cells remain under the normal cell-diff boundary.
- Scenario Manager names, comments, user metadata, stored input values,
  input/result references, and raw declarations are compared through a private
  signature only. Cached worksheet cells remain under the normal cell-diff
  boundary.
- Worksheet DrawingML regular-shape, connector, and recognized SmartArt
  graphic-frame presentation, geometry, anchors, diagram component material,
  bounded direct Diagram Data image payloads, connector endpoint targets,
  macro assignments, text links, descriptions, relationship identifiers, and
  targets are compared through private
  fingerprints only. Profiles and reports retain structural counts, never the
  underlying shape, connector, or SmartArt content.
- Native worksheet image declarations, anchors, visual properties,
  relationship identifiers/targets, and bounded direct image payloads are
  compared through private fingerprints only. Profiles and reports retain
  aggregate structural counts, never image bytes or image metadata.
- Worksheet cell-hyperlink targets, locations, display overrides, ScreenTips,
  references, relationship identifiers, and revision UIDs are compared through
  private fingerprints only. Profiles and reports retain structural counts,
  never the underlying link material, and the CLI never follows a target.
- Worksheet sparkline source formulas, date-axis sources, destination cells,
  group properties, and colour definitions are compared through private
  fingerprints only. Profiles and reports retain aggregate structural counts,
  never the underlying source or presentation material.
- Worksheet print ranges and titles, header/footer text, page values,
  printer-setting references, and raw print-layout XML are compared through
  private fingerprints only. Profiles and reports retain only aggregate
  structural counts.
- Cell-border definitions, colours, style indexes, and cell/row/column targets
  are compared through private fingerprints only. Profiles and reports retain
  only aggregate structural counts.
- Legacy Excel Note text, authors, cell associations, comment properties,
  threaded-comment placeholder links, VML visibility/layout, relationship
  identifiers, targets, and GUIDs are compared through private fingerprints
  only. Profiles and reports retain aggregate structural counts, never the
  underlying Note content or VML.
- Modern threaded-comment text, cell references, timestamps, reply links,
  mention ranges, person names/user IDs/provider IDs, relationship identifiers,
  and GUIDs are compared through private fingerprints only. Profiles and
  reports retain structural counts, never the underlying collaboration content.
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
- Portfolio comparison recursively inventories only `.xlsx` and `.xlsm` files
  under each supplied directory, identifies a workbook solely by its relative
  path, and reports additions/removals rather than guessing renames. It keeps
  roots and absolute worker paths out of portfolio output, ignores transient
  Office `~$` lock files, rejects symlinked paths and paths that
  differ only by case, and bounds each directory to 512 supported workbooks by
  default. A malformed supported file produces redacted `FF078` evidence and a
  final incomplete exit status, while remaining paths are still reported.
- Cross-workbook portfolio impact evidence is candidate-only and local to the
  supplied inventory. FormulaFence retains raw external source spellings and
  package targets only as private parser state, then resolves a direct static
  A1 source, a direct workbook-scoped or sheet-local external name, or narrow
  package-indexed forms `[N]Sheet!A1`, `[N]!Name`, and
  `[N]Sheet!LocalName`. For an indexed form, `N` must select exactly one
  document-order `externalReference`, `externalLink` part, `externalBook`, and
  external `externalLinkPath` relationship before that target may normalize to
  one exact relative candidate. A workbook-scoped consumer alias may use one
  exact static indexed spelling or one direct sheet-local spelling;
  sheet-scoped/formula consumer aliases, caches, non-static package A1 forms,
  and ambiguous package shapes are not expanded. A source name must also expand
  completely to static internal A1 destinations in that source candidate; an
  explicit source sheet selects only that local scope, never a global or other
  sheet fallback. It never opens a target path,
  searches by basename, follows an absolute/UNC/URI/escaping path, fetches
  anything, evaluates a formula, trusts cached external-link values, or emits
  the stored external path or source-name spelling in portfolio evidence.
  Ordinary source and consumer defined-name declarations remain normal profile
  context. Static ranges stay lazy. A global 100,000-state default bound emits
  `FF080` and exit status 2 rather than presenting incomplete `FF079` impact
  evidence as exhaustive.
- CLI report output is refused when it resolves to an inspected workbook or
  policy, and portfolio output is refused inside either input directory. This
  keeps a reporting request from mutating evidence or changing a portfolio's
  inventory during review.

## What a finding means

An impact count traces explicit A1-style cell dependencies available in the
baseline and candidate. It is an aid to review, not a claim that the cells will
recalculate correctly in Excel. FormulaFence also emits deterministic shortest
path samples from the changed cell to sampled downstream formulas. These paths
are explicit static dependencies, not proof of runtime evaluation. A
formula-pattern finding means both immediate peers have the same relative
formula fingerprint while the changed middle cell does not; it is a focused
review prompt, not proof of an error.

For a portfolio `FF079`, the same distinction applies across candidate
workbooks: it records that a changed source cell can reach a formula through a
bounded, explicit static graph. It does not prove that Excel can update the
link, that a workbook is open, that a source is trusted, or what value any
formula will produce.

## Deliberate limits

- Supported files are `.xlsx` and `.xlsm`; legacy `.xls` and file-encrypted or
  password-to-open workbooks are outside scope. Workbook and worksheet
  protection flags inside an otherwise readable OOXML workbook are inspected as
  operational controls, not treated as encryption.
- Portfolio mode intentionally does not support legacy `.xls`, `.xlsb`,
  templates, add-ins, or `.ods` files, infer a rename/content match across
  different paths, recursively follow a symlinked workbook, or combine cell
  dependencies between workbooks. A policy is applied independently to every
  matched path; selectors and formula/impact limits are not a portfolio-wide
  policy language. The scanner is sequential to keep resource use bounded; an
  incomplete entry is not treated as unchanged.
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
  formula sensitivity, render views, track ordinary positive widths/heights
  outside its dedicated worksheet-dimension boundary or styles, or model
  outline-display settings.
- Positive worksheet dimensions can conceal wrapped detail, reframe visible
  report context, or shift automatic pagination while leaving formulas and
  values untouched. FormulaFence privately compares raw transitional and strict
  `sheetFormatPr` default row/column/base widths, meaningful default
  `customHeight`, Office 2010 `x14ac:dyDescent` baseline adjustments, and active
  thick-border automatic adjustments; direct row
  height/custom-height/baseline/automatic-thick-border declarations; and effective raw
  `<cols>/<col>` positive width and `bestFit` state. Layered column records are
  resolved in XML order, with later records changing only the attributes they
  supply. `FF058` reports a material declaration change and
  `no_worksheet_dimension_changes` blocks it as `FFP058`. Profiles expose only
  aggregate counts; values, sheet names, row/column targets, writer hints, and
  raw XML remain private. Decimal/Boolean spellings, baseline defaults, inert
  fixed-height thick-border flags, `customWidth`, and equivalent effective range
  segmentation normalize away. Zero/hidden dimensions remain `FF036` visibility
  controls. Malformed, duplicate, unsupported, or budget-exhausted metadata is
  a coverage warning. FormulaFence does **not** calculate final AutoFit sizes,
  text overflow, merged-cell layout, exact automatic page breaks, print
  geometry, or client-specific rendering.
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
- Legacy Excel Custom Views can preserve a named alternate workbook display or
  print mode through `customWorkbookView` declarations and GUID-linked
  `customSheetView` records on every workbook sheet. They can alter hidden
  rows/columns, filters, print settings, panes, formula/gridline display,
  comments, and object visibility while ordinary cells and active views stay
  fixed. FormulaFence parses the raw workbook and supported worksheet,
  dialog-sheet, and chart-sheet declarations, reconciles each GUID privately,
  and emits `FF060`; `no_custom_workbook_view_changes` can block it as
  `FFP060`. Profiles and `FF060` details expose only structural counts for
  workbook/per-sheet views, affected sheets, hidden/filter/print/display
  settings, and unrecognized metadata. View names, GUIDs, bindings, ranges,
  filters, panes, print settings, and raw XML remain private. Coordinated GUID
  and sheet-ID/active-sheet-ID rewrites plus Boolean/default and
  unsigned-integer spelling normalize. Missing, duplicate, malformed,
  unsupported, unsafe, oversized, over-budget, or incompletely linked metadata
  is visible coverage evidence. FormulaFence does not activate/render a Custom
  View, calculate an alternate filtered result, determine final print output,
  interpret future extensions, or support Custom Views on other sheet types.
- Excel Tables can change a report's review surface through `tableStyleInfo`
  bindings/toggles, custom Table Style definitions, resolved Dxf material, and
  Table/TableColumn direct Dxf or named-cell-style references even when cell
  values, formulas, and table references remain fixed. FormulaFence compares
  those raw declarations privately and emits `FF061`; the
  `no_table_style_control_changes` rule can block it as `FFP061`. Profiles and
  `FF061` details expose only structural counts for declarations, styled/custom
  styles, Dxf/named-style assignments, banding/emphasis, and unrecognized
  metadata—never table/style names, formatting, colours, IDs, or raw XML.
  Boolean/default spelling, case-only names, `xr9:uid` revision provenance, and
  coordinated Dxf ID rewrites normalize. Missing, duplicate, malformed,
  unresolved, unsupported, oversized, or over-budget material remains visible
  coverage evidence. FormulaFence does not render resulting tables, resolve
  themes, calculate values, apply conditional formatting, cover PivotTable-only
  style regions, treat `defaultTableStyle` as an existing-table binding, or
  resolve a same-name named cell-style definition.
- Legacy shared-workbook revision headers and logs can preserve a private audit
  trail outside ordinary cells: prior/new values, locations, authors,
  timestamps, comments, formatting records, conflict-resolution material, and
  shared/tracking/retention/protection controls. FormulaFence follows the
  workbook-to-header and header-to-log relationships, fingerprints complete
  bounded raw declarations privately, and emits `FF062`; the
  `no_shared_workbook_revision_changes` rule can block it as `FFP062`. Profiles
  and `FF062` details expose only structural header/log parts and record counts,
  aggregate control counts, and unrecognized metadata—never historic values,
  locations, identities, timestamps, comments, GUIDs, relationship IDs, or raw
  XML. Equivalent Boolean/integer spelling, coordinated relationship-ID
  rewrites, and transitional/Strict relationship type spelling normalize.
  Missing, duplicate, malformed, unsafe, unsupported, oversized, or
  over-budget declarations remain visible coverage evidence. FormulaFence does
  not apply revisions, reconstruct a historical state, resolve conflicts,
  validate identity/timestamp claims, render Excel, or interpret arbitrary
  future extensions.
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
  values, model width/overflow, or compose number formats with separately
  inventoried font/fill/alignment/border or Table Style controls, quote
  prefixes, or arbitrary visual formatting. Column styles
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
  against a fill, calculate text/background contrast or values, compose font
  rendering with fill/border/alignment or other display controls, rich-text run
  rendering, separately inventoried Table Style controls, width/overflow, or
  arbitrary visual formatting. Column
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
  styles, apply separately inventoried Table Style controls, or claim arbitrary
  visual-style coverage. Column
  styles are compared only as OOXML defaults for unallocated/new cells, not as a
  claim to restyle allocated cells.
- Cell alignment can reposition, rotate, wrap, shrink, or indent an unchanged
  value, warning, or visual classification without a formula or value change.
  FormulaFence privately resolves raw `alignment` children in base
  `cellStyleXfs` and effective `cellXfs` records, follows `xfId`
  and `applyAlignment`, and compares direct cell `s`,
  `customFormat=1` row `s`, and raw `<cols>/<col style>`
  assignments. It covers horizontal/vertical placement, text rotation,
  wrapping, shrinking, indentation, relative indentation, justification, and
  reading order. A material effective declaration change emits `FF054`;
  `no_cell_alignment_changes` blocks it as `FFP054`. Profiles and
  reports expose only default/direct/row/effective-column/malformed counts;
  alignment values, style IDs, and targets remain private. Equivalent explicit
  defaults, Boolean/integer spelling, inert `mergeCell` compatibility
  material, base-XF inheritance, `applyAlignment`, and effective
  column-range splitting normalize away. Missing, duplicate, malformed, or
  unsupported metadata remains a visible coverage warning. FormulaFence does
  not calculate width, height, merged layout, overflow, final visibility,
  font/fill/conditional-format composition, or Excel rendering. Column styles
  remain OOXML defaults for unallocated/new cells, not a renderer that restyles
  allocated cells. This boundary follows Microsoft's
  [SpreadsheetML alignment definition](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oi29500/e4ad6e3e-7702-4dbe-8c44-f5a4c686c440)
  and [CellFormat alignment semantics](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oi29500/68362a4b-5589-4504-b566-e8154dce1de3).
- Cell borders can redraw a report boundary, total, exception box, or warning
  without a formula or value edit. FormulaFence privately resolves raw
  transitional and strict SpreadsheetML `<borders>/<border>` definitions, base
  `cellStyleXfs`, effective `cellXfs` records with `borderId`, `xfId`, and
  `applyBorder`, direct cell `s`, `customFormat=1` row `s`, and raw
  `<cols>/<col style>` assignments. It covers left/right/top/bottom, Office
  2010 logical start/end, diagonal/direction, outline, stored line styles, and
  stored colours. A material effective control change emits `FF057`;
  `no_cell_border_changes` blocks it as `FFP057`. Profiles and reports expose
  only default/direct/row/effective-column/unrecognized counts; definitions,
  colours, style IDs, and targets remain private. Omitted/`none` sides,
  Boolean/colour spelling, unused diagonal payload, ineffective empty
  `outline="false"`, base-XF inheritance, `applyBorder`, and equivalent
  column-range splitting normalize away. Missing, duplicate, malformed, or
  unsupported material remains a visible coverage warning. Material
  `vertical`/`horizontal` inner sides under ordinary cell styles are also
  flagged as coverage gaps because their differential-format semantics are not
  modeled here. FormulaFence does **not** resolve theme/palette colours, choose
  adjacent-cell precedence, render a final visual style, apply
  conditional-format/table/differential-style borders, calculate print output,
  or infer client behavior. Column styles remain OOXML defaults for
  unallocated/new cells, not a renderer that restyles allocated cells. This
  boundary follows OOXML's
  [`border`](https://c-rex.net/samples/ooxml/e1/part4/OOXML_P4_DOCX_border_topic_ID0EVV35.html)
  and [`xf`](https://c-rex.net/samples/ooxml/e1/part4/OOXML_P4_DOCX_xf_topic_ID0E13S6.html)
  forms, plus Microsoft's [cell-border guidance](https://support.microsoft.com/en-us/Excel/apply-or-remove-cell-borders-on-a-worksheet).
- Positive worksheet dimensions can change a reviewer's usable surface without
  an ordinary cell edit: fixed row heights can cut off wrapped text, column
  widths can reframe report fields, and sizing can move automatic page breaks.
  FormulaFence privately scans raw transitional and strict `sheetFormatPr`
  defaults (`defaultRowHeight`, `defaultColWidth`, `baseColWidth`), meaningful
  default `customHeight`, Office 2010 `x14ac:dyDescent` baseline adjustments,
  and active automatic thick-border adjustments; direct row
  `ht`/`customHeight`/`x14ac:dyDescent`/`thickTop`/`thickBot`; and raw positive
  `<cols>/<col width>`/`bestFit` declarations. It resolves overlapping columns
  in file order, preserving only changes from later present attributes. A
  material change emits `FF058`; `no_worksheet_dimension_changes` blocks it as
  `FFP058`. Profiles and reports disclose aggregate counts only; dimensions,
  targets, raw XML, and writer hints remain private. Baseline defaults,
  decimal/Boolean spelling, inert fixed-height thick-border flags,
  `customWidth`, and equivalent effective range splitting normalize away.
  Zero/hidden dimensions stay under `FF036`. Malformed, duplicate, unsupported,
  or budget-exhausted controls remain coverage warnings. FormulaFence does
  **not** compute final AutoFit sizing, wrapped/merged layout, overflow, exact
  automatic page breaks, print geometry, or client rendering.
- A stored worksheet view can change a reviewer’s surface while leaving values
  and formulas untouched: zeroes can appear blank; formulas, gridlines,
  row/column headers, outline symbols, rulers, or page margins can be
  hidden/shown; gridlines can be recoloured; the view can be right-to-left or
  page-oriented; and panes can be split/frozen. FormulaFence
  privately compares raw non-default transitional and strict SpreadsheetML
  `sheetViews/sheetView` declarations
  for those controls and emits `FF055`;
  `no_worksheet_display_control_changes` blocks it as `FFP055`.
  Profiles and reports expose only structural counts; sheet names, target
  cells, pane positions, and raw XML remain private. Omitted/default controls,
  Boolean and active custom-gridline-colour spelling, and finite non-negative
  split decimals normalize away.
  Active-cell, selection, top-left navigation, and zoom remain deliberately
  outside this boundary to avoid routine writer churn. Missing, duplicate,
  malformed, or unsupported material yields coverage evidence. FormulaFence
  does **not** render Excel, resolve the effective palette colour, calculate
  viewport geometry/final visibility, interpret extension-specific views, or
  infer client state. This boundary follows the Open XML SDK
  [`SheetView` schema surface](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.sheetview?view=openxml-3.0.1)
  and Microsoft’s [worksheet display guidance](https://learn.microsoft.com/en-us/office/dev/add-ins/excel/excel-add-ins-worksheet-display).
- A saved worksheet print layout can omit printed content, repeat different
  titles, alter print gridlines/headings/centering, reframe paper with margins
  or setup controls, change header/footer text, or insert manual page breaks
  without an ordinary cell edit. FormulaFence privately compares raw
  transitional and strict SpreadsheetML workbook print-area/print-title defined
  names plus direct worksheet `printOptions`, `pageMargins`, `pageSetup`,
  `sheetPr/pageSetUpPr`, `headerFooter`, and row/column-break declarations.
  It emits `FF056`; `no_worksheet_print_layout_changes` blocks it as `FFP056`.
  Profiles and reports expose structural counts only; print ranges,
  header/footer text, page values, printer-setting references, and raw XML stay
  private. Omitted/default, Boolean, integer, and decimal spellings normalize
  away, as do inactive first/even headers or footers, disabled first-page
  numbering, scale overridden by active fit-to-page dimensions, and automatic
  break display state. Missing, duplicate, malformed, or unsupported material
  yields coverage evidence. FormulaFence does **not** render/preview a workbook,
  calculate page geometry/counts or automatic page breaks, resolve printer or
  client defaults, inspect printer-specific `devMode` data, or cover
  custom/legacy sheet-view and extension-specific print controls. This boundary
  follows Microsoft's [print-area guidance](https://support.microsoft.com/en-us/excel/set-or-clear-a-print-area-on-a-worksheet)
  and [`PageLayout` control surface](https://learn.microsoft.com/en-us/javascript/api/excel/excel.pagelayout?view=excel-js-preview).
- A workbook-level DrawingML Theme can alter colour, font, and effect schemes
  used by themed cells, charts, and drawing objects without a local style
  change. FormulaFence inspects the raw workbook Theme binding, Theme XML, and
  direct Theme-image relationships/payloads in transitional and strict OOXML
  namespaces. A material stored control change emits `FF053`;
  `no_workbook_theme_changes` blocks it as `FFP053`. Profiles and reports
  expose only aggregate Theme-part/scheme/relationship/image and
  malformed-metadata counts; Theme XML, scheme names, colours, font names,
  image payloads, relationship IDs, and targets remain private. Writer-selected
  relationship IDs/order and equivalent internal target spelling normalize
  away. Missing, duplicate, malformed, unsafe, unbound, unreadable, oversized,
  or over-budget metadata produces a visible coverage warning; reads are
  bounded to 16 MiB per part, 64 MiB per workbook, and 512 parts. FormulaFence
  does **not** resolve effective styles, render a workbook, calculate contrast,
  decode an image, fetch a target, calculate formulas, or infer client
  behavior. This boundary follows the Open XML SDK
  [WorkbookPart](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.packaging.workbookpart?view=openxml-2.20.0)
  Theme-part surface and Microsoft's
  [conditional-formatting guidance](https://learn.microsoft.com/en-us/office/open-xml/spreadsheet/working-with-conditional-formatting).
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
- An ordinary worksheet cell can retain the same friendly value while its stored
  hyperlink changes an external/file target, in-workbook location, display
  override, or ScreenTip. FormulaFence reads raw standard SpreadsheetML
  `hyperlink` and Office 2016 `xr:hyperlink` declarations plus their selected
  worksheet relationships before the ordinary reader can normalize them. It
  privately compares binding, declaration material, location, display/ScreenTip,
  and relationship type/target/mode semantics. A material change emits
  `FF047` and `no_cell_hyperlink_changes` blocks it as `FFP047`. Profiles and
  reports expose only aggregate worksheet/hyperlink,
  location/display/ScreenTip, relationship/external-relationship, and
  malformed-metadata counts; targets, references, locations, display strings,
  ScreenTips, relationship IDs, and revision UIDs remain private.
  Writer-chosen relationship IDs/revision UIDs, relationship ordering, and
  equivalent internal target spelling normalize away. Missing, duplicate,
  malformed, unbound, unsafe, unreadable, oversized, or over-budget metadata
  becomes a visible coverage warning; raw worksheet XML reads are bounded to
  16 MiB per worksheet, 64 MiB per workbook, and 512 parts. The ordinary reader
  receives a hyperlink-removed temporary copy only after raw inspection, so
  malformed markup cannot suppress the evidence. FormulaFence does **not**
  render, resolve, fetch, follow, or test a link; inspect linked content; infer
  reputation/trust-zone/client behavior; or evaluate a `HYPERLINK()` formula.
  Stored `HYPERLINK()` calls are separately covered by `FF064`, still without
  evaluating an argument or following its result. This raw worksheet-hyperlink
  boundary follows the Open XML
  [Hyperlink](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.hyperlink?view=openxml-3.0.1)
  and Office 2016
  [Hyperlink](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.office2016.excel.hyperlink?view=openxml-3.0.1)
  definitions.
- A formula can create a link without a stored worksheet `hyperlink` element,
  request data from an intranet or Internet service, render a URL-sourced image,
  bind a real-time provider, retrieve financial history, or query a stored Cube
  connection. FormulaFence inventories stored `HYPERLINK`, `WEBSERVICE`,
  `IMAGE`, `RTD`, `STOCKHISTORY`, and all documented Cube-family calls
  (`CUBEKPIMEMBER`, `CUBEMEMBER`, `CUBEMEMBERPROPERTY`, `CUBERANKEDMEMBER`,
  `CUBESET`, `CUBESETCOUNT`, and `CUBEVALUE`), including `_xlfn.` compatibility
  spellings in cells, formula-defined names, and named `LAMBDA` bodies. It
  privately fingerprints cell, function-inventory, and relevant named-definition
  material. Public profiles and `FF064` details expose only action-cell,
  formula-defined-name, `STOCKHISTORY`, and aggregate Cube-function counts, so
  an argument-only, name-definition-only, connection/query-only, or same-count
  retarget remains reviewable without disclosing an endpoint, market symbol,
  provider, formula, query, or location. A normal cell change that reaches an
  invoking action/provider formula through FormulaFence's static dependency
  graph also emits `FF064`, covering static sources such as `HYPERLINK(A1, ...)`
  or `STOCKHISTORY(A1, ...)` without reading `A1` as an endpoint or symbol.
  Dynamic or unresolved sources remain explicit formula-coverage limits.
  `HYPERLINK` is deliberately included even when an argument appears internal:
  it can be dynamically calculated, and the ledger does not evaluate it to
  decide that. FormulaFence does **not** calculate a formula, resolve/open/fetch
  a destination, click/follow a link, authenticate, load a COM object, start an
  RTD server, contact a market provider, query a Cube, or execute a provider. A
  material change emits `FF064`; enable
  `no_formula_external_action_changes` for `FFP064`. This boundary follows
  Microsoft's [link guidance](https://support.microsoft.com/en-US/Excel/work-with-links-in-excel),
  [`WEBSERVICE` reference](https://support.microsoft.com/en-US/Excel/functions/webservice-function),
  [`IMAGE` reference](https://support.microsoft.com/en-us/excel/functions/image-function),
  [`RTD` reference](https://support.microsoft.com/en-us/excel/functions/rtd-function),
  [`STOCKHISTORY` reference](https://support.microsoft.com/en-us/office/stockhistory-function-1ac8b5b3-5f62-4d94-8ab8-7504ec7239a8),
  and [`CUBESET` reference](https://support.microsoft.com/en-us/excel/functions/cubeset-function).
- Direct DDE-style formulas can carry a cross-process link without a normal
  worksheet function or a raw `externalLink` package. FormulaFence separately
  recognizes only the lexical `application|topic!item` shape described by the
  Windows [DDE overview](https://learn.microsoft.com/en-us/windows/win32/dataxchg/about-dynamic-data-exchange),
  skipping pipes inside string literals and ordinary quoted sheet names. It
  privately fingerprints direct worksheet formulas plus formula-defined names
  and named `LAMBDA` chains; public `FF074` output has only formula-cell, link,
  and defined-name counts. A material change or statically visible input to an
  invoking named `LAMBDA` emits `FF074`; enable
  `no_formula_dde_link_changes` for `FFP074`. FormulaFence does **not**
  evaluate a formula, resolve an endpoint, look up, launch, or contact a DDE
  server, send a DDE command, or determine whether Excel's local DDE security
  settings permit any action. Raw external-link DDE/OLE metadata remains under
  `FF025`.
- Python in Excel keeps executable source separately from its `PY` formula
  placeholder. FormulaFence recognizes stored `PY` spellings, privately
  fingerprints the documented 2023 `python.xml` package part and the
  separately stored 2022 `pythonScripts.xml` compatibility contract—their
  relationships, content types, code/environment/script XML, and stored
  formula binding—then exposes only safe aggregate counts. Both physical parts
  remain independently compared when they coexist; FormulaFence does not
  choose one runtime representation or assume they agree. A
  code/package/environment change, formula-binding change, or ordinary cell
  change that statically reaches a PY formula emits `FF065`;
  `no_python_in_excel_changes` blocks it as `FFP065`. This includes a source
  such as `=_xlfn._xlws.PY(0,0,A1)` without decoding the script index,
  interpreting `A1`, or parsing source as Python. Dynamic or unresolved inputs
  remain formula-coverage limits. Relationship-ID-only rewrites normalize;
  missing, malformed, unbound, oversized, unreadable, or over-budget metadata
  remains coverage evidence. XML reads are bounded to 16 MiB per part, 64 MiB
  per workbook, and 512 parts. FormulaFence does **not** execute Python,
  evaluate a PY formula, resolve a result, contact Microsoft Cloud, or validate
  runtime package availability. Changed PY formulas and values remain in the
  ordinary semantic diff by design, so it is not a redacted source-code vault.
  This boundary follows Microsoft's [Python in Excel introduction](https://support.microsoft.com/en-US/Excel/python/introduction-to-python-in-excel)
  and the OOXML [Python part](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/151e4bcd-90a0-4d82-8b98-f16bf273e4ff)
  definition.
- Office Add-in custom functions are defined in JavaScript or TypeScript and
  exposed to Excel through a manifest namespace. They can request and stream
  external data, but a normal workbook stores only the call text—not the
  manifest, code, or runtime identity. FormulaFence therefore inventories only
  a conservative namespaced-call candidate: the direct-call classifier excludes
  known native dotted Excel functions, workbook-defined names, and `_xlfn.` /
  `_xlws.` compatibility names. Unqualified VBA, COM, or XLL UDF-shaped calls
  are covered separately by `FF075`. Candidates inside formula-defined names
  and named `LAMBDA` bodies are propagated to their invoking worksheet formulas. A candidate call or a
  normal edit that statically reaches one emits `FF066`; enable
  `no_office_custom_function_changes` for `FFP066`. The public profile exposes
  only formula-cell, call, and namespace counts; names, formulas, arguments,
  and locations stay private. FormulaFence does not evaluate a call, resolve a
  candidate to an add-in, load a manifest or add-in, execute JavaScript, or
  request a service.
  It makes no claim that a candidate maps to a particular add-in or can run.
  Dynamic or unresolved inputs remain static-coverage limits. This boundary
  follows Microsoft's [custom-functions overview](https://learn.microsoft.com/en-us/office/dev/add-ins/excel/custom-functions-overview),
  [tutorial](https://learn.microsoft.com/en-us/office/dev/add-ins/tutorials/excel-tutorial-create-custom-functions),
  and [web-data guidance](https://learn.microsoft.com/en-us/office/dev/add-ins/excel/custom-functions-web-reqs).
- A bare unknown worksheet call can resolve through VBA, a COM/Automation
  add-in, an XLL, or another registered runtime, but formula text alone does
  not prove a provider exists, is trusted, or can run. FormulaFence therefore
  inventories a conservative unqualified-runtime-function candidate rather
  than resolving or loading anything. The direct classifier accepts a bare
  identifier only after excluding workbook-defined names, local `LET`/`LAMBDA`
  bindings, qualified/dotted calls, and a stable native Excel function
  catalogue. This includes current native spellings such as `XLOOKUP`,
  `VSTACK`, `FIELDVALUE`, and `PY`; the pinned catalogue avoids third-party
  parser-version drift, while a new native function can be conservatively
  reported until FormulaFence adds it. Candidate calls inside formula-defined
  names and named `LAMBDA` bodies propagate through nested, recursive, and
  sheet-local chains to their invoking formulas. A private candidate/definition
  change or normal static input edit that reaches one emits `FF075`; enable
  `no_unqualified_runtime_function_changes` for `FFP075`. Public output has
  only formula-cell, call, and relevant definition counts; names, formulas,
  arguments, cells, provider identities, and host details stay private.
  FormulaFence does not evaluate a formula, resolve/load a VBA, COM/Automation,
  XLL, or registered provider, inspect host trust settings, or execute code.
  Stored candidate definitions remain independently reviewable even when no
  worksheet formula invokes them, while static-input paths require an actual
  inspected call. Dynamic and unresolved inputs remain static-coverage limits. This boundary
  follows Microsoft's [native function catalogue](https://support.microsoft.com/en-us/office/excel-functions-alphabetical-b3944572-255d-4efb-bb96-c6d90033e188),
  [installed UDF guidance](https://support.microsoft.com/en-us/excel/user-defined-functions-that-are-installed-with-add-ins-reference),
  [VBA custom-function guidance](https://support.microsoft.com/en-us/excel/create-custom-functions-in-excel),
  and [XLL registration/call guidance](https://learn.microsoft.com/en-us/office/client-developer/excel/accessing-xll-code-in-excel).
- Excel's worksheet-capable `REGISTER.ID` function can register a DLL or code
  resource when needed and return its registration ID. FormulaFence keeps a
  dedicated private ledger for stored worksheet and formula-defined `REGISTER.ID` calls,
  including calls held in formula-defined names and named `LAMBDA` bodies. A
  call, relevant named-definition change, or ordinary static input change emits
  `FF067`; enable `no_worksheet_code_resource_registration_changes` for
  `FFP067`. Public output exposes only formula-cell, call, and relevant
  formula-defined-name counts; module paths, procedure names, type strings,
  formulas, arguments, cells, and name identities stay private. FormulaFence
  does not evaluate a formula, resolve a path, load a DLL/XLL, inspect trust
  settings, or determine whether registration succeeds. Dynamic or unresolved
  inputs remain static-coverage limits. This is separate from raw XLM macro
  sheet program scanning: Microsoft's [`CALL` reference](https://support.microsoft.com/en-us/office/call-function-32d58445-e646-4ffd-8d5e-b45077a5e995)
  states that `CALL` is macro-sheet only. The worksheet boundary follows
  Microsoft's [`REGISTER.ID` reference](https://support.microsoft.com/en-us/office/register-id-function-f8f0af0f-fd66-4704-a0f2-87b27b175b50).
- Legacy XLM `REGISTER` can be stored in a formula-defined name or named
  `LAMBDA`, a surface not represented by ordinary macro-sheet XML. FormulaFence
  separately inventories only that stored-definition form and propagates it
  through nested/sheet-local names to invoking formula cells. Microsoft's
  [`xlfRegister` Form 1 reference](https://learn.microsoft.com/en-au/office/client-developer/excel/xlfregister-form-1)
  documents DLL-function/command registration and macro types callable from a
  defined-name definition; [`Form 2`](https://learn.microsoft.com/en-au/office/client-developer/excel/xlfregister-form-2)
  documents XLL loading and activation. A material stored definition,
  invocation, or ordinary static input change emits `FF068`; enable
  `no_formula_defined_xlm_registration_changes` for `FFP068`. Public output
  exposes only invoking-cell, call, and relevant definition counts; module
  paths, procedure names, type strings, arguments, formulas, locations, and
  name identities remain private. FormulaFence does not evaluate a formula,
  execute an XLM macro, resolve a path, load a DLL/XLL, or inspect trust
  settings. Direct worksheet `REGISTER` formulas and raw XLM macro-sheet parts
  are intentionally outside this narrow boundary; dynamic/unresolved inputs
  remain static-coverage limits.
- Legacy XLM `EVALUATE` can be stored in a formula-defined name or named
  `LAMBDA`, where it parses a supplied text expression at calculation time.
  FormulaFence separately inventories only that stored-definition form and
  propagates it through nested/sheet-local names to invoking formula cells.
  Microsoft's [Excel expression-evaluation
  reference](https://learn.microsoft.com/en-us/office/client-developer/excel/excel-worksheet-and-expression-evaluation)
  identifies `EVALUATE` as the XLM function that reduces a valid character
  string to a worksheet value. A material stored definition, invocation, or
  ordinary static argument-input change emits `FF069`; enable
  `no_formula_defined_xlm_evaluation_changes` for `FFP069`. Public output
  exposes only invoking-cell, call, and relevant definition counts; expression
  text, formulas, arguments, locations, and name identities remain private.
  FormulaFence does not evaluate text, parse a runtime-generated expression,
  execute an XLM macro, or infer dependencies inside the expression it would
  produce. It traces only the stored call's own visible static argument edge.
  Direct worksheet `EVALUATE` formulas and raw XLM macro-sheet parts are
  intentionally outside this narrow boundary; runtime-text dependencies remain
  explicit static-coverage limits.
- Selected legacy XLM action and event-dispatch calls can also be stored in a
  formula-defined name or named `LAMBDA`, outside raw macro-sheet XML.
  FormulaFence inventories only `CALL`, `EXEC`, `EXECUTE`, `RUN`, `SEND.KEYS`,
  `ON.DATA`, `ON.DOUBLECLICK`, `ON.ENTRY`, `ON.KEY`, `ON.RECALC`, `ON.SHEET`,
  `ON.TIME`, and `ON.WINDOW`, then propagates them through nested and
  sheet-local names to invoking formula cells. Microsoft's [Excel C API
  reference](https://learn.microsoft.com/en-us/office/client-developer/excel/programming-with-the-c-api-in-excel)
  describes XLM command-equivalent functions and event traps such as
  `ON.ENTRY` and `ON.TIME`; its [DLL-access
  guidance](https://learn.microsoft.com/en-us/office/client-developer/excel/how-to-access-dlls-in-excel)
  documents `CALL` and `REGISTER` as XLM macro-sheet routes to DLL functions
  or commands. A material stored definition, invocation, or ordinary static
  input change emits `FF073`; enable `no_formula_defined_xlm_action_changes`
  for `FFP073`. Public output exposes only invoking-cell, selected-call, and
  relevant definition counts; targets, handler names, formulas, arguments,
  locations, and name identities remain private. FormulaFence does not
  evaluate a formula, resolve a target/handler, load a DLL, send DDE, execute a
  macro or program, or infer whether an action succeeds. Direct worksheet
  action calls and raw XLM macro-sheet parts remain intentionally outside this
  narrow boundary; this finite inventory does not claim to interpret arbitrary
  XLM commands, and dynamic/unresolved inputs remain static-coverage limits.
- Legacy XLM GET.CELL is an XLM information function. FormulaFence separately
  inventories only calls stored in formula-defined names and named LAMBDA
  bodies, then propagates them through nested and sheet-local names to invoking
  formula cells. Microsoft's [C API
  reference](https://learn.microsoft.com/en-us/office/client-developer/excel/programming-with-the-c-api-in-excel)
  identifies GET.CELL as xlfGetCell. A material stored definition, invocation,
  or ordinary static argument-input change emits FF070; enable
  no_formula_defined_xlm_get_cell_changes for FFP070. Public output exposes
  only invoking-cell, call, and relevant definition counts; information types,
  references, formulas, arguments, locations, and name identities remain
  private. FormulaFence does not evaluate a call, determine its information
  type, resolve a dynamic reference, render formatting or display text, inspect
  comments/protection, or simulate other Excel state. Direct worksheet GET.CELL
  formulas and raw XLM macro-sheet parts are intentionally outside this narrow
  boundary; dynamic/unresolved inputs remain static-coverage limits.
- Selected legacy XLM environment-information calls GET.WORKBOOK,
  GET.WORKSPACE, and GET.DOCUMENT are separately inventoried only when stored
  in formula-defined names and named LAMBDA bodies, then propagated through
  nested and sheet-local names to invoking formula cells. Microsoft's [C API
  reference](https://learn.microsoft.com/en-us/office/client-developer/excel/programming-with-the-c-api-in-excel)
  identifies workspace information functions such as GET.CELL and
  GET.WORKBOOK; its [xlfFree
  example](https://learn.microsoft.com/en-us/office/client-developer/excel/xlfree)
  demonstrates GET.WORKSPACE returning platform information, and its [expression
  evaluation reference](https://learn.microsoft.com/en-us/office/client-developer/excel/excel-worksheet-and-expression-evaluation)
  identifies GET.DOCUMENT as an XLM information function. A material stored
  definition, invocation, or ordinary static argument-input change emits
  FF071; enable no_formula_defined_xlm_environment_information_changes for
  FFP071. Public output exposes only invoking-cell, call, and relevant
  definition counts; information types, references, formulas, arguments,
  locations, and name identities remain private. FormulaFence does not
  evaluate a call, determine its information type, resolve a dynamic reference,
  or simulate workbook/workspace/document/client/add-in/printer state. It does
  not assert that a state-only workbook change alters a stored call. Direct
  worksheet calls and raw XLM macro-sheet parts remain intentionally outside
  this narrow boundary; dynamic/unresolved inputs remain static-coverage
  limits.
- Native CELL and INFO calls can observe file/location/content or operating
  environment information outside ordinary visible precedents. Native SHEET and
  SHEETS calls can observe workbook tab position or, with an omitted SHEETS
  reference, the workbook tab count. FormulaFence inventories all four in
  worksheet formulas, formula-defined names, and named LAMBDA bodies, then
  propagates private signals through nested and sheet-local names to invoking
  formula cells. Microsoft's [CELL function
  documentation](https://support.microsoft.com/en-us/office/cell-function-51bd39a5-f338-4dbe-a33f-955d67c2b2cf)
  notes that an omitted CELL reference can use the selected cell at calculation
  time; FormulaFence therefore aggregates that subset privately without
  inferring what the selection is. Microsoft's [INFO function
  documentation](https://support.microsoft.com/en-au/office/info-function-725f259a-0e4b-49b3-8b52-58815c69acae)
  describes operating-environment information such as directory, platform, and
  calculation mode. Microsoft's [SHEET function
  documentation](https://support.microsoft.com/en-us/excel/functions/sheet-function)
  and [SHEETS function
  documentation](https://support.microsoft.com/en-us/excel/functions/sheets-function)
  document that hidden, very-hidden, macro, chart, and dialog sheets are
  included. A material stored definition, invocation, or ordinary static
  argument-input change emits FF072; with a complete raw OOXML tab catalog,
  a tab membership/order/name change also emits FF072 for stored SHEET or
  omitted-reference SHEETS calls. Visibility-only changes do not satisfy that
  condition. Enable no_formula_environment_information_changes for FFP072.
  Public output exposes only formula-cell, call, relevant definition, and
  omitted-reference counts; information types, references, formulas, arguments,
  locations, name identities, and raw tab-catalog comparison material remain
  private; ordinary sheet inventory remains normal reviewer context.
  FormulaFence does not evaluate a call, determine an information type, resolve
  a dynamic argument, infer a selected cell, calculate a result, or simulate
  file/folder/client/workspace/workbook state. Explicit SHEETS references are
  not guessed as one-sheet versus 3-D; dynamic/unresolved inputs and incomplete
  tab catalogs remain explicit static-coverage limits.
- Office 2010 worksheet sparklines live in `x14:sparklineGroups` worksheet
  extensions, outside ordinary cell values. A group can be retargeted, moved,
  or have its type, axes, display, marker, line-weight, or colour controls
  changed; a nested sparkline can change its source formula or destination
  cell. FormulaFence reads raw x14 declarations before the ordinary reader
  drops them and privately compares group membership, source/date-axis
  formulas, destinations, and visual controls. A material change emits
  `FF048` and `no_worksheet_sparkline_changes` blocks it as `FFP048`. Profiles
  and reports expose only aggregate worksheet/group/sparkline,
  source/date-axis-source, colour-control, and malformed-metadata counts;
  formulas, locations, group properties, and colours remain private.
  Equivalent local direct-range spelling, Boolean/numeric spelling, colour
  case, and declaration order normalize away. Missing, duplicate, malformed,
  unreadable, oversized, or over-budget metadata becomes a coverage warning;
  raw worksheet XML is bounded to 16 MiB per worksheet, 64 MiB per workbook,
  and 512 parts. A Sparkline Group-removed temporary reader copy is made only
  after raw inspection, so lossy reader support cannot erase the evidence.
  FormulaFence does **not** calculate source values, resolve names/external
  sources, render a sparkline, assess visual accessibility, or guarantee
  cross-version Excel rendering equivalence. This boundary follows the Open
  XML [SparklineGroup](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.office2010.excel.sparklinegroup?view=openxml-3.0.1)
  and [CT_Sparkline](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/6b28a993-e0fd-451d-860e-35097c6baa77)
  definitions.
- SpreadsheetML XML Maps can attach an embedded schema and refresh/export
  behavior to XML table columns or individual worksheet cells. A changed map
  can redirect an XPath, switch a file/connection binding, change a target
  cell, or alter append, format, sort/filter, and validation behavior without
  changing ordinary cells. FormulaFence reads raw XML Maps, table
  XML-column-property, and single-cell table declarations before ordinary
  workbook readers discard them. It privately compares schemas, map/data-
  binding controls, table/single-cell bindings, and related
  workbook/worksheet relationship targets. A material change emits FF049 and
  `no_xml_mapping_changes` blocks it as FFP049. Profiles and reports expose
  only aggregate map/schema/binding, file/connection, table, single-cell, and
  malformed-metadata counts; schemas, names, XPath expressions, identities,
  cells, connection identities, and relationship targets remain private.
  Equivalent Boolean/unsigned-integer spellings, relationship IDs/order, and
  equivalent internal target spelling stay quiet. Missing, duplicate,
  malformed, unsafe, unbound, unreadable, oversized, or over-budget metadata
  becomes a coverage warning; reads are bounded to 16 MiB per part, 64 MiB per
  workbook, and 512 parts. FormulaFence does **not** import/export XML,
  validate XML data or schemas, open a file/connection, fetch data, calculate
  a refresh, or infer Excel client behavior. This boundary follows the Open XML
  [Map](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.map?view=openxml-3.0.1),
  [XmlProperties](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.xmlproperties.xpath?view=openxml-3.0.1),
  and [SingleXmlCells](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.singlexmlcells?view=openxml-3.0.1)
  definitions.
- Excel rich data types can place linked entity values, provider-backed fields,
  web-image associations, and worksheet value-metadata bindings outside normal
  cells. FormulaFence reads Rich Value Data/structure/type/array/supporting
  property-bag/style/web-image/rich-value-relationship parts, their
  workbook/package relationships, and `XLRICHVALUE` bindings before normal
  readers can omit them. A material change emits `FF051` and
  `no_rich_data_changes` blocks it as `FFP051`. Profiles and reports expose
  aggregate part/value/structure/array/property-bag/binding/bound-cell/image/
  relationship/external-reference and malformed-metadata counts only; entity
  values, provider data, field names, identifiers, URLs, image references,
  relationship IDs, and bound-cell locations remain private. Writer-selected
  relationship IDs/order and equivalent internal targets normalize away.
  Missing, duplicate, malformed, unsafe, unreadable, oversized, or over-budget
  metadata becomes a coverage warning; reads are bounded to 16 MiB per XML
  part, 64 MiB per workbook, and 512 parts. FormulaFence does **not** contact
  providers, refresh values, calculate formulas, fetch or validate targets, or
  infer Excel client behavior. This boundary follows Microsoft's
  [Rich Value Data](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/896934fd-8df7-43f4-b154-2d39371c270d),
  [Rich Value Structure](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/d90f6d91-d868-4b94-9d26-ec3b1492cec6),
  [Rich Value Types](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/5d213b66-3196-4516-b63c-eef80d926f4a),
  and [Rich Value Web Image](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/4f3a80fd-1776-407f-8807-2497a4692dea)
  definitions.
- Generic Custom XML, workbook-bound Custom Data, and custom document
  properties can retain an add-in's workbook-specific approval, workflow, or
  integration state outside ordinary cells. FormulaFence reads generic
  `customXml/item*.xml` data, its property/schema parts and relationships,
  workbook-linked `xl/customData` property/binary parts, and
  `docProps/custom.xml` before ordinary readers can omit them. Power Query
  `DataMashup` remains exclusively under the Power Query control boundary. A
  material persisted-state change emits `FF052` and
  `no_custom_data_store_changes` blocks it as `FFP052`. Profiles and reports
  expose aggregate part/schema/relationship/payload/document-property and
  malformed-metadata counts only; custom XML, schema URIs, property names and
  values, storage IDs, binary payloads, relationship IDs, and targets remain
  private. Writer-selected relationship IDs/order and document-property `pid`
  normalize away; Custom XML `itemID` and Custom Data `id` storage identities
  are compared privately because add-ins can bind state to them. Missing,
  duplicate, malformed, unsafe, unbound, unreadable, oversized, or over-budget
  metadata becomes a coverage warning; reads are bounded to 16 MiB per part,
  64 MiB per workbook, and 512 parts. FormulaFence does **not** execute an
  add-in, resolve a property, follow or fetch a target, interpret a binary
  payload, calculate formulas, or infer Excel client behavior. This boundary
  follows Microsoft's guidance on
  [persisting add-in state](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/persisting-add-in-state-and-settings),
  [Custom Data](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/7c53f6f4-fea8-43f7-a4b0-ba6e14d0eb78),
  and [Custom Data Properties](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/1f4aa666-c966-4ecf-8399-28390399c891).
- OPC package signatures and VBA project signatures are distinct stored
  integrity/provenance surfaces. A workbook can preserve ordinary cells and
  even `xl/vbaProject.bin` while the package-root signature origin, XML signature
  envelope/signed references, optional certificate-part relationship/payload, or
  classic/Agile/V3 VBA signature payload changes. FormulaFence reads those raw
  relationships and bounded parts before normal readers can omit them. A
  material envelope change emits `FF050` and
  `no_digital_signature_changes` blocks it as `FFP050`. Profiles and reports
  expose aggregate counts only; XML signature material, reference URIs,
  certificate identities/contents, binary signature payloads, relationship IDs,
  and targets remain private. Equivalent IDs/order/internal targets and XMLDSIG
  base64 whitespace normalize away. Missing, duplicate, malformed, unsafe,
  unbound, unreadable, oversized, or over-budget metadata becomes a coverage
  warning; reads are bounded to 16 MiB per part, 64 MiB per workbook, and 512
  parts. FormulaFence does **not** validate a signature/digest/transform,
  reference coverage, certificate chain/identity/trust/expiry/revocation,
  timestamp, signed contents, or VBA code; it does not fetch certificates or
  contact trust services. Microsoft's [OPC digital-signature
  overview](https://learn.microsoft.com/en-us/previous-versions/windows/desktop/opc/open-packaging-conventions-overview)
  assigns signer/trust validation to the package consumer.
- Traditional Excel Notes are stored in worksheet-associated SpreadsheetML
  comments parts and their display declarations live in worksheet
  `legacyDrawing` VML parts. FormulaFence follows the worksheet bindings and
  privately compares author association, text/rich-text presentation, cell
  association, comment properties, Note VML visibility/layout, and relationship
  semantics. It recognizes the `tc={GUID}` legacy placeholder used to reconcile
  a threaded comment, and treats that declaration as a separate guarded surface.
  A material change emits `FF046` and `no_legacy_comment_changes` blocks it as
  `FFP046`. Profiles and reports expose only aggregate worksheet/part,
  author/comment/text/rich-text/property/placeholder, Note-shape/visibility/
  anchor, relationship, and malformed-metadata counts; Note text, author
  identities, references, VML, targets, IDs, and GUIDs remain private.
  Consistent writer-generated VML shape/comment shape/relationship IDs and
  placeholder GUIDs normalize away. Missing, duplicate, malformed, unbound,
  unsafe, unreadable, oversized, or over-budget metadata becomes a visible
  coverage warning; XML reads are bounded to 16 MiB per part, 64 MiB per
  workbook, and 512 parts. The ordinary reader uses a Note-quarantined
  temporary copy only after raw inspection, so unsafe targets and
  parser-tolerance differences cannot erase evidence. FormulaFence does
  **not** render Notes/VML, resolve authors, fetch
  targets, execute linked content, calculate client placement, or infer
  notification, permission, account, cloud, or client-visibility behavior.
  This boundary follows the Open XML
  [Comment](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.comment?view=openxml-3.0.1),
  [Authors](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.authors?view=openxml-3.0.1),
  and [LegacyDrawing](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.legacydrawing?view=openxml-3.0.1)
  definitions, plus Microsoft's [threaded-comment placeholder
  rule](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/6383f002-c90b-401c-a1d7-66b97b14cb3e).
- Modern threaded comments are stored in worksheet-associated comment parts and
  a workbook-associated persons part, outside ordinary cells. FormulaFence
  follows those bindings and privately compares comment/reply graphs, stored
  text, cell/timestamp/resolution declarations, mention range/person links,
  extension material, and person records. A material change emits `FF045` and
  can be blocked with `no_threaded_comment_changes`. It normalizes consistent
  writer-generated comment, parent, person, mention, and package
  relationship-ID rewrites by rebuilding the private graph first. Profiles and
  reports expose only aggregate worksheet/part/thread, comment/reply/resolved/
  text, mention, person, relationship, and malformed-metadata counts. It does
  **not** render comments, validate mention offsets, notify or resolve users,
  determine legacy-placeholder rendering, or infer permissions, cloud state,
  or client visibility. Missing, duplicate, malformed, unbound, unsafe,
  unreadable, oversized, or over-budget metadata becomes a visible coverage
  warning; XML reads are bounded to 16 MiB per part, 64 MiB per workbook, and
  512 parts. The boundary follows Microsoft's
  [threaded-comment overview](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/e0fb917a-1107-409a-852f-13b47aea70dc),
  [Threaded Comments part](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/66e1875d-c60a-48eb-bf88-41066d45fea8),
  [Persons part](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/1a170d26-42a2-46f0-b2b6-0ff1dec1c344),
  and [schema](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/adb84732-9fc8-48b6-bddc-6b0bcdaad940).
- Non-chart Worksheet DrawingML regular shapes (`xdr:sp`), connectors
  (`xdr:cxnSp`), nested groups (`xdr:grpSp`), and recognized SmartArt
  `xdr:graphicFrame` objects are followed from standard worksheet `drawing`
  relationships before the workbook reader can discard their declarations.
  For a non-chart graphic frame, FormulaFence recognizes the DrawingML Diagram
  `a:graphicData` URI and requires one `dgm:relIds` declaration. It supports
  transitional and Strict DrawingML, privately fingerprints supported
  anchor/layout and frame/shape/group/connector XML; the explicitly bound
  diagram data (`r:dm`), layout (`r:lo`), quick-style (`r:qs`), and colours
  (`r:cs`) parts; direct worksheet-drawing `diagramDrawing` rendering parts;
  and bounded direct internal Image targets from a Diagram Data part; connector
  `stCxn`/`endCxn` attachment semantics; macro assignments; text
  links; click/hover relationship semantics; and visible text/presentation
  declarations. Profiles expose only safe worksheet/drawing/anchor,
  shape/text/connector/group, graphic-frame/SmartArt-component, Diagram Data
  image part/fingerprinted/uninspected, connector-attachment, text paragraph/
  run, macro/text-link/hyperlink, relationship, and malformed-control counts.
  A material change emits `FF044` and can be blocked with
  `no_worksheet_drawing_shape_changes`. Consistent
  non-visual and connector endpoint ID rewrites, worksheet-DrawingML
  relationship-ID rewrites, and colour-case spelling are normalized.
  FormulaFence does **not** render DrawingML, resolve themes or contrast,
  calculate text links, execute macro assignments, retrieve external targets,
  calculate final SmartArt layout, or decode/render media. It hashes only
  bounded direct internal Diagram Data Image targets (32 MiB per image, 64 MiB
  per workbook, and 512 images), and does not follow any other component-side
  SmartArt relationship. Native pictures are handled by the separate
  worksheet-image boundary, chart frames remain in `FF030`, and unknown
  non-chart graphic-frame URI types are coverage gaps. Missing, duplicate,
  malformed, unsafe, oversized, over-budget, or unsupported metadata becomes
  visible parser-coverage evidence. XML reads are bounded to 16 MiB per part,
  64 MiB per workbook, and 512 parts. This scope follows the Open XML
  [`xdr:sp` Shape definition](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.spreadsheet.shape?view=openxml-3.0.1),
  [`xdr:cxnSp` ConnectionShape definition](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.spreadsheet.connectionshape?view=openxml-3.0.1),
  Microsoft's [Graphic Object Data](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oe376/f58e82a5-5590-4e36-b178-e12989960415),
  the OOXML [Diagram Data Part](https://ooxml.info/docs/14/14.2/14.2.4/),
  and [Diagram relationship IDs](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.diagrams.relationshipids?view=openxml-3.0.1)
  references.
- Native worksheet image controls are followed from worksheet `drawing`, direct
  `picture`, and `legacyDrawingHF` relationships before ordinary readers can
  discard those visual bindings. FormulaFence privately compares anchored
  transitional/strict DrawingML `xdr:pic` objects (including group-contained
  pictures), worksheet backgrounds, and VML-backed header/footer watermark
  images, along with their anchors, visual declarations, relationship
  semantics, and bounded direct payload hashes. Profiles and reports expose
  only safe worksheet/picture/anchor/background/header-footer/image-payload/
  relationship/malformed-control counts. Image bytes, image names/descriptions,
  visual formatting, anchors, relationship IDs/targets, and raw XML remain
  private. A material change emits `FF059` and can be blocked with
  `no_worksheet_image_changes`. Non-visual DrawingML/VML IDs and consistent
  relationship-ID rewrites normalize. FormulaFence does **not** render or
  decode media, fetch a target, resolve themes, calculate visibility, cropping,
  z-order, print pagination, or client behavior. Charts, rich-data/in-cell
  images, Theme images, ActiveX/OLE image controls,
  regular/group/connector/SmartArt drawing controls, and
  header/footer text remain in `FF030`, `FF051`, `FF053`, `FF029`, `FF044`, and
  `FF056` respectively. XML reads are bounded to 16 MiB per part, 64 MiB per
  workbook, and 512 parts; direct payload hashing is bounded to 32 MiB per
  part, 64 MiB per workbook, and 512 parts. Missing, duplicate, malformed,
  unsafe, unreadable, oversized, or over-budget material is visible coverage
  evidence. The boundary follows Open XML's
  [`xdr:pic` Picture definition](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.spreadsheet.picture?view=openxml-3.0.1),
  Microsoft's [worksheet background guidance](https://support.microsoft.com/en-us/excel/add-or-remove-a-sheet-background),
  and [header/footer watermark guidance](https://support.microsoft.com/en-us/excel/get-started/add-a-watermark-in-excel).
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
- Every canonical root or part-level OPC relationship part is also inspected
  independently for `TargetMode="External"`, including opaque relationships
  no feature-specific scanner can reach. FormulaFence retains source, type,
  endpoint, and malformed-metadata evidence only in private signatures and
  exposes aggregate relationship part/source/target plus hyperlink/image/other
  counts. A material change emits `FF063` and can be blocked with
  `no_external_relationship_changes`; relationship-ID-only rewrites normalize.
  Duplicate, orphaned, malformed, unsafe, unreadable, oversized, or
  over-budget metadata is coverage evidence. FormulaFence does **not** resolve,
  fetch, open, execute, or establish trust for any relationship target. XML
  reads are bounded to 16 MiB per part, 64 MiB per workbook, and 512 parts.
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
- Legacy XLM automatic-macro routing is separately inspected from raw workbook
  defined names. Microsoft documents four workbook automatic-macro events:
  `Auto_Open`, `Auto_Close`, `Auto_Activate`, and `Auto_Deactivate`. FormulaFence
  recognizes an optional `_xlnm.` built-in prefix and counts only a
  workbook-scoped event name whose direct internal single-cell A1 definition
  targets a sheet declared through a raw XLM macro-sheet relationship. This catches a dispatch
  add, removal, or same-count retarget without asserting that a local name,
  ordinary-sheet target, external reference, or dynamic definition will run.
  The private signature retains the name/target/definition material; profiles,
  `FF076`, and `FFP076` expose only aggregate per-event counts. `FF076` is high
  severity and `no_xlm_automatic_macro_binding_changes` makes it `FFP076` in
  CI. FormulaFence does **not** evaluate or resolve a defined name, use the
  reserved/unused `definedName@xlm` attribute, parse or execute an XLM command,
  determine macro security settings, or claim that Excel will execute a
  binding. Missing/malformed workbook or relationship metadata remains a
  parser-coverage warning. The ordinary defined-name diff retains its normal
  reviewer context rather than becoming a redacted output channel.
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
- Office Web Add-ins are read directly from the documented workbook task-pane
  relationship, `taskpanes.xml` parts, direct task-pane-to-extension bindings,
  worksheet `x15:webExtensions` entries, and active in-content DrawingML
  `we:webextensionref` frames before the workbook reader can omit them.
  FormulaFence validates worksheet `appRef` entries against definition
  bindings, skips inactive `mc:Fallback` frame branches, and leaves an active
  frame's native-picture fallback to the worksheet-image boundary. It privately
  fingerprints task-pane configuration, add-in references, auto-show
  properties, bindings, worksheet formulas, snapshots, frame placement/XML,
  and direct relationship semantics while reporting only safe structural
  counts. Add-in identities, store references, property/binding values,
  formulas, XML, snapshots, and relationship targets remain private. A
  material change emits `FF028` and can be blocked with
  `no_office_web_addin_changes`. FormulaFence does **not** install, load,
  execute, or fetch an add-in or manifest, and it never follows external
  relationships. Missing, oversized, malformed, unbound, or over-budget parts
  remain visible parser-coverage warnings. Task-pane and web-extension XML
  reads are bounded to 16 MiB per part, 32 MiB per workbook, and 64 parts;
  worksheet-binding and in-content DrawingML reads are each bounded to 16 MiB
  per part, 64 MiB per workbook, and 512 parts. Other unrecognized extension
  or graphic-frame forms remain outside this boundary.
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
  standard worksheet or chartsheet `drawing` relationships through legacy
  `c:chart` parts, direct `c:userShapes` overlays, and Office 2016+ `cx:chart`
  ChartEx parts. ChartEx `mc:AlternateContent` graphic-frame bindings are
  recognized without treating their older-client fallback shape as a second
  worksheet control. FormulaFence privately fingerprints non-cache legacy
  chart material separately from `numCache`, `strCache`, and `multiLvlStrCache`
  material, plus overlay XML, normalized relationship semantics, ChartEx XML,
  and bounded direct internal related-part payloads. Supported ChartEx direct
  relationships are style, colour-style, drawing, image, theme-override, and
  embedded package; unsupported, external, or unsafe edges remain explicit
  coverage evidence. Profiles expose only structural counts; chart formulas,
  cached values, titles, shape text, relationship targets, XML, and payload
  bytes remain private. A material change emits `FF030` and can be blocked with
  `no_chart_definition_changes`. FormulaFence does **not** calculate a series
  formula, render a chart, infer chart-to-cell impact, follow an external
  target, parse media or embedded-package formats, resolve ChartEx second-hop
  relationships, or interpret ChartEx-specific or nested-chart visualization
  semantics. Missing, malformed, orphaned, unbound, unsupported, oversized,
  over-budget, or unrecognized material remains a visible parser-coverage
  warning. Chart and overlay XML reads are bounded to 16 MiB per part, 64 MiB
  per workbook, and 512 parts; direct related payload hashes are bounded to 32
  MiB per part, 64 MiB per workbook, and 512 parts. The relationship boundary
  follows Microsoft's [ChartEx part](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-odrawxml/5d0d453e-adac-43be-a797-59b9916593dd)
  and [ChartEx relationship-ID](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-odrawxml/d8ede39e-a36c-48ad-8a17-0086a2d0889b)
  definitions.
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
  payload; OPC package XML-signature/certificate-part relationships and
  payloads; classic/Agile/V3 VBA signature payloads/relationships; XLM
  macro-sheet packages; RibbonX custom UI packages; Office Web
  Add-in task-pane packages, PivotTable view/cache-schema/shared-item/cached-
  record chains, Slicer and Timeline cache filter-definition chains, embedded
  Power Pivot/Data Model declaration/raw-payload chains, What-If Data Table and
  Scenario Manager declarations, worksheet/Table AutoFilter and row/column-
  visibility controls, material worksheet-display and worksheet print-layout
  controls, ignored-error warning suppressions, relationship-backed Named Sheet
  View and legacy Custom View controls, ordinary worksheet-cell hyperlink
  declarations/relationships,
  Office 2010 worksheet sparkline declarations, SpreadsheetML XML Map schema,
  refresh/export, table-column, single-cell, and relationship declarations,
  legacy Excel Note/comments/VML and threaded-placeholder package chains,
  modern threaded-comment/person package chains, non-chart Worksheet DrawingML
  regular/group/connector/recognized-SmartArt graphic-frame controls, native
  worksheet picture/background/header-footer image controls, legacy and ChartEx
  DrawingML chart definition/cached-presentation/overlay chains,
  relationship-backed worksheet ActiveX/form-control/legacy-VML/OLE chains, the
  protection controls above, external-data refresh controls, external-link
  packages, package-wide external OPC relationships, and private Power Query
  definition material. It does not yet
  interpret PivotTable OLAP or other extension-list semantics; deserialize or execute
  Power Pivot/Data Model content; apply Slicer/Timeline filters or model their
  worksheet/drawing view geometry/styles; ChartEx-specific visualization,
  second-hop relationship, or nested-chart semantics; future Named Sheet View extension/rich-
  sort or full differential-format semantics; unknown non-chart
  graphic-frame URI types, SmartArt rendering/final-layout behavior, or
  SmartArt component-side relationship targets other than bounded direct
  Diagram Data Image payloads; chart-to-cell impact; Ribbon image payloads;
  general VML/drawing-control
  layout beyond supported Note shapes; embedded OLE/package formats; unsupported
  worksheet Web Add-in extension or graphic-frame forms; Power Query runtime behavior or returned
  data; ordinary styles beyond direct protection assignments; complete Excel
  style-cascade results; or every OOXML part.
- The tool preserves Excel formula text and uses a limited A1-reference
  normalizer for peer-pattern detection; it is not an Excel-compatible parser
  or calculation engine.

For high-stakes use, treat FormulaFence as one control among independent review,
recalculation in the approved spreadsheet engine, input reconciliation, and an
appropriately qualified model owner.
