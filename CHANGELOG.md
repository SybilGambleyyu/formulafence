# Changelog

## 0.55.0 — 2026-07-26

- Inspect material effective cell-border controls directly from raw
  transitional and strict SpreadsheetML before ordinary readers flatten style
  inheritance: reusable `<borders>/<border>` definitions, `borderId`, base
  `cellStyleXfs`, `xfId`/`applyBorder`, direct-cell, custom-row, and column
  assignments. The boundary covers ordinary edge sides, Office 2010 logical
  start/end sides, diagonals, outline, styles, and stored colours. Profiles,
  Markdown, JSON, and SARIF expose aggregate counts only; border definitions,
  colours, style IDs, and cell/row/column targets remain private.
- Emit `FF057` for a material effective cell-border control change and add the
  fail-closed `no_cell_border_changes` policy rule (`FFP057`). This closes the
  review gap where a report boundary, total, exception box, or warning cue can
  change while ordinary cell values and formulas stay fixed.
- Normalize omitted/`none` side declarations, Boolean and colour spellings,
  unused diagonal payload, ineffective empty `outline="false"`, base-XF
  inheritance, `applyBorder`, and equivalent effective column-range splitting.
  Missing, duplicate, malformed, or unsupported metadata is an explicit
  coverage warning. FormulaFence compares stored declarations only: it does
  not render Excel, resolve theme/palette colours, choose adjacent-cell border
  precedence, apply conditional-format/table/differential-style borders,
  calculate print layout, or infer client behavior.

## 0.54.0 — 2026-07-26

- Inspect material saved worksheet print-layout controls directly from raw
  transitional and strict SpreadsheetML before ordinary readers normalize them:
  `_xlnm.Print_Area` / `_xlnm.Print_Titles` definitions, print options,
  margins, page setup/fit-to-page, headers and footers, and manual row/column
  page breaks. Profiles, Markdown, JSON, and SARIF expose structural counts
  only; print ranges, header/footer text, page values, printer-setting
  references, and raw XML remain private.
- Emit `FF056` for a material worksheet print-layout control change and add the
  fail-closed `no_worksheet_print_layout_changes` policy rule (`FFP056`). This
  closes the review gap where an unchanged workbook can print a smaller,
  reordered, reframed, or differently labelled report.
- Normalize omitted/default, Boolean, integer, and decimal spellings; paired
  print-gridline flags; inactive first/even header-footer content; disabled
  first-page numbers; the fit-to-page versus percentage-scale selection; and
  automatic-break display noise. Missing, duplicate, malformed, or unsupported
  metadata is an explicit coverage warning. FormulaFence compares stored
  declarations only: it does not render or preview Excel, calculate page
  geometry/counts or automatic pagination, resolve printer/client defaults or
  `devMode` settings, or cover custom/legacy sheet-view and extension print
  controls.

## 0.53.0 — 2026-07-26

- Inspect material raw transitional and strict SpreadsheetML worksheet-display
  controls before ordinary readers normalize them: hidden zeroes, formula display,
  gridlines/custom gridline colour, row/column headers, outline symbols, rulers,
  page whitespace/margins, right-to-left layout, page-oriented view modes, and
  split/frozen panes. Profiles, Markdown, JSON, and SARIF expose
  structural counts only; sheet names, targets, and raw view XML remain private.
- Emit `FF055` for a material worksheet-display control change and add
  the fail-closed `no_worksheet_display_control_changes` policy rule
  (`FFP055`). This closes the review gap where unchanged values can
  appear blank, controls can be obscured, or the saved workbook surface can be
  materially reframed without an ordinary cell diff.
- Normalize omitted/default controls, Boolean and active custom-gridline-colour
  spellings, finite pane-split decimals, and ordinary selection/top-left/zoom
  navigation churn. Malformed or
  unsupported display metadata produces a visible coverage warning rather than a
  silent omission. FormulaFence compares stored declarations only: it does not
  render Excel, resolve an effective palette colour, calculate viewport geometry,
  decide final visibility, inspect print settings, or compose view controls with
  styles, objects, or client state.

## 0.52.0 — 2026-07-26

- Inspect effective raw SpreadsheetML cell-alignment controls before ordinary
  readers can flatten their style inheritance: horizontal/vertical placement,
  text rotation, wrapping, shrinking, indentation, relative indentation,
  justification, and reading order across default, direct-cell, row, and
  column-style assignments. Profiles, Markdown, JSON, and SARIF expose
  structural counts only; alignment values, style IDs, and target locations
  remain private.
- Emit `FF054` for a material effective cell-alignment control change and add
  the fail-closed `no_cell_alignment_changes` policy rule (`FFP054`). This
  closes the review gap where an unchanged value, warning, or classification can
  be moved, rotated, wrapped, shrunk, or indented without a normal cell diff.
- Normalize equivalent explicit defaults, Boolean and integer spellings,
  semantically inert `mergeCell` compatibility material, base-XF inheritance,
  `applyAlignment` semantics, and effective column-range splitting.
  Missing, duplicate, malformed, or unsupported alignment metadata produces a
  visible coverage warning rather than a silent omission. FormulaFence compares
  stored declarations only: it does not compute widths/heights, merge layout,
  overflow, final text visibility, font/fill/conditional-format composition,
  or Excel client rendering.

## 0.51.0 — 2026-07-26

- Inspect the raw workbook-level DrawingML Theme before ordinary workbook
  readers can reduce it to local style references: workbook-to-Theme bindings,
  transitional and strict Theme XML colour/font/format schemes, direct
  Theme-image relationships, and bounded direct image payloads. Profiles,
  Markdown, JSON, and SARIF expose aggregate counts only; Theme XML, scheme
  names, colour values, font names, image bytes, relationship IDs, and targets
  remain private.
- Emit `FF053` for a material stored workbook-Theme control change and
  add the fail-closed `no_workbook_theme_changes` policy rule
  (`FFP053`). This closes the review gap where a colour, font,
  effect, or direct Theme-image control can change a themed cell, chart, or
  drawing appearance while ordinary cells and local style references stay
  fixed.
- Normalize writer-selected Theme relationship IDs/order and equivalent
  internal target spelling. Missing, duplicate, malformed, unsafe, unbound,
  unreadable, oversized, or over-budget metadata emits a visible coverage
  warning; reads are bounded to 16 MiB per part, 64 MiB per workbook, and 512
  parts. FormulaFence does not resolve effective styles, render a workbook,
  calculate contrast, decode an image, fetch a target, calculate formulas, or
  infer Excel client behavior.

## 0.50.0 — 2026-07-26

- Inspect raw custom workbook data stores before ordinary workbook readers can
  omit them: generic Custom XML data and property/schema declarations,
  package/item relationships, workbook-bound Custom Data Properties and opaque
  binary Custom Data payloads, and custom document properties. Power Query
  `DataMashup` Custom XML remains exclusively under the existing Power Query
  controls. Profiles, Markdown, JSON, and SARIF expose aggregate counts only;
  custom XML, property names and values, storage IDs, binary payloads,
  relationship IDs, and targets remain private.
- Emit `FF052` for a material persisted custom-data-store change and add the
  fail-closed `no_custom_data_store_changes` policy rule (`FFP052`). This
  closes the review gap where add-in state, opaque binary data, or custom
  document properties can change while ordinary cells, formulas, and Power
  Query controls remain fixed.
- Normalize writer-selected relationship IDs/order and document-property
  `pid` values. Custom XML `itemID` and Custom Data `id` storage identities are
  compared privately because add-ins can bind state to them. Missing, duplicate,
  malformed, unsafe, unbound, unreadable, oversized, or over-budget metadata
  emits a visible coverage warning; reads are bounded to 16 MiB per part, 64
  MiB per workbook, and 512 parts. FormulaFence does not execute an add-in,
  resolve a property, follow or fetch a target, interpret a binary payload,
  calculate formulas, or infer Excel client behavior.

## 0.49.0 — 2026-07-24

- Inspect raw Excel rich-data controls before ordinary workbook readers can omit
  or normalize them: Rich Value Data, structures, types, arrays, supporting
  property bags/structures, styles, web images, rich-value relationships,
  workbook/package relationships, and `XLRICHVALUE` metadata/cell bindings.
  Profiles, Markdown, JSON, and SARIF expose aggregate counts only; entity
  values, provider data, field names, identifiers, URLs, image references,
  relationship IDs, and bound-cell locations remain private.
- Emit `FF051` for a material rich-data control change and add the fail-closed
  `no_rich_data_changes` policy rule (`FFP051`). This closes the review gap
  where provider-linked entities, attached data, or external image
  associations can change while ordinary cell values and formulas stay fixed.
- Normalize writer-selected relationship IDs/order and equivalent internal
  target spelling. Missing, duplicate, malformed, unsafe, unreadable,
  oversized, or over-budget metadata emits a visible coverage warning; reads
  are bounded to 16 MiB per XML part, 64 MiB per workbook, and 512 parts.
  FormulaFence does not contact providers, refresh data, calculate formulas,
  fetch endpoints, validate target content, or infer Excel client behavior.

## 0.48.0 — 2026-07-24

- Inspect raw OPC package-signature controls before ordinary workbook readers
  can omit or normalize them: package-root signature origins, origin-to-XML
  signature relationships, XMLDSIG envelope/reference material, embedded
  certificate values, certificate-part relationships/payloads, and conventional
  VBA project signature payloads/relationships
  (`vbaProjectSignature.bin`, Agile, and V3). Profiles, Markdown, JSON, and
  SARIF expose aggregate counts only; signature XML, reference URIs,
  certificate identities/contents, binary payloads, relationship IDs, and
  relationship targets remain private.
- Emit `FF050` for material package- or VBA-signature envelope changes and add
  the fail-closed `no_digital_signature_changes` policy rule (`FFP050`).
  This closes the review gap where provenance/integrity-assurance metadata can
  be added, removed, or altered while ordinary cell values, formulas, and
  `xl/vbaProject.bin` bytes remain unchanged.
- Normalize writer-selected relationship IDs/order, equivalent internal target
  spelling, and XMLDSIG base64 whitespace. Missing, duplicate, malformed,
  unsafe, unbound, unreadable, oversized, or over-budget metadata emits a
  visible coverage warning; reads are bounded to 16 MiB per part, 64 MiB per
  workbook, and 512 parts. FormulaFence inventories envelopes only: it does
  not validate cryptography, signed-reference coverage, certificate identity or
  trust, expiry, revocation, timestamps, signed contents, or VBA-code validity.

## 0.47.0 — 2026-07-24

- Inspect raw SpreadsheetML XML Maps, XML-table column properties, and
  single-cell XML table parts before ordinary workbook readers can discard or
  normalize the mapping surface. FormulaFence privately compares embedded
  schemas, map and data-binding refresh/export behavior, mapped XPath/table/
  cell declarations, and related workbook/worksheet relationship targets while
  profiles, Markdown, JSON, and SARIF expose aggregate counts only—never
  schemas, map names, XPath expressions, table identities, target cells,
  connection identities, or relationship targets.
- Emit FF049 for a material XML-mapped workbook control change and add the
  fail-closed `no_xml_mapping_changes` policy rule (FFP049). This closes the
  review gap where a data import/export template can be redirected or have its
  refresh behavior changed without an ordinary worksheet-cell diff.
- Normalize equivalent Boolean and unsigned-integer spelling,
  writer-selected relationship IDs/order, and equivalent internal target
  spelling. Missing, duplicate, malformed, unsafe, unbound, unreadable,
  oversized, or over-budget metadata produces a visible coverage warning;
  bounded raw reads use 16 MiB per part, 64 MiB per workbook, and 512 parts.
  FormulaFence does not import/export XML, validate data against schemas, open
  map bindings, fetch data, calculate refresh results, or infer client
  behavior.

## 0.46.0 — 2026-07-24

- Inspect raw Office 2010 `x14:sparklineGroups` worksheet extensions before
  ordinary workbook readers discard them. FormulaFence privately compares
  sparkline source/date-axis formulas, destination cells, group membership,
  type/axis/display/marker controls, line weight, and colour definitions while
  profiles, Markdown, JSON, and SARIF expose aggregate counts only—never
  source formulas, locations, control values, or colour definitions.
- Emit `FF048` for a material worksheet-sparkline control change and add the
  fail-closed `no_worksheet_sparkline_changes` policy rule (`FFP048`). This
  closes the review gap where a compact trend can be retargeted or restyled
  without changing ordinary worksheet values.
- Normalize equivalent direct local-range, Boolean/numeric, colour-case, and
  declaration-order spelling. Missing, duplicate, malformed, unreadable,
  oversized, or over-budget metadata produces a visible coverage warning. A
  Sparkline Group-removed temporary reader copy is made only after raw
  inspection, so reader loss cannot suppress evidence; FormulaFence does not
  calculate, render, resolve, or fetch sparkline sources or assess visual
  accessibility.

## 0.45.1 — 2026-07-24

- Stabilize the cross-version redaction test for an unknown custom number
  format by using a fixture-specific sentinel instead of a generic numeric
  substring. This does not change FormulaFence's workbook inspection or report
  surface.

## 0.45.0 — 2026-07-24

- Inspect raw standard SpreadsheetML and Office 2016 revision worksheet-cell
  hyperlink declarations before ordinary workbook readers can normalize their
  target bindings. FormulaFence privately compares cell/range binding, external
  and internal relationship semantics, location, display override, and
  ScreenTip material while profiles, Markdown, JSON, and SARIF expose aggregate
  counts only—never targets, cell references, locations, display strings,
  ScreenTips, relationship IDs, or revision UIDs.
- Emit `FF047` for a material worksheet-cell hyperlink control change and add
  the fail-closed `no_cell_hyperlink_changes` policy rule (`FFP047`). This closes
  the review gap where a familiar cell label can redirect a reviewer to a
  different URL, file, or in-workbook destination without changing the ordinary
  cell value.
- Normalize writer-chosen relationship IDs/revision UIDs, relationship ordering,
  and equivalent internal target spelling. Missing, duplicate, malformed,
  unsafe, unbound, unreadable, oversized, or over-budget metadata produces a
  visible coverage warning. The ordinary workbook reader receives a
  hyperlink-removed temporary copy only after raw inspection, so malformed
  package metadata cannot suppress evidence; FormulaFence does not render,
  resolve, fetch, follow, reputation-check, or execute a link, inspect linked
  content, or interpret `HYPERLINK()` formulas beyond the ordinary formula
  diff.

## 0.44.0 — 2026-07-24

- Inspect raw legacy Excel Note comments parts and their worksheet-bound VML
  Note shapes before workbook readers can omit author/text or display
  declarations. FormulaFence privately compares note text and rich-text
  presentation, author association, cell binding, comment properties, Note
  visibility/layout, and relationship semantics while profiles, Markdown, JSON,
  and SARIF expose aggregate counts only—never note text, authors, locations,
  VML, targets, relationship IDs, or GUIDs.
- Recognize the documented legacy Note placeholder that Excel can retain beside
  a modern threaded comment. A consistent placeholder GUID/author rekey stays
  quiet, while a changed placeholder reconciliation declaration remains guarded
  independently of the modern thread.
- Emit FF046 for a material legacy Note or threaded-placeholder change and add
  the fail-closed no_legacy_comment_changes policy rule (FFP046). This closes
  the review gap where an instruction, review context, or its visibility/layout
  can change outside ordinary worksheet cells.
- Bound raw comments/VML XML reads to 16 MiB per part, 64 MiB per workbook, and
  512 parts. Missing, duplicate, malformed, unsafe, external, unbound, or
  oversized metadata produces a visible coverage warning. The ordinary workbook
  reader receives a temporary Note-quarantined copy after raw inspection, so
  parser tolerance cannot suppress Note evidence; FormulaFence does not render
  Notes, resolve authors, fetch targets, execute linked content, or infer
  client/cloud state.

## 0.43.0 — 2026-07-24

- Inspect raw modern Excel threaded-comment and person package parts before
  workbook readers can omit their review annotations. FormulaFence compares the
  private comment/reply graph, stored text, cell binding, timestamp, resolution
  state, mention range/person binding, extension material, and author/mentioned
  person definitions while profiles, Markdown, JSON, and SARIF expose only
  aggregate counts—never comment text, locations, timestamps, names, user IDs,
  provider IDs, relationship IDs, or GUIDs.
- Emit `FF045` for a material threaded-comment control change and add the
  fail-closed `no_threaded_comment_changes` policy rule (`FFP045`). This closes
  the review gap where an assumption, instruction, or approval reply can change
  without changing any ordinary worksheet cell.
- Rebuild comment trees and person/mention links from their private identities
  so consistent writer-chosen comment, parent, person, mention, and package
  relationship-ID rewrites stay quiet. Missing, duplicate, unsafe, unbound,
  malformed, unreadable, oversized, or over-budget metadata becomes a visible
  coverage warning rather than a silent omission. FormulaFence compares stored
  package declarations only: it does not render comments, validate mention text
  offsets, send notifications, resolve accounts, fetch targets, or inspect
  legacy note/placeholder content.

## 0.42.0 — 2026-07-24

- Inspect raw non-chart Worksheet DrawingML regular shapes (`xdr:sp`) and group
  shapes (`xdr:grpSp`) before workbook readers can discard their text-box
  presentation. FormulaFence compares anchor/layout declarations, text and
  visual XML, group nesting, macro assignments, text links, and click/hover
  relationship semantics privately; profiles, Markdown, JSON, and SARIF expose
  structural counts only, never text, formatting, anchors, formulas, macro
  names, relationship IDs, or targets.
- Emit `FF044` for a material Worksheet DrawingML shape-control change and add
  the fail-closed `no_worksheet_drawing_shape_changes` policy rule (`FFP044`).
  This catches a text-box warning whose stored cell values and concatenated
  text remain unchanged while its presentation becomes less visible.
- Normalize writer-chosen non-visual shape IDs, relationship-ID rewrites, colour
  case, and relationship target spelling while retaining meaningful z-order and
  shape/group declarations. Missing, malformed, unsupported, oversized, or
  over-budget shape metadata becomes a visible parser-coverage warning rather
  than a silent omission. FormulaFence compares stored declarations only: it
  does not render DrawingML, resolve themes, evaluate text links, execute macro
  assignments, fetch targets, inspect arbitrary media, or claim coverage for
  pictures, connectors, graphic frames, SmartArt, or other non-`xdr:sp`
  drawing objects.

## 0.41.0 — 2026-07-24

- Inspect raw shared-string and inline-string character presentation that normal
  workbook readers reduce to concatenated text. FormulaFence compares rich
  `<r>/<rPr>` property sequences, styled character boundaries, and phonetic
  presentation material privately; profiles, Markdown, JSON, and SARIF expose
  structural counts only, never text, colours, fonts, indexes, or locations.
- Emit `FF043` for a material rich-text run control change and add the
  fail-closed `no_rich_text_run_changes` policy rule (`FFP043`).
  A formatting-only change such as making a warning phrase white is detected
  even when the normal cell value remains unchanged; an ordinary text-only edit
  within the same run-property sequence remains a normal semantic cell diff.
- Normalize rich-run property ordering, colour case, explicit false Boolean
  properties, and equivalent shared-versus-inline storage. Malformed,
  unsupported, missing, or unreadable rich-text metadata becomes a visible
  parser-coverage warning rather than a silent omission. FormulaFence compares
  stored declarations only: it does not render cells, resolve theme colours,
  calculate contrast, decide visibility, or guarantee Excel rendering.

## 0.40.0 — 2026-07-24

- Inspect raw SpreadsheetML formula-result caches alongside formula text. Cache
  values, error text, per-cell digests, and formula-cell locations remain only
  in private comparison entries; profiles, Markdown, JSON, and SARIF expose
  aggregate formula/cached/missing/result-type/malformed counts only.
- Emit `FF042` when a stored formula result changes without a changed formula at
  that cell or an ordinary changed cell that reaches it through the static
  dependency graph. Add the fail-closed
  `no_formula_cached_result_changes` policy rule (`FFP042`).
- Normalize equivalent finite numeric and Boolean result spellings, keep absent
  or blank caches visible as missing rather than inventing a result, and make
  malformed or unsupported cache metadata an explicit parser-coverage warning.
  FormulaFence does not calculate or validate results, distinguish a stale
  result from a tampered one, or model volatile, dynamic, external, or
  calculation-engine dependencies; a legitimate recalculation without a static
  visible precedent can therefore require review.

## 0.39.0 — 2026-07-24

- Extend the existing filter, sort, and row/column-visibility boundary to the
  documented zero-sized states that hide worksheet content without changing a
  cell: direct row `ht="0"`, column `width="0"`, and worksheet-default
  `defaultRowHeight="0"` / `defaultColWidth="0"` controls. Retain dimensions,
  row/column targets, and raw declarations only in private signatures; profiles,
  `FF036`, and SARIF expose structural counts only.
- Resolve worksheet-default zero dimensions before direct row and layered column
  declarations, so a later positive height or width is compared as an effective
  override while an equivalent inherited zero stays quiet. Positive ordinary
  resizes remain outside this narrow concealment boundary.
- Extend `FF036` / `no_filter_visibility_changes` (`FFP036`) without adding a
  new policy switch. Invalid, negative, non-finite, or application-out-of-range
  dimensions become explicit parser-coverage warnings rather than silent
  omissions. FormulaFence does not render widths/heights, infer overflow, or
  track arbitrary nonzero layout changes.

## 0.38.0 — 2026-07-24

- Inspect raw workbook cell-fill controls that can change what a reviewer sees
  without changing a stored cell value or formula: `<fills>` definitions,
  including patterned and gradient fills, base `<cellStyleXfs>`, effective
  `<cellXfs>`, direct cell `s`, `customFormat=1` row styles, and worksheet
  `<cols>/<col style>` defaults. Retain fill colours, pattern/gradient material,
  style IDs, and targets only in private signatures; profiles, `FF041`, and
  SARIF expose structural counts only.
- Resolve fill-ID remapping, base-XF inheritance, `applyFill`, valid
  pattern-colour child ordering, semantically inert no-fill/solid-background
  declarations, and equivalent column-range splitting. Record a column fill only
  as an OOXML default for unallocated/new cells rather than claiming to
  re-render allocated cells.
- Emit `FF041` for a material cell-fill-control change and add the fail-closed
  `no_cell_fill_changes` policy rule (`FFP041`). Invalid or missing fill/style
  references, malformed definitions, invalid targets, and bounded parser
  failures become explicit coverage warnings rather than silent omissions.
  FormulaFence does not resolve theme colours or rendering, calculate
  text/background contrast, apply conditional-format differential styles or
  table styles, or model borders, alignment, rich text, width/overflow, or
  arbitrary visual formatting.

## 0.37.0 — 2026-07-24

- Inspect raw workbook cell-font controls that can change what a reviewer sees
  without changing a stored cell value or formula: `<fonts>` definitions, base
  `<cellStyleXfs>`, effective `<cellXfs>`, direct cell `s`, `customFormat=1`
  row styles, and worksheet `<cols>/<col style>` defaults. Retain font names,
  colours, effects, style IDs, and targets only in private signatures;
  profiles, `FF040`, and SARIF expose structural counts only.
- Resolve font-ID remapping, base-XF inheritance, `applyFont`, equivalent font
  child ordering, common Boolean spellings, and equivalent column-range
  splitting. Record a column font only as an OOXML default for unallocated/new
  cells rather than claiming to re-render allocated cells.
- Emit `FF040` for a material cell-font-control change and add the fail-closed
  `no_cell_font_changes` policy rule (`FFP040`). Invalid or missing font/style
  references, malformed definitions, invalid targets, and bounded parser
  failures become explicit coverage warnings rather than silent omissions.
  FormulaFence does not resolve theme colours or rendering, model fills,
  borders, alignment, rich-text runs, table styles, width/overflow, or arbitrary
  visual formatting.

## 0.36.0 — 2026-07-24

- Inspect raw workbook number-format controls that can change what a reviewer
  sees without changing a stored cell value or formula: custom `<numFmt>`
  definitions, base `<cellStyleXfs>`, effective `<cellXfs>`, direct cell `s`,
  `customFormat=1` row styles, and worksheet `<cols>/<col style>` defaults.
  Retain format codes, style IDs, and targets only in private signatures;
  profiles, `FF039`, and SARIF expose structural counts only.
- Resolve custom-format ID remapping, base-XF inheritance, and
  `applyNumberFormat`; normalize equivalent custom-ID allocation, Boolean
  spelling, and effective column-range splitting. Record column styles as
  defaults for unallocated/new cells rather than claiming to re-render allocated
  cells.
- Emit `FF039` for a material number-format-control change and add the
  fail-closed `no_number_format_changes` policy rule (`FFP039`). Invalid or
  missing format/style references, conflicting custom definitions, invalid
  targets, and bounded parser failures become explicit coverage warnings rather
  than silent omissions. FormulaFence does not render locale-specific output,
  validate format syntax, calculate values, model width/overflow, or track
  arbitrary non-number-format visual styling.

## 0.35.0 — 2026-07-25

- Inspect raw worksheet `<cols>/<col>` visibility declarations for hidden,
  outlined, and collapsed columns without relying on a workbook reader that can
  flatten or lose compressed column ranges. Retain column positions and effective
  state only in private signatures; profiles, `FF036`, and SARIF expose safe
  counts only.
- Normalize Boolean/default and unsigned-integer spellings, equivalent range
  segmentation, and layered column declarations by applying later *present*
  visibility attributes in OOXML file order. A later width/style-only record
  does not erase an existing visibility state.
- Extend `FF036` / `no_filter_visibility_changes` (`FFP036`) to block effective
  hidden, outlined, or collapsed column changes alongside existing filter, sort,
  and row-visibility controls. Make malformed column bounds, attributes, child
  markup, and bounded-update exhaustion visible coverage warnings rather than
  silently omitting the affected controls. FormulaFence does not render
  outlines, apply filters, calculate results, track widths/styles, or interpret
  outline-display settings.

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
