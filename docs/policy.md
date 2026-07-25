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
  no_xlm_macro_sheet_changes: true
  no_ribbon_customization_changes: true
  no_office_web_addin_changes: true
  no_chart_definition_changes: true
  no_pivot_table_definition_changes: true
  no_slicer_timeline_cache_changes: true
  no_power_pivot_data_model_changes: true
  no_what_if_data_table_changes: true
  no_scenario_manager_changes: true
  no_filter_visibility_changes: true
  no_ignored_error_changes: true
  no_named_sheet_view_changes: true
  no_number_format_changes: true
  no_cell_font_changes: true
  no_cell_fill_changes: true
  no_workbook_theme_changes: true
  no_cell_alignment_changes: true
  no_formula_cached_result_changes: true
  no_rich_text_run_changes: true
  no_cell_hyperlink_changes: true
  no_worksheet_sparkline_changes: true
  no_xml_mapping_changes: true
  no_digital_signature_changes: true
  no_rich_data_changes: true
  no_custom_data_store_changes: true
  no_legacy_comment_changes: true
  no_threaded_comment_changes: true
  no_worksheet_drawing_shape_changes: true
  no_worksheet_embedded_control_changes: true
  no_new_parser_warnings: true
  no_new_unresolved_references: true
  no_new_dynamic_references: true
  no_new_spill_references: true
  no_new_dynamic_array_output_references: true
  no_new_implicit_intersections: true
  no_array_formula_semantics_changes: true
  no_new_tokenization_failures: true
  no_table_definition_changes: true
  no_data_validation_changes: true
  no_conditional_formatting_changes: true
  no_protection_changes: true
  no_external_data_connection_changes: true
  no_external_link_package_changes: true
  no_power_query_changes: true
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
| `no_xlm_macro_sheet_changes` | boolean | An Excel 4.0 / XLM macro-sheet declaration, program XML, related-part relationship, or direct internal related-part payload changes. |
| `no_ribbon_customization_changes` | boolean | An Office RibbonX custom-UI package declaration, control/callback XML, or direct relationship changes. |
| `no_office_web_addin_changes` | boolean | An Office Web Add-in task-pane workbook binding, task-pane configuration, web-extension definition, or direct relationship changes. |
| `no_chart_definition_changes` | boolean | A DrawingML chart binding, chart definition, cached series data, chart-overlay shape, direct relationship, or bounded direct related payload changes. |
| `no_pivot_table_definition_changes` | boolean | A PivotTable binding/layout, cache schema, shared item, cache-record relationship, or bounded cached-record payload changes. Source and refresh controls remain under `no_external_data_connection_changes`. |
| `no_slicer_timeline_cache_changes` | boolean | A Slicer or Timeline workbook binding, cached filter state, source binding, filtered-PivotTable binding, or direct cache-part relationship changes. |
| `no_power_pivot_data_model_changes` | boolean | An embedded Power Pivot/Data Model workbook binding, `x15:dataModel` declaration, direct model-part relationship, or bounded raw model payload changes. |
| `no_what_if_data_table_changes` | boolean | An Excel What-If Data Table master changes its output range, one-/two-variable mode, orientation, input references, deleted-input state, recalculation request, or supported raw formula metadata. This is unrelated to an Excel table definition. |
| `no_scenario_manager_changes` | boolean | An Excel Scenario Manager worksheet changes its selected/shown scenario state, result-summary references, scenario definition, protection flags, comments/users, stored input values/references, deleted/undone state, or input display number formats. Scenario names, comments, users, values, and references are compared privately. |
| `no_filter_visibility_changes` | boolean | A worksheet/Table AutoFilter, stored filter criterion, filter sort state, explicitly hidden/zero-height/outlined/collapsed row, hidden/zero-width/outlined/collapsed column, or hidden/zero-dimension-by-default worksheet setting changes. Criteria, selected values, sort keys/lists, raw dimensions, table names, and row/column ranges are compared privately. |
| `no_ignored_error_changes` | boolean | A standard or Office 2010 extension ignored-error declaration changes a per-range suppression of Excel evaluation, formula-consistency, range-omission, unlocked-formula, empty-reference, list-validation, calculated-column, text-number, or two-digit-year warnings. Targets and exact suppressions are compared privately. |
| `no_named_sheet_view_changes` | boolean | A relationship-backed Excel Named Sheet View, alternate AutoFilter criterion, sort rule, or reconciled base-filter binding changes. View names, IDs, criteria, ranges, table bindings, and sort keys are compared privately. |
| `no_number_format_changes` | boolean | An effective default, direct-cell, row, or column number-format control changes. Format codes, style IDs, and cell/row/column targets are compared privately. |
| `no_cell_font_changes` | boolean | An effective default, direct-cell, row, or column font control changes. Font names, colours, effects, style IDs, and cell/row/column targets are compared privately. |
| `no_cell_fill_changes` | boolean | An effective default, direct-cell, row, or column fill control changes. Fill colours, pattern/gradient definitions, style IDs, and cell/row/column targets are compared privately. |
| `no_workbook_theme_changes` | boolean | A workbook Theme binding, stored colour/font/format scheme, direct Theme-image relationship, or direct Theme-image payload changes. Theme XML, scheme names, colours, font names, image bytes, relationship IDs, and targets are compared privately. |
| `no_cell_alignment_changes` | boolean | An effective default, direct-cell, row, or column alignment control changes. Alignment values, style IDs, and cell/row/column targets are compared privately. |
| `no_formula_cached_result_changes` | boolean | A saved formula result changes without a changed formula at that cell or a statically visible ordinary-cell precedent change. Result values, error text, result digests, and formula-cell locations are compared privately. |
| `no_rich_text_run_changes` | boolean | A shared or inline rich-text run property, styled-character boundary, or phonetic hint/property changes. Character text, format details, shared-string indexes, and locations are compared privately. |
| `no_cell_hyperlink_changes` | boolean | A standard or Office 2016 revision worksheet-cell hyperlink binding, location, display override, ScreenTip, or selected relationship target/type/mode changes. Targets, cell references, locations, display strings, ScreenTips, relationship IDs, and revision UIDs are compared privately. |
| `no_worksheet_sparkline_changes` | boolean | An Office 2010 worksheet sparkline source or date-axis formula, destination cell, group membership, type/axis/display/marker control, line weight, or colour definition changes. Source formulas, destination cells, control values, and colours are compared privately. |
| `no_xml_mapping_changes` | boolean | An XML Map schema, mapping/refresh behavior, table-column or single-cell binding, or map-related workbook/worksheet relationship changes. Schemas, map names, XPath expressions, table identities, target cells, connection identities, and relationship targets are compared privately. |
| `no_digital_signature_changes` | boolean | An OPC package signature origin/XML-signature/certificate relationship or payload, XMLDSIG envelope/reference, or VBA project signature payload/relationship changes. Signature XML, reference URIs, certificate identities and contents, binary payloads, relationship IDs, and targets are compared privately; FormulaFence inventories envelopes but does not validate cryptography or trust. |
| `no_rich_data_changes` | boolean | An Excel rich-value data/structure/type/array/property-bag/style declaration, provider-associated value, web-image/rich-value relationship, or `XLRICHVALUE` metadata/cell binding changes. Entity values, provider data, field names, identifiers, URLs, image references, relationship IDs, and bound-cell locations are compared privately. |
| `no_custom_data_store_changes` | boolean | Generic Custom XML data/property/schema material or relationships, workbook-bound Custom Data Properties or opaque binary Custom Data payloads, or custom document properties change. Custom XML, schema URIs, property names/values, storage IDs, binary payloads, relationship IDs, and targets are compared privately. Power Query `DataMashup` remains under `no_power_query_changes`. |
| `no_legacy_comment_changes` | boolean | A legacy Excel Note/comments binding, author association, Note text/rich-text/property declaration, threaded-comment placeholder reconciliation declaration, Note VML visibility/layout, or related relationship changes. Note text, authors, locations, VML, targets, IDs, and GUIDs are compared privately. |
| `no_threaded_comment_changes` | boolean | A modern Excel threaded-comment/person package binding, comment/reply graph, text, stored cell/timestamp/resolution declaration, mention range/person association, extension material, or person definition changes. Comment bodies, locations, timestamps, parent links, names, user IDs, provider IDs, relationship IDs, and GUIDs are compared privately. |
| `no_worksheet_drawing_shape_changes` | boolean | A non-chart Worksheet DrawingML regular `xdr:sp` or nested `xdr:grpSp` anchor/layout, text/presentation declaration, macro/text link, or referenced click/hover relationship changes. Shape text, formatting, formulas, anchors, IDs, and targets are compared privately. |
| `no_worksheet_embedded_control_changes` | boolean | A modern worksheet or legacy VML control/OLE binding, definition, direct relationship, or bounded direct payload changes. |
| `no_new_parser_warnings` | boolean | The candidate introduces an unsupported-workbook coverage warning. |
| `no_new_unresolved_references` | boolean | A formula adds a name, named-LAMBDA call, table reference, or other token that cannot be resolved statically. |
| `no_new_dynamic_references` | boolean | A formula adds a dynamic reference function such as `INDIRECT` or `OFFSET`. |
| `no_new_spill_references` | boolean | A formula adds a dynamic-array spill reference; FormulaFence traces its anchor but not its variable extent or blockers. |
| `no_new_dynamic_array_output_references` | boolean | A formula newly intersects a non-anchor member of an OOXML-observed dynamic-array output range. |
| `no_new_implicit_intersections` | boolean | A formula adds explicit `@` / `SINGLE()` implicit intersection, which can change which value a range or array contributes. |
| `no_array_formula_semantics_changes` | boolean | A legacy-CSE or dynamic-array formula is added, removed, or changes mode, or a legacy CSE formula's fixed output range changes. |
| `no_new_tokenization_failures` | boolean | A formula is newly introduced that the underlying formula tokenizer cannot inspect. |
| `no_table_definition_changes` | boolean | An Excel table is added, removed, moved, renamed, or has its columns/header/total-row configuration changed. |
| `no_data_validation_changes` | boolean | A worksheet data-validation control changes, including its target ranges, criteria, blank/dropdown behavior, prompts, error alert, or global prompt-disable setting. |
| `no_conditional_formatting_changes` | boolean | A worksheet conditional-formatting control changes, including its precedence, target ranges, criteria, flags, visual style, or retained OOXML extension fragment. |
| `no_protection_changes` | boolean | A workbook, worksheet, dialog-sheet, chart-sheet, protected-range, or direct cell/row/column protection control changes. |
| `no_external_data_connection_changes` | boolean | A workbook-wide external-data refresh flag, connection, linked query-table refresh control, or pivot-cache source/refresh control changes. |
| `no_external_link_package_changes` | boolean | An external-workbook, DDE, or OLE `externalLink` package definition, source binding, cached material, item behavior, or retained extension fragment changes. |
| `no_power_query_changes` | boolean | A Power Query Data Mashup formula, package definition, stable query metadata, or formula-firewall permission control changes. |
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

FormulaFence separately inventories worksheet data-validation controls. It
tracks their compact target ranges, validation type and operator, the two
criteria expressions, blank/dropdown behavior, input prompts, error alerts,
IME mode, and the worksheet-level `disablePrompts` setting. It treats omitted
OOXML defaults as their effective values (`none`, `between`, `stop`, and
`noControl`), normalizes an optional leading `=` in criterion expressions, and
joins identical rules whose target groups are serialized separately, avoiding a
writer-formatting-only control diff. Profiles omit the criterion and
message text; local reports retain the full before/after evidence. Any change
emits `FF020`; enable `no_data_validation_changes` to make it `FFP020` in CI.
FormulaFence does not evaluate a validation formula or infer whether a future
entry will pass it.

FormulaFence separately inventories worksheet conditional formatting from raw
OOXML so a reader library cannot erase an extension before it is compared. It
tracks each compact `sqref` target group, its worksheet-global precedence,
rule type/operator/criteria, `Stop If True`, average/rank/time flags,
differential style, color scale, data bar, icon set, and rule-level extension
fragments. It resolves `dxfId` to the style's actual OOXML, normalizes schema
boolean defaults, a leading `=` in criteria, non-semantic priority-number gaps,
and GUIDs that only link extension fragments. Worksheet-level conditional
formatting extensions remain opaque but are retained and diffed. Profiles omit
criteria, text rules, and raw style/extension XML; local reports retain full
evidence. Any change emits `FF021`; enable
`no_conditional_formatting_changes` to make it `FFP021` in CI. FormulaFence
does not calculate a rule, resolve its relative references for every target, or
predict a cell's final rendered format.

FormulaFence separately inventories **operational protection controls** from
raw OOXML because an editable input, an exposed formula, or a removed workbook
structure lock can matter even when no formula text changes. It compares
workbook `lockStructure`, `lockWindows`, and `lockRevision`; worksheet and
dialog-sheet action locks using their effective OOXML defaults; chart-sheet
`content`/`objects` locks; protected-range target areas; and direct serialized
cell, row, and column `locked`/`hidden` style assignments on active protected
sheets. It retains protected-range target spans compactly and does not expand a
row, column, or large styled range into cell records.

Profiles and local reports omit legacy password verifiers, modern hash/salt
values, protected-range names, and security-descriptor contents. Private
fingerprints still make a changed verifier, range name, security descriptor, or
unmodelled protection fragment visible as a change. Any protection-control
change emits `FF022`; enable `no_protection_changes` to make it `FFP022` in CI.
This is not encryption, authentication, or an access-control decision:
workbook/worksheet protection is an operational Excel control, and
FormulaFence does not determine whether an actor can edit a workbook or fully
recreate Excel's style-precedence rendering. File encryption and rights
management remain outside scope.

FormulaFence separately inventories **external-data refresh controls** from raw
OOXML because a workbook can change inputs by refreshing a connection without a
normal formula edit. It compares workbook `updateLinks`, `allowRefreshQuery`,
`refreshAllConnections`, and `saveExternalLinkValues`; connection refresh,
cache, credential, source-kind, connection-file, and parameter-refresh flags;
linked query-table refresh and growth behavior; and pivot-cache source and
refresh settings. Omitted schema defaults compare equal to their explicit
spelling. Connection names/descriptions, paths, URLs, connection strings,
commands, parameter values, SSO IDs, cached records, and opaque extension XML
never appear in profiles or reports; private fingerprints still expose material
source or identity changes. Any such control change emits `FF023`; enable
`no_external_data_connection_changes` to make it `FFP023` in CI. FormulaFence
does not execute a connection, refresh workbook data, determine source trust,
or calculate a PivotTable report.

FormulaFence separately inventories raw `xl/externalLinks/externalLink*.xml`
packages. It recognizes external-workbook, DDE, and OLE definitions; privately
binds each workbook declaration to its package part; and compares link targets,
workbook definitions, caches, DDE/OLE item behavior, and opaque extension
material. The profile exposes only safe counts: workbook targets, sheet and
defined names, DDE services/topics/items, OLE program and item names, and
cached values never enter a profile or diff. Any material package change emits
`FF025`; enable `no_external_link_package_changes` to make it `FFP025` in CI.
FormulaFence does not follow or execute these links, determine source trust, or
infer returned data.

Excel 4.0 / XLM macro sheets are separate from the VBA binary: their executable
commands live in Macro Sheet XML package parts, typically under
`xl/macrosheets/`. FormulaFence reads those parts before the workbook library
can omit their cells. It privately binds documented macro-sheet workbook
relationships to parts, fingerprints complete XML and related package
relationships, and streams direct safe internal relationship targets into
private payload fingerprints. It reports only safe counts for formula cells,
visibility, international-sheet status, related OLE/package parts, and
fingerprinted versus uninspected targets. A material change—including a
payload-only change—emits `FF026`; enable `no_xlm_macro_sheet_changes` to make
it `FFP026` in CI. XLM commands, values, relationship targets, and embedded
payloads never enter a profile or diff. FormulaFence does not execute, emulate,
resolve, or parse any of them; it never follows external targets. Direct
internal payload scanning is bounded to 32 MiB per part, 64 MiB per workbook,
and 256 parts, with an explicit coverage warning once a bound is reached.

Office RibbonX custom UI parts can bind buttons and other controls to workbook
callbacks while sitting outside ordinary worksheet XML and the VBA payload.
FormulaFence inspects documented root-package declarations and recognized
`customUI` roots for the 2006 and Office 2010-era schemas. It privately
fingerprints the complete custom-UI XML and direct package relationships,
normalizing writer-chosen relationship IDs when their semantic target is
unchanged. Profiles expose only part, control, callback-attribute, image, and
external-relationship counts; control IDs, labels, callback names, XML, and
targets never enter a profile or diff. A material change emits `FF027`; enable
`no_ribbon_customization_changes` to make it `FFP027` in CI. FormulaFence does
not execute a RibbonX callback, follow an external relationship, or parse an
image payload. Missing, oversized, malformed, unbound, version-mismatched, or
otherwise unrecognized custom-UI parts remain visible coverage warnings.
Custom-UI XML reads are bounded to 16 MiB per part, 32 MiB per workbook, and
eight parts.

Office Web Add-in task panes can bind a document to an installed add-in and
request `Office.AutoShowTaskpaneWithDocument` while remaining outside ordinary
worksheet XML, the VBA payload, and RibbonX. FormulaFence follows the bounded
chain from the workbook's documented task-pane relationship through
`taskpanes.xml`, its task-pane-to-extension bindings, and direct
`webextension*.xml` definitions. It privately fingerprints task-pane
configuration, add-in references, auto-show properties, bindings, snapshots,
and direct relationship semantics while normalizing writer-chosen relationship
IDs and equivalent internal target spellings. Profiles expose only safe counts:
parts, task panes, visible/locked panes, references, auto-show requests,
bindings, snapshots, and relationships. Add-in IDs, store references, property
values, binding values, XML, snapshot data, and relationship targets never
enter a profile or diff. A material change emits `FF028`; enable
`no_office_web_addin_changes` to make it `FFP028` in CI. FormulaFence does not
install, load, execute, or fetch an add-in or manifest, and it never follows
an external relationship. Missing, oversized, malformed, unbound, or
over-budget parts remain coverage warnings. Task-pane and web-extension XML
reads are bounded to 16 MiB per part, 32 MiB per workbook, and 64 parts.
Worksheet-scoped web-extension markup outside this task-pane chain is not yet
modeled.

DrawingML charts can change a report's series, axis, title, formatting, cached
values, or overlay annotations without changing an ordinary worksheet cell.
FormulaFence follows standard worksheet/chartsheet drawing relationships through
`c:chart` parts and direct `c:userShapes` overlays. It privately compares chart
definition material separately from `numCache`, `strCache`, and
`multiLvlStrCache` material; it also compares overlay XML, relationship
semantics, and bounded direct related payload hashes. Profiles expose safe
structural counts only. Formulas, labels, cached values, formatting, overlay
text, target paths, XML, and payload bytes never enter a profile or diff.
Writer-chosen relationship IDs and equivalent internal target spellings are
normalized. A material change emits `FF030`; enable
`no_chart_definition_changes` to make it `FFP030` in CI. FormulaFence does not
calculate a series, map chart references into downstream cell impact, render a
chart, assess visual output, follow external targets, parse media/package
formats, or interpret `chartEx`. XML reads are bounded to 16 MiB per part, 64
MiB per workbook, and 512 parts; direct payload hashes are bounded to 32 MiB
per part, 64 MiB per workbook, and 512 parts. Missing, malformed, orphaned,
unbound, oversized, or over-budget chart material remains a visible coverage
warning.

PivotTable packages can regroup, filter, aggregate, or present report material
without changing an ordinary formula cell. FormulaFence follows the bounded
workbook-cache and worksheet-PivotTable relationship graph, then privately
compares PivotTable layouts, cache schemas, shared cache items, normalized
relationships, and bounded raw cache-record payloads. Profiles expose only
safe counts for parts, fields, items, cache records, relationships, and
coverage; names, source ranges, item values, formulas, cache values, targets,
XML, and payload bytes never enter a profile or diff. Writer-chosen
relationship IDs, equivalent internal target spellings, and cache-ID
renumbering are normalized. A material change emits `FF031`; enable
`no_pivot_table_definition_changes` to make it `FFP031` in CI.

Source definitions and refresh settings remain under `FF023` /
`no_external_data_connection_changes`, so a refresh-only edit is not reported
as a PivotTable definition change. FormulaFence does not refresh a cache,
calculate or render a PivotTable, infer PivotTable-to-cell impact, fetch an
external target, or interpret OLAP or extension-list semantics. Slicer and
Timeline cache definitions are compared separately.
Missing, malformed, orphaned, unbound, oversized, or over-budget material
remains a visible coverage warning. PivotTable/cache-definition XML reads are
bounded to 16 MiB per part, 64 MiB per workbook, and 512 parts; raw cache-record
hashing is bounded to 32 MiB per part, 64 MiB per workbook, and 512 parts. A
temporary reader copy detaches cache-record relationships before the underlying
workbook library loads cells, so those raw records are not eagerly materialized;
the original workbook remains unchanged.

Slicer and Timeline caches can apply an interactive item or date filter to a
PivotTable, and a Slicer can also filter an Excel table, without changing an
ordinary worksheet cell. FormulaFence follows documented workbook cache
declarations and their explicit relationships to bounded cache XML, then
privately compares cache definitions, item selections, Timeline state/filter
material, PivotTable/table source bindings, filtered-PivotTable bindings, and
unexpected direct cache-part relationships. Profiles expose only safe
structural counts; cache names, source fields, selected values, date ranges,
PivotTable names, relationship targets, and XML remain private. A material
change emits `FF032`; enable `no_slicer_timeline_cache_changes` to make it
`FFP032` in CI. Writer-chosen relationship IDs, equivalent internal target
spellings, coordinated Slicer/Timeline PivotCache extension-ID renumbering, known optional Slicer
defaults, Boolean spellings, and Timeline GUIDs are normalized. FormulaFence
does not apply a filter, calculate or render a PivotTable/table, infer
downstream cell impact, fetch an external target, or model worksheet/drawing
Slicer or Timeline view geometry and styles. Missing, malformed, orphaned,
unbound, externally targeted, oversized, or over-budget material remains a
visible coverage warning. Cache XML reads are bounded to 16 MiB per part, 64
MiB per workbook, and 512 parts.

An embedded Power Pivot/Data Model can carry tables, relationships,
calculations, and stored values outside ordinary worksheet cells. FormulaFence
follows its explicit workbook `powerPivotData` binding and `x15:dataModel`
declaration, then privately fingerprints declaration material and bounded raw
`xl/model/*.data` payloads. A material change emits `FF033`; enable
`no_power_pivot_data_model_changes` to make it `FFP033` in CI. Profiles expose
only counts for model parts, bindings, declarations, tables, relationships,
fingerprinted payloads, and coverage. Table/column/relationship names,
connection details, DAX, stored values, targets, XML, and raw bytes remain
private. Relationship IDs, equivalent internal target spellings, and Data
Model GUIDs are normalized. FormulaFence does not deserialize the Analysis
Services payload, evaluate DAX, refresh a model, calculate/render a report,
infer model-to-cell impact, or fetch an external target. Missing, malformed,
orphaned, unbound, externally targeted, unexpected directly related,
oversized, or over-budget material remains a visible coverage warning. Raw
payload reads are bounded to 512 MiB per part, 512 MiB per workbook, and 16
parts.

Excel **What-If Data Tables** are formula-bearing sensitivity engines, distinct
from Excel tables. FormulaFence reads each worksheet's `f t="dataTable"` master
from raw OOXML and privately compares its declared output range, one-/two-input
mode, one-variable row/column orientation, input-cell references,
deleted-input flags, recalculation request, and supported generic formula
metadata. A material change emits `FF034`; enable
`no_what_if_data_table_changes` to make it `FFP034` in CI.

Profiles and `FF034` details expose only structural counts: master count,
one-/two-variable and orientation counts, declared output-cell count,
recalculation requests, deleted inputs, and malformed-definition count. Input
references, output ranges, and raw metadata remain private. Equivalent A1
case/absolute-reference spellings and Boolean spellings are normalized.
Malformed, missing, overlapping, or unsupported declarations become visible
coverage warnings. FormulaFence does not calculate a table, infer its output
formula, predict scenario values, or add its input references to the normal
dependency graph. Cached scenario-output cells remain under the ordinary
cell-diff boundary.

Excel filters and column controls can change which records or fields a reviewer
sees while leaving cell contents and formulas unchanged. FormulaFence reads worksheet-level
`<autoFilter>` / `<sortState>` metadata, AutoFilters and sort state retained in
Table Definition parts, explicit row `hidden`, `outlineLevel`, `collapsed`, and
zero `ht` attributes, `sheetFormatPr@zeroHeight` (the hidden-by-default row
optimization), and zero `defaultRowHeight` / `defaultColWidth` worksheet
dimensions. It also reads raw `<cols>/<col>` `hidden`, `outlineLevel`,
`collapsed`, and zero `width` controls, applying overlapping column declarations
in file order so only present later attributes override an earlier declaration.
A material change emits `FF036`; enable
`no_filter_visibility_changes` to make it `FFP036` in CI.

Profiles and `FF036` details contain only structural counts: worksheet/table
filters, filter columns and criterion groups, sort states/conditions,
default-hidden/default-zero-dimension sheets, explicitly hidden/zero-height/
outlined/collapsed rows, visible-row overrides, explicitly hidden/zero-width/
outlined/collapsed columns, and malformed controls. Criteria, selected values,
custom lists, table names, sort keys, raw dimensions, and row/column ranges remain
private. Local A1 case/absolute-reference spelling, Boolean/default spelling,
unsigned-integer spelling, equivalent zero-dimension spelling, and equivalent
column-range segmentation are normalized. Unsupported extensions, malformed
declarations, exhausted control-update limits, and unsafe or missing table
relationships are explicit coverage warnings.
FormulaFence does not apply a filter, evaluate `SUBTOTAL`/`AGGREGATE`, infer
which formulas are visibility-sensitive, render a report, track arbitrary
positive dimensions or styles, or model overflow or outline-display settings.

Excel's per-range ignored-error declarations can suppress warnings a reviewer
would otherwise see without changing a formula or ordinary cell. FormulaFence
reads standard `<ignoredErrors>` and Office 2010 `x14:ignoredErrors` declarations
from raw worksheet OOXML, including `evalError`, inconsistent `formula`,
`formulaRange`, `unlockedFormula`, `emptyCellReference`, `listDataValidation`,
`calculatedColumn`, `numberStoredAsText`, and `twoDigitTextYear` flags. A
material change emits `FF037`; enable `no_ignored_error_changes` to make it
`FFP037` in CI.

Profiles and `FF037` details contain only structural counts: worksheets,
standard/extension containers, suppressed-warning rules, target ranges, and
warning kinds. Target ranges and individual suppressions remain private. Local
A1 case/absolute-reference spelling, Boolean spelling, and target ordering are
normalized. Malformed or unsupported containers, extension material,
attributes, flags, targets, or child markup are explicit coverage warnings.
FormulaFence does not decide whether Excel would display a warning, calculate a
formula, repair an error, change application-level error checking, or infer
downstream impact.

Modern Excel Named Sheet Views can retain alternate filter and sort settings in
separate relationship-backed worksheet parts, leaving ordinary cells and the
active AutoFilter unchanged. FormulaFence follows the documented Named Sheet
View relationship, compares each saved-view declaration privately, and resolves
its filter target by AutoFilter UID, table ID, then worksheet-owned AutoFilter.
A material change emits `FF038`; enable `no_named_sheet_view_changes` to make
it `FFP038` in CI.

Profiles and `FF038` details expose only structural counts: worksheets, parts,
views, alternate filters, column filters, criterion groups, sort rules,
conditions, and unrecognized controls. View names, IDs, criteria, target
ranges, table bindings, table-column IDs, and sort keys remain private.
Equivalent GUID, local A1 case/absolute-reference, Boolean/default, and
unsigned-integer spellings are normalized. Missing, ambiguous, mismatched,
malformed, unsupported, oversized, or unsafe parts/bindings are explicit
coverage warnings. FormulaFence does not activate/render a saved view,
calculate its filtered result, infer formula visibility sensitivity, repair
metadata, or interpret full differential-format, future extension, or rich-sort
semantics.

Excel number formats can hide or materially reinterpret a value without changing
its stored value or formula: `;;;` can make it appear blank, while custom
sections, scaling commas, dates, percentages, literals, and text placeholders
can alter what a reviewer sees. FormulaFence reads custom `<numFmt>` definitions,
base `<cellStyleXfs>`, effective `<cellXfs>` with `xfId` and
`applyNumberFormat`, direct cell `s`, row `s` where `customFormat=1`, and raw
`<cols>/<col style>` declarations. A material change emits `FF039`; enable
`no_number_format_changes` to make it `FFP039` in CI.

Profiles and `FF039` details expose only counts for default overrides, direct
cell, row, and effective column assignments, built-in/custom assignments, and
unrecognized controls. Format codes, style indexes, and targets remain private.
Equivalent custom-format ID remapping, `applyNumberFormat` Boolean spelling,
base-XF inheritance, and effective column-range splitting are normalized.
Missing custom codes, invalid IDs/indexes/targets, conflicting definitions, and
bounded parser failures are explicit coverage warnings. FormulaFence does not
render or locale-resolve a number format, validate its syntax, calculate values,
model widths/overflow, or track visual-style properties other than separately
inventoried cell fonts, fills, and alignment. A raw column `style` is compared as a declaration/default
for unallocated/new cells; it is not treated as a renderer that restyles
allocated cells.

Cell-font controls can change the review surface without changing a stored
value or formula: a white font can make a value or warning less visible against
a matching background, while face, size, emphasis, underline, and related
effects can materially change presentation. FormulaFence reads raw `<fonts>`
definitions, base `<cellStyleXfs>`, effective `<cellXfs>` with `xfId` and
`applyFont`, direct cell `s`, row `s` where `customFormat=1`, and raw
`<cols>/<col style>` declarations. A material change emits `FF040`; enable
`no_cell_font_changes` to make it `FFP040` in CI.

Profiles and `FF040` details expose only counts for default definitions, direct
cell, row, and effective column assignments, and unrecognized controls. Font
names, colour values, effects, style indexes, and targets remain private.
Equivalent font-ID remapping, common font-child ordering, `applyFont` Boolean
spelling, base-XF inheritance, and effective column-range splitting are
normalized. Missing or malformed definitions, invalid IDs/indexes/targets, and
bounded parser failures are explicit coverage warnings. FormulaFence does not
render or resolve theme colours, decide whether a font is visible against a
fill, calculate text/background contrast or values, compose alignment with
other display controls, track borders/rich-text run rendering/table styles, or
claim arbitrary visual-style coverage. A raw column `style` is
compared as a declaration/default for unallocated/new cells; it is not treated
as a renderer that restyles allocated cells.

Cell-fill controls can change the review surface without changing a stored
value or formula: a matching solid fill can make text or an indicator less
visible, while patterns and gradients can change a reviewer's visual cues.
FormulaFence reads raw `<fills>` definitions, including `patternFill` and
`gradientFill` stops, base `<cellStyleXfs>`, effective `<cellXfs>` with `xfId`
and `applyFill`, direct cell `s`, row `s` where `customFormat=1`, and raw
`<cols>/<col style>` declarations. A material change emits `FF041`; enable
`no_cell_fill_changes` to make it `FFP041` in CI.

Profiles and `FF041` details expose only counts for default definitions, direct
cell, row, and effective column assignments, and unrecognized controls. Fill
colours, pattern types, gradient geometry/stops, style indexes, and targets
remain private. Equivalent fill-ID remapping, valid pattern-colour child
ordering, `applyFill` Boolean spelling, base-XF inheritance, semantically inert
no-fill/solid-background declarations, and effective column-range splitting are
normalized. Missing or malformed definitions, invalid IDs/indexes/targets, and
bounded parser failures are explicit coverage warnings. FormulaFence does not
resolve theme colours, render patterns or gradients, calculate text/background
contrast, evaluate conditional-format differential styles, apply table styles,
calculate values, or claim arbitrary visual-style coverage. A raw column `style`
is compared as a declaration/default for unallocated/new cells; it is not
treated as a renderer that restyles allocated cells.

Cell-alignment controls can reposition, rotate, wrap, shrink, or indent an
unchanged value, warning, or visual classification. FormulaFence reads raw
`alignment` children from base `cellStyleXfs` and effective
`cellXfs` records, follows `xfId` and `applyAlignment`, and
compares direct cell `s`, row `s` where `customFormat=1`, and raw
`<cols>/<col style>` declarations. It covers horizontal/vertical placement,
text rotation, wrapping, shrinking, indentation, relative indentation,
justification, and reading order. A material effective control change emits
`FF054`; enable `no_cell_alignment_changes` to make it `FFP054` in CI.

Profiles and `FF054` details expose only counts for default definitions,
direct cell, row, effective column, and unrecognized controls. Alignment
values, style indexes, and targets remain private. Equivalent explicit defaults,
Boolean/integer spellings, semantically inert `mergeCell` compatibility
material, base-XF inheritance, `applyAlignment`, and effective
column-range splitting are normalized. Missing, duplicate, malformed, or
unsupported alignment metadata is an explicit coverage warning rather than a
silent omission. FormulaFence compares declarations, not layout: it does not
calculate width, height, merged-cell layout, overflow, final visibility,
font/fill/conditional-format composition, or Excel client rendering. A raw
column `style` is compared as a declaration/default for unallocated/new
cells; it is not treated as a renderer that restyles allocated cells.

Workbook-level DrawingML Theme controls can change the colour, font, and effect
schemes used by themed cells, charts, and drawing objects without changing a
local cell-style reference. FormulaFence reads the raw workbook Theme binding,
Theme XML, and direct Theme-image relationships/payloads in transitional and
strict OOXML namespaces. A material stored control change emits `FF053`.
Enable `no_workbook_theme_changes` to make it `FFP053` in CI.

Profiles and `FF053` details expose only Theme-part, colour-scheme,
font-scheme, format-scheme, relationship, external-relationship, image, and
malformed-metadata counts. Theme XML, scheme names, colour values, font names,
image payloads, relationship IDs, and targets remain private. Writer-selected
relationship IDs/order and equivalent internal target spelling normalize away.
Missing, duplicate, malformed, unsafe, unbound, unreadable, oversized, or
over-budget metadata becomes a visible coverage warning; reads are bounded to
16 MiB per part, 64 MiB per workbook, and 512 parts.

This policy guards stored declarations, not rendered appearance. FormulaFence
does not resolve an effective cell/chart/drawing style, render a workbook,
calculate contrast, decode an image, fetch a target, calculate formulas, or
infer Excel client behavior. The boundary follows the Open XML SDK
[WorkbookPart](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.packaging.workbookpart?view=openxml-2.20.0)
Theme-part surface and Microsoft's
[conditional-formatting guidance](https://learn.microsoft.com/en-us/office/open-xml/spreadsheet/working-with-conditional-formatting),
which illustrates Theme-indexed colours in spreadsheet formatting.

SpreadsheetML can retain the last calculated result beside a formula in the
same `<c>` cell. FormulaFence reads raw `<f>` and `<v>` elements together,
privately fingerprints numeric, string, Boolean, and error results, and compares
the saved result without exposing its value, error text, digest, or location. A
result-cache change emits `FF042` when it has no changed formula at that cell
and no ordinary changed cell reaches it through FormulaFence's static dependency
graph. Enable `no_formula_cached_result_changes` to make it `FFP042` in CI.

Profiles and `FF042` details expose only formula-cell, cached-result,
missing-result, result-type, and malformed-metadata counts. Equivalent finite
numeric and Boolean spellings are normalized; blank or absent results remain
visible as missing caches. Unsupported or malformed cache metadata is a
coverage warning rather than a silent omission. FormulaFence does not calculate
or validate a result, decide whether it is stale or tampered, or model volatile,
dynamic, external, or calculation-engine dependencies. A normal recalculation
can therefore require review if it changes a cache without a statically visible
input edit.

SpreadsheetML can also split a string into character-level `<r>` runs. Their
`<rPr>` controls sit outside the cell-style table, so a phrase can be made less
visible while the concatenated cell value remains unchanged. FormulaFence follows
referenced shared-string items and direct inline strings, privately compares
run-property sequences, styled character boundaries, and phonetic-hint material,
and emits `FF043` for a material presentation control change. Enable
`no_rich_text_run_changes` to make it `FFP043` in CI.

Profiles and `FF043` details expose only shared-item/cell/run, inline-cell/run,
phonetic, and malformed-control counts. Text, font and colour details,
shared-string indexes, and cell locations remain private. Equivalent property
ordering, colour case, and explicit false Boolean properties are normalized.
An ordinary text edit within an unchanged run-property sequence stays a normal
cell diff; a moved styled boundary with the same text remains guarded.
Malformed, unsupported, or unreadable rich-text metadata becomes an explicit
coverage warning. FormulaFence does not render cells, resolve theme colours,
calculate contrast, determine visibility, preserve rich text, or guarantee
cross-version Excel rendering equivalence.

An ordinary worksheet-cell hyperlink can keep the same friendly cell value
while changing its external target, file target, in-workbook location, display
override, or ScreenTip. FormulaFence reads standard SpreadsheetML
`hyperlink` declarations and Office 2016 `xr:hyperlink` declarations directly
from worksheet XML, then resolves the selected raw worksheet relationship
semantics without following a target. It privately compares the cell/range
binding, declaration material, location, display/ScreenTip, and selected
relationship type, target, and mode. A material change emits `FF047`; enable
`no_cell_hyperlink_changes` to make it `FFP047` in CI.

Profiles and `FF047` details expose only worksheet/hyperlink,
location/display/ScreenTip, relationship/external-relationship, and
malformed-metadata counts. Targets, references, locations, display strings,
ScreenTips, relationship IDs, and revision UIDs remain private. Consistent
writer-generated relationship-ID/revision-UID rewrites, relationship ordering,
and equivalent internal-part target spelling stay quiet. Missing, duplicate,
unbound, malformed, unsafe, unreadable, oversized, or over-budget metadata
becomes a visible coverage warning; raw worksheet reads are bounded to 16 MiB
per worksheet, 64 MiB per workbook, and 512 parts. After raw inspection,
FormulaFence uses a hyperlink-removed reader copy so a malformed declaration
cannot turn the package evidence into a reader failure.

FormulaFence does not render, resolve, fetch, or follow a link; test target
availability; inspect linked content; infer trust-zone or client behavior; or
interpret a `HYPERLINK()` formula beyond the ordinary formula diff. This scope
follows the Open XML
[Hyperlink](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.hyperlink?view=openxml-3.0.1)
and Office 2016
[Hyperlink](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.office2016.excel.hyperlink?view=openxml-3.0.1)
definitions.

Office 2010 worksheet sparklines are stored as `x14:sparklineGroups` worksheet
extensions, not ordinary cells. A group can change its type, axes, display,
markers, colours, and optional date-axis source; each nested sparkline can
change the private source formula or destination cell without modifying the
visible data cells. FormulaFence reads the raw extension before the ordinary
reader drops it, privately compares those bindings and controls, and emits
`FF048`. Enable `no_worksheet_sparkline_changes` to block it as `FFP048`.

Profiles and `FF048` details expose only aggregate worksheet/group/sparkline,
source/date-axis-source, colour-control, and malformed-metadata counts. Source
formulas, destination cells, group properties, and colour values remain
private. Equivalent local direct-range spelling, Boolean/numeric spelling,
colour case, and declaration ordering normalize away. Missing, duplicate,
malformed, unreadable, oversized, or over-budget records become an explicit
coverage warning; raw worksheet XML is bounded to 16 MiB per worksheet, 64 MiB
per workbook, and 512 parts. FormulaFence uses a Sparkline Group-removed
temporary reader copy only after raw inspection, so a reader that omits the
extension cannot suppress the evidence.

FormulaFence does not calculate source values, resolve names or external
sources, render a sparkline, assess visual accessibility, or guarantee
cross-version Excel rendering equivalence. This scope follows the Open XML
[SparklineGroup](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.office2010.excel.sparklinegroup?view=openxml-3.0.1)
and [CT_Sparkline](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/6b28a993-e0fd-451d-860e-35097c6baa77)
definitions.

SpreadsheetML XML Maps attach a schema and map-level refresh/export behavior to
XML table columns or individual worksheet cells. A changed map can redirect
which XML field is imported or exported, switch a file or connection binding,
or alter append, format, sort/filter, and validation behavior without changing
ordinary cell values or formulas.

FormulaFence reads raw XML Maps, table XML-column-property, and single-cell
table declarations before ordinary workbook readers discard or normalize them.
It privately compares schemas, map and data-binding material, table and
single-cell bindings, and related workbook/worksheet relationship targets.
Such a change emits FF049. Enable `no_xml_mapping_changes` to block it as
FFP049.

Profiles and FF049 details expose only aggregate map/schema/binding,
file/connection, table, single-cell, and malformed-metadata counts. Schemas,
map names, XPath expressions, table identities, target cells, connection
identities, and relationship targets remain private. Equivalent Boolean and
unsigned-integer spelling, relationship IDs/order, and equivalent internal
target spelling stay quiet. Missing, duplicate, malformed, unsafe, unbound,
unreadable, oversized, or over-budget metadata becomes a coverage warning; raw
reads are bounded to 16 MiB per part, 64 MiB per workbook, and 512 parts.

FormulaFence compares only stored declarations. It does not import or export
XML, validate XML instances or schemas, open files or connections, fetch remote
data, calculate a refresh, or infer client behavior. The scope follows the
Open XML [Map](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.map?view=openxml-3.0.1),
[XmlProperties](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.xmlproperties.xpath?view=openxml-3.0.1),
and [SingleXmlCells](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.singlexmlcells?view=openxml-3.0.1)
definitions.

Excel rich data types can retain linked entity values, provider-backed fields,
web-image associations, and worksheet value-metadata bindings outside ordinary
cell values. A stored rich-data change can therefore alter a workbook's
operational data surface without a normal cell or formula diff.

FormulaFence reads raw Rich Value Data, structure, type, array, supporting
property-bag, style, web-image, and rich-value-relationship parts; their
workbook/package relationships; and `XLRICHVALUE` metadata/cell bindings. It
privately compares values, structures, web-image and rich-value endpoints, and
bindings. Such a change emits `FF051`. Enable `no_rich_data_changes` to block
it as `FFP051`.

Profiles and findings expose aggregate part, value, structure, array,
property-bag, metadata-binding, bound-cell, web-image, relationship,
external-reference, and malformed-metadata counts only. Entity values,
provider data, field names, identifiers, URLs, image references, relationship
IDs, and bound-cell locations stay private. Equivalent relationship IDs/order
and internal-target spelling normalize away. Missing, duplicate, malformed,
unsafe, unreadable, oversized, or over-budget metadata produces a coverage
warning; reads are bounded to 16 MiB per XML part, 64 MiB per workbook, and
512 parts.

This policy guards stored declarations only. It does **not** contact a
provider, refresh entity values, calculate formulas, fetch or validate an image
or relationship target, or infer Excel client behavior. The boundary follows
Microsoft's [Rich Value Data](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/896934fd-8df7-43f4-b154-2d39371c270d),
[Rich Value Structure](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/d90f6d91-d868-4b94-9d26-ec3b1492cec6),
[Rich Value Types](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/5d213b66-3196-4516-b63c-eef80d926f4a),
and [Rich Value Web Image](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/4f3a80fd-1776-407f-8807-2497a4692dea)
definitions.

Generic Custom XML, workbook-bound Custom Data, and custom document properties
are three separate places Excel add-ins can persist workbook-specific state.
They can carry an approval decision, workflow state, integration identifier, or
another add-in setting without changing a worksheet cell or formula.

FormulaFence reads generic `customXml/item*.xml` data, Custom XML
property/schema parts and relationships, workbook-linked `xl/customData`
property/binary parts, and custom document properties. Power Query
`DataMashup` Custom XML remains under the Power Query controls and is not
double counted. Such a change emits `FF052`. Enable
`no_custom_data_store_changes` to block it as `FFP052`.

Profiles and findings expose only aggregate custom-XML, schema,
relationship/external-relationship, Custom Data property/payload, custom
document-property, linked-property, and malformed-metadata counts. Custom XML,
schema URIs, property names and values, storage IDs, binary payloads,
relationship IDs, and relationship targets stay private. Equivalent
relationship IDs/order and writer-selected document-property `pid` values
normalize away. Custom XML `itemID` and Custom Data `id` storage identities
are compared privately because an add-in can bind state to them. Missing,
duplicate, malformed, unsafe, unbound, unreadable, oversized, or over-budget
metadata produces a coverage warning; reads are bounded to 16 MiB per part, 64
MiB per workbook, and 512 parts.

This policy guards persisted package state only. FormulaFence does **not**
execute an add-in, resolve a property, follow or fetch a target, interpret a
binary payload, calculate formulas, or infer Excel client behavior. The
boundary follows Microsoft's guidance on
[persisting add-in state](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/persisting-add-in-state-and-settings),
[Custom Data](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/7c53f6f4-fea8-43f7-a4b0-ba6e14d0eb78),
[Custom Data Properties](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/1f4aa666-c966-4ecf-8399-28390399c891),
and Excel's
[CustomDocumentProperties](https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.customdocumentproperties).

Package/content signatures and VBA project code signatures are separate Excel
trust surfaces. A workbook can therefore retain identical formulas, values, and
`xl/vbaProject.bin` bytes while package-signature metadata, a certificate part,
or a VBA signature payload changes. FormulaFence reads the raw OPC graph:
package-root origin, origin-to-XML-signature, XML-signature-to-certificate, and
VBA signature relationships; it privately compares XMLDSIG envelopes/signed
references, certificate-part payloads, and the conventional classic, Agile, and
V3 VBA signature binaries.

Such a change emits `FF050`. Enable `no_digital_signature_changes` to block it
as `FFP050`. Profiles and finding details expose only aggregate
origin/XML-signature, signed-reference, embedded-certificate/certificate-part,
VBA-signature, and malformed-metadata counts. Signature XML, reference URIs,
certificate identities/contents, binary signature payloads, relationship IDs,
and targets stay private. Equivalent relationship IDs/order, equivalent
internal-target spelling, and whitespace in XMLDSIG base64 values normalize
away. Missing, duplicate, malformed, unsafe, unbound, unreadable, oversized,
or over-budget metadata produces a coverage warning; reads are bounded to
16 MiB per part, 64 MiB per workbook, and 512 parts.

This policy guards stored signature envelopes only. It does **not** verify a
signature or digest, XML transforms or signed-reference coverage, certificate
chain/identity/trust/expiry/revocation, timestamps, the actual signed
contents, or VBA code validity. It does not fetch a certificate or contact a
trust service. Microsoft's [OPC digital-signature
overview](https://learn.microsoft.com/en-us/previous-versions/windows/desktop/opc/open-packaging-conventions-overview)
explicitly leaves signer/trust validation to the package consumer; see also
Excel's [workbook and VBA signing guidance](https://learn.microsoft.com/en-us/troubleshoot/microsoft-365-apps/excel/digital-signatures-code-signing).

Traditional Excel Notes are stored in worksheet-associated SpreadsheetML
comments parts, not in ordinary cells. Their author association, text, cell
binding, comment properties, and optional rich-text material sit in that part;
their visible/hidden state and layout sit in a worksheet `legacyDrawing` VML
part. A modern threaded comment can retain a legacy Note placeholder whose
author carries `tc={GUID}` for reconciliation.

FormulaFence follows these comments and VML bindings and privately compares
author association, text/presentation, comment properties, cell association,
placeholder reconciliation state, Note VML visibility/layout, and relationship
semantics. A material change emits `FF046`; enable
`no_legacy_comment_changes` to make it `FFP046` in CI. Profiles and finding
details expose only aggregate worksheet/part, author/comment/text/rich-text/
property/placeholder, Note-shape/visibility/anchor, relationship, and
malformed-metadata counts. Note bodies, authors, cell locations, raw VML,
targets, raw relationship IDs, and GUIDs remain private.

Writer-generated comment shape IDs, VML shape IDs, package relationship IDs,
and consistently rekeyed placeholder GUIDs normalize away. Missing, duplicate,
unsafe, unbound, malformed, unreadable, oversized, or over-budget metadata
becomes a coverage warning; XML reads are bounded to 16 MiB per part, 64 MiB
per workbook, and 512 parts. The ordinary workbook reader uses a
Note-quarantined copy only after the raw scanner has recorded the original
package, so parser tolerance or a bad target cannot turn a review finding into
a reader crash.

This rule compares stored package state only. It does not render Note text/VML,
resolve authors, fetch a relationship target, execute linked content, determine
client display placement, or infer notifications, permissions, account,
cloud-sync, or client-visibility behavior. The format boundary follows the
Open XML [Comment](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.comment?view=openxml-3.0.1),
[Authors](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.authors?view=openxml-3.0.1),
and [LegacyDrawing](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.legacydrawing?view=openxml-3.0.1)
definitions, plus Microsoft's [threaded-comment placeholder
rule](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/6383f002-c90b-401c-a1d7-66b97b14cb3e).

Modern Excel threaded comments live in worksheet-associated comment parts and a
workbook-associated persons part instead of ordinary cells. FormulaFence follows
those package relationships and privately compares the complete stored
comment/reply tree, comment text, cell/timestamp/resolution declarations,
mention ranges and person associations, extension material, and person records.
A material change emits `FF045`; enable
`no_threaded_comment_changes` to make it `FFP045` in CI.

Profiles and `FF045` details expose only aggregate worksheet/part/thread,
comment/reply/resolved/text, mention, person/unreferenced-person, relationship,
and malformed-metadata counts. Bodies, cell references, timestamps, parent
links, author and mentioned-person identities, and all raw IDs remain private.
FormulaFence rebuilds person, mention, and reply links before comparison so
consistent writer-generated GUID/relationship-ID rewrites do not create a
finding. It normalizes normal OOXML Boolean spellings for `done`; malformed,
unsafe, unbound, missing, unreadable, oversized, or over-budget parts become
coverage warnings. XML reads are bounded to 16 MiB per part, 64 MiB per
workbook, and 512 parts.

This rule compares stored package state only. It does not render comments,
validate mention offsets against comment text, send notifications, resolve
accounts, determine whether a legacy placeholder renders, or infer
collaboration, permissions, cloud-sync, or client-visibility behavior. The format boundary
follows Microsoft's [threaded-comment overview](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/e0fb917a-1107-409a-852f-13b47aea70dc),
[Threaded Comments part](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/66e1875d-c60a-48eb-bf88-41066d45fea8),
[Persons part](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/1a170d26-42a2-46f0-b2b6-0ff1dec1c344),
and [threaded-comment schema](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/adb84732-9fc8-48b6-bddc-6b0bcdaad940).

Non-chart Worksheet DrawingML can host a regular `xdr:sp` shape or a nested
`xdr:grpSp` shape group under a standard worksheet drawing anchor. Those shape
declarations can carry visible text, run formatting, positioning, macro
assignments, `textlink` formulas, and click/hover hyperlink relationships
outside cells. FormulaFence privately fingerprints supported anchor, regular
shape, and group XML plus referenced relationship semantics. It reports only
structural worksheet/drawing/anchor, shape/text/group, text paragraph/run,
macro/text-link/hyperlink, relationship, and malformed-control counts. A
material change emits `FF044`; enable
`no_worksheet_drawing_shape_changes` to make it `FFP044` in CI.

Text, presentation details, geometry, anchors, descriptions, macro names,
text-link formulas, relationship IDs/targets, and raw XML stay private.
Writer-chosen non-visual IDs, relationship-ID rewrites, and equivalent colour
case are normalized. Missing, malformed, unsafe, unreadable, oversized, and
over-budget metadata is a visible coverage warning; XML reads are bounded to
16 MiB per part, 64 MiB per workbook, and 512 parts. FormulaFence does not
render or assess visibility, resolve theme colours/contrast, calculate a text
link, execute a macro assignment, retrieve a target, parse/hash media, or
inspect pictures, connectors, graphic frames, SmartArt, or other non-`xdr:sp`
drawing objects.

Worksheet controls and OLE objects can bind a sheet to persisted ActiveX state,
modern or legacy form-control formulas, macro assignments, linked cells, raw
OLE data, or an external OLE link outside ordinary cells and the VBA payload.
FormulaFence starts from standard worksheet relationships, confirms modern
`<controls>` / `<oleObjects>` markup, and follows `vmlDrawing` relationships to
legacy VML `ClientData`. It inspects ActiveX `ocx` persistence parts,
`formControlPr` parts, and non-`Note` VML controls; it counts legacy macro,
linked-cell, source-range, and camera-range bindings without exposing their
values. It streams only direct safe internal ActiveX binary and OLE/package
targets into private hashes; it does not parse them. Writer-chosen relationship
IDs and equivalent internal target spellings are normalized when semantics are
unchanged. Profiles expose only structural counts for controls, bindings, OLE
behavior, relationships, and inspected versus uninspected raw payloads. Control
names, class IDs, licenses, captions, macros, formulas, ranges, OLE identities,
relationship targets, XML, and payload bytes never enter a profile or diff. A
material change emits `FF029`; enable
`no_worksheet_embedded_control_changes` to make it `FFP029` in CI.
FormulaFence does not initialize an ActiveX control, deserialize/open an OLE
object or package, render a VML drawing, read comment-note content into the
control inventory, follow an external relationship, or infer event dispatch.
Relevant XML reads are bounded to 16 MiB per part,
64 MiB per workbook, and 512 parts; direct payload hashes are bounded to 32 MiB
per part, 64 MiB per workbook, and 512 parts. Missing, malformed, orphaned,
unbound, oversized, or over-budget material remains a visible coverage warning.
VML/drawing layout, embedded payload formats, event behavior, and behavior not
reachable through this relationship-backed chain are outside this control's
scope.

Power Query stores query definitions in a `DataMashup` Custom XML part. FormulaFence
parses the documented length-prefixed container and privately compares its
`Section1.m` formula document, logical package material, stable metadata, and
formula-firewall permissions. It reports only counts and safe control state:
M text, query/source names, metadata values, embedded content, telemetry IDs,
and user-bound permission-binding blobs never enter a profile or diff. `sqmid`
telemetry and result-only refresh metadata are intentionally ignored. A material
change emits `FF024`; enable `no_power_query_changes` to make it `FFP024` in CI.
FormulaFence does not execute M, refresh a connection, assess a source, or infer
the data a query would return.

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
