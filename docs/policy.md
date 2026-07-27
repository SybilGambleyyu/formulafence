# Policy reference

FormulaFence keeps controls in a small YAML file so that the rule itself is
reviewable alongside the model. A policy is evaluated by `formulafence check`
or `formulafence portfolio --policy`; each violation is emitted as a `FFP…`
finding and makes the command exit with status `1`.

```yaml
version: 1
rules:
  no_formula_to_value: true
  no_new_external_links: true
  no_external_workbook_link_surface_changes: true
  no_new_broken_references: true
  no_macro_changes: true
  no_xlm_macro_sheet_changes: true
  no_xlm_automatic_macro_binding_changes: true
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
  no_custom_workbook_view_changes: true
  no_table_style_control_changes: true
  no_shared_workbook_revision_changes: true
  no_number_format_changes: true
  no_cell_font_changes: true
  no_cell_fill_changes: true
  no_workbook_theme_changes: true
  no_cell_alignment_changes: true
  no_cell_border_changes: true
  no_worksheet_dimension_changes: true
  no_worksheet_display_control_changes: true
  no_worksheet_print_layout_changes: true
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
  no_worksheet_image_changes: true
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
  no_external_relationship_changes: true
  no_formula_external_action_changes: true
  no_python_in_excel_changes: true
  no_office_custom_function_changes: true
  no_unqualified_runtime_function_changes: true
  no_worksheet_code_resource_registration_changes: true
  no_formula_defined_xlm_registration_changes: true
  no_formula_defined_xlm_evaluation_changes: true
  no_formula_defined_xlm_action_changes: true
  no_formula_defined_xlm_get_cell_changes: true
  no_formula_defined_xlm_environment_information_changes: true
  no_formula_environment_information_changes: true
  no_power_query_changes: true
  no_portfolio_membership_changes: true
  no_cross_workbook_impacts: true
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
| `no_external_workbook_link_surface_changes` | boolean | A private static external-workbook link ledger changes in a worksheet formula, defined name, data-validation criterion, or standard/ChartEx chart formula. This catches a same-location source or target swap that `no_new_external_links` intentionally does not treat as a new link. Chart parts with unavailable formula coverage fail closed. |
| `no_new_broken_references` | boolean | A formula adds `#REF!`. |
| `no_macro_changes` | boolean | The `xl/vbaProject.bin` payload is added, removed, or has a different SHA-256. |
| `no_xlm_macro_sheet_changes` | boolean | An Excel 4.0 / XLM macro-sheet declaration, program XML, related-part relationship, or direct internal related-part payload changes. |
| `no_xlm_automatic_macro_binding_changes` | boolean | A workbook-scoped `Auto_Open`, `Auto_Close`, `Auto_Activate`, or `Auto_Deactivate` defined name directly bound to a declared XLM macro-sheet cell is added, removed, or materially retargeted. |
| `no_ribbon_customization_changes` | boolean | An Office RibbonX custom-UI package declaration, control/callback XML, or direct relationship changes. |
| `no_office_web_addin_changes` | boolean | An Office Web Add-in task-pane or worksheet/in-content binding, configuration, web-extension definition, or direct relationship changes. |
| `no_chart_definition_changes` | boolean | A legacy DrawingML or Office 2016+ ChartEx chart binding, chart definition, cached series data, chart-overlay shape, direct relationship, or bounded direct related payload changes. |
| `no_pivot_table_definition_changes` | boolean | A PivotTable binding/layout, cache schema, shared item, cache-record relationship, or bounded cached-record payload changes. Source and refresh controls remain under `no_external_data_connection_changes`. |
| `no_slicer_timeline_cache_changes` | boolean | A Slicer or Timeline workbook binding, cached filter state, source binding, filtered-PivotTable binding, or direct cache-part relationship changes. |
| `no_power_pivot_data_model_changes` | boolean | An embedded Power Pivot/Data Model workbook binding, `x15:dataModel` declaration, direct model-part relationship, or bounded raw model payload changes. |
| `no_what_if_data_table_changes` | boolean | An Excel What-If Data Table master changes its output range, one-/two-variable mode, orientation, input references, deleted-input state, recalculation request, or supported raw formula metadata. This is unrelated to an Excel table definition. |
| `no_scenario_manager_changes` | boolean | An Excel Scenario Manager worksheet changes its selected/shown scenario state, result-summary references, scenario definition, protection flags, comments/users, stored input values/references, deleted/undone state, or input display number formats. Scenario names, comments, users, values, and references are compared privately. |
| `no_filter_visibility_changes` | boolean | A worksheet/Table AutoFilter, stored filter criterion, filter sort state, explicitly hidden/zero-height/outlined/collapsed row, hidden/zero-width/outlined/collapsed column, or hidden/zero-dimension-by-default worksheet setting changes. Criteria, selected values, sort keys/lists, raw dimensions, table names, and row/column ranges are compared privately. |
| `no_ignored_error_changes` | boolean | A standard or Office 2010 extension ignored-error declaration changes a per-range suppression of Excel evaluation, formula-consistency, range-omission, unlocked-formula, empty-reference, list-validation, calculated-column, text-number, or two-digit-year warnings. Targets and exact suppressions are compared privately. |
| `no_named_sheet_view_changes` | boolean | A relationship-backed Excel Named Sheet View, alternate AutoFilter criterion, sort rule, or reconciled base-filter binding changes. View names, IDs, criteria, ranges, table bindings, and sort keys are compared privately. |
| `no_custom_workbook_view_changes` | boolean | A legacy Excel Custom View's workbook declaration, GUID-linked per-sheet alternate display/filter/print state, or recognized sheet binding changes. View names, GUIDs, sheet bindings, ranges, filters, pane locations, print settings, and raw XML are compared privately. |
| `no_table_style_control_changes` | boolean | An Excel Table Style binding/toggle, applicable custom Table Style definition or resolved Dxf material, or direct Table/TableColumn Dxf or named-cell-style reference changes. Table/style names, formatting, colours, IDs, and raw XML are compared privately. |
| `no_shared_workbook_revision_changes` | boolean | A legacy shared-workbook revision header/log declaration, historic revision record, tracking/retention/protection control, relationship, or coverage state changes. Prior/new values, locations, author identities, timestamps, comments, GUIDs, relationship IDs, and raw XML are compared privately. |
| `no_number_format_changes` | boolean | An effective default, direct-cell, row, or column number-format control changes. Format codes, style IDs, and cell/row/column targets are compared privately. |
| `no_cell_font_changes` | boolean | An effective default, direct-cell, row, or column font control changes. Font names, colours, effects, style IDs, and cell/row/column targets are compared privately. |
| `no_cell_fill_changes` | boolean | An effective default, direct-cell, row, or column fill control changes. Fill colours, pattern/gradient definitions, style IDs, and cell/row/column targets are compared privately. |
| `no_workbook_theme_changes` | boolean | A workbook Theme binding, stored colour/font/format scheme, direct Theme-image relationship, or direct Theme-image payload changes. Theme XML, scheme names, colours, font names, image bytes, relationship IDs, and targets are compared privately. |
| `no_cell_alignment_changes` | boolean | An effective default, direct-cell, row, or column alignment control changes. Alignment values, style IDs, and cell/row/column targets are compared privately. |
| `no_cell_border_changes` | boolean | An effective default, direct-cell, row, or column border control changes. Border definitions, colours, style IDs, and cell/row/column targets are compared privately. |
| `no_worksheet_dimension_changes` | boolean | A material worksheet default row/column size, Office 2010 baseline adjustment, positive direct row/column size, AutoFit, or active thick-border automatic row-height adjustment changes. Dimension values, targets, and raw XML are compared privately. |
| `no_worksheet_display_control_changes` | boolean | A material raw worksheet view changes hidden-zero, formula-display, gridline/gridline-colour, row/column-header, outline-symbol, ruler, page-whitespace, right-to-left, non-normal-view, or split/frozen-pane controls. Sheet names, targets, pane positions, and raw view XML are compared privately. |
| `no_worksheet_print_layout_changes` | boolean | A material saved print-area/title, print-option, margin, page-setup, fit-to-page, header/footer, or manual page-break control changes. Print ranges, header/footer text, page values, printer-setting references, and raw XML are compared privately. |
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
| `no_worksheet_drawing_shape_changes` | boolean | A non-chart Worksheet DrawingML regular `xdr:sp`, connector `xdr:cxnSp`, nested `xdr:grpSp`, or recognized SmartArt `xdr:graphicFrame` anchor/layout, presentation, diagram-component or bounded Diagram Data image payload, connector attachment, macro/text-link, or referenced-relationship change. Shape and SmartArt text, formatting, formulas, anchors, IDs, component content, image bytes, and targets are compared privately. Other non-chart graphic-frame URI types are coverage gaps. |
| `no_worksheet_image_changes` | boolean | A native anchored DrawingML `xdr:pic`, worksheet `<picture>` background, header/footer VML image, related relationship, visual declaration, anchor, or bounded direct image payload changes. Image bytes, names/descriptions, anchors, formatting, IDs, targets, and raw XML are compared privately. |
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
| `no_external_relationship_changes` | boolean | Any root or part-level OPC relationship with an external target changes, including an opaque relationship that no feature-specific scanner recognizes. Source parts, types, IDs, targets, unknown metadata, and raw XML are compared privately. |
| `no_formula_external_action_changes` | boolean | A stored `HYPERLINK`, `WEBSERVICE`, `IMAGE`, `RTD`, `STOCKHISTORY`, or documented Cube-family formula call in a cell, formula-defined name, or named `LAMBDA`; its private inventory; relevant name-definition material; or a statically visible input changes. FormulaFence uses private signatures and static dependency paths; it never evaluates a formula, resolves a destination, requests content, queries a cube, or starts a provider. |
| `no_formula_dde_link_changes` | boolean | A direct lexical `application|topic!item` DDE-style formula link in a worksheet formula, formula-defined name, or named `LAMBDA`; its private invocation/definition material; or a statically visible input to an invoking named `LAMBDA` changes. Services, topics, items, formulas, locations, and identities are compared privately. FormulaFence never evaluates a formula, resolves an endpoint, looks up/launches a DDE server, or sends a DDE command. Raw `externalLink` DDE/OLE packages remain under `no_external_link_package_changes`. |
| `no_python_in_excel_changes` | boolean | Stored Python-in-Excel package code/environment/XML, a `PY` formula binding, function inventory, or a statically visible input changes. Python source, environment IDs, script indexes, formula arguments, locations, and raw XML are compared privately; FormulaFence never parses or runs Python, evaluates `PY`, or contacts the Microsoft Cloud runtime. |
| `no_office_custom_function_changes` | boolean | A namespaced Office custom-function call candidate, its private formula/call inventory, or a statically visible input changes. A candidate is not proof that an add-in is installed: FormulaFence does not load the add-in manifest or code, execute a formula, or contact a custom-function runtime. Names, namespaces, cells, formulas, and arguments are compared privately. |
| `no_unqualified_runtime_function_changes` | boolean | An unknown unqualified worksheet-call candidate, relevant formula-defined-name chain, private invocation/definition inventory, or a statically visible input changes. Candidate names, formulas, arguments, locations, and provider identities stay private. FormulaFence never evaluates a formula; resolves or loads VBA, COM/Automation, XLL, or another registered provider; or inspects host trust settings. |
| `no_worksheet_code_resource_registration_changes` | boolean | A stored worksheet or formula-defined `REGISTER.ID` call, relevant formula-defined-name chain, private call inventory, or statically visible input changes. Module paths, procedure names, type strings, formulas, arguments, locations, and name identities are compared privately. FormulaFence never evaluates a formula, resolves a path, loads a DLL/XLL, or determines whether registration succeeds. |
| `no_formula_defined_xlm_registration_changes` | boolean | A legacy XLM `REGISTER` call stored in a formula-defined name or named `LAMBDA`, its relevant definition chain/private invocation inventory, or a statically visible input changes. Module paths, procedure names, type strings, formulas, arguments, locations, and name identities are compared privately. FormulaFence never evaluates a formula, executes a macro, resolves a path, loads a DLL/XLL, or determines whether registration succeeds. |
| `no_formula_defined_xlm_evaluation_changes` | boolean | A legacy XLM `EVALUATE` call stored in a formula-defined name or named `LAMBDA`, its relevant definition chain/private invocation inventory, or a statically visible argument input changes. Expressions, formulas, arguments, locations, and name identities are compared privately. FormulaFence never evaluates the text, parses the runtime-generated expression, executes a macro, or infers dependencies inside that expression. |
| `no_formula_defined_xlm_action_changes` | boolean | A selected legacy XLM `CALL`, `EXEC`, `EXECUTE`, `RUN`, `SEND.KEYS`, or `ON.*` action/event-dispatch call stored in a formula-defined name or named `LAMBDA`; its relevant definition chain/private invocation inventory; or a statically visible input changes. Function targets, handler names, formulas, arguments, locations, and name identities are compared privately. FormulaFence never evaluates a formula, resolves a target or handler, loads a DLL, sends DDE, or runs a macro or program. |
| no_formula_defined_xlm_get_cell_changes | boolean | A legacy XLM GET.CELL call stored in a formula-defined name or named LAMBDA, its relevant definition chain/private invocation inventory, or a statically visible argument input changes. Information types, references, formulas, arguments, locations, and name identities are compared privately. FormulaFence never evaluates the call, determines its requested information type, resolves dynamic references, or simulates Excel formatting, display, comments, protection, or other workbook state. |
| no_formula_defined_xlm_environment_information_changes | boolean | A selected legacy XLM GET.WORKBOOK, GET.WORKSPACE, or GET.DOCUMENT call stored in a formula-defined name or named LAMBDA, its relevant definition chain/private invocation inventory, or a statically visible argument input changes. Information types, references, formulas, arguments, locations, and name identities are compared privately. FormulaFence never evaluates the call, determines its requested information type, resolves dynamic references, or simulates workbook, workspace, document, client, add-in, printer, or other Excel state. |
| no_formula_environment_information_changes | boolean | A native CELL, INFO, SHEET, or SHEETS call stored in a worksheet formula, formula-defined name, or named LAMBDA; its relevant definition chain/private invocation inventory; or a statically visible argument input changes. Profiles separately aggregate CELL calls without an explicit reference, SHEET calls, SHEETS calls, and SHEETS calls without an explicit reference. When a complete raw OOXML tab catalog changes, FormulaFence also raises FF072 for stored SHEET or omitted-reference SHEETS calls; all declared tabs, including hidden, chart, macro, and dialog sheets, are in scope, while visibility-only changes are not this condition. Information types, references, formulas, arguments, locations, name identities, and raw tab-catalog comparison material are compared privately; ordinary sheet inventory remains normal reviewer context. FormulaFence never evaluates the call, resolves dynamic arguments, infers a selected cell, or simulates file, folder, client, workspace, workbook, or other Excel state. |
| `no_power_query_changes` | boolean | A Power Query Data Mashup formula, package definition, stable query metadata, or formula-firewall permission control changes. |
| `no_portfolio_membership_changes` | boolean | A `formulafence portfolio` run finds a supported workbook relative path only in the baseline or only in the candidate directory. This blocks `FF077` as `FFP077`; it does not infer that same-content files with different paths are a rename. |
| `no_cross_workbook_impacts` | boolean | A changed candidate cell has one or more statically reachable formula cells in another candidate workbook through FormulaFence's exact, safely relative external-A1, external-3-D-A1, book-only external-table selector, direct workbook-scoped-name or sheet-local-name, or validated package-indexed-A1/name/table portfolio graph. A table must have an explicit static selector and exactly one case-insensitive source-table match; FormulaFence maps only its static cells. A 3-D span uses exact forward ordinary-worksheet endpoints and expands only when the inspected source candidate has a complete raw OOXML tab catalog consistent with its worksheet order. A workbook-scoped consumer alias may terminate in one supported static spelling through a finite, acyclic chain of exact unqualified non-A1 name identities, unless a same-named local consumer name shadows it. An eligible workbook-scoped formula-defined name, including a named `LAMBDA` at an actual function call, may retain every static endpoint and fixed internal input only when every external token is already a direct/package-validated endpoint and all other references are static. A bare LAMBDA name, plus broken, unresolved, dynamic, relative, recursive, local-3-D, spilled, explicitly intersected, tokenizer-failed, local, or locally shadowed definitions remains outside the bridge. An indexed form must traverse one document-order `externalReference`, one `externalBook`, and one external `externalLinkPath` relationship; any source name must expand completely to static internal A1 destinations in the already-inspected candidate, and an explicit source sheet selects only that local scope. This blocks `FF079` as `FFP079` without resolving a file, trusting a cache, evaluating a formula, or inferring an unresolved target. |
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

FormulaFence separately maintains a private **static external-workbook
link-surface ledger** across worksheet formulas, workbook/sheet-local defined
names, data-validation criteria, and standard DrawingML/ChartEx chart formula
elements. It detects a material source or target swap at the same cell or
object, which is deliberately broader than `no_new_external_links` but does
not evaluate a formula, resolve/open a source workbook, refresh data, or trust
a cached value. Profiles and `FF081` / `FFP081` evidence expose only surface
and endpoint counts; source paths, workbook/sheet/name identities, formulas,
validation ranges, and chart-part identities remain inside this ledger's
private signatures. A chart part whose formula coverage is unavailable is
retained as opaque coverage evidence and makes
`no_external_workbook_link_surface_changes` fail closed.
Text-built references and unsupported formula syntax are not inferred; retain
`no_new_tokenization_failures` for newly unavailable formula coverage.

FormulaFence also inspects every canonical OPC relationship part—not just
relationships reached through a recognized workbook feature—for
`TargetMode="External"`. This boundary catches remote hyperlink, image, and
opaque endpoints introduced into arbitrary package parts. It exposes only
aggregate relationship part/source/target and hyperlink/image/other counts;
source part paths, relationship types and IDs, targets, unknown metadata, and
raw XML remain private. A material endpoint, type, source, or coverage change
emits `FF063`; `no_external_relationship_changes` makes it `FFP063` in CI.
Relationship-ID-only rewrites normalize. Duplicate, orphaned, malformed,
unsafe, unreadable, oversized, or over-budget relationship metadata stays
visible as coverage evidence. FormulaFence does not resolve, open, fetch, or
trust any target; it bounds relationship XML to 16 MiB per part, 64 MiB per
workbook, and 512 parts.

FormulaFence separately inventories stored `HYPERLINK`, `WEBSERVICE`, `IMAGE`,
`RTD`, `STOCKHISTORY`, and documented Cube-family formula calls (including
`_xlfn.` compatibility spellings) in cells, formula-defined names, and named
`LAMBDA` bodies. The Cube family is `CUBEKPIMEMBER`, `CUBEMEMBER`,
`CUBEMEMBERPROPERTY`, `CUBERANKEDMEMBER`, `CUBESET`, `CUBESETCOUNT`, and
`CUBEVALUE`. This is intentionally broader than proven remote access: a
`HYPERLINK` destination may be in the current workbook, a Cube connection can
refer to an offline cube, and a function argument may be calculated. Microsoft
documents a HYPERLINK destination as a text string or cell reference in its
[link guidance](https://support.microsoft.com/en-US/Excel/work-with-links-in-excel),
while `WEBSERVICE` calls a URL,
[`IMAGE`](https://support.microsoft.com/en-us/excel/functions/image-function)
uses an HTTPS source, and
[`RTD`](https://support.microsoft.com/en-us/excel/functions/rtd-function)
uses a COM-automation data provider.
[`STOCKHISTORY`](https://support.microsoft.com/en-us/office/stockhistory-function-1ac8b5b3-5f62-4d94-8ab8-7504ec7239a8)
retrieves financial history, and the documented
[`CUBESET`](https://support.microsoft.com/en-us/excel/functions/cubeset-function)
and [`CUBEVALUE`](https://support.microsoft.com/en-us/excel/functions/cubevalue-function)
references describe stored connections and server-backed retrieval.

The public profile and `FF064` details contain only action-cell,
formula-defined-name, `STOCKHISTORY`, and aggregate Cube-function counts;
formulas, name identities, arguments, destinations, market symbols,
connections, query expressions, provider names, results, and cell locations
stay in private comparison signatures. A destination-only, connection/query,
or named-definition-only change can therefore emit `FF064` when every public
count is unchanged. FormulaFence also emits `FF064` when a normal cell change
reaches one of these formula cells through its static dependency graph,
covering a source such as `=HYPERLINK(A1, ...)` or `=STOCKHISTORY(A1, ...)`
without reading or evaluating an effective URL or symbol. Dynamic or unresolved
arguments remain parser-coverage boundaries rather than a claim that every
indirect source is tracked.
`no_formula_external_action_changes` makes this `FFP064` in CI. FormulaFence
does not calculate, resolve, fetch, open, follow, click, authenticate to,
query, or execute any function/provider. Its ordinary semantic diff
intentionally remains a review artifact with changed formulas, so it is not a
substitute for the ledger's minimised action details.

## Direct DDE-style formula links

FormulaFence separately inventories direct formula syntax that has an
application, a pipe outside quoted text, a topic, and an `!item` boundary, such as the
Excel DDE example `='Quote'|'NYSE'!ZAXX` documented in the Windows [Dynamic
Data Exchange overview](https://learn.microsoft.com/en-us/windows/win32/dataxchg/about-dynamic-data-exchange).
The scanner deliberately ignores pipes inside double-quoted string literals and
single-quoted ordinary sheet names, so `='cmd|/C calc'!A0` is not treated as a
DDE link. It inventories worksheet formulas plus formula-defined names and
named `LAMBDA` bodies, following nested, recursive, and sheet-local
definitions to invoking formula cells. It does not use the syntax to resolve a
service, topic, or item.

The public `formula_dde_links` profile object and `FF074` contain only
formula-cell, link, and defined-name counts. Private signatures make a
same-count endpoint, formula, invocation, or definition-chain change visible
without exposing it. A normal edit that statically reaches an invoking named
`LAMBDA` produces `FF074` too. Direct DDE syntax can make the generic formula
tokenizer fail; FormulaFence inventories it before that parser boundary and
retains the ordinary tokenization warning as separate coverage evidence.

`no_formula_dde_link_changes` turns `FF074` into `FFP074`. FormulaFence never
calculates a formula, resolves an endpoint, looks up or starts a DDE server,
sends a command, or infers whether a server is present, trusted, or permitted
by Excel. Excel's [DDE security settings](https://learn.microsoft.com/en-us/troubleshoot/microsoft-365-apps/excel/security-settings)
distinguish server lookup from the not-recommended server-launch option; this
policy does not attempt to reproduce either setting. Raw OOXML external-link
DDE/OLE parts are separately guarded by `no_external_link_package_changes`.

Python in Excel is a distinct executable-code boundary. Microsoft documents
that its Python runtime runs in the Microsoft Cloud, and the OOXML standard
stores code separately from the `PY` formula that references it. FormulaFence
recognizes `PY` formula spellings and inventories the documented workbook
2023 `python.xml` part plus the separately stored 2022 `pythonScripts.xml`
compatibility contract, including their relationships and content types,
directly from the package. It privately fingerprints bounded
source/environment/script XML and the stored PY formula binding, while the
public profile and `FF065` show only safe physical-package, formula-cell, call,
script, environment, initialization, and coverage counts. When both package
contracts are present, FormulaFence inventories each stored part independently
rather than assuming they agree or selecting a runtime implementation.

`FF065` is emitted for a code/environment/package change, a changed PY formula
binding, or a normal cell edit that reaches a PY formula through the static
dependency graph. This catches a source such as
`=_xlfn._xlws.PY(0,0,A1)` without decoding its script index or interpreting
`A1` as Python input. Relationship-ID-only rewrites normalize. Missing,
malformed, unbound, oversized, unreadable, or over-budget metadata remains a
coverage warning; XML reads are bounded to 16 MiB per part, 64 MiB per workbook,
and 512 parts. `no_python_in_excel_changes` makes this `FFP065` in CI.
FormulaFence does not parse or execute Python, evaluate a formula, resolve a
result, contact Microsoft Cloud, or validate runtime package availability.
Dynamic or unresolved formula inputs remain explicit coverage limits. The
ordinary semantic diff still displays changed PY formulas and values for review,
so it is not a redacted code archive. This boundary follows Microsoft's
[Python in Excel introduction](https://support.microsoft.com/en-US/Excel/python/introduction-to-python-in-excel)
and the OOXML [Python part](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/151e4bcd-90a0-4d82-8b98-f16bf273e4ff).

## Namespaced Office custom-function candidates

Office Add-in custom functions are registered from a manifest and exposed in
Excel as namespaced formulas such as `=CONTOSO.ADD(10,200)`. They can request
or stream web data, but the manifest, JavaScript/TypeScript code, and runtime
are outside a normal workbook. FormulaFence therefore keeps a private ledger
of a conservative formula candidate, not a claim that a particular add-in is
present or executable.

The direct-call classifier admits only namespaced callable tokens that are not
known native dotted Excel functions or workbook-defined names. It excludes
`_xlfn.` and `_xlws.` compatibility names. Unqualified VBA, COM, and XLL
UDF-shaped calls are handled by the separate `FF075` boundary. A candidate found
in a formula-defined name or named
`LAMBDA` body is propagated to its invoking worksheet formulas instead.
Public profiles and `FF066` contain only formula-cell, call, and namespace
counts; names, namespaces, locations, formulas, and arguments stay private. A
same-count callable or argument change is consequently visible via
a private signature, and a normal cell edit that statically reaches a candidate
call emits `FF066` too. Dynamic and unresolved formula inputs remain coverage
limits.

`no_office_custom_function_changes` turns `FF066` into `FFP066`. FormulaFence
does not evaluate a formula, resolve a candidate to an add-in, read an external
manifest, install or load an add-in, execute JavaScript, or make a network
request. The ordinary semantic diff still retains changed formulas and
formula-defined names for review, so it is not a redacted candidate ledger.
This boundary follows Microsoft's
[custom-functions overview](https://learn.microsoft.com/en-us/office/dev/add-ins/excel/custom-functions-overview),
[custom-functions tutorial](https://learn.microsoft.com/en-us/office/dev/add-ins/tutorials/excel-tutorial-create-custom-functions),
and [web-data guidance](https://learn.microsoft.com/en-us/office/dev/add-ins/excel/custom-functions-web-reqs).

## Unqualified runtime-function candidates

Excel can bind a bare unknown formula call such as `=MYUDF(A1)` to a VBA UDF,
COM/Automation add-in, XLL, or another registered runtime. The formula alone
does not prove which provider, if any, will resolve that name. FormulaFence
therefore inventories a candidate rather than loading a provider or claiming it
can execute.

The classifier accepts only a bare identifier and excludes known native Excel
functions, qualified/dotted calls, workbook-defined names, and local
`LET`/`LAMBDA` bindings. Its native-function catalogue is a stable FormulaFence
snapshot derived from Microsoft's alphabetical reference, with explicit
compatibility additions for runtime-only/native serialization spellings. This
avoids third-party parser-version drift; a new native Excel function can be a
conservative candidate until FormulaFence adds it to that catalogue. `PY` stays
under the separate Python-in-Excel boundary, namespaced Office calls stay under
`FF066`, and direct `REGISTER.ID` stays under `FF067`.

Candidates stored in formula-defined names and named `LAMBDA` bodies are
propagated through nested, recursive, and sheet-local definition chains to
their invoking worksheet formulas. Public profiles and `FF075` expose only
formula-cell, call, and relevant definition counts. Names, cells, formulas,
arguments, provider identities, and host details remain private; private
signatures preserve same-count candidate and definition changes. A normal edit
that statically reaches a candidate emits `FF075` too. Dynamic or unresolved
inputs remain coverage limits. A stored candidate definition is compared even
when no worksheet formula currently invokes it, but it does not create a
synthetic static-input path.

`no_unqualified_runtime_function_changes` turns `FF075` into `FFP075`.
FormulaFence does not evaluate a formula; resolve or load VBA, COM/Automation,
an XLL, or another registered runtime; inspect host trust settings; or execute
code. The ordinary semantic diff still retains changed formulas and definitions
for reviewer context, so it is not a redacted candidate ledger. For a shared
artifact, the separate output-only `--redact-unqualified-runtime-functions`
mode hides direct bare-call material, exact changed static inputs, and changed
resolved name-chain evidence without changing the comparison, policy, or exit
status. This boundary
follows Microsoft's [native function catalogue](https://support.microsoft.com/en-us/office/excel-functions-alphabetical-b3944572-255d-4efb-bb96-c6d90033e188),
[installed UDF guidance](https://support.microsoft.com/en-us/excel/user-defined-functions-that-are-installed-with-add-ins-reference),
[VBA custom-function guidance](https://support.microsoft.com/en-us/excel/create-custom-functions-in-excel),
and [XLL registration/call guidance](https://learn.microsoft.com/en-us/office/client-developer/excel/accessing-xll-code-in-excel).

## Worksheet code-resource registrations

Microsoft's [`REGISTER.ID` reference](https://support.microsoft.com/en-us/office/register-id-function-f8f0af0f-fd66-4704-a0f2-87b27b175b50)
documents that the function returns a DLL/code-resource registration ID and
registers the resource when necessary; unlike `REGISTER`, it can be used in a
worksheet. FormulaFence therefore treats stored `REGISTER.ID` expressions as a
separate boundary from namespaced add-in candidates.

Calls in formula-defined names and named `LAMBDA` bodies are propagated to the
worksheet formulas that invoke them. The public profile and `FF067` expose only
formula-cell, call, and relevant formula-defined-name counts. Module paths,
procedure names, type strings, formulas, arguments, cells, and name identities
stay private. Same-count call or definition changes remain visible via private
signatures; a normal cell edit that statically reaches a registration emits
`FF067` too. Dynamic and unresolved inputs remain coverage limits.

`no_worksheet_code_resource_registration_changes` turns `FF067` into `FFP067`.
FormulaFence does not evaluate a formula, resolve a module path, load a DLL or
XLL, inspect host security settings, or determine whether a registration
succeeds. The ordinary semantic diff still retains changed formulas and defined
names for general review, so it is not a redacted registration ledger. For a
shared artifact, the separate output-only
`--redact-worksheet-code-resource-registrations` mode hides direct
`REGISTER.ID` material, exact changed static inputs, and changed resolved
name-chain evidence without changing comparison, policy, or exit status. XLM
macro-sheet `CALL`/`REGISTER` program material remains within the separate raw
XLM boundary; Microsoft's [`CALL` reference](https://support.microsoft.com/en-us/office/call-function-32d58445-e646-4ffd-8d5e-b45077a5e995)
states that `CALL` is available only from an Excel macro sheet.

## Formula-defined XLM registrations

Microsoft's [`xlfRegister` Form 1 reference](https://learn.microsoft.com/en-au/office/client-developer/excel/xlfregister-form-1)
identifies `REGISTER` as the Excel XLM equivalent used to register a DLL
function or command, including macro types callable from a defined-name
definition. Its [`Form 2` reference](https://learn.microsoft.com/en-au/office/client-developer/excel/xlfregister-form-2)
documents the form that loads and activates an XLL. FormulaFence therefore
tracks stored `REGISTER` calls only while inspecting a formula-defined name or
named `LAMBDA`, then propagates that private marker through nested and
sheet-local names to an invoking worksheet formula.

The public profile and `FF068` expose only invocation-cell, call, and relevant
formula-defined-name counts. Module paths, procedure names, type strings,
formulas, arguments, cells, and name identities stay private. Same-count
definition or invocation changes remain visible through private signatures; a
normal cell edit that statically reaches an invoking formula emits `FF068` as
well. Uninvoked stored definitions still appear as a count, while dynamic or
unresolved inputs remain coverage limits.

`no_formula_defined_xlm_registration_changes` turns `FF068` into `FFP068`.
FormulaFence does not evaluate a formula, execute a macro, resolve a module
path, load a DLL/XLL, inspect host security settings, or determine whether a
registration succeeds. Direct worksheet `REGISTER` formulas and raw XLM
macro-sheet parts are deliberately outside this narrow stored-definition
boundary; the latter remain under `FF026`. The ordinary semantic diff still
retains changed formulas and definitions for local review, so it is not a
redacted registration ledger. For a shared artifact, the separate output-only
`--redact-formula-defined-xlm-registrations` mode hides direct stored
`REGISTER` material, exact changed static inputs, and changed resolved
name-chain evidence without changing comparison, policy, or exit status.

## Formula-defined XLM expression evaluation

Microsoft's [Excel expression-evaluation
reference](https://learn.microsoft.com/en-us/office/client-developer/excel/excel-worksheet-and-expression-evaluation)
identifies `EVALUATE` as an XLM function that reduces a valid character string
to a worksheet value. FormulaFence therefore inventories stored `EVALUATE`
calls only while inspecting formula-defined names and named `LAMBDA` bodies,
then propagates their private marker through nested and sheet-local names to an
invoking worksheet formula.

The public profile and `FF069` expose only invocation-cell, call, and relevant
formula-defined-name counts. Expressions, formulas, arguments, cells, and name
identities stay private. Same-count definition or invocation changes remain
visible through private signatures; a normal cell edit that statically reaches
an invoking formula emits `FF069` as well. Uninvoked stored definitions still
appear as a count.

`no_formula_defined_xlm_evaluation_changes` turns `FF069` into `FFP069`.
FormulaFence does not evaluate a formula or its text argument, parse a
runtime-generated expression, execute a macro, or infer dependencies embedded
inside that expression. It traces only the stored call's own statically visible
argument edge. Direct worksheet `EVALUATE` formulas and raw XLM macro-sheet
parts remain deliberately outside this narrow stored-definition boundary; the
latter remain under `FF026`.

The ordinary semantic diff retains local-review formula and defined-name
evidence. For a shared artifact, the separate output-only
`--redact-formula-defined-xlm-evaluations` mode hides direct stored `EVALUATE`
material, exact changed static inputs, and changed private resolved name-chain
evidence without changing comparison, policy evaluation, or exit status. It
does not evaluate or parse runtime-generated expression text, so text-only
dependencies remain an explicit coverage limit.

## Formula-defined XLM actions and event dispatch

FormulaFence inventories the deliberately finite set of legacy XLM action and
event-dispatch spellings `CALL`, `EXEC`, `EXECUTE`, `RUN`, `SEND.KEYS`,
`ON.DATA`, `ON.DOUBLECLICK`, `ON.ENTRY`, `ON.KEY`, `ON.RECALC`, `ON.SHEET`,
`ON.TIME`, and `ON.WINDOW` only while inspecting formula-defined names and
named `LAMBDA` bodies. Microsoft's [Excel C API
reference](https://learn.microsoft.com/en-us/office/client-developer/excel/programming-with-the-c-api-in-excel)
describes XLM command-equivalent functions and event traps including
`ON.ENTRY` and `ON.TIME`; its [DLL-access
guidance](https://learn.microsoft.com/en-us/office/client-developer/excel/how-to-access-dlls-in-excel)
documents `CALL` and `REGISTER` as XLM macro-sheet routes to DLL functions or
commands. This is not an attempt to interpret all XLM commands.

The public profile and `FF073` expose only invocation-cell, selected-call, and
relevant formula-defined-name counts. Function targets, handler names,
formulas, arguments, cells, and name identities stay private. Same-count
definition or invocation changes remain visible through private signatures; a
normal cell edit that statically reaches an invoking formula emits `FF073` as
well. Uninvoked stored definitions still appear as a count.

`no_formula_defined_xlm_action_changes` turns `FF073` into `FFP073`.
FormulaFence does not evaluate a formula, resolve an action target or event
handler, load a DLL, send DDE, run a macro or program, or determine whether an
action succeeds. Direct worksheet action formulas and raw XLM macro-sheet
parts remain deliberately outside this narrow stored-definition boundary; the
latter remain under `FF026`. Workbook-defined callables shadow the native
spelling rather than being classified as legacy XLM actions.

The ordinary semantic diff retains local-review formula and defined-name
evidence. For a shared artifact, the separate output-only
`--redact-formula-defined-xlm-actions` mode hides direct stored selected-action
material, exact changed static inputs, and changed private resolved name-chain
evidence without changing comparison, policy evaluation, or exit status. It
does not resolve a target or handler, execute an action, or reconstruct a
dynamically assembled target.

## Formula-defined XLM GET.CELL information

Microsoft identifies
[GET.CELL / xlfGetCell](https://learn.microsoft.com/en-us/office/client-developer/excel/programming-with-the-c-api-in-excel)
as an XLM information function. FormulaFence inventories stored GET.CELL calls
only while inspecting formula-defined names and named LAMBDA bodies, then
propagates their private marker through nested and sheet-local names to an
invoking worksheet formula.

The public profile and FF070 expose only invocation-cell, call, and relevant
formula-defined-name counts. Information types, references, formulas,
arguments, cells, and name identities stay private. Same-count definition or
invocation changes remain visible through private signatures; a normal cell
edit that statically reaches an invoking formula emits FF070 as well.
Uninvoked stored definitions still appear as a count.

The no_formula_defined_xlm_get_cell_changes rule turns FF070 into FFP070.
FormulaFence does not evaluate a formula or information call, determine an
information type, resolve a dynamic reference, render display/formatting,
inspect comments or protection, or simulate Excel state. It traces only the
stored call's statically visible argument edge. Direct worksheet GET.CELL
formulas and raw XLM macro-sheet parts remain deliberately outside this narrow
stored-definition boundary; the latter remain under FF026.

## Formula-defined XLM environment information

Microsoft's [C API
reference](https://learn.microsoft.com/en-us/office/client-developer/excel/programming-with-the-c-api-in-excel)
identifies workspace information functions such as GET.CELL and GET.WORKBOOK.
FormulaFence additionally selects GET.WORKSPACE and GET.DOCUMENT because
Microsoft documents GET.WORKSPACE returning platform information in its
[xlfFree example](https://learn.microsoft.com/en-us/office/client-developer/excel/xlfree)
and GET.DOCUMENT as an XLM information function in the [expression-evaluation
reference](https://learn.microsoft.com/en-us/office/client-developer/excel/excel-worksheet-and-expression-evaluation).
FormulaFence inventories those selected calls only while inspecting
formula-defined names and named LAMBDA bodies, then propagates their private
marker through nested and sheet-local names to an invoking worksheet formula.

The public profile and FF071 expose only invocation-cell, call, and relevant
formula-defined-name counts. Information types, references, formulas,
arguments, cells, and name identities stay private. Same-count definition or
invocation changes remain visible through private signatures; a normal cell
edit that statically reaches an invoking formula emits FF071 as well.
Uninvoked stored definitions still appear as a count.

The no_formula_defined_xlm_environment_information_changes rule turns FF071
into FFP071. FormulaFence does not evaluate a formula or information call,
determine an information type, resolve a dynamic reference, or simulate
workbook, workspace, document, client, add-in, printer, or other Excel state.
It traces only the stored call's statically visible argument edge. A
state-only workbook change is not asserted to change a stored call. Direct
worksheet calls and raw XLM macro-sheet parts remain deliberately outside this
narrow stored-definition boundary; the latter remain under FF026.

## Native CELL, INFO, SHEET, and SHEETS information

Microsoft's [CELL function
documentation](https://support.microsoft.com/en-us/office/cell-function-51bd39a5-f338-4dbe-a33f-955d67c2b2cf)
states that CELL returns formatting, location, or content information, and that
an omitted optional reference uses the selected cell at calculation time.
Microsoft's [INFO function
documentation](https://support.microsoft.com/en-au/office/info-function-725f259a-0e4b-49b3-8b52-58815c69acae)
lists operating-environment information such as directory, calculation mode,
platform, and workbook-count values. Microsoft's [SHEET function
documentation](https://support.microsoft.com/en-us/excel/functions/sheet-function)
documents sheet-number behavior, including its optional value, while the
[SHEETS function documentation](https://support.microsoft.com/en-us/excel/functions/sheets-function)
documents that an omitted reference counts the sheets in the containing
workbook. Both include hidden, very-hidden, macro, chart, and dialog sheets.
FormulaFence inventories native CELL, INFO, SHEET, and SHEETS calls in worksheet
formulas, formula-defined names, and named LAMBDA bodies, then propagates
private signals through nested and sheet-local names to an invoking formula.

The public profile and FF072 expose only formula-cell, call, relevant
formula-defined-name, omitted-CELL-reference, SHEET, SHEETS, and omitted-
SHEETS-reference counts. Information types, references, formulas, arguments,
cells, name identities, and raw tab-catalog comparison material stay private;
ordinary sheet inventory remains normal reviewer context. Same-count definition
or invocation changes remain visible through private signatures; a normal cell
edit that statically reaches an invoking formula emits FF072 as well. Uninvoked
stored definitions still appear as a count.

The no_formula_environment_information_changes rule turns FF072 into FFP072.
FormulaFence does not evaluate a formula or information call, determine an
information type, resolve dynamic references or arguments, infer the selected
cell, inspect a file/folder/client/workspace state, or simulate any of those
states. It traces only stored, ordinary static argument edges. When it can read
the raw OOXML tab catalog completely, it privately compares its all-tab member
order for SHEET calls and SHEETS calls whose reference is omitted. An addition,
removal, reorder, or tab-name change emits FF072 because Excel may then
calculate from different workbook-structure information. It does not calculate
an individual result, resolve an explicit SHEET/SHEETS argument, or infer
whether a non-omitted SHEETS reference is one-sheet or 3-D; visibility-only
changes are not a tab-catalog condition because Excel includes hidden tabs.
Incomplete raw tab metadata remains a parser coverage warning.

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

FormulaFence separately inventories **XLM automatic-macro bindings**. Excel's
backward-compatible workbook automatic-macro API names four events:
`Auto_Open`, `Auto_Close`, `Auto_Activate`, and `Auto_Deactivate`. This is a
different control plane from macro-sheet program XML: an unchanged XLM sheet can
be added to, removed from, or retargeted in that automatic dispatch path through
a special workbook defined name. FormulaFence reads raw `xl/workbook.xml`,
normalizes the optional SpreadsheetML `_xlnm.` built-in-name prefix, and counts
only a workbook-scoped special name whose stored definition is a direct,
internal single-cell A1 reference to a sheet declared through a raw
`xlMacrosheet` or `xlIntlMacrosheet` relationship. It deliberately excludes
sheet-local names,
ordinary-sheet targets, external references, and non-direct/dynamic name
formulas rather than guessing at a target.

A material add, removal, or same-count target/definition change emits `FF076`;
enable `no_xlm_automatic_macro_binding_changes` for `FFP076` in CI. The public
profile and `FF076`/`FFP076` details expose only total and per-event counts.
Name spellings, target cells, and stored definitions stay in a private
signature. The ordinary defined-name diff remains intentionally readable and
is not redacted by this specialized ledger. FormulaFence does not evaluate or
resolve a defined name, inspect the reserved/unused `definedName@xlm`
attribute, parse or execute an XLM command, infer macro-security/trust settings,
or claim that a binding will run. This scope follows Microsoft's
[RunAutoMacros documentation](https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.runautomacros),
[automatic-macro enumeration](https://learn.microsoft.com/en-us/office/vba/api/excel.xlrunautomacro),
and [defined-name compatibility notes](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oi29500/16c3c118-f358-493d-a99f-4c85ca834c00).

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

Office Web Add-ins can bind a document through a task pane, a worksheet's
documented `x15:webExtensions` extension, or an in-content DrawingML
`we:webextensionref` frame while remaining outside ordinary cell values, the
VBA payload, and RibbonX. FormulaFence follows the bounded workbook-to-
`taskpanes.xml`-to-`webextension*.xml` chain, validates worksheet `appRef`
bindings against definition bindings, and follows direct worksheet-DrawingML
web-extension references in the active `mc:Choice` branch. It fingerprints
task-pane configuration, add-in references, auto-show properties, bindings,
snapshots, active frame XML/placement, and direct relationship semantics while
normalizing writer-chosen relationship IDs and equivalent internal target
spellings. The static native-picture fallback of an in-content frame remains
under the native worksheet-image boundary. Profiles expose only safe counts:
parts, task panes, worksheet bindings, in-content references, snapshots, and
relationships. Add-in IDs, store references, property values, binding values,
worksheet formulas, frame XML, snapshot data, and relationship targets never
enter a profile or diff. A material change emits `FF028`; enable
`no_office_web_addin_changes` to make it `FFP028` in CI. FormulaFence does not
install, load, execute, or fetch an add-in or manifest, and it never follows
an external relationship. Missing, oversized, malformed, unbound, or
over-budget parts remain coverage warnings. Task-pane and web-extension XML
reads are bounded to 16 MiB per part, 32 MiB per workbook, and 64 parts;
worksheet-binding and in-content DrawingML scans are each bounded to 16 MiB
per part, 64 MiB per workbook, and 512 parts. Other unrecognized extension or
graphic-frame forms remain outside this boundary.

DrawingML charts can change a report's series, axis, title, formatting, cached
values, or overlay annotations without changing an ordinary worksheet cell.
FormulaFence follows standard worksheet/chartsheet drawing relationships through
legacy `c:chart` parts and Office 2016+ `cx:chart` ChartEx parts, including
ChartEx frames stored inside `mc:AlternateContent` with an older-client
fallback. It privately compares legacy chart definition material separately
from `numCache`, `strCache`, and `multiLvlStrCache` material; it fingerprints
ChartEx XML and bounded direct ChartEx style, colour-style, drawing, image,
theme-override, and embedded-package payloads. It also compares legacy overlay
XML, relationship semantics, and bounded direct related payload hashes.
Profiles expose safe structural counts only. Formulas, labels, cached values,
formatting, overlay text, target paths, XML, and payload bytes never enter a
profile or diff. Writer-chosen relationship IDs and equivalent internal target
spellings are normalized. A material change emits `FF030`; enable
`no_chart_definition_changes` to make it `FFP030` in CI. FormulaFence does not
calculate a series, map chart references into downstream cell impact, render a
chart, assess visual output, follow external targets, parse media/package
formats, resolve ChartEx second-hop relationships, or interpret ChartEx-specific
visualization semantics. XML reads are bounded to 16 MiB per part, 64 MiB per
workbook, and 512 parts; direct payload hashes are bounded to 32 MiB per part,
64 MiB per workbook, and 512 parts. Missing, malformed, orphaned, unbound,
unsupported, oversized, or over-budget chart material remains a visible
coverage warning. The boundary follows Microsoft's [ChartEx part](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-odrawxml/5d0d453e-adac-43be-a797-59b9916593dd)
and [ChartEx relationship-ID](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-odrawxml/d8ede39e-a36c-48ad-8a17-0086a2d0889b)
definitions.

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
positive dimensions outside its dedicated worksheet-dimension boundary or
styles, or model overflow or outline-display settings.

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

Legacy Excel Custom Views persist a named alternate workbook display/print mode
in `<customWorkbookView>` and a GUID-linked `<customSheetView>` on each
worksheet, dialog sheet, or chart sheet. That alternate state can hide
rows/columns, preserve a filter, alter print settings, panes, gridlines,
formula display, comments, or object visibility while cells and the active
worksheet view remain unchanged. FormulaFence reads and reconciles the raw
workbook/sheet declarations privately. A material change emits `FF060`; enable
`no_custom_workbook_view_changes` to make it `FFP060` in CI.

Profiles and `FF060` details expose only structural counts: workbook views,
per-sheet views, sheets with alternate state, per-sheet
hidden/filter/print/display settings, and unrecognized metadata. Names, GUIDs,
bindings, ranges, filter
criteria, pane locations, print settings, and raw XML remain private.
Coordinated GUID and sheet-ID/active-sheet-ID rewrites plus equivalent
Boolean/default and unsigned-integer spellings normalize. Transitional and
Strict SpreadsheetML worksheets/dialog sheets and chart sheets are supported.
Missing, duplicate, malformed, unsupported, unsafe, oversized, over-budget, or
incompletely linked declarations are explicit coverage warnings. FormulaFence
does not activate/render a Custom View, calculate an alternate filtered result,
determine final print output, interpret future extensions, or support Custom
Views on other sheet types.

Excel Tables can carry a presentation declaration independently of the table
reference and cell contents. `tableStyleInfo` selects a built-in or custom
style and turns headers, totals, row/column banding, or first/last-column
emphasis on or off. Custom `<tableStyle>` records resolve their Dxf material
through `styles.xml`; Tables and TableColumns can also directly select a Dxf
for data/header/total/border regions or a named cell style. FormulaFence reads
those declarations before the ordinary reader flattens them. A material change
emits `FF061`; `no_table_style_control_changes` makes it `FFP061` in CI.

Profiles and `FF061` details expose only structural counts: declarations,
styled tables, custom styles/elements, direct Dxf/named-cell-style assignments,
banding/emphasis, and unrecognized metadata. Table names, style names,
formatting, colours, IDs, and raw XML remain private. Boolean/default spelling,
case-only style names, Excel's `xr9:uid` style revision provenance, and
coordinated Dxf reordering/ID rewriting normalize. Transitional and Strict
SpreadsheetML are supported. Missing, duplicate, malformed, unresolved,
unsupported, oversized, or over-budget controls are explicit coverage warnings.
FormulaFence does not render the resulting Table, resolve themes, apply
conditional formatting, or cover PivotTable-only style regions.
`defaultTableStyle` is a new-table preference rather than an existing-table
binding, and a same-name named cell-style definition is not resolved under this
Table Style rule.

Legacy shared-workbook revision history can persist outside ordinary worksheet
cells in relationship-backed `revisionHeaders` and `revisionLog` parts. The
records can retain prior/new values, locations, authors, timestamps, comments,
formatting edits, conflict-resolution material, and shared/tracking/retention/
protection controls. FormulaFence fingerprints complete bounded declarations
privately, follows workbook-to-header and header-to-log relationships, and
emits `FF062`; `no_shared_workbook_revision_changes` makes it `FFP062` in CI.

Profiles and `FF062` details expose only header/log part and record counts,
aggregate shared/tracked/history-retention/protection state, and coverage
counts. Historic values, locations, identities, timestamps, comments, GUIDs,
relationship IDs, and raw XML remain private. Equivalent Boolean/integer
spelling, coordinated relationship-ID changes, and transitional/Strict
relationship type spelling normalize. Missing, duplicate, malformed, unsafe,
unsupported, oversized, or over-budget declarations are explicit coverage
warnings. FormulaFence does not apply revisions, reconstruct historical
workbook state, resolve conflicts, validate identity/timestamp claims, render
Excel, or interpret arbitrary future extensions.

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
other display controls, track borders/rich-text run rendering or separately
inventoried Table Style controls, or claim arbitrary visual-style coverage. A raw column `style` is
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
contrast, evaluate conditional-format differential styles, apply separately
inventoried Table Style controls, calculate values, or claim arbitrary visual-style coverage. A raw column `style` is compared as a declaration/default for unallocated/new cells; it is not
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

Cell-border controls can redraw a report edge, total, exception box, or warning
without an ordinary cell edit. FormulaFence reads raw transitional and strict
SpreadsheetML `<borders>/<border>` definitions, base `cellStyleXfs`, effective
`cellXfs` records with `borderId`, `xfId`, and `applyBorder`, direct cell `s`,
`customFormat=1` row `s`, and raw `<cols>/<col style>` declarations. It covers
left/right/top/bottom, Office 2010 logical start/end, diagonal/direction,
outline, stored line styles, and stored colours. A material effective control
change emits `FF057`; enable `no_cell_border_changes` to make it `FFP057` in
CI.

Profiles and `FF057` details expose only counts for default definitions, direct
cell, row, effective column, and unrecognized controls. Border definitions,
colours, style indexes, and targets remain private. Omitted/`none` sides,
Boolean/colour spelling, unused diagonal payload, ineffective empty
`outline="false"`, base-XF inheritance, `applyBorder`, and equivalent
column-range splitting are normalized. Missing, duplicate, malformed, or
unsupported material is an explicit coverage warning. Material `vertical` or
`horizontal` inner sides under ordinary cell styles are likewise coverage
warnings because their differential-format semantics are not modeled here.
FormulaFence compares declarations, not final rendering: it does not resolve
theme/palette colours, choose adjacent-cell precedence, apply
conditional-format/table/differential-style borders, calculate print output,
or infer Excel client behavior. A raw column `style` is compared as a
declaration/default for unallocated/new cells; it is not treated as a renderer
that restyles allocated cells. This boundary follows OOXML's
[`border`](https://c-rex.net/samples/ooxml/e1/part4/OOXML_P4_DOCX_border_topic_ID0EVV35.html)
and [`xf`](https://c-rex.net/samples/ooxml/e1/part4/OOXML_P4_DOCX_xf_topic_ID0E13S6.html)
forms, plus Microsoft's [cell-border guidance](https://support.microsoft.com/en-us/Excel/apply-or-remove-cell-borders-on-a-worksheet).

Worksheet dimensions can change the usable report surface while leaving cells
unchanged: a fixed positive row height can cut off wrapped text, a column width
can reframe visible fields, and a size change can move automatic page breaks.
FormulaFence reads raw transitional and strict SpreadsheetML `sheetFormatPr`
defaults (`defaultRowHeight`, `defaultColWidth`, and `baseColWidth`), explicit
default `customHeight`, Office 2010 `x14ac:dyDescent` baseline adjustments, and
active `thickTop`/`thickBottom` automatic-height adjustments. It also reads
direct row `ht`, `customHeight`, `x14ac:dyDescent`, `thickTop`, and `thickBot`
declarations, plus raw `<cols>/<col>` positive `width` and
`bestFit` state. Overlapping columns resolve in file order: a later declaration
overrides only a width or AutoFit attribute it actually supplies. A material
change emits `FF058`; enable `no_worksheet_dimension_changes` to make it
`FFP058` in CI.

Profiles and `FF058` details expose counts only: default row/column controls,
baseline/automatic border-adjustment sheets, direct row heights/baseline/border
adjustments, effective positive-width and best-fit columns, and unrecognized
controls. Sheet names, dimension values, row/column targets, raw XML, and
writer hints remain private.
Decimal and Boolean spelling, ordinary baseline defaults, inert thick-border
flags under a fixed custom height, inert `customWidth`, and equivalent
effective-column range splitting normalize away. Zero/hidden dimensions remain
in the `FF036` visibility boundary; a change from zero to a positive width can
therefore trigger both controls. Malformed, duplicate, unsupported, or
budget-exhausted metadata is a coverage warning. FormulaFence compares stored
declarations rather than rendering a workbook: it does not calculate final
AutoFit sizes, text overflow, merged-cell layout, automatic page geometry, or
client-specific display. This boundary follows Microsoft's
[row-height and column-width guidance](https://support.microsoft.com/en-us/excel/change-the-column-width-and-row-height),
[wrapped-text guidance](https://support.microsoft.com/en-us/excel/wrap-text-in-a-cell-in-excel),
and [automatic page-break guidance](https://support.microsoft.com/en-US/Excel/insert-move-or-delete-page-breaks-in-a-worksheet),
plus OOXML's [`sheetFormatPr`](https://c-rex.net/samples/ooxml/e1/part4/OOXML_P4_DOCX_sheetFormatPr_topic_ID0EVAG5.html),
[`row`](https://c-rex.net/samples/ooxml/e1/part4/OOXML_P4_DOCX_row_topic_ID0EIKD5.html),
and [`col`](https://c-rex.net/samples/ooxml/e1/part4/OOXML_P4_DOCX_col_topic_ID0ELFQ4.html)
definitions, plus Microsoft's [`dyDescent` extension documentation](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/f11dfda4-46de-4035-8418-d76b0d3898f1).

Worksheet-display controls can change a reviewer's saved surface without a
formula or value edit. FormulaFence reads raw transitional and strict
SpreadsheetML `sheetViews/sheetView` declarations and compares non-default
`showZeros`, `showFormulas`, `showGridLines`, custom gridline-colour
(`defaultGridColor`/`colorId`), `showRowColHeaders`, `showOutlineSymbols`,
`showRuler`, `showWhiteSpace`, and `rightToLeft` flags; non-normal view modes;
and material split/frozen pane state. A material declaration change emits
`FF055`; enable
`no_worksheet_display_control_changes` to make it `FFP055`
in CI.

Profiles and `FF055` details expose only structural counts for
hidden-zero, formula, gridline/gridline-colour, header, outline, ruler,
page-whitespace, direction, view-mode, pane, and malformed controls. Sheet
names, target cells, pane positions, and raw XML remain private.
Omitted/default controls, Boolean and active custom-gridline-colour spellings,
and finite non-negative pane-split decimal spellings are normalized. Active-cell,
selection, top-left navigation, and zoom state are intentionally excluded so
ordinary navigation does not become a policy event. Missing, duplicate,
malformed, or unsupported material is an explicit coverage warning rather than
a silent omission. FormulaFence compares declarations, not rendering: it does
not resolve the effective palette colour, calculate viewport geometry or final
visibility, interpret extension-specific views, or infer Excel client state.
The boundary
follows the Open XML SDK [`SheetView` schema
surface](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.sheetview?view=openxml-3.0.1)
and Microsoft’s [worksheet display
guidance](https://learn.microsoft.com/en-us/office/dev/add-ins/excel/excel-add-ins-worksheet-display).

Worksheet print-layout controls can change what gets printed even when every
ordinary cell and formula stays fixed. FormulaFence reads raw transitional and
strict SpreadsheetML workbook `_xlnm.Print_Area` / `_xlnm.Print_Titles`
defined names and direct worksheet `printOptions`, `pageMargins`, `pageSetup`,
`sheetPr/pageSetUpPr`, `headerFooter`, `rowBreaks`, and `colBreaks`
declarations. A material declaration change emits `FF056`; enable
`no_worksheet_print_layout_changes` to make it `FFP056` in CI.

Profiles and `FF056` details expose only counts for print areas/titles,
gridlines/headings/centering, margins, page setup, headers/footers, manual
row/column breaks, and unrecognized controls. Print ranges, header/footer
text, page values, printer-setting relationship IDs, and raw XML remain
private. Omitted/default, Boolean, integer, and decimal spellings are
normalized; so are the documented semantic no-ops where both print-gridline
flags must be true, inactive first/even header-footer sections are ignored,
`firstPageNumber` is disabled, scale is overridden by fit-to-page dimensions,
or an automatic page-break display declaration changes. Missing, duplicate,
malformed, or unsupported material is an explicit coverage warning rather than
a silent omission. FormulaFence compares stored declarations, not rendered
paper: it does not calculate page geometry/counts or automatic pagination,
resolve printer/client defaults or printer-specific `devMode` settings, or
cover custom/legacy sheet-view and extension print controls. This boundary
follows Microsoft's [print-area guidance](https://support.microsoft.com/en-us/excel/set-or-clear-a-print-area-on-a-worksheet)
and [`PageLayout` control surface](https://learn.microsoft.com/en-us/javascript/api/excel/excel.pagelayout?view=excel-js-preview).

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
evaluate a `HYPERLINK()` formula. Formula calls themselves are separately
covered by `FF064`, still without evaluating an argument or following its
result. This raw worksheet-hyperlink scope follows the Open XML
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

Non-chart Worksheet DrawingML can host a regular `xdr:sp` shape, a connector
`xdr:cxnSp`, a nested `xdr:grpSp` group, or a SmartArt `xdr:graphicFrame`
under a standard worksheet drawing anchor. Regular shapes can carry visible
text, run formatting, positioning, macro assignments, `textlink` formulas,
and click/hover hyperlink relationships outside cells. A connector can add or
alter a visual process link by changing its geometry, style, anchor, or
`stCxn`/`endCxn` endpoint attachment.

For a non-chart graphic frame, FormulaFence recognizes the DrawingML Diagram
`a:graphicData` URI and requires exactly one `dgm:relIds` declaration. It
privately fingerprints supported anchor, shape, connector, group, and frame
XML; the four explicitly bound SmartArt components—data (`r:dm`), layout
(`r:lo`), quick style (`r:qs`), and colours (`r:cs`)—and direct
worksheet-drawing `diagramDrawing` rendering parts; and bounded direct
internal Image targets from a Diagram Data part; connector endpoint target
semantics; and referenced relationships. Transitional and Strict DrawingML are
supported. It reports only structural worksheet/drawing/anchor,
shape/text/connector/group, graphic-frame/SmartArt-component, Diagram Data
image part/fingerprinted/uninspected, connector-attachment, text paragraph/run,
macro/text-link/hyperlink, relationship, and malformed-control counts. A
material change emits `FF044`; enable `no_worksheet_drawing_shape_changes` to
make it `FFP044` in CI.

Text, presentation details, geometry, anchors, descriptions, diagram content
and IDs, connector target IDs, macro names, text-link formulas, relationship
IDs/targets, image names/bytes, and raw XML stay private. Consistent non-visual
and connector-endpoint ID rewrites, worksheet-DrawingML relationship-ID
rewrites, and equivalent colour case are normalized. Missing, duplicate,
malformed, unsafe, unreadable, oversized, over-budget, unsupported, or external
Diagram Data image material is visible coverage evidence; XML reads are bounded
to 16 MiB per part, 64 MiB per workbook, and 512 parts, while direct Diagram
Data image hashing is bounded to 32 MiB per image, 64 MiB per workbook, and
512 images. FormulaFence does not render or assess visibility, resolve theme
colours/contrast, calculate a text link, execute a macro assignment, retrieve a
target, calculate final SmartArt layout, or decode/render media. It hashes only
bounded direct internal Image targets from a Diagram Data part and does not
follow hyperlinks, second-hop targets, or component-side relationships from
other SmartArt parts; those edges remain coverage gaps. Chart graphic frames
remain under `FF030`, native `xdr:pic` objects remain in the separate image
rule, and other non-chart graphic-frame URI types are coverage gaps. The
boundary follows
Open XML's [`xdr:cxnSp` ConnectionShape definition](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.spreadsheet.connectionshape?view=openxml-3.0.1),
Microsoft's [Graphic Object Data](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oe376/f58e82a5-5590-4e36-b178-e12989960415),
the OOXML [Diagram Data Part](https://ooxml.info/docs/14/14.2/14.2.4/),
and [Diagram relationship IDs](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.diagrams.relationshipids?view=openxml-3.0.1)
references.

Native worksheet images can alter a spreadsheet's review and print surface even
when every stored cell is unchanged. FormulaFence follows a worksheet's
`drawing`, direct `<picture>`, and `legacyDrawingHF` relationships to inspect
anchored transitional/strict DrawingML `xdr:pic` objects (including grouped
pictures), worksheet background images, and VML-backed header/footer watermark
images. It compares private anchor/picture/VML declarations, relationship
semantics, and bounded direct image-payload hashes. A material change emits
`FF059`; enable `no_worksheet_image_changes` to make it `FFP059` in CI.

Profiles and `FF059` details expose only worksheet, picture/anchor, background,
header/footer, image-payload, relationship, and malformed-control counts. Image
bytes, image names/descriptions, visual formatting, anchors, relationship
IDs/targets, and raw XML remain private. Non-visual DrawingML/VML IDs and
consistent relationship-ID rewrites normalize. Missing, duplicate, malformed,
unsafe, unreadable, oversized, or over-budget material is a visible coverage
warning. XML reads are bounded to 16 MiB per part, 64 MiB per workbook, and 512
parts; direct image-payload hashing is bounded to 32 MiB per part, 64 MiB per
workbook, and 512 parts.

FormulaFence does not render or decode media, fetch external targets, resolve
themes, calculate visibility/cropping/z-order, compose controls, or calculate
final pagination. Charts remain under `FF030`, rich-data/in-cell images under
`FF051`, Theme images under `FF053`, ActiveX/OLE image controls under `FF029`,
regular/group/connector/SmartArt drawing controls under `FF044`, and
header/footer text under `FF056`.

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

## Portfolio policies

`formulafence portfolio BASELINE_DIRECTORY CANDIDATE_DIRECTORY --policy
formulafence.yml` applies the same YAML policy independently to each pair of
supported workbook files at the same relative path. `protected_cells`,
`allowed_changes`, and formula/impact limits are therefore per workbook, not a
portfolio-wide aggregate. This supports a repository of structurally similar
models; path-specific policy routing is deliberately not guessed.

An added or removed relative path emits `FF077` even without a policy. Enable
`no_portfolio_membership_changes` to convert it to `FFP077`. A supported file
that cannot be inspected emits redacted `FF078` evidence and makes the portfolio
command return `2`, because no semantic comparison can safely be claimed for
that entry. If that file is also newly added or removed, its known `FF077` /
`FFP077` membership evidence remains present. The report still records the
remaining files. Office `~$` lock
files are ignored; legacy `.xls`, `.xlsb`, templates, add-ins, and `.ods` files
cause an explicit unsupported-format error rather than being omitted.

Candidate-only portfolio analysis also builds a separate static dependency graph
across a deliberately narrow subset of external A1 formulas and 3-D A1 spans,
book-only external structured-table selectors, direct workbook-scoped or
sheet-local external names, and package-indexed external A1/name/table forms.
An A1 link such as `=[Inputs.xlsx]Data!B2`, a 3-D span
such as `=[Inputs.xlsx]Jan:Mar!B2`, a workbook-scoped name link such as
`=[Inputs.xlsx]InputRange`, a sheet-local name link such as
`=[Inputs.xlsx]Data!LocalInput`, or Excel package links such as
`=[1]Data!B2`, `=[1]Jan:Mar!B2`, `=[1]!InputRange`, and
`=[1]Data!LocalInput`, plus selector-bearing table links such as
`='..\\inputs\\source.xlsx'!Sales[Amount]` and `=[1]!Sales[#Data]`, are
eligible only when their workbook spelling is an exact relative path from the consuming
workbook to another already-inspected candidate path. For the package forms,
`[1]` is not a filename: FormulaFence requires it to select one document-order
`externalReference`, one declared `externalLink` part with exactly one
`externalBook`, and one external `externalLinkPath` relationship before
treating that private target as the workbook spelling. A workbook-scoped
consumer alias may terminate in one exact indexed static spelling, or one exact
direct A1, workbook-scoped-name, sheet-local, or selector-bearing table
spelling. It may reach that terminal through a finite, acyclic chain whose
intermediate definitions are exactly one unqualified, non-A1 name identity
(with or without a leading `=`). An eligible workbook-scoped formula-defined
name can also retain static endpoint inputs, such as `=SUM(ExternalInput)`,
`=SUM('..\\inputs\\[Inputs.xlsx]Data'!$B$2:$B$4)`, or a named
`LAMBDA` such as `=LAMBDA(value,SUM(value,Inputs!$B$2,ExternalInput))`.
This is input-edge extraction, not calculation: each external token must be an
already parsed direct or package-validated endpoint; every remaining reference
must be static; and the definition must have no broken/unresolved/tokenizer-
failed, dynamic, relative, recursive, local-3-D, spill, or explicit-
intersection form. Each eligible definition carries both its endpoint models
and fixed internal inputs to a caller. A named LAMBDA does so only at a function
call, never through a bare name. Eligible global formula names and named
LAMBDAs may call each other. A same-named sheet-local consumer definition
shadows the workbook alias. A source name must expand
completely to static internal A1 destinations in the source candidate; a
sheet-local spelling selects only the exact source sheet's local-name scope,
with no global or other-sheet fallback. This may include a safely resolved
formula-defined source alias, but never evaluates the name. Backslash and slash
relative forms are normalized only in memory; path resolution never touches the
filesystem. Case is matched in the same case-insensitive way Excel uses for
workbook and sheet names. Direct A1 cells, ranges, whole rows, whole columns,
safe name destinations, and resolved table selector bounds remain lazy edges,
so a range is not expanded into millions of cells. A table is resolved only
when the source candidate has exactly one case-insensitive table-name match;
its `#All`, `#Data`, `#Headers`, `#Totals`, column, or contiguous-column
selector becomes fixed source-cell bounds. A 3-D span is expanded only after the exact source candidate
exposes a complete raw OOXML tab catalog consistent with its inspected
ordinary-worksheet order; both endpoints must exist uniquely and in forward
order, and every included worksheet receives the same static A1 bounds.

Absolute, UNC, URI, or portfolio-escaping paths; malformed or ambiguous
package declarations; DDE/OLE/non-workbook package links; package A1 forms
that are not one static destination; bare or source-sheet-qualified table
forms, `@`/`#This Row`, unsupported selectors, missing/colliding source tables;
sheet-scoped consumer aliases; consumer formula-defined bridges with broken,
unresolved, tokenizer-failed, dynamic, relative, local-3-D, spilled, or
explicit-intersection semantics; missing bridges, or cyclic exact aliases;
external-link cache values; missing, dynamic, relative, cyclic, external, 3-D,
malformed, otherwise non-statically-expanded, unknown-scope, or
wrong-scope source names; unknown source sheets;
3-D spans with missing, reversed, non-worksheet, or incomplete/inconsistent-tab-
catalog endpoints; unreadable targets; and basename-only near matches are not
resolved or guessed.
FormulaFence never opens, downloads, calculates, refreshes, trusts a cache, or
otherwise follows an external link.
When a changed source cell reaches a formula in another candidate workbook,
`FF079` supplies only reviewed relative workbook identities, logical Excel
cells, counts, and bounded shortest-path samples; stored external paths,
source-name identities, table identities, and selectors stay private in
portfolio evidence. Ordinary source and consumer defined-name declarations
remain normal review context. A name
declaration change remains its ordinary defined-name review event rather than
an `FF079` source root. Enable
`no_cross_workbook_impacts` to make each such item `FFP079`.

The graph has one global `--max-link-impact` bound for source-to-node traversal
states (100,000 by default). Reaching it emits critical `FF080` with the bound,
makes the portfolio report incomplete, and returns status `2`; any partial
`FF079` evidence remains visible. This is separate from the per-workbook
`max_downstream_impact` policy rule.

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
tokenizer-unsupported definitions remain unresolved at their call site, except
for the narrowly eligible global named-LAMBDA external-endpoint bridge above.
That bridge retains only fully static endpoint and internal-input edges at a
real function call, never through a bare name; `no_new_unresolved_references`
can make newly introduced remaining instances a hard CI failure.

Static internal 3-D A1 references such as `Jan:Mar!B2:B10` are expanded across
every tab between their endpoints in workbook order. FormulaFence records the
cells that use them in a profile. If sheet insertion, removal, or movement
changes the resolved span while the formula text remains the same, it emits
`FF014`; `no_3d_reference_scope_changes` turns that condition into `FFP014`.
Exact static external 3-D A1 spans are separately eligible for the
candidate-only portfolio graph only with an exact source candidate and complete
consistent tab catalog. Other external 3-D forms, malformed or endpoint-missing
spans, and non-A1 forms stay visible as unresolved coverage rather than being
assigned to a synthetic sheet.

## Exit status

| Status | Meaning |
| --- | --- |
| `0` | No policy violation (and no selected `--fail-on` threshold reached). |
| `1` | One or more policy violations, or `--fail-on` was reached. |
| `2` | Invalid policy, unreadable workbook (including a portfolio entry), unsupported format, incomplete portfolio scan, or output error. |

## Suggested rollout

1. Start with `formulafence diff approved.xlsx candidate.xlsx --format markdown`.
2. Commit a policy that only protects headline outputs and bans new broken/external links.
3. Run `check` in non-blocking CI for a few review cycles and tune `max_*` limits.
4. Make the check required once the report matches the team's real review process.

Do not use an allow-list as a substitute for reviewing a material change. It is
most useful for separating designated input blocks from calculation and output
areas.
