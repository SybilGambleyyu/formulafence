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
or model PivotTable layout semantics.

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
external target, or interpret OLAP, extension-list, or slicer semantics.
Missing, malformed, orphaned, unbound, oversized, or over-budget material
remains a visible coverage warning. PivotTable/cache-definition XML reads are
bounded to 16 MiB per part, 64 MiB per workbook, and 512 parts; raw cache-record
hashing is bounded to 32 MiB per part, 64 MiB per workbook, and 512 parts. A
temporary reader copy detaches cache-record relationships before the underlying
workbook library loads cells, so those raw records are not eagerly materialized;
the original workbook remains unchanged.

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
