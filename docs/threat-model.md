# Scope and threat model

FormulaFence is a static change-assurance layer. It answers whether an Excel
workbook's inspectable structure changed in a risky way; it does not certify
financial correctness or replace model review.

## Safety properties

- Workbook content stays on the machine running FormulaFence. The CLI makes no
  network requests.
- It loads formulas as text with `data_only=False`; it does not calculate them.
- It never executes VBA, XLM macro sheets, RibbonX callbacks, DDE, external
  links, Power Query, Office Web Add-in code, or worksheet ActiveX/OLE code.
- VBA payloads, XLM macro-sheet source material, RibbonX control/callback
  material, Office Web Add-in task-pane material, and worksheet control/OLE
  material are compared through private fingerprints only.
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
  refresh data, establish source trust, or model PivotTable layout semantics.
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
  Add-in task-pane packages, relationship-backed worksheet ActiveX/form-control/
  legacy-VML/OLE chains, the protection controls above, external-data refresh controls,
  external-link packages, and private Power Query definition material. It does
  not yet diff chart definitions, PivotTable layout or cached data, Ribbon image
  payloads, VML/drawing control layout or comment content, embedded OLE/package formats,
  worksheet-scoped Web Add-in markup, Power Query runtime behavior or returned
  data, ordinary styles beyond direct protection assignments, complete Excel
  style-cascade results, or every OOXML part.
- The tool preserves Excel formula text and uses a limited A1-reference
  normalizer for peer-pattern detection; it is not an Excel-compatible parser
  or calculation engine.

For high-stakes use, treat FormulaFence as one control among independent review,
recalculation in the approved spreadsheet engine, input reconciliation, and an
appropriately qualified model owner.
