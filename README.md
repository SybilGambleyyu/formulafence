# FormulaFence

FormulaFence is a local-first spreadsheet change-assurance CLI. It makes `.xlsx`
changes reviewable in CI: compare workbook semantics, trace downstream formula
impact, detect high-risk edits, and enforce a small policy file before a model
is shared or merged.

It never executes formulas or macros, and it does not upload workbook contents.

> Status: early alpha. The first release supports `.xlsx` and `.xlsm` inspection
> with formula-aware diffs, dependency impact, policy checks, Markdown/JSON/SARIF
> reports, and deterministic evidence metadata.

## Why

Spreadsheets remain the operating surface for financial models, planning,
operations, and research. Git sees an Excel workbook as a binary blob; ordinary
file diffs do not answer the questions a reviewer actually has:

- Did a formula become a hard-coded number?
- Which outputs now depend on the changed cell?
- Did a formula stop following the pattern used by its peers?
- Was a hidden sheet, macro payload, external link, or calculation setting changed?
- Does the change comply with the model's review policy?

FormulaFence is intentionally a guardrail, not a spreadsheet calculation engine
or hosted document system. It produces evidence a team can inspect in its normal
Git and CI workflow.

## Design basis

Spreadsheet controls already have useful specialist tools: [Git XL](https://www.xltrail.com/git-xl)
makes workbook diffs readable in Git, while [ExceLint](https://www.microsoft.com/en-us/research/publication/excelint-automatically-finding-spreadsheet-formula-errors/)
demonstrated static detection of formula anomalies. FormulaFence concentrates on
the merge boundary between those approaches: a local semantic diff coupled to
downstream impact and a review policy that can fail CI. It is complementary to,
not a replacement for, source control, model audit, or recalculation in Excel.

## Quick start

```bash
# Install the pinned public release directly from GitHub.
python -m pip install https://github.com/SybilGambleyyu/formulafence/releases/download/v0.24.0/formulafence-0.24.0-py3-none-any.whl

# Readable review report
formulafence diff baseline.xlsx candidate.xlsx --format markdown

# Enforce a policy in CI (non-zero when a rule fails)
formulafence check baseline.xlsx candidate.xlsx --policy formulafence.yml --format sarif --output results.sarif
```

FormulaFence is not yet published to PyPI; the direct release URL above avoids
an ambiguous package-name install. See [GitHub Releases](https://github.com/SybilGambleyyu/formulafence/releases)
for the current version.

Create `formulafence.yml`:

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
  max_changed_formulas: 20
  max_downstream_impact: 100

# Cells whose contents must not change without an explicit policy edit.
protected_cells:
  - Dashboard!B12
  - Dashboard!B18

# Optional: limit ordinary cell edits to designated input areas.
allowed_changes:
  - Inputs!B2:B100
```

## What the first release checks

| Capability | What it catches |
| --- | --- |
| Semantic cell diff | Formula/value additions, removals, and changes—not ZIP/XML noise |
| Impact trace | Downstream formula cells and deterministic shortest dependency-path samples, including cross-sheet, static named ranges, formula-defined names, static named `LAMBDA` calls, `LET`/inline-`LAMBDA`, Excel-table, 3-D worksheet references, fixed legacy CSE result members, and currently observed dynamic-array result members |
| Formula-pattern break | An edited formula that no longer matches equal neighboring formulas |
| Workbook controls | Sheet visibility, defined names, Excel-table definitions, data-validation, conditional-formatting, operational protection, external-data refresh, external-link-package, XLM macro-sheet, Office RibbonX, Office Web Add-in task-pane, worksheet ActiveX/form-control/OLE, and Power Query controls; array-formula mode/fixed-output range, static 3-D-reference scope, calculation settings, and VBA payload changes |
| Formula hazards | New external-workbook references and `#REF!` formulas |
| Coverage changes | New parser warnings, unresolved formula references (including unsupported table syntax), dynamic-reference functions (`INDIRECT`/`OFFSET`), dynamic-array spill references, explicit implicit intersection, and formula-tokenization failures |
| Policy as code | Protected cells, allowed edit areas, bans, and change/impact limits |
| CI output | Deterministic JSON, reviewer-friendly Markdown, and SARIF |

See [the policy reference](docs/policy.md) for the configuration contract and
[the threat model](docs/threat-model.md) for important limits. The
[external validation notes](docs/validation.md) record an independently
maintained financial-model compatibility check.

For Excel tables, FormulaFence statically resolves a table name, a single
column or contiguous column range, and the `#All`, `#Data`, `#Headers`, and
`#Totals` regions. It also resolves row-scoped `@` / `#This Row` references
when their formula location proves the row: unqualified forms such as
`[@[Sales Amount]]` and `[Sales Amount]` only inside a table data cell, plus
qualified forms such as `Sales[@Amount]` and
`Sales[[#This Row],[Amount]:[Rate]]` anywhere on that table's data row. Header,
total, cross-sheet, ambiguous, and complex bracket-escape cases remain explicit
coverage notes. The supported subset follows Excel's documented
[structured-reference semantics](https://support.microsoft.com/en-us/excel/using-structured-references-with-excel-tables).

FormulaFence expands a formula-defined name when its complete definition can be
statically resolved to internal dependencies or a constant. It follows nested
workbook and sheet-local definitions, so a formula such as
`=DiscountedValue`, where `DiscountedValue` uses `TaxRate`, receives edges to
the underlying inputs. Relative references, cycles, external links,
`INDIRECT`/`OFFSET`, 3-D spans inside a name definition, and syntax the formula
tokenizer cannot inspect remain unresolved at the use site rather than being
treated as a safe dependency. This builds on Excel's documented support for
[names that represent formulas and constants](https://support.microsoft.com/en-us/excel/names-in-formulas).

FormulaFence recognizes the lexical names introduced by inline `LET` and
`LAMBDA` expressions, including nested lambdas supplied to functions such as
`REDUCE`. Those local variables no longer masquerade as unresolved workbook
names; the static A1, named, and table references around them remain graph
edges. The implementation follows Excel's documented
[LET scope](https://support.microsoft.com/en-us/excel/functions/let-function)
and [LAMBDA parameter syntax](https://support.microsoft.com/en-us/excel/functions/lambda-function).

FormulaFence traces the static anchor behind a direct internal spilled-array
reference such as `=SUM(Inputs!B2#)`, including Excel-compatible OOXML spelling
such as `=SUM(_xlfn.ANCHORARRAY(Inputs!B2))`. That creates an edge from the
anchor cell to the consumer, so changes to the anchor formula and its statically
visible inputs reach the consumer. Profiles retain every spill token, and a
diff emits `FF015`; `no_new_spill_references` can make new instances a CI
failure. FormulaFence does not invent the spill extent or a dependency on every
possible blocking cell. External, 3-D, range, named, implicit-intersection, and
malformed spill forms remain explicit coverage limits. A formula-defined name
that contains a spill reference also remains unresolved at its call site rather
than hiding that dynamic boundary behind a static name edge.

FormulaFence also inventories Excel's explicit implicit-intersection operator:
the display form `@A1:A3`, function forms such as `@INDEX(...)`, and persisted
OOXML `_xlfn.SINGLE(...)`. When a direct static A1 cell or range has one
unambiguous documented row/column intersection, it adds that single-cell
edge—for example, `=_xlfn.SINGLE(Inputs!B2:B4)` in row 3
depends on `Inputs!B3`, not every input in the range. Other expressions retain
their visible static input edges without evaluating Excel. Profiles record each
recognized explicit use; a new use emits `FF017`, and
`no_new_implicit_intersections` can require review in CI. This is distinct from
Excel-table `[@Column]` syntax, which follows the table-specific resolver.
Formula-defined names containing explicit implicit intersection remain
unresolved at call sites because their selected cell depends on the caller's
position. The scope follows Microsoft's
[implicit-intersection guidance](https://support.microsoft.com/en-us/excel/implicit-intersection-operator)
and [XlsxWriter's documentation](https://xlsxwriter.readthedocs.io/working_with_formulas.html)
that persisted `@` behavior uses `SINGLE()`.

FormulaFence also recognizes a multi-cell **legacy CSE** array formula as a
fixed output range when its OOXML array anchor has no dynamic-array cell
metadata. It does not inflate that range into virtual cells. Instead, if an
ordinary formula reads a member such as `Model!B2`, FormulaFence adds a compact
alias from the CSE anchor (for example `Model!B1`) to that consumer. Thus a
change to an input of `{=LEN(Inputs!A1:A3)}` can reach a downstream `=B2*10`
calculation even though Excel stores the formula only at `B1`. The profile
lists each declared fixed range. An OOXML dynamic-array marker (`cm` resolving
to `XLDAPR`/`fDynamic`) records the current serialized output range as an
**observed** spill surface. If a static formula reads a non-anchor member such
as `Model!B2`, FormulaFence adds the same compact anchor-to-consumer graph edge
and records the relationship in the profile. It never calls that range fixed:
Excel can grow, shrink, or block a spill during recalculation. A newly observed
output-member relationship emits `FF019`; enable
`no_new_dynamic_array_output_references` for `FFP019`. Unknown array metadata
stays visible as a parser coverage warning with no aliases. FormulaFence emits
`FF018` when a legacy-CSE or dynamic-array formula is added, removed, or changes
mode, or when a legacy CSE fixed output range changes; enable
`no_array_formula_semantics_changes` for `FFP018`. This follows Microsoft's
[dynamic-versus-legacy array guidance](https://support.microsoft.com/en-us/excel/dynamic-array-formulas-and-spilled-array-behavior)
and [XlsxWriter's documented CSE/dynamic serialization behavior](https://xlsxwriter.readthedocs.io/working_with_formulas.html).

FormulaFence inventories **data-validation controls** as reviewable workbook
semantics. A validation can restrict a whole column, select a list from another
sheet, enforce a custom formula, hide a list arrow, or change whether Excel
shows a prompt or blocks an invalid entry. The profile records each target range
and its effective control settings without exposing validation formulas or
prompt/error text; a local change report includes the full before/after rule.
It normalizes OOXML defaults such as `operator=between` and `errorStyle=stop`,
the optional leading `=` in a criterion, and equivalent grouping of identical
target rules, so equivalent writers do not create noise. Any changed validation
control emits `FF020`; enable
`no_data_validation_changes` for `FFP020`. FormulaFence does not evaluate a
validation formula or predict whether a user entry will be accepted. The scope
follows Microsoft's [data-validation guidance](https://support.microsoft.com/en-US/Excel/get-started/apply-data-validation-to-cells)
and [openpyxl's documented range model](https://openpyxl.readthedocs.io/en/stable/validation.html).

FormulaFence also inventories **conditional-formatting controls** directly from
worksheet OOXML, because a priority move, `Stop If True` toggle, exception
formula, or changed red/green threshold can alter what reviewers see without
altering a normal formula cell. It records compact target ranges, worksheet-wide
precedence, criteria, rule flags, differential styles, color scales, data bars,
and icon sets. Differential styles are resolved from their actual OOXML rather
than their unstable `dxfId`, while omitted boolean defaults, optional leading
`=` formula spelling, priority-number gaps, and extension GUID links are
normalized to avoid writer-only noise. Excel-2010 extension fragments are retained as
opaque local evidence even when the parser cannot model them. Profiles redact
criteria, text rules, and raw style/extension XML; local change reports retain
the full before/after control. Any change emits `FF021`; enable
`no_conditional_formatting_changes` for `FFP021`. FormulaFence does not
calculate a conditional-formatting formula, resolve a relative rule for every
target cell, or predict the final visual display. The scope follows Microsoft's
[conditional-formatting precedence guidance](https://support.microsoft.com/en-us/excel/use-conditional-formatting-to-highlight-information-in-excel)
and [OOXML conditional-formatting model](https://learn.microsoft.com/en-us/office/open-xml/spreadsheet/working-with-conditional-formatting).

FormulaFence also inventories **operational protection controls** directly from
OOXML: workbook structure/windows/revision locks; worksheet and dialog-sheet
permissions; chart-sheet content/object locks; protected-range target areas;
and explicit cell, row, and column `locked`/`hidden` assignments when a normal
sheet is protected. This catches a newly editable input or newly hidden formula
without expanding styled rows or columns into cells. It normalizes effective
OOXML defaults for sheet actions, so writers that omit or explicitly serialize
the same defaults do not create a diff. The profile and every report redact
legacy verifiers, password hashes, salts, protected-range names, and security
descriptors; FormulaFence retains only private comparison fingerprints plus
safe presence metadata. Any change emits `FF022`; enable
`no_protection_changes` for `FFP022`. These controls are not file encryption,
identity enforcement, or a security guarantee: workbook and worksheet
protection are an operational review surface, and FormulaFence does not decide
who can successfully edit a file or fully emulate Excel's style cascade. The
scope follows Microsoft's [workbook protection guidance](https://support.microsoft.com/en-US/Excel/protect-a-workbook),
[worksheet protection guidance](https://support.microsoft.com/en-us/excel/protect-a-worksheet),
and [protection-and-security overview](https://support.microsoft.com/en-US/Excel/protection-and-security-in-excel).

FormulaFence also inventories **external-data refresh controls** directly from
OOXML before a reader library can omit them: workbook-wide external-link and
refresh-on-open flags; connection refresh schedules, background behavior,
credential/cache flags, source-kind metadata, connection-file reload policy,
and parameter-triggered refreshes; linked query-table refresh/growth behavior;
and pivot-cache source and refresh controls. This catches a workbook that can
change its inputs on open or during use even when its ordinary formulas did not
change. Profiles and reports intentionally omit connection names and
descriptions, paths, URLs, connection strings, commands, parameter values, SSO
identifiers, cached records, and opaque extension XML; private fingerprints
still expose a material source or identity change. Any change emits `FF023`;
enable `no_external_data_connection_changes` for `FFP023`. FormulaFence never
opens a connection or refreshes data, does not assess source trust or actual
returned values, and does not model PivotTable layout semantics. The scope follows Microsoft's
[external-data refresh guidance](https://support.microsoft.com/en-us/excel/refresh-an-external-data-connection-in-excel)
and the [SpreadsheetML Connections part](https://c-rex.net/samples/ooxml/e1/Part1/OOXML_P1_Fundamentals_Connections_topic_ID0EQLGK.html).

FormulaFence separately inventories raw **external-link packages**
(`xl/externalLinks/externalLink*.xml`). It recognizes external-workbook,
DDE, and OLE definitions; binds `externalReferences` declarations to their
package parts privately; and compares their source relationship, definition,
cached-data, item-behavior, and unmodelled-extension material. Profiles expose
only safe counts—never workbook targets, sheet/defined names, DDE services or
items, OLE program/item names, or cached values. A material package change
emits `FF025`; enable `no_external_link_package_changes` for `FFP025`.
FormulaFence never follows or executes these links, establishes source trust,
or infers returned data. The package shape follows the
[SpreadsheetML `externalLink` definition](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.externallink?view=openxml-3.0.1).

FormulaFence separately inventories **Excel 4.0 / XLM macro sheets**. Unlike
VBA, this executable automation is stored in raw macro-sheet XML parts (usually
`xl/macrosheets/*.xml`), not `xl/vbaProject.bin`. FormulaFence binds the
documented `xlMacrosheet` and `xlIntlMacrosheet` workbook relationships to
their parts, privately fingerprints complete macro XML and related package
relationships, then streams direct safe internal relationship targets into
private payload fingerprints. Profiles expose only safe counts for sheets,
formula cells, visibility, related OLE/package parts, and fingerprinted versus
uninspected internal targets. A material change emits `FF026`; enable
`no_xlm_macro_sheet_changes` for `FFP026`. FormulaFence never executes,
emulates, resolves, or parses XLM commands, relationship targets, or embedded
objects. It never follows external targets; direct internal payload scanning is
bounded to 32 MiB per part, 64 MiB per workbook, and 256 parts, after which a
coverage warning remains visible. The package shape follows Microsoft's [Macro Sheet
part](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-offmacro/b8bee527-ef5a-4734-bb8c-6eae4166b6c9)
and [International Macro Sheet
part](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-offmacro/450634cb-ca5a-4350-9edb-940a90707f49).

FormulaFence separately inventories **Office RibbonX custom UI** package parts.
RibbonX can bind controls to workbook callback names even when no ordinary cell
or `vbaProject.bin` payload changes. FormulaFence reads the documented root
package relationships and `customUI` XML for the 2006 and Office 2010-era
schemas, privately fingerprints full control XML plus direct image
relationships, and reports only safe part, control, callback-attribute, and
relationship counts. A material change emits `FF027`; enable
`no_ribbon_customization_changes` for `FFP027`. Control IDs, labels, callback
names, XML, and image targets never enter profiles or reports. FormulaFence
does not execute callbacks, follow external relationships, or parse image
payloads. Custom-UI XML reads are bounded to 16 MiB per part, 32 MiB per
workbook, and eight parts; a bound or malformed part becomes a coverage
warning. The package forms follow Microsoft's [Ribbon Extensibility
Part](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-customui/52faf7b6-fecc-48d9-96db-ee80a631a5ac)
and [Ribbon and Backstage Customizations
part](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-customui2/452a58ae-cb0a-4926-83f8-fb1cbaa6114c)
specifications; its `onLoad`, `loadImage`, and `onAction` callback surface is
documented in the [Custom UI schema](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-customui2/a232628d-f6fb-4630-a463-459989a68e7a).

FormulaFence separately inventories document-linked **Office Web Add-in task
panes**. A workbook can declare a task-pane part, bind it to a web-extension
definition, and request `Office.AutoShowTaskpaneWithDocument` even though no
ordinary cell, VBA payload, or RibbonX control changes. FormulaFence follows
the bounded package chain from the workbook relationship through
`taskpanes.xml` and its direct `webextension*.xml` definitions. It compares
task-pane configuration, visible/locked state, add-in references, auto-show
properties, bindings, snapshots, and direct relationship semantics privately;
profiles expose only safe counts. A material change emits `FF028`; enable
`no_office_web_addin_changes` for `FFP028`. Add-in IDs, store references,
property values, binding values, snapshot data, XML, and relationship targets
never enter profiles or reports. FormulaFence does not install, load, execute,
or fetch an add-in or manifest, and it never follows an external relationship.
Task-pane and web-extension XML reads are bounded to 16 MiB per part, 32 MiB
per workbook, and 64 parts; malformed, unbound, oversized, or over-budget
parts remain explicit coverage warnings. Worksheet-scoped web-extension markup
outside this task-pane chain is not yet modeled. The package surface follows
Microsoft's [Taskpane Web Extension File](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-owexml/3d04f8ce-65f2-4dc3-bafa-636d0a7e41a1)
and [Web Extension](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-owexml/56fe5a64-dd6d-422c-beac-19d72dd10ade)
specifications.

FormulaFence also inventories relationship-backed **worksheet embedded controls
and OLE objects**. A worksheet can bind `<control>` markup to persisted ActiveX
state, nested form-control properties, or raw OLE/package data without changing
an ordinary cell or the VBA payload. FormulaFence confirms the worksheet
relationship chain, then compares control/OLE declarations, ActiveX `ocx`
persistence XML, `formControlPr` bindings, direct related relationships, and
bounded private hashes of ActiveX binary and direct OLE/package payloads. A
material change emits `FF029`; enable
`no_worksheet_embedded_control_changes` for `FFP029`. Profiles expose only safe
counts. Control names, class IDs, licenses, macros, formulas/ranges, OLE
identities, relationship targets, XML, and payload bytes never enter profiles
or reports. FormulaFence never initializes a control, deserializes or opens an
OLE object/package, renders a drawing surface, follows an external target, or
infers control event dispatch. Relevant XML reads are bounded to 16 MiB per
part, 64 MiB per workbook, and 512 parts; direct raw payload hashes are bounded
to 32 MiB per part, 64 MiB per workbook, and 512 parts. Malformed, orphaned,
unbound, oversized, or over-budget material remains an explicit coverage
warning. This scope follows Microsoft's guidance on [sheet ActiveX
controls](https://learn.microsoft.com/en-us/office/vba/excel/concepts/controls-dialogboxes-forms/using-activex-controls-on-sheets),
the [`ocx` ActiveX persistence
schema](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oe376/b30a660a-95eb-4716-b201-a46aae788610),
and [form-control properties](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/3d054a6d-4f94-4082-837a-f939fd8d4a45).

FormulaFence also inventories **Power Query Data Mashup** custom XML parts.
It reads the documented length-prefixed container, fingerprints the embedded
`Section1.m` formula document and logical package material privately, and
compares structural query metadata plus formula-firewall permissions. Its
profile exposes only safe counts and controls—not M text, query names, source
locations, metadata values, embedded content, telemetry IDs, or user-bound
permission bindings. `sqmid` telemetry and result-only refresh metadata are
ignored to avoid writer and refresh noise. A material query-definition or
semantic-control change emits `FF024`; enable `no_power_query_changes` for
`FFP024`. FormulaFence does not execute M, refresh a connection, or infer the
returned data. The implementation follows Microsoft's
[Query Definition File Format](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-qdeff/27b1dd1e-7de8-45d9-9c84-dfcc7a802e37),
which stores Power Query definitions in a Custom XML part.

FormulaFence also follows a call to a workbook- or worksheet-local defined name
when its complete definition is one statically resolvable `LAMBDA` expression.
For `=ToCelsius(A2)`, the caller keeps its explicit `A2` edge and gains the
function's static internal dependencies; nested named-LAMBDA calls and
formula-defined names that call a named LAMBDA are resolved the same way. It
recognizes both human-authored formulas and the `_xlfn.LAMBDA` / `_xlpm.` /
`_xlop.` OOXML spelling produced by Excel-compatible writers. Definition scope
and worksheet-local precedence are preserved. Relative, cyclic, external,
dynamic, 3-D, tokenizer-unsupported, or otherwise non-static LAMBDAs remain a
visible unresolved reference at each call site. Spill extents and blockers,
plus arbitrary VBA, add-in, or other custom functions, remain coverage limits
rather than inferred dependencies.

If the underlying formula tokenizer cannot inspect a formula at all,
FormulaFence records the affected cell in the profile and reports `FF016` when
it is newly introduced. `no_new_tokenization_failures` turns that condition
into a CI failure instead of silently omitting its dependency graph.

FormulaFence also expands internal static 3-D A1 references such as
`Jan:Mar!B2:B10` over every worksheet tab between the named endpoints. Profiles
identify those formula cells. If an unchanged 3-D formula's tab span changes
because a sheet is inserted, removed, or moved, FormulaFence emits `FF014`; the
`no_3d_reference_scope_changes` policy rule can make that a hard boundary.
Unknown endpoints and non-A1 constructs remain explicit coverage cases rather
than invented dependencies; external 3-D references remain external-link
hazards. This follows Excel's
[3-D-reference behavior](https://support.microsoft.com/en-us/excel/create-a-3-d-reference-to-the-same-cell-range-on-multiple-worksheets).

## Development

```bash
python -m venv .venv
.venv/bin/python -m pip install -e '.[dev]'
.venv/bin/python -m pytest
.venv/bin/python -m ruff check .
```

## Safety and scope

FormulaFence reads workbook structure only. It does not recalculate formulas,
run VBA, follow external links, or claim that a workbook's numbers are correct.
It is a review and control layer; human review remains essential for material
models.

## License

[MIT](LICENSE)
