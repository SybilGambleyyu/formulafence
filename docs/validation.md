# External validation notes

FormulaFence's test suite builds small fixtures that isolate individual risks.
Those tests are necessary but insufficient for confidence in an Office-file
reader, so each release should also be exercised on independently maintained
workbooks without copying their contents into this repository.

## Excel Named Sheet Views — 2026-07-25

FormulaFence 0.34.0 was validated with controlled raw-OOXML `.xlsx` packages
containing one relationship-backed Named Sheet View part, two private named
views, two alternate filters, two column filters and criterion groups, and two
sort rules/conditions. The suite verifies safe profile counts, a zero-change
self-diff, `FF038` for a criterion-only saved-view change, and `FFP038` under
`no_named_sheet_view_changes`. A second controlled fixture uses a table-owned
AutoFilter and exercises Excel's table-ID reconciliation fallback.

Equivalent GUID, local A1 case/absolute-reference, Boolean/default, and
unsigned-integer spellings are exercised without a finding. An out-of-range
alternate-view column identifier produces an explicit parser-coverage warning,
`FF010`, and `FF038` rather than a silent omission. View names, IDs, criteria,
target ranges, table bindings, table-column IDs, and sort keys are verified
absent from JSON, Markdown, ordinary reports, and SARIF. The controlled
criterion-only mutation has no ordinary cell, formula, or active-AutoFilter
change and is invisible to the published 0.33.0 wheel (`0` changes, no
findings); a freshly installed 0.34.0 wheel emits exactly one
`named_sheet_views_changed` change and `FF038` (then `FFP038` with the starter
policy).

As an independent package-compatibility check, a fresh 0.34.0 wheel profiled
LibreOffice's public
[`NamedSheetViews.xlsx`](https://raw.githubusercontent.com/LibreOffice/core/6e6bf902f0e4849e4fdb180e5a9e859028e40a1e/sc/qa/unit/data/xlsx/NamedSheetViews.xlsx)
fixture at commit `6e6bf902f0e4849e4fdb180e5a9e859028e40a1e`. The downloaded
workbook SHA-256 was
`896f863f92dc5fc05ce7b038272261106a0d40f6fb77abff7bc149880346eaef`.
FormulaFence found one worksheet and relationship-backed part, two views/two
alternate filters, two column filters and criterion groups, two sort
rules/conditions, and no Named Sheet View coverage warning. This validates
static declaration and reconciliation comparison plus data minimisation—not
whether Excel/LibreOffice will activate or render a saved view, calculate a
filtered result, repair metadata, infer formula visibility sensitivity, or
interpret differential-format, extension, or rich-sort behavior. The boundary
follows Microsoft's [Named Sheet Views part](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/4b4d6448-d997-4ebe-9153-5c2c67d16972),
[`CT_NsvFilter`](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/e132d9cc-c711-4fb3-aa28-e7356a791b1c),
and [reconciliation](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/dd6b2cb8-b5b3-43b1-a5bd-dccdd9c0864a)
definitions.

## Excel ignored-error controls — 2026-07-25

FormulaFence 0.33.0 was validated with controlled raw-OOXML `.xlsx` packages
containing three standard `ignoredError` declarations and one Office 2010
`x14:ignoredError` declaration. Together they suppress private evaluation,
inconsistent-formula, omitted-range, unlocked-formula, empty-reference,
list-validation, calculated-column, text-number, and two-digit-year warnings
across five private target ranges. The suite verifies safe profile counts, a
zero-change self-diff, `FF037` for a target-only standard or Office 2010
extension change, and `FFP037` under `no_ignored_error_changes`.

Equivalent local A1 case/absolute-reference, Boolean, and target-order spellings
are exercised without a finding. A nonlocal target produces an explicit parser
coverage warning, `FF010`, and `FF037` rather than a silent omission. Target
ranges and individual suppressions are verified absent from JSON, Markdown,
ordinary reports, and SARIF. The controlled target-only mutation has no ordinary
cell or formula change and is invisible to the published 0.32.0 wheel, while
0.33.0 emits only `ignored_error_controls_changed` / `FF037`.

As an independent package-compatibility check, FormulaFence profiled the public
XlsxWriter [`ignore_errors.py`](https://github.com/jmcnamara/XlsxWriter/blob/main/examples/ignore_errors.py)
example, generated locally with XlsxWriter 3.2.9 and not bundled with this
repository. The resulting `ignore_errors.xlsx` SHA-256 was
`57d059a43c6d01602199e0dbac5030fa38489936df7bfd6392474a01122a0eca`.
FormulaFence found one standard container, two suppressed-warning rules, two
target ranges, one evaluation-error suppression, one number-stored-as-text
suppression, and no ignored-error coverage warning. This validates static
declaration comparison and data minimisation—not whether Excel would show a
warning, formula evaluation, error repair, application-level error-checking
configuration, or downstream-impact inference. The boundary follows the OOXML
[`ignoredError`](https://c-rex.net/samples/ooxml/e1/Part4/OOXML_P4_DOCX_ignoredError_topic_ID0EVK24.html)
definition and Microsoft's Office 2010 [`ignoredErrors`](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/0d164d85-23bf-4d43-87c5-9fcde148aabe)
documentation.

## Excel filters and row visibility — 2026-07-25

FormulaFence 0.32.0 was validated with controlled raw-OOXML `.xlsx` packages
containing one worksheet AutoFilter, one Table Definition-part AutoFilter, two
private criterion groups, two private sort conditions, two explicitly hidden
outlined rows, one collapsed outline marker, and a `zeroHeight` hidden-by-default
sheet with an explicit visible-row override. The suite verifies the safe profile
counts, a zero-change self-diff, `FF036` when only a worksheet criterion, a
table criterion, or a raw row-hidden flag changes, and `FFP036` under
`no_filter_visibility_changes`.

Equivalent local A1 case/absolute-reference, Boolean/default, and unsigned
integer spellings are exercised without a finding. An out-of-range unsigned
filter-column identifier produces an explicit parser-coverage warning, `FF010`,
and `FF036` rather than a silent omission. Filter criteria, selected values,
custom sort lists, and row/range references are verified absent from JSON,
Markdown, ordinary reports, and SARIF. The controlled criterion-only mutation
is invisible to the published 0.31.0 wheel: it has no ordinary cell or formula
change, whereas 0.32.0 emits only `filter_visibility_controls_changed` /
`FF036`.

As an independent package-compatibility check, FormulaFence profiled the public
XlsxWriter [`autofilter.py`](https://github.com/jmcnamara/XlsxWriter/blob/main/examples/autofilter.py)
example and its [`autofilter_data.txt`](https://github.com/jmcnamara/XlsxWriter/blob/main/examples/autofilter_data.txt)
input, generated locally with XlsxWriter 3.2.9 and not bundled with this
repository. The resulting `autofilter.xlsx` SHA-256 was
`ff09b2a3f580fbca170fd94acba46d344f7916ec892f421a460f02404160d2ba`.
FormulaFence found seven worksheet AutoFilters, seven filter columns and
criterion groups, 163 explicitly hidden rows, and no visibility-control
coverage warning. This validates static OOXML declaration comparison and data
minimisation—not filter application, recalculation, `SUBTOTAL`/`AGGREGATE`
correctness, formula-sensitivity inference, or rendering. The boundary follows
Microsoft's [SUBTOTAL documentation](https://support.microsoft.com/en-us/excel/functions/subtotal-function)
and the Open XML [`autoFilter`](https://c-rex.net/samples/ooxml/e1/part4/OOXML_P4_DOCX_autoFilter_topic_ID0EIDM4.html),
[`filterColumn`](https://c-rex.net/samples/ooxml/e1/Part4/OOXML_P4_DOCX_filterColumn_topic_ID0ELVP5.html),
and [`sheetFormatPr`](https://c-rex.net/samples/ooxml/e1/part4/OOXML_P4_DOCX_sheetFormatPr_topic_ID0EVAG5.html)
definitions.

## Excel Scenario Manager — 2026-07-25

FormulaFence 0.31.0 was validated with controlled raw-OOXML `.xlsx` packages
containing two worksheet-local scenarios, four private stored inputs, one locked
scenario, one hidden scenario, comments/users, selected/shown scenario state,
summary references, and an input display number format. The suite verifies safe
profile counts, a zero-change self-diff, `FF035` for a stored-input-only change,
and `FFP035` under `no_scenario_manager_changes`. The same controlled mutation
is invisible to the published 0.30.0 wheel: its JSON report contains no changes
or findings, while 0.31.0 emits only `scenario_manager_changed` / `FF035`.

Equivalent local A1 case/absolute-reference, Boolean, and unsigned-integer
spellings are exercised alongside omitted schema-default false flags. A scenario
name duplicated on a different worksheet remains valid because Scenario Manager
is worksheet-scoped. A malformed input reference produces an explicit parser
coverage warning, `FF010`, and `FF035` rather than a silent omission. Scenario
names, comments, users, stored values, changing-cell references, and summary
references are verified absent from JSON, Markdown, ordinary reports, and
SARIF.

As an independent package-compatibility check, FormulaFence profiled the public
[`scenario.xlsx` example](http://carltoncollins.com/scenario.xlsx) linked by the
[Journal of Accountancy Scenario Manager article](https://www.journalofaccountancy.com/issues/2018/nov/excel-scenario-manager/),
downloaded locally and not bundled with this repository. The downloaded workbook
SHA-256 was `087e7cc6c64c42c66f26049e66334c2cb0df20042f6b3a89d363e1ee44ca631d`.
FormulaFence found one Scenario Manager worksheet, six scenarios, 18 stored
inputs, six locked scenarios, six comments/users, six display number formats,
one summary reference, selected/shown scenario state, and no Scenario Manager
coverage warning. Changing only one stored input in a temporary copy emitted
`FF035` and `FFP035`; the replacement value was absent from JSON output. This
validates static OOXML declaration comparison and data minimisation—not scenario
application, formula calculation, result correctness, scenario-summary
generation, dependency inference, or Excel rendering. The declaration boundary
follows Open XML [`scenarios`](https://c-rex.net/samples/ooxml/e1/part4/OOXML_P4_DOCX_scenarios_topic_ID0EVDF5.html),
[`scenario`](https://c-rex.net/samples/ooxml/e1/part4/OOXML_P4_DOCX_scenario_topic_ID0E5WE5.html),
and [`inputCells`](https://c-rex.net/samples/ooxml/e1/Part4/OOXML_P4_DOCX_inputCells_topic_ID0EE624.html).

## Excel What-If Data Tables — 2026-07-24

FormulaFence 0.30.0 was validated with controlled `.xlsx` fixtures containing
three raw-OOXML `f t="dataTable"` masters: a one-variable column table, a
one-variable row table, and a two-variable table. The test suite verifies the
safe profile counts, a zero-change self-diff, `FF034` on a private input-reference
change, and `FFP034` under `no_what_if_data_table_changes`. It also checks that
equivalent lowercase/absolute A1 and Boolean spellings do not create a finding,
while a deleted input is recorded and a malformed input reference becomes an
explicit parser-coverage warning. Raw input references and output ranges are
verified absent from JSON, Markdown, ordinary reports, and SARIF.

As an independent package-compatibility check, FormulaFence inspected the
public [`sensitivity2d.xlsx` fixture](https://github.com/witanlabs/witan-vs-openpyxl/blob/8a7f538b13b98f7098102bfdc779b8920f63e403/fixtures/sensitivity2d.xlsx)
from the pinned `witan-vs-openpyxl` source revision, downloaded locally and not
bundled with this repository. The downloaded workbook SHA-256 was
`bc5f9efa6ca78ebd986d81b3cf372acc2a72a6a1326178118d0158f988427bcc`.
FormulaFence found one two-variable master, 25 declared output cells, one
recalculation request, and no Data Table coverage warning. Comparing the exact
file to itself produced no changes, which independently confirms stable handling
of the reader's `DataTableFormula` object representation. This validates static
OOXML declaration comparison and data minimisation—not scenario calculation,
cached-output correctness, output-formula inference, or downstream-impact
analysis. The declaration boundary follows the Open XML
[`f` (Formula) specification](https://c-rex.net/samples/ooxml/e1/part4/OOXML_P4_DOCX_f_topic_ID0E6TY4.html).

## Embedded Power Pivot/Data Model — 2026-07-24

FormulaFence 0.29.0 was validated with controlled raw-OOXML `.xlsx` packages
that add an `x15:dataModel` workbook declaration, two private model-table
records, one private model relationship, an explicit workbook
`powerPivotData` binding, and a harmless opaque `xl/model/item.data` payload to
a small ordinary workbook. The payload was never deserialized or opened in
Office. Changing only the raw payload or only the declaration emitted `FF033`;
`no_power_pivot_data_model_changes` emitted `FFP033`. Synthetic table,
relationship, connection, column, and payload values were verified absent from
JSON, Markdown, ordinary reports, and SARIF.

The controlled suite rewrote a workbook relationship ID, used an equivalent
internal target spelling, and regenerated writer GUIDs in model metadata; those
equivalent representations produced no finding. Moving the binding,
externalizing it, adding an unexpected direct relationship on the model part,
and lowering the payload-size limit each produced `FF033` or a visible coverage
warning. External targets were not fetched or exposed. A model-free workbook
did not consume the Data Model payload budget.

As an independent package-compatibility check, FormulaFence profiled Microsoft's
public [Customer Profitability Excel sample](https://github.com/MicrosoftDocs/powerbi-docs/blob/main/powerbi-docs/create-reports/sample-datasets.md),
downloaded locally from the pinned
[`powerbi-desktop-samples` source revision](https://github.com/microsoft/powerbi-desktop-samples/tree/f66e8c775a1426504254f7a061b8fed482601800)
and not bundled with this repository. The downloaded workbook SHA-256 was
`76f21c59d631e95bbad5489350695a46d903061aedb179e88b72f038772666d4`.
FormulaFence found one embedded model payload and workbook binding, one Data
Model declaration, nine model tables, eight model relationships, and no
Data-Model coverage warning. The source workbook's unrelated external-data
coverage notes remained visible. This validates static, relationship-aware
comparison and data minimisation—not Analysis Services deserialization, DAX
evaluation, refresh, report calculation/rendering, model-to-cell impact,
external-target retrieval, or the semantic correctness of model data. Production
raw payload reads are bounded to 512 MiB per part, 512 MiB per workbook, and 16
parts. The declaration boundary follows the Open XML
[`x15:dataModel` reference](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.linq.x15.datamodel?view=openxml-3.0.1).

## Slicer and Timeline cache filter state — 2026-07-24

FormulaFence 0.28.0 was validated with controlled raw-OOXML `.xlsx` packages
that add documented workbook Slicer-cache and Timeline-cache declarations,
their explicit workbook relationships, one Pivot-backed Slicer cache, one
table-Slicer cache, and one Pivot-backed Timeline cache to a harmless existing
PivotTable fixture. The caches contain private source names, item selections,
date-state/filter values, and filtered-PivotTable names; the workbook was never
opened in Office.

The controlled package exercises both documented `x14` and `x15`
`slicerCaches` workbook containers, the `x14:pivotCacheDefinition` extension
identifier used by Pivot-backed filters, and the `extLst`-scoped table-Slicer
binding. As an independent compatibility check, FormulaFence also profiled the
unmodified `TestSlicer.xlsm` generated by upstream
[Excelize](https://github.com/qax-os/excelize/tree/32931c30d9195445c1f5bdca00eaf29176cd2c54)
test code. It found all four real Slicer cache parts, two table and two
PivotCache bindings, 13 items, 11 selections, and no FormulaFence
Slicer/Timeline coverage warning. That workbook is library-generated rather
than an Office-authored validation sample, so it validates package compatibility
only—not Excel rendering or filter application.

Changing only a Slicer item selection and Timeline state emitted `FF032` with
separate private Slicer and Timeline material flags, while the PivotTable
definition remained unchanged. The policy
`no_slicer_timeline_cache_changes` emitted `FFP032`. Profiles exposed only safe
counts for cache parts, bindings, items/selections, Timeline states/filters, and
relationship coverage. Synthetic names, fields, selected values, date markers,
targets, XML, Markdown, JSON, ordinary reports, and SARIF were verified absent.

The suite rewrote workbook relationship IDs, used equivalent internal target
spellings, coordinated a Slicer/Timeline PivotCache extension-ID renumbering
across both cache forms, spelled out optional Slicer defaults, changed Boolean
spellings, and regenerated a Timeline GUID. Those equivalent representations
produced no finding. Moving a cache target, externalizing a target, corrupting a
cache root, and reducing the XML limit each produced either `FF032` or an
explicit coverage warning; external targets were never fetched or exposed.
Production cache XML reads are bounded to 16 MiB per part, 64 MiB per workbook,
and 512 parts. This validates static, relationship-aware comparison and data
minimisation—not filter application, PivotTable/table calculation or rendering,
downstream-impact analysis, external-target retrieval, or worksheet/drawing
Slicer and Timeline view geometry/styles. The fixture follows Microsoft's
[Slicer Cache part](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/7dbb4481-b021-45cc-8bd4-6094b566a5ff),
[Timeline Cache part](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/29a0f58c-d942-4641-8ed0-4f02010326f2),
[Slicer-to-PivotCache relationship](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/2a393f85-21f9-4a27-a2b7-4867223f4b9a),
and [Timeline cache definition](https://learn.microsoft.com/en-sg/openspecs/office_standards/ms-xlsx/f45ff6ef-fb62-4e19-8e8c-822e3be9ef75).

## Public cap-table model — 2026-07-24

We inspected the public [Foresight Cap Table and Exit Waterfall Tool](https://github.com/foresighthq/cap-table-tool),
which its repository licenses under [CC BY-NC-SA 4.0](https://creativecommons.org/licenses/by-nc-sa/4.0/).
The workbook was used locally for compatibility validation only and is not
bundled with FormulaFence.

```bash
formulafence profile 'Cap Table and Exit Waterfall by Foresight.xlsx' --format markdown
```

Observed profile:

| Measure | Result |
| --- | ---: |
| Sheets | 18 |
| Non-empty cells | 6,623 |
| Formula cells | 4,228 |
| VBA payload | absent |

The model contains an OOXML extension that the underlying parser does not fully
support. FormulaFence now captures that as an explicit coverage note rather than
allowing a dependency warning to leak into CI logs. This is a useful result, not
a pass/fail claim about the model itself: unsupported workbook features should
remain visible to the reviewer.

## Named and dynamic reference coverage — 2026-07-24

The same local profile under FormulaFence 0.3.0 found 20 workbook or
sheet-scoped defined names, no unresolved range tokens, and **36 formula cells
using `INDIRECT`**. Those dynamic cells are now explicit inspection-coverage
notes rather than silent holes in an impact trace. The result does not judge the
model's use of `INDIRECT`; it gives a reviewer or policy author a concrete,
machine-readable scope for the limitation.

## Formula-defined names — 2026-07-24

FormulaFence 0.7.0 adds conservative expansion for defined names whose
definitions are formulas. [Microsoft documents that a defined name can
represent a cell, range, formula, or constant](https://support.microsoft.com/en-us/excel/names-in-formulas),
including reusable formulas in modern Excel. In a controlled local workbook,
`DiscountedValue` expanded through another name (`TaxRate`) to two `Inputs`
cells; changing the rate reached the formula that used `=DiscountedValue`. A
sheet-local `LocalMetric` definition also resolved both from its own sheet and
through an explicitly qualified use from another sheet. A named constant was
recognized without inventing a cell edge.

The same fixture confirmed the boundary: relative definitions, `OFFSET`,
cycles, 3-D spans inside a definition, and spill syntax rejected by the
underlying tokenizer remained unresolved at the consuming formula. FormulaFence
does not try to infer those paths. The public Foresight cap-table workbook was
re-profiled after the change with the same 4,228 formula cells, 36 `INDIRECT`
cells, zero unresolved formula-reference cells, and one parser warning. This
is a graph-coverage validation, not a claim to calculate Excel results.

## LET and inline LAMBDA scope — 2026-07-24

FormulaFence 0.8.0 distinguishes formula-local variables from workbook names.
[Microsoft documents that `LET` names apply only within the function's
scope](https://support.microsoft.com/en-us/excel/functions/let-function) and
that [LAMBDA parameters apply to its final calculation](https://support.microsoft.com/en-us/excel/functions/lambda-function).
The parser now follows those scopes without evaluating formulas, including
nested `LAMBDA` expressions inside higher-order functions.

The controlled fixture used `=LET(rate,Inputs!B2,amount,Inputs!B3,amount*(1-rate))`.
It produced real edges from both inputs to the calculation and zero unresolved
tokens; changing the rate reached both that calculation and its dashboard
output. Unit coverage also reproduces Microsoft's `LET` example, nested
shadowing, an inline LAMBDA call, and `REDUCE(...,LAMBDA(...))`. The public
Foresight workbooks contain no `LET` or `LAMBDA` formulas, so they remain a
compatibility regression check rather than evidence for the new syntax.

This is lexical static inspection, not an Excel evaluator. Spilled ranges and
arbitrary custom-function calls remain explicit limits.

## Named LAMBDA calls and OOXML serialization — 2026-07-24

FormulaFence 0.9.0 expands a named `LAMBDA` call only when the complete
definition is statically visible and internal. [Microsoft documents that a
LAMBDA moved into Name Manager becomes a reusable named function, callable like
a native Excel function](https://support.microsoft.com/en-us/excel/functions/lambda-function).
The externally maintained
[Vertex42 LAMBDA Library template](https://www.vertex42.com/lambda/templates.html)
was downloaded locally for compatibility validation only and is not bundled
with FormulaFence. The inspected file had SHA-256
`24e62d67f177ca02f9f8b6dc0381ccee88f12ba8ac9a6902ab7f25a98d0f5b71`.

| Measure | Result |
| --- | ---: |
| Sheets | 11 |
| Non-empty cells | 8,284 |
| Formula cells | 933 |
| Defined names | 121 |
| Top-level named LAMBDAs | 116 |
| Statically resolved named LAMBDAs | 99 |
| Unsafe or unsupported named LAMBDAs left visible | 17 |
| Formula cells with unresolved coverage notes | 36 |

The template uses the OOXML spellings `_xlfn.LAMBDA`, `_xlpm.`, and `_xlop.`,
and stores formula-defined names without a leading `=`. FormulaFence recognizes
those forms, including nested named-function calls. The 99 safe definitions in
this library have no static worksheet-cell inputs, so they resolve to an empty
internal dependency set; the remaining 17 stay explicit because their bodies
contain unsupported or non-static constructs. This is expected fail-closed
coverage behavior, not a judgment about the library.

The controlled graph fixture uses the same serialized notation for a
`ToCelsius` function that reads `Inputs!B2`, an `AdjustedCelsius` function that
calls it and reads `Inputs!B3`, and a formula-defined name that calls the latter.
Changing `Inputs!B2` reached all three model callers and the dashboard output.
Worksheet-local functions with the same name correctly shadowed workbook-level
functions, while dynamic and recursive named LAMBDAs remained unresolved at the
call site. No formula was evaluated during either check.

## Dynamic-array spill references and tokenizer coverage — 2026-07-24

Microsoft documents `A1#` as a reference to the whole spilled range rooted at
`A1`, whose size can grow or shrink. Its
[array-formula guidance](https://support.microsoft.com/en-us/excel/guidelines-and-examples-of-array-formulas)
uses `=D9#` as the equivalent of a concrete output range. Excel-compatible
writers do not necessarily store that display syntax verbatim:
[XlsxWriter documents](https://xlsxwriter.readthedocs.io/working_with_formulas.html)
that `F2#` is emitted as `ANCHORARRAY(F2)` in OOXML.

FormulaFence 0.10.0 accepts both a direct internal A1 anchor such as
`=SUM(Inputs!B2#)` and OOXML `_xlfn.ANCHORARRAY(Inputs!B2)`. It adds an anchor
edge to the graph but records the spill token at the consumer. This is a
deliberate partial edge: changing the anchor formula or its visible inputs
reaches consumers, while the variable spill extent and potential blocking cells
remain coverage limits. Formula-defined names containing a spill reference are
not expanded, so callers continue to receive an unresolved coverage note rather
than silently inheriting a partial graph.

The controlled fixture has a literal spill consumer and an OOXML-style consumer.
Changing the first anchor's `SEQUENCE` formula reached that consumer and its
dashboard output; the profile listed both spill sites and no unresolved or
tokenization-failure cells. An interoperability workbook generated with the
independently maintained XlsxWriter 3.2.9 package had SHA-256
`2d650855c229b2901dd0885242fcc1941d817990bd8acd69d340a76c6c93aa64` and
stored these exact worksheet formulas:

| Cell | Stored formula |
| --- | --- |
| `F2` | `_xlfn.UNIQUE(B2:B5)` |
| `H2` | `_xlfn.ANCHORARRAY(F2)` |
| `J2` | `COUNTA(_xlfn.ANCHORARRAY(F2))` |

The profile found two spill-reference cells (`H2`, `J2`), zero unresolved
references, and zero tokenizer failures; `F2` had both spill consumers as
direct dependents. The compatibility fixture is generated locally and is not
bundled with FormulaFence. It verifies OOXML serialization and graph behavior,
not Excel calculation results.

Finally, an unsupported malformed form such as `=SUM(Inputs!B2#1)` now appears
as a tokenizer-failure cell, emits `FF016` when newly introduced, and can be
blocked by `no_new_tokenization_failures`. This ensures a parser failure cannot
silently erase a formula's dependency coverage.

## Explicit implicit intersection — 2026-07-24

Excel documents `@` as explicit implicit intersection: a range contributes the
cell on the formula's row or column, while an array contributes its top-left
value. Its [Formula versus Formula2 guidance](https://learn.microsoft.com/en-us/office/vba/excel/concepts/cells-and-ranges/range-formula-vs-formula2)
also distinguishes old implicit-intersection evaluation from modern array
evaluation. The literal `@` is not necessarily what reaches an `.xlsx` file:
[XlsxWriter documents](https://xlsxwriter.readthedocs.io/working_with_formulas.html)
that Excel stores the persisted form as `SINGLE()` / `_xlfn.SINGLE()`.

FormulaFence 0.11.0 records literal direct `@A1:A3`, `@` applied to functions,
and persisted `_xlfn.SINGLE(...)`; profiles list each site and a newly added one
emits `FF017`. For a direct static A1 range with an unambiguous formula-location
intersection it adds only the selected cell edge. This is deliberately narrower
than formula evaluation: complex function results keep their visible static
inputs, while
formula-defined names containing explicit intersection remain unresolved at a
call site because the caller position changes the selection.

An interoperability workbook generated with independently maintained XlsxWriter
3.2.9 had SHA-256
`de7ee24b194ceb1c58ace589d3725876bad2aa3d4e45c5b1f9cba428d3837067` and
stored these formulas as one-cell array formulas:

| Cell | Stored formula | Selected dependency |
| --- | --- | --- |
| `B2` | `_xlfn.SINGLE(A1:A3)` | `A2` |
| `B3` | `_xlfn.SINGLE(A1:A3)` | `A3` |

FormulaFence profiled the workbook with three formula cells, two
implicit-intersection sites, no dependency from `A1` to either `SINGLE` caller,
and a direct `B2 → C2` downstream edge. The fixture is generated locally and
not bundled; it verifies OOXML serialization and static graph behavior, not
Excel's calculated values. Separate unit fixtures cover literal `@A1:A3`,
serialized `SINGLE`, horizontal and rectangular direct ranges, string safety,
unsupported `@A1#`, dynamic `@OFFSET`, table `[@Column]` separation, and
formula-defined-name containment.

## Fixed CSE and observed dynamic-array output aliases — 2026-07-24

Microsoft distinguishes a legacy Ctrl+Shift+Enter array's fixed output range
from a dynamic array whose spill can resize. [XlsxWriter documents the same
distinction](https://xlsxwriter.readthedocs.io/working_with_formulas.html):
`write_array_formula()` writes a static CSE range, while
`write_dynamic_array_formula()` writes a dynamic array.

Two otherwise identical workbooks were generated locally with independently
maintained XlsxWriter 3.2.9. Each had `Inputs!A1:A3`, an array formula at
`Model!B1`, an ordinary `=B2*10` consumer at `Model!C2`, and a cross-sheet
`=SUM(Model!B2:B3)` consumer at `Dashboard!B2`. The fixtures are not bundled.

| Form | Baseline SHA-256 | Stored anchor | FormulaFence result |
| --- | --- | --- | --- |
| CSE | `fdf2dba62f5390cfc88137470b22709415651d88b9ced9a116fe5ba8e601ca42` | `<f t="array" ref="B1:B3">LEN(Inputs!A1:A3)</f>` | Fixed legacy range `B1:B3`; anchor reaches `Model!C2` and `Dashboard!B2`. |
| Dynamic | `55155034f7115a417cb26ba815890d5bad7540a5f90c26ec5a1de93cd91a45c4` | `<c r="B1" cm="1"><f t="array" ref="B1:B3">…` plus `XLDAPR` `fDynamic="1"` metadata | Observed range `B1:B3`; anchor reaches `Model!C2` and `Dashboard!B2`. |

Changing `Inputs!A2` in the XlsxWriter dynamic candidate (SHA-256
`81d06efe44c9b023d5d5f2b6f349fd9eb5a9b497997d3d5836bb643be746b3a1`)
produced the exact paths
`Inputs!A2 → Model!B1 → Model!C2` and
`Inputs!A2 → Model!B1 → Dashboard!B2`; the CSE fixture produces the same
paths. The dynamic links are explicitly **observed**, not fixed: a profile
records the current OOXML range and each formula reading a non-anchor member,
but FormulaFence never predicts a later spill size or blocker. A separate
`B1:B3`-to-`B1:B4` dynamic fixture with an unchanged `=B4*10` at `Model!C4`
emitted `FF019` at `Model!C4`, because the new observed extent reached that
formula; it did not emit `FF018` for a fixed-range change. A compact-range
fixture declared `B1:XFD1048576`; FormulaFence retained eight stored cells
while linking known consumers, demonstrating that it does not materialize an
output node for every result cell. This validates OOXML classification and
static graph behavior, not Excel calculation results.

## Data-validation controls — 2026-07-24

Microsoft documents data validation as a worksheet control that can restrict
entries, show input guidance, and show an error alert. An interoperability
workbook generated locally with independently maintained XlsxWriter 3.2.9 had
SHA-256 `1e0e94a26e521b8a4e80214e0ad03ea1bc8f9e5e34286ce494a54b35e5c33132`.
It contained a list control targeting `Inputs!B2:B1048576` with
`Limits!$A$2:$A$4` as its source and a decimal control targeting
`Inputs!C2:C100` with lower and upper bounds from `Limits`.

The serialized XlsxWriter rules omitted the schema-default `operator=between`
and `errorStyle=stop` attributes, and stored criteria without a leading `=`.
A matching locally generated openpyxl workbook used explicit defaults and
leading-equals criterion text. FormulaFence produced equal data-validation
snapshots for both workbooks and no `FF020`, demonstrating that these harmless
writer representations do not create a control diff. The controlled suite also
splits one identical rule into separate OOXML target groups without a diff.
Changing the target range or disabling the error alert emitted `FF020`; the
`no_data_validation_changes` policy emitted `FFP020`.

The profile retained two compact rules and two target ranges rather than
materializing the full-column target as cells. It deliberately redacted the
criteria and prompt/error text from the profile, while the local JSON diff kept
full before/after evidence. This validates OOXML representation and
change-detection behavior, not whether Excel will evaluate a validation formula
or accept a particular user entry. The scope follows Microsoft's
[data-validation guidance](https://support.microsoft.com/en-US/Excel/get-started/apply-data-validation-to-cells)
and [openpyxl's documented validation range model](https://openpyxl.readthedocs.io/en/stable/validation.html).

## Conditional-formatting controls and precedence — 2026-07-24

Microsoft documents that conditional-formatting rules are evaluated by
precedence, that conflicts use the higher rule, and that `Stop If True` prevents
lower-priority rules from taking effect. The OOXML model makes that priority
global to the worksheet, not merely to one target range. FormulaFence therefore
records a compact target group for each rule and its normalized precedence,
rather than treating a color change as ordinary cell style noise.

We profiled Microsoft's public [Conditional Formatting examples workbook](https://support.microsoft.com/en-us/excel/use-conditional-formatting-to-highlight-information-in-excel), downloaded locally for compatibility validation only and not bundled with FormulaFence. Its SHA-256 was
`b73a6de84668d9f728967f31bd3240eba2d766b667021a0adb378a19df70f887`.
The workbook had 16 sheets, 1,677 non-empty cells, 346 formula cells, **38**
conditional-formatting rules, and 38 compact target ranges. The inventory found
`aboveAverage`, `cellIs`, `colorScale`, `containsText`, `dataBar`,
`duplicateValues`, `expression`, `iconSet`, `timePeriod`, and `top10` rules.
It retained no raw criterion formula or text rule in the profile. Its one
parser warning concerned an unrelated header/footer parse limitation, not
conditional formatting.

Two independently maintained XlsxWriter 3.2.9 fixtures exercised the harder
extension boundary. The baseline SHA-256
`98d1aba217995085824ccf91129bf729027c69732fffa3b84936c636a6e59942`
used an Excel-2010 data bar with a black axis; the candidate SHA-256
`a7f84aa204af9f9012b3cdfba91075e25499db8a078e32b3981766cb4f870d9d`
changed only that axis to red. openpyxl issued the same unsupported-extension
warnings for both files, but FormulaFence retained one worksheet extension
fragment in each snapshot and emitted `FF021` for the axis-color change.
Equivalent fixtures with a different extension GUID, explicit false boolean
defaults, leading `=` criteria, non-contiguous raw priority numbers, and a
reordered `dxfs` style table produced no conditional-formatting diff.

This validates OOXML control and extension change detection, not Excel display
calculation. FormulaFence does not decide whether a condition is true, map a
relative formula over every target cell, or emulate the final interaction with
manual formats and all overlapping rules. The policy control is
`no_conditional_formatting_changes` (`FFP021`).

## Protection controls and redaction — 2026-07-24

Microsoft distinguishes workbook structure protection and worksheet editing
controls from file encryption. FormulaFence 0.16.0 therefore treats them as a
reviewable operational surface, not proof that a workbook is secure. We
profiled Excel Easy's public [Protect Sheet example](https://www.excel-easy.com/examples/protect-sheet.html)
and [Protect Workbook example](https://www.excel-easy.com/examples/protect-workbook.html),
downloaded only for local compatibility validation and not bundled with the
project. Their SHA-256 values were respectively
`579328b0579140919ae0ccc468d1d4ebb4b840229d47d25a1a038f44415af855` and
`2cd04503595ecb6cfb425f26890a6edf274bbb11d3ec159fb1983b2ace280330`.

The sheet example has active worksheet protection and a modern SHA-512 verifier
with 100,000 iterations; FormulaFence reported one protected sheet, no parser
warning, and no raw verifier/hash/salt value in the profile. The workbook
example has active structure protection with the same modern verifier shape;
FormulaFence reported the structure lock and no protected-sheet declaration.
This checks current Excel-produced OOXML representation, not password strength
or access enforcement.

Controlled fixtures additionally covered legacy verifiers, a protected range
with a credential and security descriptor, explicit unlocked and hidden cell
styles, row/column style assignments, and a protected chart sheet. Omitted and
explicit schema-default worksheet action flags compared equal. Changing only a
modern verifier, a protected-range name, a security descriptor, a structure
lock, an action lock, or an unlocked cell emitted `FF022`; the policy emitted
`FFP022`. A redaction check read the raw OOXML verifier, salt, protected-range
name, and security descriptor, then verified that none appeared in JSON,
Markdown, or SARIF artifacts. This validates comparison and data-minimisation
behavior, not Excel authentication, file encryption, rights management, or the
complete style-cascade result.

## Power Query Data Mashup controls and redaction — 2026-07-24

FormulaFence 0.18.0 was profiled against DecimalTurn's public
[Power Query `.xlsm` example](https://github.com/DecimalTurn/VBA-StackOverflow-Demos/tree/7100c961fec96435adfb402fdd7e6c59c0af4f43/demos/answers/79461277),
downloaded only for local compatibility validation and not bundled with the
project. The downloaded `79461277.xlsm` had SHA-256
`eed5eb994b426fde70d68e20f59bab2e0c02dd5fe6a620a4d9ffff9777bd1bc9`.

The workbook carries a `customXml/item1.xml` `DataMashup` part, one embedded
`Section1.m` formula document, Data Mashup metadata and permissions, and a
QueryTable reached through an Excel table relationship rather than a direct
worksheet relationship. FormulaFence reported the safe structural counts,
formula-firewall state, and linked query-table control without printing its M
formula, query/source identity, or metadata values. The file's pre-existing
unmodelled Connections-container warning remained a visible coverage note.
On a local, non-distributed copy, a change confined to `Section1.m` emitted
only `FF024` with a private formula-material flag; the inserted M-code marker
did not appear in the JSON report.

Controlled raw-OOXML fixtures changed M text, stable metadata, and firewall
settings, producing `FF024` and `FFP024` under `no_power_query_changes`.
Separate changes to only `sqmid` telemetry, result-time metadata, and the
user-bound permission-binding blob did not create a query-control diff. The
fixtures placed synthetic URLs, query names, package content, embedded content,
metadata values, IDs, and permission material in the raw package, then verified
that none entered JSON, Markdown, or SARIF. This validates static comparison
and redaction—not M execution, connection refresh, source trust, or returned
data correctness.

## Worksheet embedded-control and OLE guardrails — 2026-07-24

FormulaFence 0.24.0 was validated with controlled raw-OOXML `.xlsx` packages
following a worksheet control chain: one `<control>` bound to an
`xl/activeX/activeX1.xml` persistence part and its raw binary target, nested
`controlPr` markup bound to an `xl/ctrlProps/ctrlProp1.xml` form-control part,
and both one embedded and one externally linked `<oleObject>`. The raw binary
and OLE payloads were harmless fixture bytes; the workbook was never opened in
Office. FormulaFence only inspected bounded package parts before the ordinary
workbook reader loaded the file.

Changing private control macro/link material, OLE auto-load behavior, ActiveX
class/license material, a form-control formula range, and a direct OLE target
emitted `FF029` with the corresponding private material flags. A change to only
the raw OLE payload emitted `FF029` with only the private payload-material flag.
Synthetic control names, macros, class/license values, formulas/ranges, OLE
program/link values, relationship targets, and payload markers were verified
absent from JSON, Markdown, and SARIF output. The
`no_worksheet_embedded_control_changes` policy produced `FFP029`.

The controlled suite also covered `mc:AlternateContent` duplicate control
markup, relationship-ID-only rewrites, equivalent internal target spellings, an
unexpected ActiveX root, oversized XML/raw payloads, and deliberately lowered
XML and payload byte/part budgets. Duplicate fallback markup was not
double-counted; identifier and spelling rewrites did not produce `FF029`; an
ordinary worksheet with no relevant relationship did not consume the control
XML budget; malformed or bounded-out material remained explicitly visible as a
coverage warning. Production limits are 16 MiB per relevant XML part, 64 MiB
per workbook, and 512 parts; direct payload hashes are limited to 32 MiB per
part, 64 MiB per workbook, and 512 parts. This validates static
comparison and data minimisation—not ActiveX initialization, OLE/package
deserialization, Office rendering, event dispatch, source trust, or embedded
payload behavior. The fixture shape follows Microsoft's guidance for [sheet
ActiveX controls](https://learn.microsoft.com/en-us/office/vba/excel/concepts/controls-dialogboxes-forms/using-activex-controls-on-sheets),
the [`ocx` persistence schema](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oe376/b30a660a-95eb-4716-b201-a46aae788610),
and [form-control properties](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/3d054a6d-4f94-4082-837a-f939fd8d4a45).

## Legacy VML worksheet-control guardrails — 2026-07-24

FormulaFence 0.25.0 was validated with controlled raw-OOXML `.xlsx` packages
containing a standard worksheet-to-`vmlDrawing` relationship. The VML drawing
contained a `Button` with `FmlaMacro`, a `Drop` with `FmlaLink`, `FmlaRange`,
and `FmlaTxbx`, a `GBox` with `FmlaGroup`, a `Pict` with `FmlaPict`, an image
relationship, and a separate `Note` comment shape. All macros, formula bindings, captions, relationship
targets, and image bytes were harmless fixture values; no workbook was opened
in Office.

Changing the VML macro/binding material and direct presentation relationship
emitted `FF029` with private legacy-VML definition and relationship flags.
Profiles exposed only safe counts for VML parts, controls, macro assignments,
cell links, source ranges, and camera ranges. Synthetic macro names, bindings,
captions, note text, relationship targets, and image markers were verified
absent from JSON, Markdown, report, and SARIF output. The same
`no_worksheet_embedded_control_changes` policy produced `FFP029`.

The controlled suite also changed only the adjacent VML `Note` comment,
renumbered worksheet and VML relationship IDs, rewrote equivalent internal
target spellings, corrupted the VML root, and reduced the XML per-part limit.
Comment-only edits and identifier/path spelling churn produced no control
finding; malformed or bounded-out VML material remained visible as a coverage
warning. This validates static, relationship-aware comparison and data
minimisation—not VML rendering, comment parsing, macro execution, formula
evaluation, image decoding, source trust, or event behavior. The fixture uses
the documented VML [`ClientData`](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.vml.spreadsheet.clientdata?view=openxml-3.0.1)
structure and Microsoft’s notes on [`FmlaMacro`](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oi29500/fdd83507-7a57-4bf1-b844-66f551ee55b9),
[`FmlaRange`](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oi29500/3d0c9716-88c5-4af3-b63a-feef60b8ebd8),
and camera ranges.

## DrawingML chart definitions and cached presentation data — 2026-07-24

FormulaFence 0.26.0 was validated with controlled raw-OOXML `.xlsx` packages
starting from a generated bar chart and its standard worksheet-to-drawing-to-
chart relationship chain. The fixture then added numeric and string series
caches, a chart `userShapes` overlay with private text, and an overlay image
relationship. All chart text, formulas, cached values, relationship targets,
and image bytes were harmless fixture values; the workbook was never opened in
Office. FormulaFence only inspected bounded package parts before the ordinary
workbook reader loaded the file.

Changing only private chart definition material, an overlay shape, or a direct
related presentation payload emitted `FF030` with the matching private
definition, overlay, relationship, or payload-material flag. Changing only a
cached series value emitted `FF030` with only the cached-series-material flag.
Profiles exposed safe structural counts for host sheets, chart parts, series,
references, caches, overlays, relationships, and bounded payloads. Synthetic
formula text, cached values, titles, shape text, relationship targets, XML, and
payload markers were verified absent from JSON, Markdown, and SARIF output. The
`no_chart_definition_changes` policy produced `FFP030`.

The controlled suite also covered a chartsheet chart chain, relationship-ID-
only rewrites, equivalent internal target spellings, an unexpected chart root,
an externally targeted overlay relation, and deliberately lowered chart-XML and
related-payload budgets. Identifier and path-spelling churn produced no chart
finding; the external target was counted without being fetched or exposed;
malformed or bounded-out material remained explicitly visible as a coverage
warning. Production chart
and overlay XML reads are bounded to 16 MiB per part, 64 MiB per workbook, and
512 parts; direct related payload hashes are limited to 32 MiB per part, 64 MiB
per workbook, and 512 parts. This validates static, relationship-aware
comparison and data minimisation—not series-formula calculation, chart
rendering, chart-to-cell impact analysis, external-target retrieval, media or
embedded-package parsing, source trust, or modern `chartEx`/nested-chart
semantics. The fixture follows the OOXML [chart-part model](https://c-rex.net/samples/ooxml/e1/Part1/OOXML_P1_Fundamentals_Chart_topic_ID0ELZLM.html),
the documented [number-reference cache](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.charts.numberreference?view=openxml-3.0.1),
and the [chart user-shapes relationship](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.charts.usershapesreference?view=openxml-2.20.0).

## PivotTable definitions and cached report material — 2026-07-24

FormulaFence 0.27.0 was validated with controlled raw-OOXML `.xlsx` packages
that openpyxl can load and preserve but does not author. The fixture follows a
standard worksheet → PivotTable → pivot-cache-definition → cache-records chain,
plus the workbook-level cache declaration. It includes a report location,
fields, items, data field, cache schema, shared cache items, source definition,
and raw cache records. All labels, source ranges, item values, and record values
were harmless redaction sentinels; the workbook was never opened in Office.

Changing only private layout, cache-schema, shared-item, and cache-record
material emitted `FF031` with the matching private flags. Moving the direct
cache-record relationship emitted relationship and raw-payload flags. Changing
only `refreshOnLoad` emitted the existing `FF023` and no `FF031`, preserving the
source/refresh boundary. Profiles exposed only structural counts, and synthetic
names, source ranges, item values, XML, and record markers were verified absent
from JSON, Markdown, ordinary reports, and SARIF. The
`no_pivot_table_definition_changes` policy produced `FFP031`.

The controlled suite also renumbered relationship IDs and cache IDs, rewrote
equivalent internal target spellings, corrupted the PivotTable root, and
lowered both XML and cache-record limits. Identifier and path-spelling churn
produced no finding; malformed or bounded-out material remained explicitly
visible as a coverage warning. Production PivotTable/cache-definition XML reads
are bounded to 16 MiB per part, 64 MiB per workbook, and 512 parts; raw cache
record hashes are bounded to 32 MiB per part, 64 MiB per workbook, and 512
parts. The suite also made the underlying reader's record parser fail if called;
FormulaFence still loaded the workbook by detaching cache-record relationships
in a temporary reader copy. This validates static, relationship-aware
comparison and data minimisation—not PivotTable refresh/calculation/rendering,
PivotTable-to-cell impact analysis, external-target retrieval, source trust, or OLAP,
extension-list, and slicer semantics. The fixture follows the OOXML [Pivot
Table Part](https://c-rex.net/samples/ooxml/e1/Part1/OOXML_P1_Fundamentals_Pivot_topic_ID0ELLAM.html),
[Pivot Cache Definition Part](https://c-rex.net/samples/ooxml/e1/Part1/OOXML_P1_Fundamentals_Pivot_topic_ID0E1TAM.html),
and [Pivot Cache Records Part](https://c-rex.net/samples/ooxml/e1/Part1/OOXML_P1_Fundamentals_Pivot_topic_ID0EV2AM.html).

## Office Web Add-in task-pane controls and redaction — 2026-07-24

FormulaFence 0.23.0 was validated with controlled raw-OOXML `.xlsx` packages
following the documented workbook-to-task-pane-to-web-extension chain:
`xl/_rels/workbook.xml.rels` declared `taskpanes.xml`, the task-pane part bound
one `webextension` definition, and the definition included a store reference,
alternate reference, private property, binding, snapshot relationship, and
`Office.AutoShowTaskpaneWithDocument=true`. The fixture was never opened in
Office, and FormulaFence only read its bounded package XML before the ordinary
workbook reader loaded the file.

Changing **only** the auto-show property emitted `FF028` with a private
`web_extension_definition_material_changed` flag even though the workbook's
cells, VBA payload, RibbonX surface, and task-pane counts were otherwise
unchanged. Synthetic add-in IDs, store references, property values, binding
values, XML, snapshot target, and external relationship endpoint were verified
absent from JSON, Markdown, and SARIF output. The
`no_office_web_addin_changes` policy produced `FFP028`.

The controlled suite also covered task-pane configuration and direct
relationship changes, relationship-ID-only rewrites, and equivalent internal
target spellings. Identifier and spelling rewrites produced no Web Add-in
finding; changed task-pane configuration or a snapshot relationship emitted
`FF028`. An unexpected definition root and deliberately lowered per-part,
aggregate-byte, and part-count limits each surfaced an explicit coverage
warning and remained diff-visible. Production task-pane and web-extension XML
reads are bounded to 16 MiB per part, 32 MiB per workbook, and 64 parts. This
validates static comparison and data minimisation—not add-in installation,
manifest retrieval, Office.js execution, task-pane rendering, source trust, or
worksheet-scoped Web Add-in markup. The fixture shape follows Microsoft's
[Taskpane Web Extension File](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-owexml/3d04f8ce-65f2-4dc3-bafa-636d0a7e41a1),
[Web Extension](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-owexml/56fe5a64-dd6d-422c-beac-19d72dd10ade),
and [automatic task-pane sample](https://learn.microsoft.com/en-us/samples/officedev/office-add-in-samples/excel-add-in-create-spreadsheet-from-web-page/).

## Office RibbonX custom UI controls and redaction — 2026-07-24

FormulaFence 0.22.0 was validated with controlled raw-OOXML `.xlsx` packages
using the documented root-package Ribbon Extensibility relationship to a
`customUI` part. Each fixture had a custom tab, group, and button; an `onLoad`
callback; an `onAction` callback; and one explicit image relationship. The
fixture was never opened in Office, and FormulaFence only read its bounded
package parts before the ordinary workbook reader loaded the workbook.

A change to **only** the button's private `onAction` callback emitted `FF027`
with a private `ribbon_definition_material_changed` flag even though every
public control and callback count stayed the same. This is the exact blind
spot the guard is intended to close: a workbook UI callback can change without
a worksheet-cell or VBA-payload diff. Synthetic callback names, labels, XML,
image target, and image payload markers were verified absent from JSON,
Markdown, and SARIF output. `no_ribbon_customization_changes` added `FFP027`.

The controlled suite also covered the 2006 schema plus the 2009/07 and
2007/10 Office 2010-era roots, an image-target change, relationship-ID-only
rewrites, and equivalent internal target spellings. Identifier and spelling
rewrites produced no RibbonX finding; a changed image relationship produced
`FF027`. An unexpected root or a deliberately lowered part-size limit produced
an explicit coverage warning and remained diff-visible. Separate lowered
aggregate byte and part-count budgets also left an explicit coverage warning.
Production reads are bounded to 16 MiB per part, 32 MiB per workbook, and
eight parts. This validates static
comparison and redaction—not callback execution, Office UI rendering, image
decoding, source trust, or runtime macro behavior. The fixture shape follows
Microsoft's [Ribbon Extensibility Part](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-customui/52faf7b6-fecc-48d9-96db-ee80a631a5ac)
and [Ribbon and Backstage Customizations
part](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-customui2/452a58ae-cb0a-4926-83f8-fb1cbaa6114c)
specifications.

## Excel 4.0 / XLM macro-sheet controls and redaction — 2026-07-24

FormulaFence 0.21.0 was validated with a controlled, macro-enabled OOXML
package shaped according to Microsoft's documented Macro Sheet and
International Macro Sheet relationship/content-type definitions. It included a
very-hidden macro sheet, two macro formula cells, one internal embedded-object
relationship, one external linked-object relationship, and one embedded-package
relationship; no VBA binary was present.

As an independent compatibility check, SheetJS Community Edition 0.20.3 read
the package as a macro sheet (`!type: "macro"`) while reporting no VBA blob.
The ordinary Python workbook reader retained the sheet tab but exposed zero
formula cells. FormulaFence's raw preflight reported one macro-sheet part, two
macro formula cells, one external relationship, two OLE-object relationships,
one package relationship, and two fingerprinted internal related parts with no
parser warnings.

On a local, non-distributed copy, changing private macro formula text, cell
material, an embedded-object program identifier, related-part targets, and the
hidden state emitted `FF026` with private program, relationship, and
workbook-binding flags; `no_xlm_macro_sheet_changes` added `FFP026`. The
synthetic command arguments, values, identifiers, targets, and extension
payload did not appear in JSON, Markdown, or SARIF. Rewriting only relationship
identifiers did not create an XLM control diff. An unexpected macro-sheet root
produced an explicit coverage warning and still produced `FF026`.

A separate local mutation changed only the bytes of an internal embedded OLE
payload while retaining the macro XML and every relationship target. It emitted
`FF026` with the private `related_part_payload_material_changed` flag; neither
the baseline nor candidate bytes or paths entered JSON, Markdown, or SARIF.
The regression suite also lowers the per-part byte, aggregate byte, and part
count budgets to confirm that each bound emits a coverage warning without
parsing payloads. This validates static package comparison and redaction—not
XLM execution or emulation, embedded-object execution or parsing, source
trust, or Excel runtime behavior.

## External-link package controls and redaction — 2026-07-24

FormulaFence 0.19.0 was profiled against Apache POI's public
[`ref2-56737.xlsx` fixture](https://github.com/apache/poi/blob/0d6d4872c491b1f230f51c6878e57407c60ae697/test-data/spreadsheet/ref2-56737.xlsx),
downloaded only for local compatibility validation and not bundled with the
project. The downloaded file had SHA-256
`7ee59e3710f1aa75cbc6585ac6548f8ce3b3bca04a4cbebb079c455773bce344`.

The workbook has two external-workbook package parts. FormulaFence reported
two package parts, five external sheet names, four external defined names,
five cached sheets, seven cached cells, and one cached refresh error, with no
parser warnings. Its profile did not print either workbook target, sheet or
defined name, or cached value.

On a local, non-distributed copy, changing only an `externalLink` relationship
target emitted `FF025` with a private source-material flag; the inserted target
marker did not appear in JSON, Markdown, or SARIF. Controlled raw-OOXML
fixtures additionally covered external-workbook, DDE, and OLE definitions,
workbook-declaration rebinding, cached material, item flags, and opaque
extension data. They produced `FF025` and `FFP025` under
`no_external_link_package_changes` while verifying that synthetic targets,
names, services, program IDs, cached values, and extension payloads were never
serialized. This validates static package comparison and redaction—not opening
or executing external-workbook, DDE, or OLE links, source trust, or returned
data correctness.

## External-data refresh controls and redaction — 2026-07-24

FormulaFence 0.17.0 was profiled against Mullins Lab's public
[external-data workbook](https://github.com/MullinsLab/excel-external-data),
downloaded only for local compatibility validation and not bundled with the
project. The downloaded `external-data-blank.xlsx` had SHA-256
`b194aa281d64f1b5cf7f953a328adca211d67245c6b2d0fe64b5245c352a7b68`.

The file carries an OOXML Connections part and a linked QueryTable part that the
previous inventory did not expose. FormulaFence reported one text-import
connection and one linked query-table control, including their safe refresh,
background, cache, and growth behavior metadata. It did not emit the
connection's name or source filename. The existing workbook-reader warning for
an unsupported extension remained visible as a coverage note rather than being
silently discarded.

Controlled raw-OOXML fixtures then covered workbook-wide refresh flags; OLE DB
and web connections; refresh-on-open, periodic, background, password/cache,
connection-file, SSO-presence, and parameter-change controls; a linked query
table; and an external pivot cache. Explicit and omitted connection defaults
compared equal. Changing source material, an identity, a refresh setting, or a
linked control emitted `FF023`; the policy emitted `FFP023`. Redaction tests
placed synthetic paths, URLs, passwords, connection strings, commands, names,
parameter values, SSO IDs, and extension payloads in the raw package, then
verified that none appeared in JSON, Markdown, or SARIF. This validates static
control comparison and data minimisation—not a live refresh, source trust, or
returned data correctness.

## Public structured-reference example — 2026-07-24

FormulaFence 0.6.0 was also profiled against the public
[Excel Easy structured-reference example](https://www.excel-easy.com/examples/structured-references.html).
The downloaded workbook was used locally for compatibility validation only and
is not bundled with FormulaFence. Its profile reported one Excel table, 79
non-empty cells, 17 formula cells (15 with structured references), and no
unresolved formula-reference tokens.

On a local, non-distributed copy, changing one table data cell traced **16
downstream formula cells**, including the table's total and an output formula
outside the table. This validates that FormulaFence turns the supported table
forms into real dependency edges while still reporting unsupported forms as
coverage notes.

## Current-row structured references — 2026-07-24

FormulaFence 0.5.0 adds context-bound table-row edges without evaluating a
formula. [Microsoft documents `@` and `#This Row` as references to the
formula's row](https://support.microsoft.com/en-us/excel/using-structured-references-with-excel-tables),
while noting that the same syntax in a header or total row returns an error. We
also checked the adjacent-cell case against
[ClosedXML's independently maintained structured-reference test](https://github.com/ClosedXML/ClosedXML/blob/4e89dcedd83cad553e84d2d97f77fc3d7deb630f/ClosedXML.Tests/Excel/CalcEngine/StructuredReferenceTests.cs),
which exercises `TableName[[#This Row],…]` from a cell beside the table on a
data row.

In a controlled local workbook, a `Sales` table used all three common
calculated-column spellings: `[@[Sales Amount]]`, `[Sales Amount]`, and
`Sales[[#This Row],[Sales Amount]]`. A neighboring cell used the qualified
`#This Row` form. The baseline had zero unresolved reference tokens. Replacing
the first row's input value produced exactly three downstream formula cells:
the matching calculated-column cell, the adjacent qualified-reference cell,
and the external `SUM(Sales[Value])` output. FormulaFence did not report the
other two table rows as impacted. This validates graph precision for the
supported subset; it does not claim to recalculate or certify Excel results.

## 3-D worksheet references — 2026-07-24

FormulaFence 0.6.0 adds static expansion for internal 3-D A1 references. The
[Microsoft 3-D-reference documentation](https://support.microsoft.com/en-us/excel/create-a-3-d-reference-to-the-same-cell-range-on-multiple-worksheets)
defines a reference such as `Sales:Marketing!B3` over every worksheet tab
between those endpoints, and specifies that inserting or moving tabs can change
the calculation.

In a controlled local workbook with `Jan`, `Feb`, `Mar`, and `Summary` tabs,
`=SUM(Jan:Mar!B2)` created dependency edges from all three period inputs to the
summary, with zero unresolved reference tokens. Changing `Feb!B2` reached the
summary. Inserting a new period tab between the endpoints also reached the
summary, while moving `Feb` outside the `Jan:Mar` span produced `FF014` and the
optional `no_3d_reference_scope_changes` policy produced `FFP014`. This is a
static graph validation, not a claim to calculate Excel results.

## Controlled local change

On a local, non-distributed copy, we replaced the formula in
`'5 - Exit Waterfall'!O6` with a numeric value and ran the starter policy. The
check identified a `formula_to_value` change, traced **330 downstream formula
cells**, and failed both the formula-override and default impact-limit controls
(`FFP001`, `FFP009`). The end-to-end check completed in approximately two
seconds in the release environment. This is a compatibility demonstration, not
a performance guarantee or an assertion about the source model's correctness.
