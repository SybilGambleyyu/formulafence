# External validation notes

FormulaFence's test suite builds small fixtures that isolate individual risks.
Those tests are necessary but insufficient for confidence in an Office-file
reader, so each release should also be exercised on independently maintained
workbooks without copying their contents into this repository.

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
