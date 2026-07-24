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
