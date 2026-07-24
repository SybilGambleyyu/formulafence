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

## Controlled local change

On a local, non-distributed copy, we replaced the formula in
`'5 - Exit Waterfall'!O6` with a numeric value and ran the starter policy. The
check identified a `formula_to_value` change, traced **330 downstream formula
cells**, and failed both the formula-override and default impact-limit controls
(`FFP001`, `FFP009`). The end-to-end check completed in approximately two
seconds in the release environment. This is a compatibility demonstration, not
a performance guarantee or an assertion about the source model's correctness.
