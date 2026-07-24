# Scope and threat model

FormulaFence is a static change-assurance layer. It answers whether an Excel
workbook's inspectable structure changed in a risky way; it does not certify
financial correctness or replace model review.

## Safety properties

- Workbook content stays on the machine running FormulaFence. The CLI makes no
  network requests.
- It loads formulas as text with `data_only=False`; it does not calculate them.
- It never executes VBA, DDE, external links, Power Query, or add-in code.
- Macro payloads are reported by cryptographic hash only.
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

- Supported files are `.xlsx` and `.xlsm`; legacy `.xls` and password-protected
  workbooks are outside scope.
- Ordinary workbook and sheet-local names with static A1 destinations are
  resolved into the dependency graph. `INDIRECT`, `OFFSET`, named formulas,
  structured table references, dynamic array behavior, cube functions, add-ins,
  and custom functions cannot always be statically resolved. FormulaFence flags
  newly introduced unresolved tokens and `INDIRECT`/`OFFSET` use, but does not
  fabricate dependencies for them.
- Explicit external-workbook references are detected. References assembled from
  text or macro code are not.
- It inventories sheet visibility, defined names, calculation settings, and the
  VBA payload. It does not yet diff chart definitions, PivotTables, Power Query,
  styles, data validations, protection settings, or every OOXML part.
- The tool preserves Excel formula text and uses a limited A1-reference
  normalizer for peer-pattern detection; it is not an Excel-compatible parser
  or calculation engine.

For high-stakes use, treat FormulaFence as one control among independent review,
recalculation in the approved spreadsheet engine, input reconciliation, and an
appropriately qualified model owner.
