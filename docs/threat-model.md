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
- It inventories sheet visibility, defined names, calculation settings, and the
  VBA payload. It does not yet diff chart definitions, PivotTables, Power Query,
  ordinary styles, protection settings, or every OOXML part.
- The tool preserves Excel formula text and uses a limited A1-reference
  normalizer for peer-pattern detection; it is not an Excel-compatible parser
  or calculation engine.

For high-stakes use, treat FormulaFence as one control among independent review,
recalculation in the approved spreadsheet engine, input reconciliation, and an
appropriately qualified model owner.
