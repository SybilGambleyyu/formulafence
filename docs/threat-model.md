# Scope and threat model

FormulaFence is a static change-assurance layer. It answers whether an Excel
workbook's inspectable structure changed in a risky way; it does not certify
financial correctness or replace model review.

## Safety properties

- Workbook content stays on the machine running FormulaFence. The CLI makes no
  network requests.
- HTML review artifacts contain only inline styles and a fixed local filtering
  script; they load no remote assets and make no browser network requests.
  Workbook-derived text is HTML-escaped before rendering, so it is review
  evidence rather than executable page content. The report still reflects the
  sharing boundary selected by the caller; HTML is not a general secret scrubber.
- It loads formulas as text with `data_only=False`; it does not calculate them.
- The single-workbook formula lint is static. Its copied-formula signal compares
  only already-loaded relative formula fingerprints and reports an interruption
  only after two matching immediate peers plus a third contiguous peer support
  the same copied pattern. Its aggregate-range signal accepts only a pure
  `SUM`, `AVERAGE`, `MIN`, `MAX`, or `COUNT` with one direct same-sheet,
  one-dimensional A1 range, followed on that row or column by a bounded run of
  at least two literal numeric cells before the aggregate formula. It ignores
  short/ambiguous patterns, computed/multi-range/named/external/3-D aggregate
  expressions, nonnumeric gaps, tokenizer failures, and all declared or
  unclassified array-formula territory. Its formula-protection signal accepts
  only an explicit direct-cell unlocked assignment for an ordinary formula on
  an active worksheet-protection declaration; it does not infer row, column,
  default-style, or allowed-edit-range precedence. Its calculation-freshness
  signal accepts only a workbook with at least one formula whose stored
  calculation properties explicitly combine `calcMode=manual` and
  `calcCompleted=false`; it does not infer a stale result from manual mode
  alone or claim a particular result is mathematically wrong. Its
  error-checking-suppression signal accepts only recognized stored per-range
  declarations and emits aggregate warning-category, suppression-rule, and
  target-range counts. It never exposes target ranges, decides whether a prompt
  would apply, or claims that a suppression was incorrect. Its Table
  calculated-column signal accepts only a stored scalar table master formula
  and an interior data cell bracketed by two immediate eligible formulas that
  match its fingerprint. It skips first/last data rows, array territory,
  uninspectable formulas, and longer or ambiguous exception runs; the stored
  master formula and table identity remain private. Its conditional-aggregate
  range-shape signal accepts only native `SUMIFS`, `COUNTIFS`, `AVERAGEIFS`,
  `MAXIFS`, and `MINIFS` calls (optionally with `@`) plus the exact OOXML
  `_xlfn.MAXIFS`/`_xlfn.MINIFS` serializations, with valid arity and relevant
  arguments that are each one bounded, internal direct A1 cell/range or
  whole-column reference. It then compares dimensions without calculation.
  Names, Tables, external/3-D/full-row/union references, computed or dynamic
  expressions, spills, implicit intersection, malformed formulas, explicit
  broken references, and all array territory stay outside the boundary. It
  emits only a location plus aggregate call and mismatched-range counts;
  formulas, range spellings, and Table identities remain private.
  Its `SUMPRODUCT` range-shape signal accepts only the unqualified native
  spelling (optionally with `@`) with at least two comma-separated arguments,
  each one bounded, internal direct A1 cell/range or whole-column reference.
  It compares dimensions without calculation. Names, Tables,
  external/3-D/full-row/union references, computed or dynamic expressions,
  spills, implicit intersection, malformed formulas, explicit broken
  references, and all array territory stay outside the boundary. It emits only
  a location plus aggregate call and mismatched-array counts; formulas, range
  spellings, and source sheets remain private.
  Its `MMULT` matrix-dimension signal accepts only the unqualified native
  spelling (optionally with `@`) with exactly two comma-separated arguments,
  each one bounded, internal direct A1 cell/range or whole-column reference.
  It compares the first array's column count with the second array's row count
  without calculation or inspecting cell values. Names, Tables,
  external/3-D/full-row/union references, computed or dynamic expressions,
  spills, implicit intersection, malformed formulas, explicit broken
  references, and all array territory stay outside the boundary. It emits only
  a location plus aggregate call and incompatible-matrix-pair counts; formulas,
  range spellings, and source sheets remain private.
  Its legacy-lookup return-index signal accepts only unqualified native
  `VLOOKUP` and `HLOOKUP` spellings (optionally with `@`) with exactly three or
  four comma-separated arguments. Its table argument must be one bounded,
  internal direct A1 cell/range or whole-column reference and its return index
  one direct positive integer literal. It compares a `VLOOKUP` index with table
  width or an `HLOOKUP` index with table height, without calculation or
  inspecting lookup/table values. Names, Tables, external/3-D/full-row/union
  references, computed or dynamic expressions, spills, implicit intersection,
  malformed formulas, nonliteral or nonpositive indices, explicit broken
  references, and all array territory stay outside the boundary. It emits only
  a location plus aggregate call and out-of-range-literal-index counts;
  formulas, range spellings, and source sheets remain private.
  Its `CHOOSE` literal-index signal accepts only the unqualified native spelling
  (optionally with `@`) with one through 254 nonempty value arguments and one
  direct bare nonnegative decimal index. It reports only a zero index or one
  above the supplied value-argument count, without inspecting selected values.
  Computed, signed, decimal, array, dynamic, malformed, explicit-broken-
  reference, and namespaced forms, plus all array territory, stay outside the
  boundary. It emits only a location plus aggregate call and out-of-range-
  literal-index counts; formulas, values, and source sheets remain private.
  Its `RANDBETWEEN` literal-bound signal accepts only the unqualified native
  spelling (optionally with `@`) with exactly two direct decimal integer
  literals, each optionally preceded by one unary `+` or `-`. It reports only
  when the bottom literal is greater than the top literal, without calculating
  a random value. Decimal/scientific, computed, reference, array, malformed,
  explicit-broken-reference, and namespaced forms, plus all array territory,
  stay outside the boundary. It emits only a location plus aggregate call and
  inverted-literal-bound counts; formulas, literal values, and source sheets
  remain private.
  Its `SUBTOTAL` function-code signal accepts only the unqualified native
  spelling (optionally with `@`) with a direct bare nonnegative decimal
  function number and one through 254 nonempty reference arguments. It reports
  only codes outside Excel's documented 1–11 and 101–111 families, without
  inspecting references or calculating a subtotal. Computed, signed, decimal,
  array, malformed, explicit-broken-reference, and namespaced forms, plus all
  array territory, stay outside the boundary. It emits only a location plus
  aggregate call and unsupported-literal-function-code counts; formulas,
  function-code values, references, and source sheets remain private.
  Its direct
  circular-reference signal accepts only an ordinary formula with a resolved
  scalar static dependency directly back to itself while workbook `iterate` is
  absent or false (the OOXML default). Its separate multi-cell signal accepts
  only a strongly connected component of at least two eligible ordinary formula
  cells in that same resolved scalar graph. It does not expand a range or
  evaluate a formula to close a cycle. Both signals stay quiet for enabled
  iteration; the multi-cell signal also excludes dynamic-reference, 3-D, spill,
  explicit-intersection, tokenizer-failure, and all array territory. Incomplete
  array metadata makes the command fail closed. Its JSON, Markdown, and SARIF
  evidence contains only locations, peer coordinates where copied-pattern
  evidence needs them, static range coordinates, the two calculation-status
  flags, aggregate error-checking suppression counts, Table exception kinds,
  conditional-aggregate and `SUMPRODUCT` mismatch counts, `MMULT`
  incompatible-matrix-pair counts, legacy-lookup out-of-range-literal-index
  counts, `CHOOSE` out-of-range-literal-index counts, `RANDBETWEEN`
  inverted-literal-bound counts, `SUBTOTAL` unsupported-literal-function-code
  counts, direct- or multi-cell-static scope, and a multi-cell component size,
  never formula text, cell values, cached results, ignored-error target ranges,
  direct conditional-aggregate, `SUMPRODUCT`, `MMULT`, legacy-lookup ranges,
  `CHOOSE` values, `RANDBETWEEN` literal values, `SUBTOTAL` function-code or
  reference material, Table identities, or Table master formulas.
- It never executes VBA, XLM macro sheets, Python-in-Excel scripts, RibbonX
  callbacks, DDE, external links, Power Query, Power Pivot/DAX, Office Web
  Add-in or custom-function code, or worksheet ActiveX/OLE code; it does not
  contact a Python-in-Excel Microsoft Cloud or custom-function runtime.
- VBA payloads, XLM macro-sheet source material, RibbonX control/callback
  material, Office Web Add-in task-pane/worksheet/in-content material, and worksheet control/OLE
  material are compared through private fingerprints only.
- Embedded Power Pivot/Data Model declarations and bounded raw model payloads
  are compared through private fingerprints only; table names, relationships,
  DAX, stored values, and connection details are never emitted.
- The dedicated Python-in-Excel ledger compares package source, environment
  definitions/identifiers, script indexes, formula arguments and locations,
  and raw XML through private fingerprints only. Its public ledger retains
  aggregate package, formula-call, script, environment, initialization, and
  coverage counts. Generic semantic cell reports deliberately retain changed
  formula and value evidence for local review; the separate
  `--redact-python-in-excel` rendering boundary covers direct PY source and
  exact changed static PY inputs when that artifact is shared.
- The dedicated namespaced custom-function ledger compares candidate names,
  namespaces, cells, formulas, arguments, and relevant formula-defined-name
  chains through private signatures only. Its public inventory retains aggregate
  formula-cell, call, namespace, and relevant-definition counts. Generic
  semantic cell reports retain local-review evidence by default; the separate
  `--redact-office-custom-functions` rendering boundary covers direct call
  material, exact changed static inputs, and changed private name-chain evidence
  when that artifact is shared. A matching formula is not evidence that an
  Office Add-in is installed, trusted, or runnable.
- The dedicated unqualified runtime-function ledger compares unknown bare-call
  candidates, arguments, cells, and relevant formula-defined-name chains through
  private signatures only. Its public inventory retains aggregate formula-cell,
  call, and relevant-definition counts. Generic semantic reports retain local
  reviewer evidence by default; the separate
  `--redact-unqualified-runtime-functions` rendering boundary covers direct
  bare-call material, exact changed static inputs, and changed private
  name-chain evidence when that artifact is shared. A candidate is not proof a
  VBA, COM/Automation, XLL, or other provider is installed, trusted, or runnable.
- The dedicated worksheet code-resource registration ledger compares stored
  `REGISTER.ID` formulas, module/procedure/type arguments, cells, and relevant
  formula-defined-name chains through private signatures only. Its public
  inventory retains aggregate formula-cell, call, and relevant-definition
  counts. Generic semantic reports retain local reviewer evidence by default;
  the separate `--redact-worksheet-code-resource-registrations` rendering
  boundary covers direct registration material, exact changed static inputs,
  and changed private name-chain evidence when that artifact is shared. A
  stored formula is not proof a DLL/code resource is available, trusted, or can
  be registered.
- What-If Data Table output ranges, input-cell references, and raw formula
  metadata are compared through a private signature only. Cached scenario-output
  cells remain under the normal cell-diff boundary.
- Scenario Manager names, comments, user metadata, stored input values,
  input/result references, and raw declarations are compared through a private
  signature only. Cached worksheet cells remain under the normal cell-diff
  boundary.
- Worksheet DrawingML regular-shape, connector, and recognized SmartArt
  graphic-frame presentation, geometry, anchors, diagram component material,
  bounded direct Diagram Data image payloads, connector endpoint targets,
  macro assignments, text links, descriptions, relationship identifiers, and
  targets are compared through private
  fingerprints only. Profiles and reports retain structural counts, never the
  underlying shape, connector, or SmartArt content.
- Native worksheet image declarations, anchors, visual properties,
  relationship identifiers/targets, and bounded direct image payloads are
  compared through private fingerprints only. Profiles and reports retain
  aggregate structural counts, never image bytes or image metadata.
- Worksheet cell-hyperlink targets, locations, display overrides, ScreenTips,
  references, relationship identifiers, and revision UIDs are compared through
  private fingerprints only. Profiles and reports retain structural counts,
  never the underlying link material, and the CLI never follows a target.
- Worksheet sparkline source formulas, date-axis sources, destination cells,
  group properties, and colour definitions are compared through private
  fingerprints only. Profiles and reports retain aggregate structural counts,
  never the underlying source or presentation material.
- Worksheet print ranges and titles, header/footer text, page values,
  printer-setting references, and raw print-layout XML are compared through
  private fingerprints only. Profiles and reports retain only aggregate
  structural counts.
- Cell-border definitions, colours, style indexes, and cell/row/column targets
  are compared through private fingerprints only. Profiles and reports retain
  only aggregate structural counts.
- Legacy Excel Note text, authors, cell associations, comment properties,
  threaded-comment placeholder links, VML visibility/layout, relationship
  identifiers, targets, and GUIDs are compared through private fingerprints
  only. Profiles and reports retain aggregate structural counts, never the
  underlying Note content or VML.
- Modern threaded-comment text, cell references, timestamps, reply links,
  mention ranges, person names/user IDs/provider IDs, relationship identifiers,
  and GUIDs are compared through private fingerprints only. Profiles and
  reports retain structural counts, never the underlying collaboration content.
- Filter criteria, selected values, custom sort lists, table names, sort keys,
  and row/range references are compared through a private signature only.
- Ignored-error target ranges and exact per-range warning suppressions are
  compared through a private signature only.
- Protection credential material is never emitted: legacy verifiers, modern
  hashes/salts, protected-range names, and security descriptors are compared
  through private fingerprints and reported only as safe presence/change metadata.
- External-data source material is never emitted: connection names/descriptions,
  paths, URLs, connection strings, commands, parameter values, SSO identifiers,
  cached records, and opaque extension XML remain private comparison evidence.
- A policy is bounded and parsed before a policy-enforcing command opens a
  workbook input. FormulaFence opens it once through a file descriptor,
  requests nonblocking mode where the host provides it, verifies that descriptor
  is a regular file, and reads bounded source bytes from that descriptor rather
  than reopening its pathname. On hosts with nonblocking descriptor opens, a
  post-check FIFO or device replacement therefore fails closed instead of
  stalling a runner. The policy accepts one UTF-8 YAML document with ordinary
  mappings, lists, and scalars; duplicate keys, anchors, aliases, and merge
  keys fail closed rather than silently changing a reviewed control.
- A CLI report and an `init --force` starter policy are written first to a
  private temporary file in the requested destination directory, then published
  by atomic replacement of the final directory entry. A final-component symlink
  or hard link swapped in after report-input validation is consequently replaced
  rather than followed into an inspected input. Ordinary `init` publishes that
  complete temporary file with an atomic no-replace link operation instead: a
  regular file, symlink, or hard link that appears at the final name causes a
  refusal and remains unchanged. This protects final-entry publication races,
  not a hostile parent-directory owner or a caller's broader workspace
  permissions.
- For every workbook snapshot, FormulaFence opens one regular source and makes
  one bounded private copy before archive preflight. The archive inventory,
  semantic-reader gate, downstream raw scanners, ordinary reader, and snapshot
  hash all consume that same copy, while the public snapshot path remains the
  caller-supplied path. A pathname replacement after materialization therefore
  cannot make an earlier preflight describe different bytes from later evidence.
  This establishes an internally coherent inspection artifact; it does not
  lock a source against an in-place producer write or defend against another
  process able to modify FormulaFence's private temporary file under the same
  operating-system identity. Use atomic producer handoff and isolated CI
  workspace permissions when provenance or writer integrity is in scope.
- Before any raw OOXML member scanner or `openpyxl` reader runs, FormulaFence
  performs a bounded, fail-closed inventory of that private ZIP container. It
  reads only ZIP headers and central-directory metadata at this stage; it does
  not extract members. The inventory requires one canonical single-disk package
  with stored or deflated members, checks central and local-header consistency,
  rejects duplicate/case-colliding or unsafe member paths, ZIP Unicode-path
  aliases, encrypted and special-file members, and overlapping payloads. It
  bounds source, directory, member, aggregate-expansion, and compression-ratio
  resources.
  A rejected package is never handed to a raw OOXML scanner or workbook reader.
- After the structural ZIP inventory, FormulaFence applies a separate
  semantic-reader resource preflight before it runs downstream raw OOXML
  scanners or starts `openpyxl`. It streams the reader-visible package
  manifest, workbook metadata, canonical styles, the manifest-selected
  shared-string table plus any distinct raw rich-text relationship-selected
  table (with their narrow fallbacks), and workbook-selected sheet parts. Every
  XML/relationship part is capped at 64
  MiB and aggregate XML material at 256 MiB. Each streamed reader part is also
  capped at 4,000,000 elements and 256 nesting levels. The gate limits every
  physical XML opening tag to 128 KiB before ElementTree receives its first
  start callback and constructs a complete attribute map. The same streaming
  lexical gate caps each non-character-data token at 128 KiB: comments,
  processing instructions, declarations, closing tags, and entity references
  cannot make the parser retain an unbounded lexical value before an ordinary
  stream event. It covers UTF-8/ASCII-compatible, UTF-16, and UTF-32
  punctuation, preserves quoted delimiters, and leaves CDATA content to the
  separate character-data bound. The shared defused parser explicitly forbids
  document-type declarations. Bounded raw XML structure scans and in-memory
  OOXML root reads use the same lexical gate before tree construction. A custom
  ElementTree target additionally limits each decoded
  character-data node to 1 MiB before the ordinary tree builder can retain and
  join more parser chunks. It covers ordinary text, tails, and CDATA through
  the shared semantic-reader stream, bounded raw scans, and in-memory root
  reads; the allowance resets at XML markup boundaries rather than treating a
  document's unrelated text nodes as one value. It also limits
  populated SpreadsheetML cell records and shared-string entries to 500,000
  each, row-dimension declarations to 16,384 across selected ordinary
  worksheet parts, column-dimension declarations to 16,384, and direct
  column-dimension containers to 4,096 across those parts, effective `cellXfs`
  styles to 65,490, cell text to 32,767 characters, and stored
  formula/defined-name text to 8,192 characters. Simple shared strings retain
  that broad entry allowance, while each complete `si` item is capped at
  32,768 XML elements; complex/rich items share a 65,536-element budget, and
  ignored opaque direct `sst` children allow 32,768 elements per selected table
  and 65,536 in aggregate. This protects both the ordinary reader, which must
  retain an item it is interpreting, and FormulaFence's raw rich-text scanner.
  That scanner processes one direct `si` item at a time and releases completed
  unrelated root children. It also matches the documented
  [Worksheet root-child grammar](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.worksheet?view=openxml-3.0.1)
  for selected transitional and Strict worksheet parts. A direct subtree rooted
  at any other child allows 32,768 XML elements per worksheet and 65,536 in
  aggregate before raw worksheet scanners or the ordinary reader can retain it.
  A named SpreadsheetML `extLst` is itself an arbitrary extension container,
  so every extension-list subtree in those selected worksheets separately
  allows 32,768 XML elements per worksheet and 65,536 in aggregate. Ordinary
  `sheetData` and other named base controls remain on their existing specific
  budgets; these narrow counters are allocation boundaries for opaque root and
  extension content, not a claim that they make foreign markup invalid.
  Relationship-selected Chartsheet and Dialogsheet parts are non-grid control
  parts, so their complete XML trees separately allow 32,768 elements per part
  and 65,536 in aggregate before raw control scanners or the workbook reader
  can materialize them. This includes documented `extLst` content and opaque
  markup; chart DrawingML remains under its dedicated relationship-selected
  structural bound. The bootstrap `xl/workbook.xml` reader likewise constructs
  one complete package tree. Its documented
  [Workbook `extLst` location](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.workbook?view=openxml-3.0.1)
  and every nested local-name `extLst` subtree allow 32,768 elements; a foreign
  direct workbook-root subtree has the same separate bound. The gate follows
  the parser's local-name dispatch for `workbook` and `extLst`, covering
  alternate namespaces without constraining named Workbook controls or their
  established catalog limits. A row
  counts only when it has an unqualified attribute other than `r` or `spans`,
  the condition that makes `openpyxl` retain a `RowDimension`; namespaced
  extension attributes do not count. Every reader-visible `col` counts because
  `openpyxl` dispatches it before resolving its attributes, while raw dimension
  scanners retain direct `cols` containers. The complete `xl/styles.xml`
  reader tree is separately bounded: every local-name `extLst` subtree, foreign
  direct root subtree, foreign root local name, and ignored direct child in a
  named catalog permits 32,768 elements. Every materialized direct style record
  permits 32,768 non-extension descendants and those records share 262,144
  elements. Repeated known stylesheet containers and every reader-materialized
  number-format, font, fill, fill-child, gradient-stop, border, base-XF,
  named-style, differential-style, palette, table-style, table-style-element,
  and extension catalog retain their existing 4,096-record bounds; effective
  `cellXfs` retains its separate 65,490-style ceiling. The counters follow the
  stylesheet reader's local-name and nested-sequence behavior, including
  alternate namespaces and unexpected direct nested records. It separately
  caps the reader-materialized bootstrap catalogs at 4,096 content-type
  declarations, 4,096 workbook relationships, 512 workbook sheet declarations, and 100,000
  direct workbook defined-name declarations, including repeated declarations of
  one relationship target. It separately caps direct relationship-backed
  external-reference and pivot-cache declarations at 4,096 each, and caps
  direct workbook book-view, custom-workbook-view, function-group, smart-tag-
  type, and web-publish-object catalogs at 4,096 each. Direct legacy custom
  sheet-view declarations are capped at 4,096 in aggregate across the selected
  worksheet, chart-sheet, and dialog-sheet parts before FormulaFence's raw
  Custom View scanner builds per-view records. It also follows direct internal
  `drawing` relationships from selected transitional or Strict worksheet parts,
  then streams each unique XML target before shape, native-image, in-content
  Office Web Add-in, worksheet-chart, or ordinary workbook readers can
  materialize it. Those shared Worksheet DrawingML targets allow 32,768 XML
  elements per part and 65,536 in aggregate. A successfully parsed structural
  overage is rejected by the stable safety preflight; malformed, missing, or
  non-XML optional targets retain their downstream coverage behavior, and
  orphan DrawingML parts are not selected. It also streams every canonical
  `xl/tables/*.xml` Table Definition part, because the raw Table Style scanner
  inventories canonical table parts even when they are orphaned, and every safe
  direct internal worksheet `table` relationship target, because raw filter,
  Named Sheet View, external-data, XML Mapping, and ordinary workbook readers
  can materialize it. The standard transitional and Strict relationship forms
  plus noncanonical safe package targets are covered. Those shared Table
  Definition XML targets allow 32,768 elements per part and 65,536 in
  aggregate. A successfully parsed structural overage is rejected by the
  stable safety preflight; malformed, missing, and non-XML optional targets
  preserve their downstream coverage behavior. These are CI allocation limits
  rather than OOXML validity limits. The catalog counters follow the reader's
  local-name behavior so alternate-namespace entries cannot evade a safety
  boundary. FormulaFence requires
  `defusedxml` for its XML parser,
  which also enables `openpyxl`'s defused XML path in the supported
  installation. This prevents valid-but-impractical documents from allocating
  an unbounded complete workbook model; it is not a malware classifier or a
  substitute for an isolated CI runner.
- It uses sparse cell storage rather than walking every coordinate in a workbook's
  declared used rectangle.
- Parser warnings from unsupported OOXML extensions are captured in the profile
  as coverage notes; FormulaFence does not silently discard them from its report.
- Portfolio comparison recursively inventories only `.xlsx` and `.xlsm` files
  under each supplied directory, identifies a workbook solely by its relative
  path, and reports additions/removals rather than guessing renames. It keeps
  roots and absolute worker paths out of portfolio output, ignores transient
  Office `~$` lock files, rejects symlinked paths and paths that differ only by
  case, and bounds each directory to 512 supported workbooks plus 32,768 raw
  filesystem entries by default. The raw-entry budget is applied before paths
  are retained or sorted, so arbitrary non-workbook files and directories cannot
  create an unbounded inventory; both limits are caller-configurable. Direct
  directory enumeration propagates every subtree read error, so a successful
  report never silently treats an unreadable branch as empty. For every retained
  workbook, the later private-copy read verifies the observed regular file's
  identity and state, requesting a no-follow final component where the host
  supports it. A post-inventory in-place rewrite, regular-file replacement, or
  symlink substitution produces redacted `FF078` evidence and a final incomplete
  exit status, while remaining paths are still reported.
- Snapshot construction has a separate 2,000,000-edge static local
  dependency-graph ceiling by default (`--max-dependency-edges`). It counts
  retained direct/range dependencies and additional fixed-CSE or observed
  dynamic-array output aliases; in a directory portfolio one independent pool
  spans each baseline/candidate side. This prevents a compact formula-defined
  name from multiplying retained local graph state at many callers, without
  expanding a range into cells or changing the candidate-only cross-workbook
  graph boundary. An overage stops with status 2 before a partial artifact is
  published.
- Formula-defined-name sensitive-call propagation has a separate 1,000,000-
  state ceiling by default (`--max-formula-defined-name-states`). It reserves
  direct ledger entries, direct name-marker dependencies, and inherited
  component ledgers before retaining them; in a directory portfolio one
  independent pool spans each baseline/candidate side. This prevents a compact
  acyclic chain from repeatedly materializing action, DDE, custom-function,
  registration, XLM, or environment-information prefixes while preserving
  distinct runtime calls. An overage stops with status 2 before a partial
  artifact is published.
- Local `diff` / `check` impact analysis has a 100,000-state aggregate default
  budget. A state is a changed source or one reachable local dependency state;
  a portfolio shares one pool across every matched workbook. This prevents a
  broad edit set from multiplying the per-source graph bound into impractical
  CI CPU or retained evidence. The command stops with status 2 rather than
  emitting a partial impact report, and reconstructs shortest paths only for
  the fixed review sample instead of every reachable path prefix.
- Before it builds a public `profile` object, the CLI applies a separate
  100,000-record default ceiling (`--max-profile-records`). It counts every
  serialised profile-list item, including nested table columns, control ranges,
  token/function sequences, and dynamic-array references. A profile inventory
  overage returns status 2 before a new large Python object graph or an output
  path exists; this is intentionally distinct from reader state and rendered
  artifact bytes, and a reviewer can deliberately choose a larger positive
  budget for a complete known inventory.
- Rendered `profile`, `diff`, `check`, and `portfolio` artifacts have a
  separate 32 MiB UTF-8 default ceiling (`--max-report-bytes`). It applies to
  profile JSON/Markdown plus comparison JSON, Markdown, HTML, and SARIF before
  FormulaFence writes or replaces the requested output, so a highly compressible
  workbook cannot turn bounded reader state into an impractical CI artifact.
  JSON/SARIF count incremental encoding, Markdown streams lines into the shared
  budget, and HTML counts escaped review entries; an overage returns status 2
  without publishing a partial file. A caller may set a larger positive limit
  only for an intentionally reviewed artifact.
- Cross-workbook portfolio impact evidence is candidate-only and local to the
  supplied inventory. FormulaFence retains raw external source spellings and
  package targets only as private parser state, then resolves a direct static
  A1 source, an exact static external 3-D A1 span, a direct workbook-scoped or
  sheet-local external name, a direct book-only table selector, or narrow
  package-indexed forms `[N]Sheet!A1`, `[N]First:Last!A1`, `[N]!Name`,
  `[N]Sheet!LocalName`, and `[N]!Table[Column]`. For an indexed form, `N` must
  select exactly one
  document-order `externalReference`, `externalLink` part, `externalBook`, and
  external `externalLinkPath` relationship before that target may normalize to
  one exact relative candidate. A workbook-scoped consumer alias may terminate
  in one exact static indexed spelling or one direct A1, workbook-scoped-name,
  sheet-local, or selector-bearing table spelling. It may reach that terminal
  through a finite, acyclic chain whose intermediate definitions are each one
  unqualified non-A1 name identity; a same-named sheet-local consumer
  definition shadows the workbook alias. An eligible workbook-scoped
  formula-defined name may also retain static endpoint inputs, such as
  `=SUM(ExternalInput)`, `=SUM('..\\inputs\\[Inputs.xlsx]Data'!$B$2:$B$4)`,
  or a named `LAMBDA` such as
  `=LAMBDA(value,SUM(value,Inputs!$B$2,ExternalInput))`. This extracts input
  edges rather than calculating a name: every external token must already be
  a direct/package-validated endpoint, every other reference must be static,
  and broken, unresolved, tokenizer-failed, dynamic, relative, recursive,
  local-3-D, spilled, explicitly intersected, local, and locally shadowed
  definitions stay outside the bridge. A named LAMBDA retains its endpoint and
  fixed internal inputs only at a function call, never through a bare name.
  Eligible global formula names and named LAMBDAs may call each other. A 3-D span is expanded only when its
  source candidate has a complete raw OOXML tab catalog consistent with the
  inspected ordinary-worksheet order, and only from exact forward endpoints.
  Sheet-scoped aliases, formula-defined bridges with broken, unresolved,
  tokenizer-failed, dynamic, relative, local-3-D, spilled, or explicitly
  intersected semantics, missing or cyclic exact aliases, caches, non-static
  package A1 forms, ambiguous package shapes, bare or
  source-sheet-qualified table forms, `@`/`#This Row`, unsupported selectors,
  missing/colliding source tables, and spans with missing, reversed,
  non-worksheet, or inconsistent-tab-catalog endpoints are not expanded. A
  source table must be the sole case-insensitive match in the inspected
  candidate; FormulaFence then maps only its static selector bounds to source
  cells. A source name must also expand completely to static internal A1
  destinations in that source candidate; an explicit source sheet selects only
  that local scope, never a global or other sheet fallback. It never opens a target path,
  searches by basename, follows an absolute/UNC/URI/escaping path, fetches
  anything, evaluates a formula, trusts cached external-link values, or emits
  the stored external path, source-name spelling, table identity, or selector
  in portfolio evidence.
  Ordinary source and consumer defined-name declarations remain normal profile
  context. Static ranges stay lazy. A global 100,000-state default bound emits
  `FF080` and exit status 2 rather than presenting incomplete `FF079` impact
  evidence as exhaustive.
- CLI report output is refused when it resolves to an inspected workbook or
  policy, and portfolio output is refused inside either input directory. This
  keeps a reporting request from mutating evidence or changing a portfolio's
  inventory during review.

## What a finding means

An impact count traces explicit A1-style cell dependencies available in the
baseline and candidate. It is an aid to review, not a claim that the cells will
recalculate correctly in Excel. FormulaFence also emits deterministic shortest
path samples from the changed cell to sampled downstream formulas. These paths
are explicit static dependencies, not proof of runtime evaluation. A
formula-pattern finding means both immediate peers have the same relative
formula fingerprint while the changed middle cell does not. An aggregate-range
finding means an accepted static aggregate range ends before a short contiguous
run of numeric literals. A formula-protection finding means a stored direct
cell style makes an ordinary formula editable despite active sheet protection.
A calculation-freshness finding means a formula workbook records manual mode
and incomplete calculation before save. An error-checking-suppression finding
means the workbook stores recognized Excel prompt suppressions; its aggregate
counts do not reveal targets, determine whether a prompt applies, or judge the
suppression. A Table-calculated-column finding means an
interior data cell differs while its immediate neighboring rows match the
stored scalar Table master; it retains neither that master nor the Table
identity and does not determine whether the exception was intentional. A
direct-circular-reference finding means FormulaFence's resolved scalar static
dependency returns to its own ordinary formula cell while
calculation iteration is disabled. A multi-cell static-circular-reference
finding means the formula belongs to a component of at least two ordinary
formula cells connected through those resolved scalar dependencies; its
reported component size never reveals formulas or peer-edge details. An
explicit-broken-reference finding means a stored formula tokenized an actual
`#REF!` error operand, rather than merely containing those characters as text.
A saved-broken-reference finding means a well-formed formula cache recorded a
broken-reference error at its last calculation; the literal-formula finding
takes precedence at the same location. Each is a focused review prompt, not
proof that a formula should change, a cached result is current, or an error will
occur at calculation time.

For a portfolio `FF079`, the same distinction applies across candidate
workbooks: it records that a changed source cell can reach a formula through a
bounded, explicit static graph. It does not prove that Excel can update the
link, that a workbook is open, that a source is trusted, or what value any
formula will produce.

## Deliberate limits

- Supported files are `.xlsx` and `.xlsm`; legacy `.xls` and file-encrypted or
  password-to-open workbooks are outside scope. Workbook and worksheet
  protection flags inside an otherwise readable OOXML workbook are inspected as
  operational controls, not treated as encryption.
- Source package safety limits are fixed at 1 GiB compressed source bytes,
  32 MiB central-directory metadata, 4,096 members, 1,024 UTF-8 bytes per
  member name, 512 MiB expanded bytes per member, 768 MiB expanded bytes in
  aggregate, and a 1,000:1 maximum member compression ratio. FormulaFence
  accepts only stored or deflated, canonical, single-disk ZIP members. It
  rejects duplicate and case-colliding names, traversal or non-canonical paths,
  ZIP Unicode-path aliases, encrypted/symlink/special-file members,
  inconsistent local headers, and overlapping payloads before it reads workbook
  content. These are resource and
  interpretation limits for untrusted CI input, not a claim to detect every
  hostile document or to make an untrusted runner safe.
- Policy source is independently capped at 1 MiB, 4,096 composed YAML nodes,
  64 nesting levels, 4,096 characters per scalar, and 512 selectors in each
  selector list. It is one UTF-8 YAML document read from one verified regular
  descriptor, requesting nonblocking mode where available; FormulaFence rejects
  duplicate mapping keys, anchors, aliases, and merge keys so a committed policy
  cannot hide a last-wins or inherited control change.
- After that ZIP-only pass, the semantic-reader limit is 64 MiB for each
  XML/relationship part, 256 MiB for aggregate XML material, 4,000,000 XML
  elements and 256 nesting levels for every reader-visible part it streams,
  and a 128 KiB physical opening-tag limit before an XML parser can allocate
  one complete attribute map. It then limits 500,000 populated SpreadsheetML
  cell records, 16,384 reader-materialized
  row-dimension declarations across reader-selected ordinary worksheet parts,
  16,384 column-dimension declarations and 4,096 direct column-dimension
  containers across those parts, 500,000 shared-string entries, and 65,490
  effective `cellXfs` styles. It separately allows 32,768 XML elements per
  complete shared-string `si` item, 65,536 across complex/rich items, and
  32,768 opaque direct `sst`-child elements per selected table with 65,536 in
  aggregate; simple shared strings still use the 500,000-entry budget. Each
  selected transitional or Strict worksheet also allows 32,768 XML elements in
  a direct opaque root subtree and 65,536 across those subtrees. Every
  SpreadsheetML `extLst` subtree in those worksheets separately allows 32,768
  XML elements per worksheet and 65,536 in aggregate; ordinary `sheetData` and
  other named base controls keep their existing specialized budgets.
  Relationship-selected Chartsheet and Dialogsheet parts also each allow 32,768
  complete XML elements, with 65,536 across that non-grid-sheet inventory,
  before raw control scanners or the workbook reader can materialize them.
  Their documented `extLst` and opaque content share that boundary while chart
  DrawingML remains separately relationship-bounded. The bootstrap
  `xl/workbook.xml` part separately permits 32,768 elements in a foreign direct
  root subtree and 32,768 in every local-name `extLst` subtree, including a
  nested or alternate-namespace extension list; named Workbook controls retain
  their existing catalog budgets. A row
  declaration counts only when it has an
  unqualified attribute other than `r` or `spans`, which is the `openpyxl`
  `RowDimension` allocation trigger; a namespace-qualified extension attribute
  does not count. Every reader-visible `col` counts because `openpyxl`
  dispatches it before examining attributes, and raw dimension scanners retain
  direct `cols` containers. The complete `xl/styles.xml` reader tree is
  separately bounded: every local-name `extLst` subtree, foreign direct root
  subtree, foreign root local name, and ignored direct child in a named catalog
  permits 32,768 elements. Every materialized direct style record permits
  32,768 non-extension descendants and those records share 262,144 elements.
  Repeated known stylesheet containers plus number-format, font, fill,
  fill-child, gradient-stop, border, base-XF, named-style, differential-style,
  palette, table-style, table-style-element, and extension records retain their
  4,096 bounds; effective `cellXfs` retains its separate 65,490-style ceiling.
  The gate follows `openpyxl`'s local-name and nested-sequence behavior so
  alternate namespaces and repeated ordinary-looking children cannot bypass
  those style bounds. It also rejects a cell text value above 32,767 characters
  or stored formula/defined-name text above 8,192 characters, using
  [Excel's published specifications and limits](https://support.microsoft.com/en-US/Excel/excel-specifications-and-limits)
  for the text, formula, and style compatibility ceilings. It caps manifest
  declarations and workbook relationships at 4,096 each, workbook sheet
  declarations at 512, and direct workbook defined-name declarations at
  100,000, counting reader-materialized catalog entries rather than only
  conventional namespace-qualified tags or unique target parts. Direct
  relationship-backed external-reference and pivot-cache declarations are each
  capped at 4,096 so repetitions cannot reuse a target unboundedly. FormulaFence
  caps direct merged-cell declarations across reader-selected ordinary worksheet
  parts at 4,096 before `openpyxl` expands them. Each declared merge range and
  their aggregate expanded coordinate area are limited to 100,000 cells, and a
  merge reference is limited to 256 characters. This prevents a tiny worksheet
  XML part from allocating a `MergedCell` object for every coordinate of an
  impractically large range. FormulaFence also caps direct data-validation
  declarations, conditional-formatting declarations and rules, and Scenario
  Manager containers, scenarios, and input-cell records at 4,096 each across
  reader-selected ordinary worksheet parts. Each of their `sqref` references is
  limited to 128 KiB and 4,096 target ranges, with 8,192 target ranges in
  aggregate for each catalog, before `openpyxl` materializes a `CellRange` per
  target. Data-validation and conditional-formatting formula fields also share
  the 8,192-character stored-formula bound. FormulaFence
  follows the bounded sheet relationships plus the ordinary reader's
  manifest-selected and raw rich-text relationship-selected shared strings, and
  streams them with `defusedxml` before it creates the complete workbook model.
  A valid workbook above those limits is intentionally rejected rather than
  partially inspected; split high-volume data workbooks
  before sending them through this CI-oriented reader.
- Portfolio mode intentionally does not support legacy `.xls`, `.xlsb`,
  templates, add-ins, or `.ods` files, infer a rename/content match across
  different paths, recursively follow a symlinked workbook, or combine cell
  dependencies between workbooks. A policy is applied independently to every
  matched path; selectors and formula/impact limits are not a portfolio-wide
  policy language. The scanner is sequential to keep resource use bounded; an
  incomplete entry is not treated as unchanged.
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
  Before parsing the canonical `xl/metadata.xml` mapping, FormulaFence streams
  it under a 16 MiB / 32,768-element boundary. A successfully streamed overage
  never reaches the metadata tree parser, makes array-formula classification
  unavailable, and receives no aliases; raw worksheet `c`/`f` bindings are
  streamed rather than retained as a second worksheet tree. Array formulas with
  absent, malformed, unrecognized, or unavailable metadata mappings are
  reported as coverage notes. A private fallback fingerprint keeps a material
  unavailable-metadata change visible as `FF018` without exposing raw XML.
  FormulaFence also reports adding, removing, or changing mode, plus a fixed
  CSE output-range change, as `FF018`; it does not calculate either array form.
- Excel What-If Data Tables are distinct from Excel tables. FormulaFence reads
  each worksheet `f t="dataTable"` master directly from OOXML and privately
  compares its declared output range, one-/two-variable form, one-variable
  orientation, input references, deleted-input flags, recalculation request,
  and supported generic formula metadata. A material change emits `FF034` and
  can be blocked with `no_what_if_data_table_changes`. Profiles and `FF034`
  details expose only structural counts, never those references or ranges.
  Equivalent A1 case/absolute-reference and Boolean spellings are normalized.
  Missing, malformed, overlapping, or unsupported declarations remain visible
  coverage warnings. FormulaFence does not calculate scenarios, infer their
  output formula, predict recalculation results, or add Data Table inputs to
  the ordinary dependency graph; cached scenario-output cells remain ordinary
  cell values under the normal diff boundary.
- Excel Scenario Manager controls are distinct from Data Tables. FormulaFence
  reads each worksheet's raw `<scenarios>` declaration and privately compares
  current/shown selection state, summary references, scenario names,
  locked/hidden flags, declared input counts, comments/users, input references,
  stored values, deleted/undone flags, and display number formats. A material
  change emits `FF035` and can be blocked with
  `no_scenario_manager_changes`. Profiles and `FF035` details expose only
  structural counts, never names, comments, users, values, or references.
  Equivalent local A1 case/absolute-reference, Boolean, and unsigned-integer
  spellings plus schema-default false flags are normalized. Missing, malformed,
  duplicate-within-worksheet, or unsupported declarations remain visible
  coverage warnings. FormulaFence does not show/apply a scenario, calculate its
  result, infer a scenario-to-formula dependency, or fetch an external target;
  cached worksheet cells remain ordinary cell values under the normal diff
  boundary.
- Excel AutoFilters and row/column visibility can change which records or
  fields are shown, and which values vertical `SUBTOTAL` formulas include,
  without changing a formula or ordinary cell value. FormulaFence reads
  worksheet `<autoFilter>` and `<sortState>` elements, the same controls in
  relationship-backed Table Definition parts, explicit row `hidden`,
  `outlineLevel`, `collapsed`, and zero `ht` attributes,
  `sheetFormatPr@zeroHeight`, zero worksheet-default row/column dimensions, and
  raw `<cols>/<col>` `hidden`, `outlineLevel`, `collapsed`, and zero `width`
  declarations. Column declarations are applied in file order and only
  attributes present in a later declaration override earlier effective state,
  including a positive-width override of an inherited zero width. Criteria,
  selected values, custom sort lists, sort keys, raw dimension values, and row/column
  ranges remain private; profiles and `FF036` expose structural counts only.
  Local A1 case/absolute-reference, Boolean/default, unsigned-integer,
  equivalent zero-dimension, and equivalent column-range spellings are
  normalized.
  Unsupported extensions, malformed controls, exhausted column-update limits,
  and unsafe/missing table relationships remain visible coverage warnings, and
  `no_filter_visibility_changes` can block the change as `FFP036`.
  FormulaFence does not apply filters, calculate `SUBTOTAL`/`AGGREGATE`, infer
  formula sensitivity, render views, track ordinary positive widths/heights
  outside its dedicated worksheet-dimension boundary or styles, or model
  outline-display settings.
- Positive worksheet dimensions can conceal wrapped detail, reframe visible
  report context, or shift automatic pagination while leaving formulas and
  values untouched. FormulaFence privately compares raw transitional and strict
  `sheetFormatPr` default row/column/base widths, meaningful default
  `customHeight`, Office 2010 `x14ac:dyDescent` baseline adjustments, and active
  thick-border automatic adjustments; direct row
  height/custom-height/baseline/automatic-thick-border declarations; and effective raw
  `<cols>/<col>` positive width and `bestFit` state. Layered column records are
  resolved in XML order, with later records changing only the attributes they
  supply. `FF058` reports a material declaration change and
  `no_worksheet_dimension_changes` blocks it as `FFP058`. Profiles expose only
  aggregate counts; values, sheet names, row/column targets, writer hints, and
  raw XML remain private. Decimal/Boolean spellings, baseline defaults, inert
  fixed-height thick-border flags, `customWidth`, and equivalent effective range
  segmentation normalize away. Zero/hidden dimensions remain `FF036` visibility
  controls. Malformed, duplicate, unsupported, or budget-exhausted metadata is
  a coverage warning. FormulaFence does **not** calculate final AutoFit sizes,
  text overflow, merged-cell layout, exact automatic page breaks, print
  geometry, or client-specific rendering.
- Excel ignored-error declarations can suppress evaluation, inconsistent-formula,
  omitted-range, unlocked-formula, empty-reference, list-validation,
  calculated-column, text-number, and two-digit-year warnings without changing
  a cell or formula. FormulaFence reads standard `<ignoredErrors>` and Office
  2010 `x14:ignoredErrors` declarations, privately compares local target ranges
  and enabled warning flags, and emits `FF037`; `no_ignored_error_changes` can
  block the change as `FFP037`. Profiles and `FF037` details expose only
  structural counts, never target ranges or individual suppressions. Equivalent
  local A1 case/absolute-reference, Boolean, and target-order spellings are
  normalized. Malformed or unsupported containers, extension material,
  attributes, flags, targets, and child markup remain visible coverage warnings.
  FormulaFence does not decide whether Excel would display a warning, calculate
  a formula, repair an error, alter application-level error checking, or infer
  a suppressed warning's downstream impact.
- Modern Excel Named Sheet Views retain alternate filter and sort settings in
  relationship-backed worksheet parts, potentially changing a saved report view
  without changing ordinary cells or the active AutoFilter. FormulaFence follows
  those parts, privately compares view definitions, and resolves each filter by
  AutoFilter UID, table ID, then worksheet-owned AutoFilter. It emits `FF038`;
  `no_named_sheet_view_changes` can block the change as `FFP038`. Profiles and
  `FF038` details expose only counts for parts, views, alternate filters,
  columns, criterion groups, sort rules/conditions, and unrecognized controls;
  names, IDs, criteria, ranges, table bindings, table-column IDs, and sort keys
  remain private. Equivalent GUID, local A1 case/absolute-reference,
  Boolean/default, and unsigned-integer spellings are normalized. Missing,
  ambiguous, mismatched, malformed, unsupported, oversized, or unsafe
  parts/bindings remain visible coverage warnings. FormulaFence does not
  activate/render a saved view, calculate a filtered result, infer formula
  visibility sensitivity, repair metadata, or interpret future extension/rich-
  sort or full differential-format semantics.
- Legacy Excel Custom Views can preserve a named alternate workbook display or
  print mode through `customWorkbookView` declarations and GUID-linked
  `customSheetView` records on every workbook sheet. They can alter hidden
  rows/columns, filters, print settings, panes, formula/gridline display,
  comments, and object visibility while ordinary cells and active views stay
  fixed. FormulaFence parses the raw workbook and supported worksheet,
  dialog-sheet, and chart-sheet declarations, reconciles each GUID privately,
  and emits `FF060`; `no_custom_workbook_view_changes` can block it as
  `FFP060`. Profiles and `FF060` details expose only structural counts for
  workbook/per-sheet views, affected sheets, hidden/filter/print/display
  settings, and unrecognized metadata. View names, GUIDs, bindings, ranges,
  filters, panes, print settings, and raw XML remain private. Coordinated GUID
  and sheet-ID/active-sheet-ID rewrites plus Boolean/default and
  unsigned-integer spelling normalize. Missing, duplicate, malformed,
  unsupported, unsafe, oversized, over-budget, or incompletely linked metadata
  is visible coverage evidence. FormulaFence does not activate/render a Custom
  View, calculate an alternate filtered result, determine final print output,
  interpret future extensions, or support Custom Views on other sheet types.
- Excel Tables can change a report's review surface through `tableStyleInfo`
  bindings/toggles, custom Table Style definitions, resolved Dxf material, and
  Table/TableColumn direct Dxf or named-cell-style references even when cell
  values, formulas, and table references remain fixed. FormulaFence compares
  those raw declarations privately and emits `FF061`; the
  `no_table_style_control_changes` rule can block it as `FFP061`. Profiles and
  `FF061` details expose only structural counts for declarations, styled/custom
  styles, Dxf/named-style assignments, banding/emphasis, and unrecognized
  metadata—never table/style names, formatting, colours, IDs, or raw XML.
  Boolean/default spelling, case-only names, `xr9:uid` revision provenance, and
  coordinated Dxf ID rewrites normalize. Missing, duplicate, malformed,
  unresolved, unsupported, oversized, or over-budget material remains visible
  coverage evidence. FormulaFence does not render resulting tables, resolve
  themes, calculate values, apply conditional formatting, cover PivotTable-only
  style regions, treat `defaultTableStyle` as an existing-table binding, or
  resolve a same-name named cell-style definition.
- Legacy shared-workbook revision headers and logs can preserve a private audit
  trail outside ordinary cells: prior/new values, locations, authors,
  timestamps, comments, formatting records, conflict-resolution material, and
  shared/tracking/retention/protection controls. FormulaFence follows the
  workbook-to-header and header-to-log relationships, and streams raw revision
  XML before it privately canonicalizes complete bounded declarations. Each
  revision XML part allows 32,768 elements and the complete scan allows 65,536,
  alongside the existing 16 MiB per-part, 64 MiB aggregate, and 512-part
  byte/count limits. A successfully streamed overage emits `FF010`/`FF062`; the
  `no_shared_workbook_revision_changes` rule can block it as `FFP062`. Profiles
  and `FF062` details expose only structural header/log parts and record counts,
  aggregate control counts, and unrecognized metadata—never historic values,
  locations, identities, timestamps, comments, GUIDs, relationship IDs, or raw
  XML. Equivalent Boolean/integer spelling, coordinated relationship-ID
  rewrites, and transitional/Strict relationship type spelling normalize.
  Missing, duplicate, malformed, unsafe, unsupported, oversized, or
  over-budget declarations remain visible coverage evidence; a structural
  overage retains a private streamed content fingerprint so same-size changes
  remain diff-visible. FormulaFence does not apply revisions, reconstruct a
  historical state, resolve conflicts, validate identity/timestamp claims,
  render Excel, or interpret arbitrary future extensions.
- Excel number formats can hide or materially reinterpret an unchanged stored
  value: `;;;` can display it as blank, while custom sections, scaling commas,
  dates, percentages, literals, and text placeholders can change the review
  surface. FormulaFence privately resolves custom `<numFmt>` codes, base
  `<cellStyleXfs>`, effective `<cellXfs>` with `xfId` and
  `applyNumberFormat`, direct cell `s`, `customFormat=1` row `s`, and raw
  `<cols>/<col style>` assignments. It emits `FF039`; the
  `no_number_format_changes` policy rule can block it as `FFP039`. Profiles and
  `FF039` details expose only counts for default/direct/row/effective-column
  assignments, built-in/custom classes, and malformed controls—never codes,
  style IDs, or targets. Equivalent custom-ID remapping, Boolean spelling,
  base-XF inheritance, and effective column-range splitting are normalized.
  Missing custom definitions, invalid IDs/indexes/targets, conflicting
  definitions, and bounded parser failures remain coverage warnings. FormulaFence
  does not render locale-specific output, validate format syntax, calculate
  values, model width/overflow, or compose number formats with separately
  inventoried font/fill/alignment/border or Table Style controls, quote
  prefixes, or arbitrary visual formatting. Column styles
  are compared only as OOXML defaults for unallocated/new cells, not as a claim
  to restyle allocated cells.
- Excel cell fonts can make an unchanged value or warning less visible, such as
  a white font against a matching background, without changing a formula or
  value. FormulaFence privately resolves raw `<fonts>` records, base
  `<cellStyleXfs>`, effective `<cellXfs>` with `xfId` and `applyFont`, direct
  cell `s`, `customFormat=1` row `s`, and raw `<cols>/<col style>` assignments.
  It emits `FF040`; `no_cell_font_changes` can block it as `FFP040`. Profiles
  and `FF040` details expose only default-definition, direct/row/effective-column,
  and malformed-control counts—never font names, colour values, effects, style
  IDs, or targets. Equivalent font-ID remapping, common font-child ordering,
  Boolean spelling, base-XF inheritance, and effective column-range splitting
  are normalized. Missing or malformed definitions, invalid IDs/indexes/targets,
  and bounded parser failures remain visible coverage warnings. FormulaFence
  does not render or resolve theme colours, decide whether a font is visible
  against a fill, calculate text/background contrast or values, compose font
  rendering with fill/border/alignment or other display controls, rich-text run
  rendering, separately inventoried Table Style controls, width/overflow, or
  arbitrary visual formatting. Column
  styles are compared only as OOXML defaults for unallocated/new cells, not as a
  claim to restyle allocated cells.
- Excel cell fills can make unchanged text, warnings, or input/output cues less
  visible without changing a formula or value. FormulaFence privately resolves
  raw `<fills>` definitions, including patterned and gradient fills, base
  `<cellStyleXfs>`, effective `<cellXfs>` with `xfId` and `applyFill`, direct
  cell `s`, `customFormat=1` row `s`, and raw `<cols>/<col style>` assignments.
  It emits `FF041`; `no_cell_fill_changes` can block it as `FFP041`. Profiles
  and `FF041` details expose only default-definition, direct/row/effective-column,
  and malformed-control counts—never fill colours, pattern types, gradient
  geometry/stops, style IDs, or targets. Equivalent fill-ID remapping, valid
  pattern-colour child ordering, Boolean spelling, base-XF inheritance,
  semantically inert no-fill/solid-background declarations, and effective
  column-range splitting are normalized. Missing or malformed definitions,
  invalid IDs/indexes/targets, and bounded parser failures remain visible
  coverage warnings. FormulaFence does not resolve theme colours, render fills,
  calculate text/background contrast, evaluate conditional-format differential
  styles, apply separately inventoried Table Style controls, or claim arbitrary
  visual-style coverage. Column
  styles are compared only as OOXML defaults for unallocated/new cells, not as a
  claim to restyle allocated cells.
- Cell alignment can reposition, rotate, wrap, shrink, or indent an unchanged
  value, warning, or visual classification without a formula or value change.
  FormulaFence privately resolves raw `alignment` children in base
  `cellStyleXfs` and effective `cellXfs` records, follows `xfId`
  and `applyAlignment`, and compares direct cell `s`,
  `customFormat=1` row `s`, and raw `<cols>/<col style>`
  assignments. It covers horizontal/vertical placement, text rotation,
  wrapping, shrinking, indentation, relative indentation, justification, and
  reading order. A material effective declaration change emits `FF054`;
  `no_cell_alignment_changes` blocks it as `FFP054`. Profiles and
  reports expose only default/direct/row/effective-column/malformed counts;
  alignment values, style IDs, and targets remain private. Equivalent explicit
  defaults, Boolean/integer spelling, inert `mergeCell` compatibility
  material, base-XF inheritance, `applyAlignment`, and effective
  column-range splitting normalize away. Missing, duplicate, malformed, or
  unsupported metadata remains a visible coverage warning. FormulaFence does
  not calculate width, height, merged layout, overflow, final visibility,
  font/fill/conditional-format composition, or Excel rendering. Column styles
  remain OOXML defaults for unallocated/new cells, not a renderer that restyles
  allocated cells. This boundary follows Microsoft's
  [SpreadsheetML alignment definition](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oi29500/e4ad6e3e-7702-4dbe-8c44-f5a4c686c440)
  and [CellFormat alignment semantics](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oi29500/68362a4b-5589-4504-b566-e8154dce1de3).
- Cell borders can redraw a report boundary, total, exception box, or warning
  without a formula or value edit. FormulaFence privately resolves raw
  transitional and strict SpreadsheetML `<borders>/<border>` definitions, base
  `cellStyleXfs`, effective `cellXfs` records with `borderId`, `xfId`, and
  `applyBorder`, direct cell `s`, `customFormat=1` row `s`, and raw
  `<cols>/<col style>` assignments. It covers left/right/top/bottom, Office
  2010 logical start/end, diagonal/direction, outline, stored line styles, and
  stored colours. A material effective control change emits `FF057`;
  `no_cell_border_changes` blocks it as `FFP057`. Profiles and reports expose
  only default/direct/row/effective-column/unrecognized counts; definitions,
  colours, style IDs, and targets remain private. Omitted/`none` sides,
  Boolean/colour spelling, unused diagonal payload, ineffective empty
  `outline="false"`, base-XF inheritance, `applyBorder`, and equivalent
  column-range splitting normalize away. Missing, duplicate, malformed, or
  unsupported material remains a visible coverage warning. Material
  `vertical`/`horizontal` inner sides under ordinary cell styles are also
  flagged as coverage gaps because their differential-format semantics are not
  modeled here. FormulaFence does **not** resolve theme/palette colours, choose
  adjacent-cell precedence, render a final visual style, apply
  conditional-format/table/differential-style borders, calculate print output,
  or infer client behavior. Column styles remain OOXML defaults for
  unallocated/new cells, not a renderer that restyles allocated cells. This
  boundary follows OOXML's
  [`border`](https://c-rex.net/samples/ooxml/e1/part4/OOXML_P4_DOCX_border_topic_ID0EVV35.html)
  and [`xf`](https://c-rex.net/samples/ooxml/e1/part4/OOXML_P4_DOCX_xf_topic_ID0E13S6.html)
  forms, plus Microsoft's [cell-border guidance](https://support.microsoft.com/en-us/Excel/apply-or-remove-cell-borders-on-a-worksheet).
- Positive worksheet dimensions can change a reviewer's usable surface without
  an ordinary cell edit: fixed row heights can cut off wrapped text, column
  widths can reframe report fields, and sizing can move automatic page breaks.
  FormulaFence privately scans raw transitional and strict `sheetFormatPr`
  defaults (`defaultRowHeight`, `defaultColWidth`, `baseColWidth`), meaningful
  default `customHeight`, Office 2010 `x14ac:dyDescent` baseline adjustments,
  and active automatic thick-border adjustments; direct row
  `ht`/`customHeight`/`x14ac:dyDescent`/`thickTop`/`thickBot`; and raw positive
  `<cols>/<col width>`/`bestFit` declarations. It resolves overlapping columns
  in file order, preserving only changes from later present attributes. A
  material change emits `FF058`; `no_worksheet_dimension_changes` blocks it as
  `FFP058`. Profiles and reports disclose aggregate counts only; dimensions,
  targets, raw XML, and writer hints remain private. Baseline defaults,
  decimal/Boolean spelling, inert fixed-height thick-border flags,
  `customWidth`, and equivalent effective range splitting normalize away.
  Zero/hidden dimensions stay under `FF036`. Malformed, duplicate, unsupported,
  or budget-exhausted controls remain coverage warnings. FormulaFence does
  **not** compute final AutoFit sizing, wrapped/merged layout, overflow, exact
  automatic page breaks, print geometry, or client rendering.
- A stored worksheet view can change a reviewer’s surface while leaving values
  and formulas untouched: zeroes can appear blank; formulas, gridlines,
  row/column headers, outline symbols, rulers, or page margins can be
  hidden/shown; gridlines can be recoloured; the view can be right-to-left or
  page-oriented; and panes can be split/frozen. FormulaFence
  privately compares raw non-default transitional and strict SpreadsheetML
  `sheetViews/sheetView` declarations
  for those controls and emits `FF055`;
  `no_worksheet_display_control_changes` blocks it as `FFP055`.
  Profiles and reports expose only structural counts; sheet names, target
  cells, pane positions, and raw XML remain private. Omitted/default controls,
  Boolean and active custom-gridline-colour spelling, and finite non-negative
  split decimals normalize away.
  Active-cell, selection, top-left navigation, and zoom remain deliberately
  outside this boundary to avoid routine writer churn. Missing, duplicate,
  malformed, or unsupported material yields coverage evidence. FormulaFence
  does **not** render Excel, resolve the effective palette colour, calculate
  viewport geometry/final visibility, interpret extension-specific views, or
  infer client state. This boundary follows the Open XML SDK
  [`SheetView` schema surface](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.sheetview?view=openxml-3.0.1)
  and Microsoft’s [worksheet display guidance](https://learn.microsoft.com/en-us/office/dev/add-ins/excel/excel-add-ins-worksheet-display).
- A saved worksheet print layout can omit printed content, repeat different
  titles, alter print gridlines/headings/centering, reframe paper with margins
  or setup controls, change header/footer text, or insert manual page breaks
  without an ordinary cell edit. FormulaFence privately compares raw
  transitional and strict SpreadsheetML workbook print-area/print-title defined
  names plus direct worksheet `printOptions`, `pageMargins`, `pageSetup`,
  `sheetPr/pageSetUpPr`, `headerFooter`, and row/column-break declarations.
  It emits `FF056`; `no_worksheet_print_layout_changes` blocks it as `FFP056`.
  Profiles and reports expose structural counts only; print ranges,
  header/footer text, page values, printer-setting references, and raw XML stay
  private. Omitted/default, Boolean, integer, and decimal spellings normalize
  away, as do inactive first/even headers or footers, disabled first-page
  numbering, scale overridden by active fit-to-page dimensions, and automatic
  break display state. Missing, duplicate, malformed, or unsupported material
  yields coverage evidence. FormulaFence does **not** render/preview a workbook,
  calculate page geometry/counts or automatic page breaks, resolve printer or
  client defaults, inspect printer-specific `devMode` data, or cover
  custom/legacy sheet-view and extension-specific print controls. This boundary
  follows Microsoft's [print-area guidance](https://support.microsoft.com/en-us/excel/set-or-clear-a-print-area-on-a-worksheet)
  and [`PageLayout` control surface](https://learn.microsoft.com/en-us/javascript/api/excel/excel.pagelayout?view=excel-js-preview).
- A workbook-level DrawingML Theme can alter colour, font, and effect schemes
  used by themed cells, charts, and drawing objects without a local style
  change. FormulaFence inspects the raw workbook Theme binding, Theme XML, and
  direct Theme-image relationships/payloads in transitional and strict OOXML
  namespaces. A material stored control change emits `FF053`;
  `no_workbook_theme_changes` blocks it as `FFP053`. Profiles and reports
  expose only aggregate Theme-part/scheme/relationship/image and
  malformed-metadata counts; Theme XML, scheme names, colours, font names,
  image payloads, relationship IDs, and targets remain private. Writer-selected
  relationship IDs/order and equivalent internal target spelling normalize
  away. Missing, duplicate, malformed, unsafe, unbound, unreadable, oversized,
  or over-budget metadata produces a visible coverage warning; reads are
  bounded to 16 MiB per part, 64 MiB per workbook, and 512 parts. FormulaFence
  also streams Theme and Theme-relationship XML before tree materialization:
  32,768 elements per XML part and 65,536 across the complete Theme scan.
  These are reader-allocation and coverage limits, not workbook-validity
  limits; a well-formed structural overage produces visible `FF053` coverage
  evidence. Direct Theme-image payloads remain byte-bounded rather than being
  treated as XML. FormulaFence does **not** resolve effective styles, render a
  workbook, calculate contrast, decode an image, fetch a target, calculate
  formulas, or infer client behavior. This boundary follows the Open XML SDK
  [WorkbookPart](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.packaging.workbookpart?view=openxml-2.20.0)
  Theme-part surface and Microsoft's
  [conditional-formatting guidance](https://learn.microsoft.com/en-us/office/open-xml/spreadsheet/working-with-conditional-formatting).
- SpreadsheetML can retain a formula's last calculated result beside its
  formula text in the same `<c>` cell. That lets a workbook save a different
  displayed result without changing the ordinary formula text. FormulaFence
  reads raw `<f>` and `<v>` elements together and privately fingerprints
  numeric, string, Boolean, and error result material. It emits `FF042` only
  when a result cache changes without a changed formula at that cell and without
  an ordinary changed cell reaching it through the static dependency graph;
  `no_formula_cached_result_changes` can block it as `FFP042`.
  Profiles and reports expose only formula/cached/missing/type/malformed counts,
  never result values, error text, digests, or locations. Equivalent finite
  numeric and Boolean spellings are normalized; blank results remain visible as
  missing caches. Unsupported or malformed metadata becomes an explicit coverage
  warning. FormulaFence does not calculate or validate results, determine
  whether they are stale or tampered, or model volatile, dynamic, external, or
  calculation-engine dependencies. A legitimate recalculation without a
  statically visible input edit can therefore still require review.
- SpreadsheetML shared strings and inline strings can split one displayed cell
  value into character-level `<r>` runs. Their `<rPr>` formatting can hide or
  alter the emphasis of a phrase while the normal cell reader still returns the
  same concatenated text. FormulaFence follows referenced shared-string items
  and direct inline strings, privately compares run-property sequences, styled
  character boundaries, and phonetic runs/properties, and emits `FF043`.
  `no_rich_text_run_changes` blocks it as `FFP043`. Profiles and report
  details expose only aggregate shared-item/cell/run, inline-cell/run, phonetic,
  and malformed-control counts; text, font/colour material, shared-string
  indexes, and locations remain private. Equivalent property ordering, colour
  case, and explicit false Boolean properties are normalized. A normal text
  edit inside an unchanged run-property sequence remains a normal cell diff,
  while a moved styled boundary with unchanged text is guarded. Malformed,
  unsupported, or unreadable metadata becomes a coverage warning. FormulaFence
  does not render a cell, resolve theme colours, calculate contrast, determine
  whether text is visible, preserve rich text, or guarantee cross-version Excel
  rendering equivalence. This boundary follows Microsoft's
  [shared-string-table guidance](https://learn.microsoft.com/en-us/office/open-xml/spreadsheet/working-with-the-shared-string-table),
  the Open XML `r` [rich-text-run definition](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.run?view=openxml-3.0.1),
  and `is` [inline-string definition](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.inlinestring?view=openxml-3.0.1).
- An ordinary worksheet cell can retain the same friendly value while its stored
  hyperlink changes an external/file target, in-workbook location, display
  override, or ScreenTip. FormulaFence reads raw standard SpreadsheetML
  `hyperlink` and Office 2016 `xr:hyperlink` declarations plus their selected
  worksheet relationships before the ordinary reader can normalize them. It
  privately compares binding, declaration material, location, display/ScreenTip,
  and relationship type/target/mode semantics. A material change emits
  `FF047` and `no_cell_hyperlink_changes` blocks it as `FFP047`. Profiles and
  reports expose only aggregate worksheet/hyperlink,
  location/display/ScreenTip, relationship/external-relationship, and
  malformed-metadata counts; targets, references, locations, display strings,
  ScreenTips, relationship IDs, and revision UIDs remain private.
  Writer-chosen relationship IDs/revision UIDs, relationship ordering, and
  equivalent internal target spelling normalize away. Missing, duplicate,
  malformed, unbound, unsafe, unreadable, oversized, or over-budget metadata
  becomes a visible coverage warning; raw worksheet XML reads are bounded to
  16 MiB per worksheet, 64 MiB per workbook, and 512 parts. The ordinary reader
  receives a hyperlink-removed temporary copy only after raw inspection, so
  malformed markup cannot suppress the evidence. FormulaFence does **not**
  render, resolve, fetch, follow, or test a link; inspect linked content; infer
  reputation/trust-zone/client behavior; or evaluate a `HYPERLINK()` formula.
  Stored `HYPERLINK()` calls are separately covered by `FF064`, still without
  evaluating an argument or following its result. This raw worksheet-hyperlink
  boundary follows the Open XML
  [Hyperlink](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.hyperlink?view=openxml-3.0.1)
  and Office 2016
  [Hyperlink](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.office2016.excel.hyperlink?view=openxml-3.0.1)
  definitions.
- A formula can create a link without a stored worksheet `hyperlink` element,
  request data from an intranet or Internet service, render a URL-sourced image,
  bind a real-time provider, retrieve financial history, or query a stored Cube
  connection. FormulaFence inventories stored `HYPERLINK`, `WEBSERVICE`,
  `IMAGE`, `RTD`, `STOCKHISTORY`, and all documented Cube-family calls
  (`CUBEKPIMEMBER`, `CUBEMEMBER`, `CUBEMEMBERPROPERTY`, `CUBERANKEDMEMBER`,
  `CUBESET`, `CUBESETCOUNT`, and `CUBEVALUE`), including `_xlfn.` compatibility
  spellings in cells, formula-defined names, and named `LAMBDA` bodies. It
  privately fingerprints cell, function-inventory, and relevant named-definition
  material. Public profiles and `FF064` details expose only action-cell,
  formula-defined-name, `STOCKHISTORY`, and aggregate Cube-function counts, so
  an argument-only, name-definition-only, connection/query-only, or same-count
  retarget remains reviewable without disclosing an endpoint, market symbol,
  provider, formula, query, or location. A normal cell change that reaches an
  invoking action/provider formula through FormulaFence's static dependency
  graph also emits `FF064`, covering static sources such as `HYPERLINK(A1, ...)`
  or `STOCKHISTORY(A1, ...)` without reading `A1` as an endpoint or symbol.
  Dynamic or unresolved sources remain explicit formula-coverage limits.
  `HYPERLINK` is deliberately included even when an argument appears internal:
  it can be dynamically calculated, and the ledger does not evaluate it to
  decide that. FormulaFence does **not** calculate a formula, resolve/open/fetch
  a destination, click/follow a link, authenticate, load a COM object, start an
  RTD server, contact a market provider, query a Cube, or execute a provider. A
  material change emits `FF064`; enable
  `no_formula_external_action_changes` for `FFP064`. This boundary follows
  Microsoft's [link guidance](https://support.microsoft.com/en-US/Excel/work-with-links-in-excel),
  [`WEBSERVICE` reference](https://support.microsoft.com/en-US/Excel/functions/webservice-function),
  [`IMAGE` reference](https://support.microsoft.com/en-us/excel/functions/image-function),
  [`RTD` reference](https://support.microsoft.com/en-us/excel/functions/rtd-function),
  [`STOCKHISTORY` reference](https://support.microsoft.com/en-us/office/stockhistory-function-1ac8b5b3-5f62-4d94-8ab8-7504ec7239a8),
  and [`CUBESET` reference](https://support.microsoft.com/en-us/excel/functions/cubeset-function).
- Direct DDE-style formulas can carry a cross-process link without a normal
  worksheet function or a raw `externalLink` package. FormulaFence separately
  recognizes only the lexical `application|topic!item` shape described by the
  Windows [DDE overview](https://learn.microsoft.com/en-us/windows/win32/dataxchg/about-dynamic-data-exchange),
  skipping pipes inside string literals and ordinary quoted sheet names. It
  privately fingerprints direct worksheet formulas plus formula-defined names
  and named `LAMBDA` chains; public `FF074` output has only formula-cell, link,
  and defined-name counts. A material change or statically visible input to an
  invoking named `LAMBDA` emits `FF074`; enable
  `no_formula_dde_link_changes` for `FFP074`. FormulaFence does **not**
  evaluate a formula, resolve an endpoint, look up, launch, or contact a DDE
  server, send a DDE command, or determine whether Excel's local DDE security
  settings permit any action. Raw external-link DDE/OLE metadata remains under
  `FF025`.
- Python in Excel can retain executable source in its `PY` formula and in
  related workbook package material. FormulaFence recognizes stored `PY`
  spellings, privately fingerprints the documented 2023 `python.xml` package
  part and the separately stored 2022 `pythonScripts.xml` compatibility
  contract—including relationships, content types, code/environment/script XML,
  and stored formula binding—then exposes only safe aggregate counts. Both
  physical parts
  remain independently compared when they coexist; FormulaFence does not
  choose one runtime representation or assume they agree. A
  code/package/environment change, formula-binding change, or ordinary cell
  change that statically reaches a PY formula emits `FF065`;
  `no_python_in_excel_changes` blocks it as `FFP065`. This includes a source
  such as `=_xlfn._xlws.PY(0,0,A1)` without decoding the script index,
  interpreting `A1`, or parsing source as Python. Dynamic or unresolved inputs
  remain formula-coverage limits. Relationship-ID-only rewrites normalize;
  missing, malformed, unbound, oversized, unreadable, or over-budget metadata
  remains coverage evidence. XML reads are bounded to 16 MiB per part, 64 MiB
  per workbook, and 512 parts. Before private package XML is materialized,
  FormulaFence streams 32,768 elements per part and 65,536 across the complete
  Python-in-Excel scan; a successfully parsed structural overage becomes visible
  `FF010`/`FF065` coverage evidence. FormulaFence does **not** execute Python,
  evaluate a PY formula, resolve a result, contact Microsoft Cloud, or validate
  runtime package availability. Changed PY formulas and values remain in the
  ordinary semantic diff by design, so it is not a redacted source-code vault.
  `--redact-python-in-excel` is the separate output-only sharing boundary for
  direct stored PY material and exact changed static cells that reach a PY
  formula. It follows Microsoft's [PY function reference](https://support.microsoft.com/en-us/excel/functions/py-function),
  [Python in Excel introduction](https://support.microsoft.com/en-US/Excel/python/introduction-to-python-in-excel),
  and the OOXML [Python part](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/151e4bcd-90a0-4d82-8b98-f16bf273e4ff)
  definition.
- Office Add-in custom functions are defined in JavaScript or TypeScript and
  exposed to Excel through a manifest namespace. They can request and stream
  external data, but a normal workbook stores only the call text—not the
  manifest, code, or runtime identity. FormulaFence therefore inventories only
  a conservative namespaced-call candidate: the direct-call classifier excludes
  known native dotted Excel functions, workbook-defined names, and `_xlfn.` /
  `_xlws.` compatibility names. Unqualified VBA, COM, or XLL UDF-shaped calls
  are covered separately by `FF075`. Candidates inside formula-defined names
  and named `LAMBDA` bodies are propagated to their invoking worksheet formulas.
  A candidate call or a normal edit that statically reaches one emits `FF066`; enable
  `no_office_custom_function_changes` for `FFP066`. The public profile exposes
  only formula-cell, call, namespace, and relevant formula-defined-name counts;
  names, formulas, arguments, and locations stay private. FormulaFence does
  not evaluate a call, resolve a candidate to an add-in, load a manifest or
  add-in, execute JavaScript, or request a service.
  It makes no claim that a candidate maps to a particular add-in or can run.
  Dynamic or unresolved inputs remain static-coverage limits. For a shared
  artifact, `--redact-office-custom-functions` separately hides direct call
  material, exact static input evidence, and changed resolved name-chain
  evidence without changing comparison or policy facts. This boundary
  follows Microsoft's [custom-functions overview](https://learn.microsoft.com/en-us/office/dev/add-ins/excel/custom-functions-overview),
  [tutorial](https://learn.microsoft.com/en-us/office/dev/add-ins/tutorials/excel-tutorial-create-custom-functions),
  and [web-data guidance](https://learn.microsoft.com/en-us/office/dev/add-ins/excel/custom-functions-web-reqs).
- A bare unknown worksheet call can resolve through VBA, a COM/Automation
  add-in, an XLL, or another registered runtime, but formula text alone does
  not prove a provider exists, is trusted, or can run. FormulaFence therefore
  inventories a conservative unqualified-runtime-function candidate rather
  than resolving or loading anything. The direct classifier accepts a bare
  identifier only after excluding workbook-defined names, local `LET`/`LAMBDA`
  bindings, qualified/dotted calls, and a stable native Excel function
  catalogue. This includes current native spellings such as `XLOOKUP`,
  `VSTACK`, `FIELDVALUE`, and `PY`; the pinned catalogue avoids third-party
  parser-version drift, while a new native function can be conservatively
  reported until FormulaFence adds it. Candidate calls inside formula-defined
  names and named `LAMBDA` bodies propagate through nested, recursive, and
  sheet-local chains to their invoking formulas. A private candidate/definition
  change or normal static input edit that reaches one emits `FF075`; enable
  `no_unqualified_runtime_function_changes` for `FFP075`. Public output has
  only formula-cell, call, and relevant definition counts; names, formulas,
  arguments, cells, provider identities, and host details stay private.
  FormulaFence does not evaluate a formula, resolve/load a VBA, COM/Automation,
  XLL, or registered provider, inspect host trust settings, or execute code.
  Stored candidate definitions remain independently reviewable even when no
  worksheet formula invokes them, while static-input paths require an actual
  inspected call. Dynamic and unresolved inputs remain static-coverage limits.
  The separate `--redact-unqualified-runtime-functions` shared-artifact mode
  hides direct bare-call material, exact changed static inputs, and changed
  private name-chain evidence after comparison and policy evaluation without
  resolving a provider or changing any comparison fact. This boundary
  follows Microsoft's [native function catalogue](https://support.microsoft.com/en-us/office/excel-functions-alphabetical-b3944572-255d-4efb-bb96-c6d90033e188),
  [installed UDF guidance](https://support.microsoft.com/en-us/excel/user-defined-functions-that-are-installed-with-add-ins-reference),
  [VBA custom-function guidance](https://support.microsoft.com/en-us/excel/create-custom-functions-in-excel),
  and [XLL registration/call guidance](https://learn.microsoft.com/en-us/office/client-developer/excel/accessing-xll-code-in-excel).
- Excel's worksheet-capable `REGISTER.ID` function can register a DLL or code
  resource when needed and return its registration ID. FormulaFence keeps a
  dedicated private ledger for stored worksheet and formula-defined `REGISTER.ID` calls,
  including calls held in formula-defined names and named `LAMBDA` bodies. A
  call, relevant named-definition change, or ordinary static input change emits
  `FF067`; enable `no_worksheet_code_resource_registration_changes` for
  `FFP067`. Public output exposes only formula-cell, call, and relevant
  formula-defined-name counts; module paths, procedure names, type strings,
  formulas, arguments, cells, and name identities stay private. FormulaFence
  does not evaluate a formula, resolve a path, load a DLL/XLL, inspect trust
  settings, or determine whether registration succeeds. Dynamic or unresolved
  inputs remain static-coverage limits. The separate
  `--redact-worksheet-code-resource-registrations` shared-artifact mode hides
  direct `REGISTER.ID` material, exact changed static inputs, and changed
  private name-chain evidence after comparison and policy evaluation without
  resolving a module/provider or changing any comparison fact. This is separate
  from raw XLM macro-sheet program scanning: Microsoft's [`CALL` reference](https://support.microsoft.com/en-us/office/call-function-32d58445-e646-4ffd-8d5e-b45077a5e995)
  states that `CALL` is macro-sheet only. The worksheet boundary follows
  Microsoft's [`REGISTER.ID` reference](https://support.microsoft.com/en-us/office/register-id-function-f8f0af0f-fd66-4704-a0f2-87b27b175b50).
- Legacy XLM `REGISTER` can be stored in a formula-defined name or named
  `LAMBDA`, a surface not represented by ordinary macro-sheet XML. FormulaFence
  separately inventories only that stored-definition form and propagates it
  through nested/sheet-local names to invoking formula cells. Microsoft's
  [`xlfRegister` Form 1 reference](https://learn.microsoft.com/en-au/office/client-developer/excel/xlfregister-form-1)
  documents DLL-function/command registration and macro types callable from a
  defined-name definition; [`Form 2`](https://learn.microsoft.com/en-au/office/client-developer/excel/xlfregister-form-2)
  documents XLL loading and activation. A material stored definition,
  invocation, or ordinary static input change emits `FF068`; enable
  `no_formula_defined_xlm_registration_changes` for `FFP068`. Public output
  exposes only invoking-cell, call, and relevant definition counts; module
  paths, procedure names, type strings, arguments, formulas, locations, and
  name identities remain private. FormulaFence does not evaluate a formula,
  execute an XLM macro, resolve a path, load a DLL/XLL, or inspect trust
  settings. Direct worksheet `REGISTER` formulas and raw XLM macro-sheet parts
  are intentionally outside this narrow boundary; dynamic/unresolved inputs
  remain static-coverage limits. The separate
  `--redact-formula-defined-xlm-registrations` shared-artifact mode hides direct
  stored `REGISTER` material, exact changed static inputs, and changed private
  name-chain evidence after comparison and policy evaluation without executing
  or resolving a registration target, and without changing any comparison fact.
- Legacy XLM `EVALUATE` can be stored in a formula-defined name or named
  `LAMBDA`, where it parses a supplied text expression at calculation time.
  FormulaFence separately inventories only that stored-definition form and
  propagates it through nested/sheet-local names to invoking formula cells.
  Microsoft's [Excel expression-evaluation
  reference](https://learn.microsoft.com/en-us/office/client-developer/excel/excel-worksheet-and-expression-evaluation)
  identifies `EVALUATE` as the XLM function that reduces a valid character
  string to a worksheet value. A material stored definition, invocation, or
  ordinary static argument-input change emits `FF069`; enable
  `no_formula_defined_xlm_evaluation_changes` for `FFP069`. Public output
  exposes only invoking-cell, call, and relevant definition counts; expression
  text, formulas, arguments, locations, and name identities remain private.
  FormulaFence does not evaluate text, parse a runtime-generated expression,
  execute an XLM macro, or infer dependencies inside the expression it would
  produce. It traces only the stored call's own visible static argument edge.
  Direct worksheet `EVALUATE` formulas and raw XLM macro-sheet parts are
  intentionally outside this narrow boundary; runtime-text dependencies remain
  explicit static-coverage limits. The separate
  `--redact-formula-defined-xlm-evaluations` shared-artifact mode hides direct
  stored `EVALUATE` material, exact changed static inputs, and changed private
  name-chain evidence after comparison and policy evaluation without evaluating
  a formula or text argument, parsing a runtime-generated expression, or
  changing any comparison fact.
- Selected legacy XLM action and event-dispatch calls can also be stored in a
  formula-defined name or named `LAMBDA`, outside raw macro-sheet XML.
  FormulaFence inventories only `CALL`, `EXEC`, `EXECUTE`, `RUN`, `SEND.KEYS`,
  `ON.DATA`, `ON.DOUBLECLICK`, `ON.ENTRY`, `ON.KEY`, `ON.RECALC`, `ON.SHEET`,
  `ON.TIME`, and `ON.WINDOW`, then propagates them through nested and
  sheet-local names to invoking formula cells. Microsoft's [Excel C API
  reference](https://learn.microsoft.com/en-us/office/client-developer/excel/programming-with-the-c-api-in-excel)
  describes XLM command-equivalent functions and event traps such as
  `ON.ENTRY` and `ON.TIME`; its [DLL-access
  guidance](https://learn.microsoft.com/en-us/office/client-developer/excel/how-to-access-dlls-in-excel)
  documents `CALL` and `REGISTER` as XLM macro-sheet routes to DLL functions
  or commands. A material stored definition, invocation, or ordinary static
  input change emits `FF073`; enable `no_formula_defined_xlm_action_changes`
  for `FFP073`. Public output exposes only invoking-cell, selected-call, and
  relevant definition counts; targets, handler names, formulas, arguments,
  locations, and name identities remain private. FormulaFence does not
  evaluate a formula, resolve a target/handler, load a DLL, send DDE, execute a
  macro or program, or infer whether an action succeeds. Direct worksheet
  action calls and raw XLM macro-sheet parts remain intentionally outside this
  narrow boundary; this finite inventory does not claim to interpret arbitrary
  XLM commands, and dynamic/unresolved inputs remain static-coverage limits.
  The separate `--redact-formula-defined-xlm-actions` shared-artifact mode
  hides direct selected-action material, exact changed static inputs, and
  changed private name-chain evidence after comparison and policy evaluation
  without resolving a target/handler, executing an action, or changing any
  comparison fact.
- Legacy XLM GET.CELL is an XLM information function. FormulaFence separately
  inventories only calls stored in formula-defined names and named LAMBDA
  bodies, then propagates them through nested and sheet-local names to invoking
  formula cells. Microsoft's [C API
  reference](https://learn.microsoft.com/en-us/office/client-developer/excel/programming-with-the-c-api-in-excel)
  identifies GET.CELL as xlfGetCell. A material stored definition, invocation,
  or ordinary static argument-input change emits FF070; enable
  no_formula_defined_xlm_get_cell_changes for FFP070. Public output exposes
  only invoking-cell, call, and relevant definition counts; information types,
  references, formulas, arguments, locations, and name identities remain
  private. FormulaFence does not evaluate a call, determine its information
  type, resolve a dynamic reference, render formatting or display text, inspect
  comments/protection, or simulate other Excel state. Direct worksheet GET.CELL
  formulas and raw XLM macro-sheet parts are intentionally outside this narrow
  boundary; dynamic/unresolved inputs remain static-coverage limits. The
  separate `--redact-formula-defined-xlm-get-cell-calls` shared-artifact mode
  hides direct stored GET.CELL material, exact changed static inputs, and
  changed private name-chain evidence after comparison and policy evaluation
  without evaluating a call, resolving a dynamic reference, or changing any
  comparison fact.
- Selected legacy XLM environment-information calls GET.WORKBOOK,
  GET.WORKSPACE, and GET.DOCUMENT are separately inventoried only when stored
  in formula-defined names and named LAMBDA bodies, then propagated through
  nested and sheet-local names to invoking formula cells. Microsoft's [C API
  reference](https://learn.microsoft.com/en-us/office/client-developer/excel/programming-with-the-c-api-in-excel)
  identifies workspace information functions such as GET.CELL and
  GET.WORKBOOK; its [xlfFree
  example](https://learn.microsoft.com/en-us/office/client-developer/excel/xlfree)
  demonstrates GET.WORKSPACE returning platform information, and its [expression
  evaluation reference](https://learn.microsoft.com/en-us/office/client-developer/excel/excel-worksheet-and-expression-evaluation)
  identifies GET.DOCUMENT as an XLM information function. A material stored
  definition, invocation, or ordinary static argument-input change emits
  FF071; enable no_formula_defined_xlm_environment_information_changes for
  FFP071. Public output exposes only invoking-cell, call, and relevant
  definition counts; information types, references, formulas, arguments,
  locations, and name identities remain private. FormulaFence does not
  evaluate a call, determine its information type, resolve a dynamic reference,
  or simulate workbook/workspace/document/client/add-in/printer state. It does
  not assert that a state-only workbook change alters a stored call. Direct
  worksheet calls and raw XLM macro-sheet parts remain intentionally outside
  this narrow boundary; dynamic/unresolved inputs remain static-coverage
  limits. The separate
  `--redact-formula-defined-xlm-environment-information-calls`
  shared-artifact mode hides direct selected-call material, exact changed static
  inputs, and changed private name-chain evidence after comparison and policy
  evaluation without evaluating a call, simulating Excel state, or changing any
  comparison fact.
- Native CELL and INFO calls can observe file/location/content or operating
  environment information outside ordinary visible precedents. Native SHEET and
  SHEETS calls can observe workbook tab position or, with an omitted SHEETS
  reference, the workbook tab count. FormulaFence inventories all four in
  worksheet formulas, formula-defined names, and named LAMBDA bodies, then
  propagates private signals through nested and sheet-local names to invoking
  formula cells. Microsoft's [CELL function
  documentation](https://support.microsoft.com/en-us/office/cell-function-51bd39a5-f338-4dbe-a33f-955d67c2b2cf)
  notes that an omitted CELL reference can use the selected cell at calculation
  time; FormulaFence therefore aggregates that subset privately without
  inferring what the selection is. Microsoft's [INFO function
  documentation](https://support.microsoft.com/en-au/office/info-function-725f259a-0e4b-49b3-8b52-58815c69acae)
  describes operating-environment information such as directory, platform, and
  calculation mode. Microsoft's [SHEET function
  documentation](https://support.microsoft.com/en-us/excel/functions/sheet-function)
  and [SHEETS function
  documentation](https://support.microsoft.com/en-us/excel/functions/sheets-function)
  document that hidden, very-hidden, macro, chart, and dialog sheets are
  included. A material stored definition, invocation, or ordinary static
  argument-input change emits FF072; with a complete raw OOXML tab catalog,
  a tab membership/order/name change also emits FF072 for stored SHEET or
  omitted-reference SHEETS calls. Visibility-only changes do not satisfy that
  condition. Enable no_formula_environment_information_changes for FFP072.
  Public output exposes only formula-cell, call, relevant definition, and
  omitted-reference counts; information types, references, formulas, arguments,
  locations, name identities, and raw tab-catalog comparison material remain
  private; ordinary sheet inventory remains normal reviewer context.
  FormulaFence does not evaluate a call, determine an information type, resolve
  a dynamic argument, infer a selected cell, calculate a result, or simulate
  file/folder/client/workspace/workbook state. Explicit SHEETS references are
  not guessed as one-sheet versus 3-D; dynamic/unresolved inputs and incomplete
  tab catalogs remain explicit static-coverage limits. The separate
  `--redact-formula-environment-information` shared-artifact mode hides direct
  stored native-call material, exact changed static inputs, and changed private
  name-chain evidence after comparison and policy evaluation without evaluating
  a call, simulating Excel state, or changing any comparison fact.
- Office 2010 worksheet sparklines live in `x14:sparklineGroups` worksheet
  extensions, outside ordinary cell values. A group can be retargeted, moved,
  or have its type, axes, display, marker, line-weight, or colour controls
  changed; a nested sparkline can change its source formula or destination
  cell. FormulaFence reads raw x14 declarations before the ordinary reader
  drops them and privately compares group membership, source/date-axis
  formulas, destinations, and visual controls. A material change emits
  `FF048` and `no_worksheet_sparkline_changes` blocks it as `FFP048`. Profiles
  and reports expose only aggregate worksheet/group/sparkline,
  source/date-axis-source, colour-control, and malformed-metadata counts;
  formulas, locations, group properties, and colours remain private.
  Equivalent local direct-range spelling, Boolean/numeric spelling, colour
  case, and declaration order normalize away. Missing, duplicate, malformed,
  unreadable, oversized, or over-budget metadata becomes a coverage warning;
  raw worksheet XML is bounded to 16 MiB per worksheet, 64 MiB per workbook,
  and 512 parts. A Sparkline Group-removed temporary reader copy is made only
  after raw inspection, so lossy reader support cannot erase the evidence.
  FormulaFence does **not** calculate source values, resolve names/external
  sources, render a sparkline, assess visual accessibility, or guarantee
  cross-version Excel rendering equivalence. This boundary follows the Open
  XML [SparklineGroup](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.office2010.excel.sparklinegroup?view=openxml-3.0.1)
  and [CT_Sparkline](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/6b28a993-e0fd-451d-860e-35097c6baa77)
  definitions.
- SpreadsheetML XML Maps can attach an embedded schema and refresh/export
  behavior to XML table columns or individual worksheet cells. A changed map
  can redirect an XPath, switch a file/connection binding, change a target
  cell, or alter append, format, sort/filter, and validation behavior without
  changing ordinary cells. FormulaFence reads raw XML Maps, table
  XML-column-property, and single-cell table declarations before ordinary
  workbook readers discard them. It privately compares schemas, map/data-
  binding controls, table/single-cell bindings, and related
  workbook/worksheet relationship targets. A material change emits FF049 and
  `no_xml_mapping_changes` blocks it as FFP049. Profiles and reports expose
  only aggregate map/schema/binding, file/connection, table, single-cell, and
  malformed-metadata counts; schemas, names, XPath expressions, identities,
  cells, connection identities, and relationship targets remain private.
  Equivalent Boolean/unsigned-integer spellings, relationship IDs/order, and
  equivalent internal target spelling stay quiet. Missing, duplicate,
  malformed, unsafe, unbound, unreadable, oversized, or over-budget metadata
  becomes a coverage warning; reads are bounded to 16 MiB per part, 64 MiB per
  workbook, and 512 parts. Before an XML Map, mapped-table, or single-cell
  table tree is materialized, FormulaFence streams 32,768 elements per part and
  65,536 across the complete inventory; a successfully parsed structural
  overage becomes visible `FF010`/`FF049` coverage evidence. FormulaFence does
  **not** import/export XML,
  validate XML data or schemas, open a file/connection, fetch data, calculate
  a refresh, or infer Excel client behavior. This boundary follows the Open XML
  [Map](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.map?view=openxml-3.0.1),
  [XmlProperties](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.xmlproperties.xpath?view=openxml-3.0.1),
  and [SingleXmlCells](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.singlexmlcells?view=openxml-3.0.1)
  definitions.
- Excel rich data types can place linked entity values, provider-backed fields,
  web-image associations, and worksheet value-metadata bindings outside normal
  cells. FormulaFence reads Rich Value Data/structure/type/array/supporting
  property-bag/style/web-image/rich-value-relationship parts, their
  workbook/package relationships, and `XLRICHVALUE` bindings before normal
  readers can omit them. A material change emits `FF051` and
  `no_rich_data_changes` blocks it as `FFP051`. Profiles and reports expose
  aggregate part/value/structure/array/property-bag/binding/bound-cell/image/
  relationship/external-reference and malformed-metadata counts only; entity
  values, provider data, field names, identifiers, URLs, image references,
  relationship IDs, and bound-cell locations remain private. Writer-selected
  relationship IDs/order and equivalent internal targets normalize away.
  Missing, duplicate, malformed, unsafe, unreadable, oversized, or over-budget
  metadata becomes a coverage warning; reads are bounded to 16 MiB per XML
  part, 64 MiB per workbook, and 512 parts. Before raw rich-data package XML is
  materialized, FormulaFence streams 32,768 elements per part and 65,536 across
  the inventory; a successfully parsed structural overage becomes visible
  `FF010`/`FF051` coverage evidence. Its worksheet-binding pass streams only
  required cell attributes after the shared semantic-reader preflight, so it
  does not retain a second worksheet tree. FormulaFence does **not** contact
  providers, refresh values, calculate formulas, fetch or validate targets, or
  infer Excel client behavior. This boundary follows Microsoft's
  [Rich Value Data](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/896934fd-8df7-43f4-b154-2d39371c270d),
  [Rich Value Structure](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/d90f6d91-d868-4b94-9d26-ec3b1492cec6),
  [Rich Value Types](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/5d213b66-3196-4516-b63c-eef80d926f4a),
  and [Rich Value Web Image](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/4f3a80fd-1776-407f-8807-2497a4692dea)
  definitions.
- Generic Custom XML, workbook-bound Custom Data, and custom document
  properties can retain an add-in's workbook-specific approval, workflow, or
  integration state outside ordinary cells. FormulaFence reads generic
  `customXml/item*.xml` data, its property/schema parts and relationships,
  workbook-linked `xl/customData` property/binary parts, and
  `docProps/custom.xml` before ordinary readers can omit them. Power Query
  `DataMashup` remains exclusively under the Power Query control boundary. A
  material persisted-state change emits `FF052` and
  `no_custom_data_store_changes` blocks it as `FFP052`. Profiles and reports
  expose aggregate part/schema/relationship/payload/document-property and
  malformed-metadata counts only; custom XML, schema URIs, property names and
  values, storage IDs, binary payloads, relationship IDs, and targets remain
  private. Writer-selected relationship IDs/order and document-property `pid`
  normalize away; Custom XML `itemID` and Custom Data `id` storage identities
  are compared privately because add-ins can bind state to them. Missing,
  duplicate, malformed, unsafe, unbound, unreadable, oversized, or over-budget
  metadata becomes a coverage warning; reads are bounded to 16 MiB per part,
  64 MiB per workbook, and 512 parts. Before any custom-state XML tree is
  materialized, FormulaFence streams its complete structure with a
  32,768-element per-part limit and a 65,536-element aggregate limit. These
  limits bound CI allocation and make a well-formed overage visible `FF052`
  coverage evidence; they do not define Excel file validity. Opaque binary
  Custom Data remains byte-bounded rather than being interpreted as XML. The
  Power Query scanner receives only `DataMashup` parts safely classified by the
  same bounded custom-state pass, so a rejected generic custom XML tree is not
  materialized again. FormulaFence does **not** execute an add-in, resolve a
  property, follow or fetch a target, interpret a binary payload, calculate
  formulas, or infer Excel client behavior. This boundary follows Microsoft's
  guidance on
  [persisting add-in state](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/persisting-add-in-state-and-settings),
  [Custom Data](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/7c53f6f4-fea8-43f7-a4b0-ba6e14d0eb78),
  and [Custom Data Properties](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/1f4aa666-c966-4ecf-8399-28390399c891).
- OPC package signatures and VBA project signatures are distinct stored
  integrity/provenance surfaces. A workbook can preserve ordinary cells and
  even `xl/vbaProject.bin` while the package-root signature origin, XML signature
  envelope/signed references, optional certificate-part relationship/payload, or
  classic/Agile/V3 VBA signature payload changes. FormulaFence reads those raw
  relationships and bounded parts before normal readers can omit them. A
  material envelope change emits `FF050` and
  `no_digital_signature_changes` blocks it as `FFP050`. Profiles and reports
  expose aggregate counts only; XML signature material, reference URIs,
  certificate identities/contents, binary signature payloads, relationship IDs,
  and targets remain private. Equivalent IDs/order/internal targets and XMLDSIG
  base64 whitespace normalize away. Missing, duplicate, malformed, unsafe,
  unbound, unreadable, oversized, or over-budget metadata becomes a coverage
  warning; reads are bounded to 16 MiB per part, 64 MiB per workbook, and 512
  parts. Before an XMLDSIG envelope is materialized, FormulaFence streams
  32,768 elements per part and 65,536 across the signature inventory; a
  successfully parsed structural overage becomes visible `FF010`/`FF050`
  coverage evidence. Certificate and VBA-signature binary payloads remain
  byte-bounded rather than being interpreted as XML. FormulaFence does **not**
  validate a signature/digest/transform,
  reference coverage, certificate chain/identity/trust/expiry/revocation,
  timestamp, signed contents, or VBA code; it does not fetch certificates or
  contact trust services. Microsoft's [OPC digital-signature
  overview](https://learn.microsoft.com/en-us/previous-versions/windows/desktop/opc/open-packaging-conventions-overview)
  assigns signer/trust validation to the package consumer.
- Traditional Excel Notes are stored in worksheet-associated SpreadsheetML
  comments parts and their display declarations live in worksheet
  `legacyDrawing` VML parts. FormulaFence follows the worksheet bindings and
  privately compares author association, text/rich-text presentation, cell
  association, comment properties, Note VML visibility/layout, and relationship
  semantics. It recognizes the `tc={GUID}` legacy placeholder used to reconcile
  a threaded comment, and treats that declaration as a separate guarded surface.
  A material change emits `FF046` and `no_legacy_comment_changes` blocks it as
  `FFP046`. Profiles and reports expose only aggregate worksheet/part,
  author/comment/text/rich-text/property/placeholder, Note-shape/visibility/
  anchor, relationship, and malformed-metadata counts; Note text, author
  identities, references, VML, targets, IDs, and GUIDs remain private.
  Consistent writer-generated VML shape/comment shape/relationship IDs and
  placeholder GUIDs normalize away. Missing, duplicate, malformed, unbound,
  unsafe, unreadable, oversized, or over-budget metadata becomes a visible
  coverage warning; XML reads are bounded to 16 MiB per part, 64 MiB per
  workbook, and 512 parts. Before a comments or Note-VML tree is materialized,
  FormulaFence streams complete structure: 32,768 elements per part and 65,536
  across the legacy-Note scan. This CI allocation boundary turns a well-formed
  structural overage into visible `FF010`/`FF046` coverage evidence rather
  than a workbook-validity verdict. The ordinary reader uses a Note-quarantined
  temporary copy only after raw inspection, so unsafe targets and
  parser-tolerance differences cannot erase evidence. FormulaFence does
  **not** render Notes/VML, resolve authors, fetch
  targets, execute linked content, calculate client placement, or infer
  notification, permission, account, cloud, or client-visibility behavior.
  This boundary follows the Open XML
  [Comment](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.comment?view=openxml-3.0.1),
  [Authors](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.authors?view=openxml-3.0.1),
  and [LegacyDrawing](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.legacydrawing?view=openxml-3.0.1)
  definitions, plus Microsoft's [threaded-comment placeholder
  rule](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/6383f002-c90b-401c-a1d7-66b97b14cb3e).
- Modern threaded comments are stored in worksheet-associated comment parts and
  a workbook-associated persons part, outside ordinary cells. FormulaFence
  follows those bindings and privately compares comment/reply graphs, stored
  text, cell/timestamp/resolution declarations, mention range/person links,
  extension material, and person records. A material change emits `FF045` and
  can be blocked with `no_threaded_comment_changes`. It normalizes consistent
  writer-generated comment, parent, person, mention, and package
  relationship-ID rewrites by rebuilding the private graph first. Profiles and
  reports expose only aggregate worksheet/part/thread, comment/reply/resolved/
  text, mention, person, relationship, and malformed-metadata counts. It does
  **not** render comments, validate mention offsets, notify or resolve users,
  determine legacy-placeholder rendering, or infer permissions, cloud state,
  or client visibility. Missing, duplicate, malformed, unbound, unsafe,
  unreadable, oversized, or over-budget metadata becomes a visible coverage
  warning; XML reads are bounded to 16 MiB per part, 64 MiB per workbook, and
  512 parts. Before materializing comment or person trees, FormulaFence streams
  complete XML structure: 32,768 elements per part and 65,536 across the
  threaded-comment scan. This is a CI allocation and coverage boundary, not a
  workbook-validity limit; a well-formed structural overage becomes visible
  `FF010`/`FF045` evidence. After raw inspection, the temporary ordinary-reader
  copy removes threaded-comment and person relationship bindings so no current
  or future underlying reader can re-materialize a rejected tree; the original
  package and private raw evidence remain intact. The boundary follows Microsoft's
  [threaded-comment overview](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/e0fb917a-1107-409a-852f-13b47aea70dc),
  [Threaded Comments part](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/66e1875d-c60a-48eb-bf88-41066d45fea8),
  [Persons part](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/1a170d26-42a2-46f0-b2b6-0ff1dec1c344),
  and [schema](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/adb84732-9fc8-48b6-bddc-6b0bcdaad940).
- Non-chart Worksheet DrawingML regular shapes (`xdr:sp`), connectors
  (`xdr:cxnSp`), nested groups (`xdr:grpSp`), and recognized SmartArt
  `xdr:graphicFrame` objects are followed from standard worksheet `drawing`
  relationships before the workbook reader can discard their declarations.
  For a non-chart graphic frame, FormulaFence recognizes the DrawingML Diagram
  `a:graphicData` URI and requires one `dgm:relIds` declaration. It supports
  transitional and Strict DrawingML, privately fingerprints supported
  anchor/layout and frame/shape/group/connector XML; the explicitly bound
  diagram data (`r:dm`), layout (`r:lo`), quick-style (`r:qs`), and colours
  (`r:cs`) parts; direct worksheet-drawing `diagramDrawing` rendering parts;
  and bounded direct internal Image targets from a Diagram Data part; connector
  `stCxn`/`endCxn` attachment semantics; macro assignments; text
  links; click/hover relationship semantics; and visible text/presentation
  declarations. Profiles expose only safe worksheet/drawing/anchor,
  shape/text/connector/group, graphic-frame/SmartArt-component, Diagram Data
  image part/fingerprinted/uninspected, connector-attachment, text paragraph/
  run, macro/text-link/hyperlink, relationship, and malformed-control counts.
  A material change emits `FF044` and can be blocked with
  `no_worksheet_drawing_shape_changes`. Consistent
  non-visual and connector endpoint ID rewrites, worksheet-DrawingML
  relationship-ID rewrites, and colour-case spelling are normalized.
  FormulaFence does **not** render DrawingML, resolve themes or contrast,
  calculate text links, execute macro assignments, retrieve external targets,
  calculate final SmartArt layout, or decode/render media. It hashes only
  bounded direct internal Diagram Data Image targets (32 MiB per image, 64 MiB
  per workbook, and 512 images), and does not follow any other component-side
  SmartArt relationship. Native pictures are handled by the separate
  worksheet-image boundary, chart frames remain in `FF030`, and unknown
  non-chart graphic-frame URI types are coverage gaps. Missing, duplicate,
  malformed, unsafe, oversized, over-budget, or unsupported metadata becomes
  visible parser-coverage evidence. XML reads are bounded to 16 MiB per part,
  64 MiB per workbook, and 512 parts. This scope follows the Open XML
  [`xdr:sp` Shape definition](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.spreadsheet.shape?view=openxml-3.0.1),
  [`xdr:cxnSp` ConnectionShape definition](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.spreadsheet.connectionshape?view=openxml-3.0.1),
  Microsoft's [Graphic Object Data](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oe376/f58e82a5-5590-4e36-b178-e12989960415),
  the OOXML [Diagram Data Part](https://ooxml.info/docs/14/14.2/14.2.4/),
  and [Diagram relationship IDs](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.diagrams.relationshipids?view=openxml-3.0.1)
  references.
- Native worksheet image controls are followed from worksheet `drawing`, direct
  `picture`, and `legacyDrawingHF` relationships before ordinary readers can
  discard those visual bindings. FormulaFence privately compares anchored
  transitional/strict DrawingML `xdr:pic` objects (including group-contained
  pictures), worksheet backgrounds, and VML-backed header/footer watermark
  images, along with their anchors, visual declarations, relationship
  semantics, and bounded direct payload hashes. Profiles and reports expose
  only safe worksheet/picture/anchor/background/header-footer/image-payload/
  relationship/malformed-control counts. Image bytes, image names/descriptions,
  visual formatting, anchors, relationship IDs/targets, and raw XML remain
  private. A material change emits `FF059` and can be blocked with
  `no_worksheet_image_changes`. Non-visual DrawingML/VML IDs and consistent
  relationship-ID rewrites normalize. FormulaFence does **not** render or
  decode media, fetch a target, resolve themes, calculate visibility, cropping,
  z-order, print pagination, or client behavior. Charts, rich-data/in-cell
  images, Theme images, ActiveX/OLE image controls,
  regular/group/connector/SmartArt drawing controls, and
  header/footer text remain in `FF030`, `FF051`, `FF053`, `FF029`, `FF044`, and
  `FF056` respectively. XML reads are bounded to 16 MiB per part, 64 MiB per
  workbook, and 512 parts; direct payload hashing is bounded to 32 MiB per
  part, 64 MiB per workbook, and 512 parts. Missing, duplicate, malformed,
  unsafe, unreadable, oversized, or over-budget material is visible coverage
  evidence. The boundary follows Open XML's
  [`xdr:pic` Picture definition](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.spreadsheet.picture?view=openxml-3.0.1),
  Microsoft's [worksheet background guidance](https://support.microsoft.com/en-us/excel/add-or-remove-a-sheet-background),
  and [header/footer watermark guidance](https://support.microsoft.com/en-us/excel/get-started/add-a-watermark-in-excel).
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
  refresh data, establish source trust, or calculate/render a PivotTable report.
- Before private parsing, raw `xl/connections*.xml` parts are streamed through
  32,768-element per-part and 65,536-element Connections-scan structural
  limits, in addition to 16 MiB per-part, 64 MiB aggregate, and 512-part read
  limits. A valid structural overage is represented only by private
  opaque coverage evidence, produces `FF010`, and remains materially
  diff-visible via `FF023`; malformed input retains the ordinary parser warning.
  This allocation boundary applies specifically to Connections XML rather than
  every raw external-data package reader.
- Before private parsing, selected raw query-table relationship targets are
  streamed through 32,768-element per-part and 65,536-element shared-scan
  structural limits, in addition to 16 MiB per part, 64 MiB aggregate, and
  512-part read limits. A private snapshot template is reused when one target
  is bound from more than one worksheet, avoiding repeated recursive opaque
  metadata canonicalization. A valid structural overage is represented only by
  private opaque coverage evidence, produces `FF010`, and remains materially
  diff-visible via `FF023`; malformed input retains the ordinary parser
  warning. This allocation boundary applies specifically to the raw query-table
  reader rather than every table or external-data reader.
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
- Before private parsing, selected raw
  `xl/externalLinks/externalLink*.xml` parts and the direct `.rels` parts used
  by the external-link inventory or package-indexed resolver are streamed
  through 32,768-element per-part and 65,536-element shared-scan structural
  limits, in addition to 16 MiB per part, 64 MiB aggregate, and 512-part read
  limits. A valid structural overage is represented only by private opaque
  coverage evidence, produces `FF010`, and remains materially diff-visible via
  `FF025`; malformed input retains the ordinary parser warning. This allocation
  boundary applies specifically to these raw external-link readers rather than
  every package relationship reader.
- A separate private static external-workbook link-surface ledger covers literal
  endpoints stored in worksheet formulas, workbook/sheet-local defined names,
  data-validation criteria, and standard DrawingML/ChartEx chart formula
  elements. It catches same-location source/target swaps without evaluating a
  formula, opening/resolving a source, refreshing data, or trusting a cache.
  Reports expose only surface/endpoint counts; source paths, workbook/sheet/name
  identities, formulas, ranges, and chart-part identities stay inside the
  ledger's private signature. A material change emits `FF081` and can be blocked with
  `no_external_workbook_link_surface_changes`; chart parts with unavailable
  formula coverage make that guard fail closed.
- Generic semantic output intentionally remains normal local reviewer context,
  so `diff`, `check`, and `portfolio` offer an explicit
  `--redact-external-workbook-links` rendering mode for artifacts that cross
  that boundary. It walks serialized output values only after comparison and
  policy evaluation, replacing a whole value that contains a parser-recognized
  literal static external-workbook endpoint. A conservative lexical fallback
  also hides visibly embedded bracketed/dynamic endpoint literals without
  evaluating a formula. The mode does not mutate snapshots, policy facts, or
  exit status; it does not follow a link, recover a value assembled from text
  fragments, or claim to redact unrelated sensitive workbook material. It is
  opt-in so ordinary local reports retain their established detailed evidence.
- The separate `--redact-formula-external-actions` rendering mode is for the
  stored `FF064` formula-action/provider and `FF074` direct-DDE boundary. After
  comparison and policy evaluation, it replaces serialized direct action/DDE
  formula material and the before/after evidence for a changed action/DDE cell
  or an exact private static input recorded as reaching one. A formula-defined
  name can route to an action through another name without spelling a native
  action itself, so a relevant private definition-chain change conservatively
  redacts changed defined-name before/after values as well. The mode never
  exposes the private cell/name set used for that decision; it does not mutate
  a snapshot, finding, policy fact, or exit status; and it does not calculate a
  formula, contact a provider or DDE server, follow a link, or reconstruct a
  dynamically assembled destination. It is not a general secret scrubber and
  does not replace the external-workbook-link sharing boundary.
- The separate `--redact-python-in-excel` rendering mode is for direct stored
  `PY` source/material and exact changed static cells recorded by the private
  dependency graph as reaching an inventoried PY formula. It runs after
  comparison and policy evaluation, replaces whole direct PY formula strings
  and before/after evidence for those changed formula/input cells, and never
  exposes the private source-cell set used for that decision. It does not parse
  or execute Python, calculate a formula, contact Microsoft Cloud, reconstruct
  a runtime value, mutate a snapshot/policy fact/exit status, or claim to
  redact arbitrary workbook material. It does not replace the external-link or
  formula-action sharing boundaries.
- The separate `--redact-office-custom-functions` rendering mode is for direct
  stored namespaced `FF066` call material, exact changed static cells recorded
  by the private dependency graph as reaching an inventoried call, and changed
  formula-defined-name evidence when the private resolved custom-function chain
  changed. It runs after comparison and policy evaluation, replaces whole direct
  formula strings and before/after evidence for those changed cells or names,
  and never exposes the private cell/name set used for that decision. It does
  not evaluate a formula, load a manifest or add-in, execute JavaScript, contact
  a custom-function runtime, reconstruct a runtime argument, mutate a
  snapshot/policy fact/exit status, or claim to redact arbitrary workbook
  material. It does not replace the external-link, formula-action, or Python
  sharing boundaries.
- The separate `--redact-unqualified-runtime-functions` rendering mode is for
  direct stored bare `FF075` candidate material, exact changed static cells
  recorded by the private dependency graph as reaching an inventoried call, and
  changed formula-defined-name evidence when the private resolved runtime-name
  chain changed. It runs after comparison and policy evaluation, replaces whole
  direct formula strings and before/after evidence for those changed cells or
  names, and never exposes the private cell/name set used for that decision.
  It does not evaluate a formula, resolve or load VBA, COM/Automation, XLL, or
  another provider, execute code, contact a runtime, reconstruct an argument,
  mutate a snapshot/policy fact/exit status, or claim to redact arbitrary
  workbook material. It does not replace the external-link, formula-action,
  Python, or Office custom-function sharing boundaries.
- The separate `--redact-worksheet-code-resource-registrations` rendering mode
  is for direct stored `FF067` `REGISTER.ID` material, exact changed static
  cells recorded by the private dependency graph as reaching an inventoried
  registration, and changed formula-defined-name evidence when the private
  resolved registration chain changed. It runs after comparison and policy
  evaluation, replaces whole direct formula strings and before/after evidence
  for those changed cells or names, and never exposes the private cell/name set
  used for that decision. It does not evaluate a formula, resolve a module path,
  load a DLL/XLL, inspect host trust settings, execute code, contact a provider,
  reconstruct an argument, mutate a snapshot/policy fact/exit status, or claim
  to redact arbitrary workbook material. It does not replace the external-link,
  formula-action, Python, Office custom-function, or unqualified runtime-
  function sharing boundaries.
- Every canonical root or part-level OPC relationship part is also inspected
  independently for `TargetMode="External"`, including opaque relationships
  no feature-specific scanner can reach. FormulaFence retains source, type,
  endpoint, and malformed-metadata evidence only in private signatures and
  exposes aggregate relationship part/source/target plus hyperlink/image/other
  counts. A material change emits `FF063` and can be blocked with
  `no_external_relationship_changes`; relationship-ID-only rewrites normalize.
  Duplicate, orphaned, malformed, unsafe, unreadable, oversized, or
  over-budget metadata is coverage evidence. FormulaFence does **not** resolve,
  fetch, open, execute, or establish trust for any relationship target. XML
  reads are bounded to 16 MiB per part, 64 MiB per workbook, and 512 parts.
- Excel 4.0 / XLM macro sheets are read directly from their raw Macro Sheet XML
  package parts before a workbook library can omit their executable cells.
  FormulaFence binds the documented workbook relationships to those parts,
  privately fingerprints accepted XML plus related-part relationships, and
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
- Before private tree parsing of a selected XLM macro-sheet part, FormulaFence
  streams its XML structure. The macro-sheet scanner permits 32,768 elements
  per part and 65,536 across its scan, alongside 16 MiB per part, 64 MiB
  aggregate, and 512 parts. A valid structural overage becomes opaque private
  program evidence with a streamed content fingerprint and explicit coverage
  warning, keeping `FF010` and `FF026` diff-visible without retaining raw macro
  XML. After raw scanning, the temporary ordinary-workbook reader gets an
  empty worksheet replacement for selected macro targets; Custom View
  sanitization excludes them as well. An invalid ordinary-sheet relationship
  alias to the same target is treated as XLM, with explicit coverage evidence.
  Thus neither secondary reader materializes the original program. Malformed
  macro XML reached before a structural overage retains its ordinary parser
  diagnostic. This boundary applies to the named raw macro-sheet XML readers,
  not every legacy workbook XML part.
- Legacy XLM automatic-macro routing is separately inspected from raw workbook
  defined names. Microsoft documents four workbook automatic-macro events:
  `Auto_Open`, `Auto_Close`, `Auto_Activate`, and `Auto_Deactivate`. FormulaFence
  recognizes an optional `_xlnm.` built-in prefix and counts only a
  workbook-scoped event name whose direct internal single-cell A1 definition
  targets a sheet declared through a raw XLM macro-sheet relationship. This catches a dispatch
  add, removal, or same-count retarget without asserting that a local name,
  ordinary-sheet target, external reference, or dynamic definition will run.
  The private signature retains the name/target/definition material; profiles,
  `FF076`, and `FFP076` expose only aggregate per-event counts. `FF076` is high
  severity and `no_xlm_automatic_macro_binding_changes` makes it `FFP076` in
  CI. FormulaFence does **not** evaluate or resolve a defined name, use the
  reserved/unused `definedName@xlm` attribute, parse or execute an XLM command,
  determine macro security settings, or claim that Excel will execute a
  binding. Missing/malformed workbook or relationship metadata remains a
  parser-coverage warning. The ordinary defined-name diff retains its normal
  reviewer context rather than becoming a redacted output channel.
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
- Office Web Add-ins are read directly from the documented workbook task-pane
  relationship, `taskpanes.xml` parts, direct task-pane-to-extension bindings,
  worksheet `x15:webExtensions` entries, and active in-content DrawingML
  `we:webextensionref` frames before the workbook reader can omit them.
  FormulaFence validates worksheet `appRef` entries against definition
  bindings, skips inactive `mc:Fallback` frame branches, and leaves an active
  frame's native-picture fallback to the worksheet-image boundary. It privately
  fingerprints task-pane configuration, add-in references, auto-show
  properties, bindings, worksheet formulas, snapshots, frame placement/XML,
  and direct relationship semantics while reporting only safe structural
  counts. Add-in identities, store references, property/binding values,
  formulas, XML, snapshots, and relationship targets remain private. A
  material change emits `FF028` and can be blocked with
  `no_office_web_addin_changes`. FormulaFence does **not** install, load,
  execute, or fetch an add-in or manifest, and it never follows external
  relationships. Missing, oversized, malformed, unbound, or over-budget parts
  remain visible parser-coverage warnings. Task-pane and web-extension XML
  reads are bounded to 16 MiB per part, 32 MiB per workbook, and 64 parts;
  worksheet-binding and in-content DrawingML reads are each bounded to 16 MiB
  per part, 64 MiB per workbook, and 512 parts. Other unrecognized extension
  or graphic-frame forms remain outside this boundary.
- PivotTable packages are followed through the bounded workbook cache and
  worksheet PivotTable relationship graph, then privately fingerprinted as
  view layout, cache-schema, shared-item, normalized relationship, and bounded
  raw cache-record material. Profiles expose only structural counts; names,
  source ranges, fields, item values, formulas, cache records, targets, XML,
  and payload bytes remain private. A material change emits `FF031` and can be
  blocked with `no_pivot_table_definition_changes`. Cache source and refresh
  settings remain under `FF023` / `no_external_data_connection_changes` so a
  refresh-only edit remains distinct. Relationship IDs, equivalent internal
  target spellings, and cache-ID renumbering are normalized. FormulaFence does
  **not** refresh a cache, calculate or render a PivotTable, infer
  PivotTable-to-cell impact, fetch an external target, or interpret OLAP,
  or extension-list semantics. Slicer and Timeline cache definitions are
  compared separately. Missing, malformed, orphaned, unbound,
  oversized, or over-budget material remains a visible parser-coverage warning.
  PivotTable/cache-definition XML reads are bounded to 16 MiB per part, 64 MiB
  per workbook, and 512 parts; raw cache-record hashes are bounded to 32 MiB per
  part, 64 MiB per workbook, and 512 parts. FormulaFence detaches cache-record
  relationships in a temporary reader copy before the underlying workbook
  library loads cells, so raw records are not eagerly materialized; the original
  workbook is never changed.
- Slicer and Timeline cache definitions are followed from their documented
  workbook extension declarations through explicit workbook relationships to
  bounded cache XML. FormulaFence privately fingerprints Slicer item selection,
  Timeline state/filter material, cache definitions, PivotTable/table source
  bindings, filtered-PivotTable bindings, and unexpected direct cache-part
  relationships while exposing only structural counts. Cache names, source
  fields, selected values, date ranges, PivotTable names, relationship targets,
  and XML remain private. A material change emits `FF032` and can be blocked
  with `no_slicer_timeline_cache_changes`. Relationship IDs, equivalent
  internal target spellings, coordinated Slicer/Timeline PivotCache extension-ID renumbering, known
  optional Slicer defaults, Boolean spellings, and Timeline GUIDs are
  normalized. FormulaFence does **not** apply a filter, calculate/render a
  PivotTable or table, infer downstream impact, fetch an external target, or
  model worksheet/drawing Slicer or Timeline view geometry/styles. Missing,
  malformed, orphaned, unbound, externally targeted, oversized, or over-budget
  material remains a visible parser-coverage warning. Cache XML reads are
  bounded to 16 MiB per part, 64 MiB per workbook, and 512 parts. Before a
  cache tree is materialized, FormulaFence also streams 16,384 XML elements per
  part and 32,768 across the complete Slicer/Timeline cache scan. This is a CI
  allocation boundary above Excel's documented 10,000 displayed filter
  drop-down items, not a validity limit; a well-formed structural overage stays
  visible as coverage evidence rather than being recursively canonicalized.
- Embedded Power Pivot/Data Model packages are followed from the workbook's
  explicit `powerPivotData` relationship and `x15:dataModel` declaration.
  FormulaFence privately fingerprints declaration material, normalized workbook
  relationship semantics, and bounded raw `xl/model/*.data` payloads while
  exposing only model-part, binding, declaration, table, relationship, payload,
  and coverage counts. Table/column/relationship names, connection details,
  DAX, stored values, targets, XML, and raw bytes remain private. A material
  change emits `FF033` and can be blocked with
  `no_power_pivot_data_model_changes`. Relationship IDs, equivalent internal
  target spellings, and GUIDs in Data Model metadata are normalized. FormulaFence
  does **not** deserialize the Analysis Services payload, evaluate DAX, refresh
  the model, calculate/render a report, infer model-to-cell impact, or fetch an
  external target. Missing, malformed, orphaned, unbound, externally targeted,
  unexpected directly related, oversized, or over-budget material remains a
  visible parser-coverage warning. Raw payload reads are bounded to 512 MiB per
  part, 512 MiB per workbook, and 16 parts.
- DrawingML chart definitions and cached presentation data are followed from
  standard worksheet or chartsheet `drawing` relationships through legacy
  `c:chart` parts, direct `c:userShapes` overlays, and Office 2016+ `cx:chart`
  ChartEx parts. ChartEx `mc:AlternateContent` graphic-frame bindings are
  recognized without treating their older-client fallback shape as a second
  worksheet control. FormulaFence privately fingerprints non-cache legacy
  chart material separately from `numCache`, `strCache`, and `multiLvlStrCache`
  material, plus overlay XML, normalized relationship semantics, ChartEx XML,
  and bounded direct internal related-part payloads. Supported ChartEx direct
  relationships are style, colour-style, drawing, image, theme-override, and
  embedded package; unsupported, external, or unsafe edges remain explicit
  coverage evidence. Profiles expose only structural counts; chart formulas,
  cached values, titles, shape text, relationship targets, XML, and payload
  bytes remain private. A material change emits `FF030` and can be blocked with
  `no_chart_definition_changes`. FormulaFence does **not** calculate a series
  formula, render a chart, infer chart-to-cell impact, follow an external
  target, parse media or embedded-package formats, resolve ChartEx second-hop
  relationships, or interpret ChartEx-specific or nested-chart visualization
  semantics. Missing, malformed, orphaned, unbound, unsupported, oversized,
  over-budget, or unrecognized material remains a visible parser-coverage
  warning. Chart and overlay XML reads are bounded to 16 MiB per part, 64 MiB
  per workbook, and 512 parts; direct related payload hashes are bounded to 32
  MiB per part, 64 MiB per workbook, and 512 parts. The relationship boundary
  follows Microsoft's [ChartEx part](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-odrawxml/5d0d453e-adac-43be-a797-59b9916593dd)
  and [ChartEx relationship-ID](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-odrawxml/d8ede39e-a36c-48ad-8a17-0086a2d0889b)
  definitions.
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
  workbook, and 512 parts. Before private XML canonicalization, FormulaFence
  streams complete structure: 32,768 elements per part and 65,536 across the
  embedded-control scan. A well-formed structural overage becomes visible
  `FF010`/`FF029` coverage evidence. Direct payload hashing is bounded to 32
  MiB per part, 64 MiB per workbook, and 512 parts. Missing, malformed,
  orphaned, unbound, oversized, or over-budget material remains a visible
  parser-coverage warning. VML/drawing layout, embedded payload formats, and
  behavior outside this relationship-backed chain are not modeled.
- Power Query Data Mashup custom XML is inspected without serializing its M
  formulas or data/source material. FormulaFence privately compares the
  `Section1.m` formula document, logical package content, stable query metadata,
  and formula-firewall permissions. Profiles expose only structural counts and
  safe controls; query names, locations, metadata values, embedded content,
  telemetry IDs, and user-bound permission bindings remain private. `sqmid`
  telemetry and result-only refresh metadata are intentionally ignored. A
  material change emits `FF024` and can be blocked with
  `no_power_query_changes`. FormulaFence does **not** execute M, refresh a
  query, establish source trust, or infer returned values. The outer Custom XML
  boundary does not limit decoded Data Mashup metadata or permission XML, so
  FormulaFence streams each document before private parsing: 32,768 elements
  per document and 65,536 across the Power Query scan. A successfully parsed
  overage is visible `FF010`/`FF024` coverage evidence; malformed input retains
  its established diagnostic. Every nested Data Mashup ZIP central directory is
  separately preflighted before Python can materialize its entry catalog: 768
  KiB source, 1 KiB raw member names, and 512 parts shared across logical
  packages and metadata embedded-content ZIPs in the scan. ZIP64, multi-disk,
  malformed, and filename-rewriting metadata (Unicode-path aliases, NULs, or
  platform separators) is retained as coverage evidence. The logical package
  alone is read after that preflight: stored/deflated
  entries only, 16 MiB per member, 64 MiB declared expanded data across the
  scan, and a 1,000:1 maximum member ratio. An unsafe nested ZIP is not read or
  cataloged beyond the boundary; FormulaFence preserves a private opaque
  fingerprint and a coverage warning instead of treating it as inspected
  content.
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
  subset and stay visible as coverage gaps. In candidate portfolio mode, the
  separately described global named-LAMBDA bridge may retain fully static
  external endpoints and fixed internal inputs only at an actual function call;
  a bare LAMBDA name never creates an edge.
- Static internal 3-D A1 references such as `Jan:Mar!B2:B10` are expanded over
  every worksheet in the endpoint tab span. FormulaFence compares the resolved
  span when the same 3-D formula survives a workbook change, because moving,
  adding, or removing tabs can change its semantics. Exact static external 3-D
  A1 spans are separately eligible only for a candidate source with a complete
  consistent tab catalog; malformed, endpoint-missing, non-A1, and all other
  external 3-D forms remain explicit coverage gaps.
- Explicit external-workbook references are detected. References assembled from
  text or macro code are not.
- The link-surface ledger covers only literal static endpoints in worksheet
  formulas, defined names, data-validation criteria, and standard/ChartEx chart
  formulas. It does not infer text-built links, evaluate formulas, follow an
  endpoint, or establish source trust. Conditional-formatting external-cell
  references are excluded because the SpreadsheetML conditional-formatting
  formula grammar forbids them; ordinary shapes, text boxes, OLE/DDE, and OPC
  relationships remain under their dedicated control boundaries.
- A formula that the underlying tokenizer cannot inspect is recorded by cell
  location in the profile, and a newly introduced one emits `FF016`; its graph
  is deliberately omitted rather than partially guessed.
- It inventories sheet visibility, defined names, calculation settings, the VBA
  payload; OPC package XML-signature/certificate-part relationships and
  payloads; classic/Agile/V3 VBA signature payloads/relationships; XLM
  macro-sheet packages; RibbonX custom UI packages; Office Web
  Add-in task-pane packages, PivotTable view/cache-schema/shared-item/cached-
  record chains, Slicer and Timeline cache filter-definition chains, embedded
  Power Pivot/Data Model declaration/raw-payload chains, What-If Data Table and
  Scenario Manager declarations, worksheet/Table AutoFilter and row/column-
  visibility controls, material worksheet-display and worksheet print-layout
  controls, ignored-error warning suppressions, relationship-backed Named Sheet
  View and legacy Custom View controls, ordinary worksheet-cell hyperlink
  declarations/relationships,
  Office 2010 worksheet sparkline declarations, SpreadsheetML XML Map schema,
  refresh/export, table-column, single-cell, and relationship declarations,
  legacy Excel Note/comments/VML and threaded-placeholder package chains,
  modern threaded-comment/person package chains, non-chart Worksheet DrawingML
  regular/group/connector/recognized-SmartArt graphic-frame controls, native
  worksheet picture/background/header-footer image controls, legacy and ChartEx
  DrawingML chart definition/cached-presentation/overlay chains,
  relationship-backed worksheet ActiveX/form-control/legacy-VML/OLE chains, the
  protection controls above, external-data refresh controls, external-link
  packages, package-wide external OPC relationships, and private Power Query
  definition material. It does not yet
  interpret PivotTable OLAP or other extension-list semantics; deserialize or execute
  Power Pivot/Data Model content; apply Slicer/Timeline filters or model their
  worksheet/drawing view geometry/styles; ChartEx-specific visualization,
  second-hop relationship, or nested-chart semantics; future Named Sheet View extension/rich-
  sort or full differential-format semantics; unknown non-chart
  graphic-frame URI types, SmartArt rendering/final-layout behavior, or
  SmartArt component-side relationship targets other than bounded direct
  Diagram Data Image payloads; chart-to-cell impact; Ribbon image payloads;
  general VML/drawing-control
  layout beyond supported Note shapes; embedded OLE/package formats; unsupported
  worksheet Web Add-in extension or graphic-frame forms; Power Query runtime behavior or returned
  data; ordinary styles beyond direct protection assignments; complete Excel
  style-cascade results; or every OOXML part.
- The tool preserves Excel formula text and uses a limited A1-reference
  normalizer for peer-pattern detection; it is not an Excel-compatible parser
  or calculation engine.

For high-stakes use, treat FormulaFence as one control among independent review,
recalculation in the approved spreadsheet engine, input reconciliation, and an
appropriately qualified model owner.
