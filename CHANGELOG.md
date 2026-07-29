# Changelog

## 0.192.0 — 2026-07-28

- Add medium-severity `FF091` to `formulafence lint`: it reports recognized
  stored Excel error-checking suppressions whose review prompts may be hidden.
- Reuse FormulaFence's hardened ignored-error inventory and retain only
  aggregate warning categories, suppression-rule counts, and target-range
  counts. Individual ranges, formulas, and values remain private; malformed or
  unsupported material remains parser coverage evidence rather than being
  exposed as `FF091` detail.

## 0.191.0 — 2026-07-28

- Add high-severity `FF090` to `formulafence lint`: it reports an ordinary
  formula in a proven multi-cell static circular-reference component when
  workbook calculation iteration is disabled.
- Reuse only FormulaFence's resolved scalar static dependency graph, with an
  iterative strongly connected-component traversal. Direct self references
  remain under `FF087`; ranges are not expanded or used to close a cycle, and
  dynamic-reference, 3-D, spill, explicit-intersection, array, and
  tokenizer-failure territory remains quiet. JSON, Markdown, and SARIF retain
  only each affected location, the disabled-iteration fact, scope, and
  component size—never
  formula text, peer edges, or cached values.

## 0.190.0 — 2026-07-28

- Add high-severity `FF089` to `formulafence lint`: it reports a formula only
  when its well-formed saved result is an exact broken-reference error. The
  finding describes the last saved display state rather than evaluating the
  formula or claiming its current result is unchanged.
- Keep the cache boundary narrow and private. Other saved error kinds, missing
  or malformed cache records, and locations already covered by critical
  `FF088` remain quiet. JSON, Markdown, and SARIF retain only the affected
  location and no formula text, error value, or cached value.

## 0.189.0 — 2026-07-28

- Add critical-severity `FF088` to `formulafence lint`: it reports a stored
  formula only when tokenization exposes an actual `#REF!` error operand.
- Replace the shared broken-reference substring heuristic with the same exact
  token boundary. Formula text literals, quoted worksheet names, and
  tokenization failures remain quiet. The lint never evaluates a formula and
  JSON, Markdown, and SARIF retain only the location, never formula text or a
  cached value.

## 0.188.0 — 2026-07-28

- Add high-severity `FF087` to `formulafence lint`: it reports an ordinary
  formula only when FormulaFence's resolved scalar static dependency returns
  directly to that same formula cell while workbook calculation iteration is
  disabled (including the omitted OOXML default).
- Keep circular-reference inference deliberately narrow. Enabled iteration,
  indirect cycles, static ranges, dynamic references, spill references,
  explicit intersection, and array territory remain quiet. JSON, Markdown, and
  SARIF retain only the formula location, disabled-iteration fact, and
  direct-static scope—never formula text or cached values.

## 0.187.0 — 2026-07-28

- Add medium-severity `FF086` to `formulafence lint`: it reports a workbook
  containing formulas only when stored calculation properties explicitly
  combine `calcMode=manual` with `calcCompleted=false`, so the file records
  incomplete calculation before save.
- Keep the calculation-freshness boundary deliberately narrow. Manual mode
  alone, automatic mode, completed or omitted completion metadata, and
  formula-free workbooks remain quiet; the finding does not claim a particular
  cached result is stale or incorrect. JSON, Markdown, and SARIF retain only
  the two configuration flags and no formula text or cached values.

## 0.186.0 — 2026-07-28

- Add medium-severity `FF085` to `formulafence lint`: it reports an ordinary
  formula with an explicit direct-cell `locked=false` assignment on an actively
  protected worksheet, retaining only its location and the `direct_cell`
  protection scope.
- Keep protection inference deliberately narrow. Row, column, default-style,
  and allowed-edit-range precedence remain quiet until FormulaFence can model
  their complete effective state without guessing. JSON, Markdown, and SARIF
  remain formula-free; the shared 10,000-finding cap stays fail closed.

## 0.185.0 — 2026-07-28

- Add `FF084`, a static, single-workbook review signal for a pure local
  `SUM`, `AVERAGE`, `MIN`, `MAX`, or `COUNT` whose direct one-dimensional A1
  range stops before a bounded run of at least two literal numeric cells on the
  same row or column. The finding is medium severity and retains only the
  aggregate function, orientation, and range coordinates.
- Keep the detector intentionally quiet for named, table, external, 3-D,
  multiple-range, computed, array-territory, nonnumeric-gap, one-cell-gap, and
  tokenizer-failure cases. Add a 128-cell default inspection bound configurable
  with `--max-aggregate-omission-gap-cells`; the existing 10,000 total
  formula-lint finding cap remains fail closed.

## 0.184.0 — 2026-07-28

- Add `formulafence lint WORKBOOK`, a conservative single-workbook audit for
  interruptions inside copied formula blocks. It emits `FF082` for blank,
  stored-error, manual-value, or text interruptions and `FF083` for a formula
  outlier only when two matching immediate peers and a third contiguous peer
  establish the same relative-copy fingerprint.
- Make the CI gate deliberate: blanks and stored errors are high, manual values
  and formula outliers are medium, and text markers are low. JSON, Markdown,
  and SARIF reports retain only affected and peer coordinates, never formula
  text; array territory, tokenizer failures, short patterns, incomplete array
  metadata, oversized candidate sets, and oversized artifacts all fail closed
  or remain quiet as appropriate.

## 0.183.0 — 2026-07-28

- Let bounded worksheet readers reuse a private snapshot XML root after they
  have independently read and enforced their own part and aggregate budgets.
  Office Web Add-in, cell-hyperlink, sparkline, and native-image readers now
  receive deep copies rather than reparsing the same safe worksheet payload.
- Reuse remains tied to the private materialized source and the existing 128
  KiB-per-payload, 4 MiB-total, 16 KiB-per-tree, 2,048-elements-per-tree, and
  8,192-total-elements limits. Source, nested-mutation, character-guard, and
  specialized-reader regressions hold the boundary.

## 0.182.0 — 2026-07-28

- Reuse small parsed OOXML roots as private, element-accounted trees during one
  stable workbook snapshot. Each raw reader receives a deep copy, so reader
  mutation remains isolated while repeated worksheet XML parsing is avoided.
- Cache a tree only after its payload is already within the existing validated
  payload cache, at most 16 KiB and 2,048 elements per part, with 8,192
  retained elements in total. Derived trees and catalogs are discarded if the
  character-data parser limit changes; byte, per-tree, aggregate, source, and
  nested-mutation regressions hold the boundary.

## 0.181.0 — 2026-07-28

- Reuse bounded, immutable OOXML relationship and worksheet-part catalogs for
  one private workbook snapshot. Raw readers no longer repeatedly parse the
  same workbook relationship catalog and reconstruct the same standard,
  visual, and display worksheet maps.
- Keep the catalog cache tied to the existing validated-payload limits, cap
  retained relationship records at 2,048 and sheet records at 512, return fresh
  mutable maps to readers, and discard derived entries if the character-data
  parser limit changes. Source-isolation, record-budget, mutation-isolation,
  and guard-renewal regressions cover the boundary.

## 0.180.0 — 2026-07-28

- Keep worksheet-dimension column state implicit when a visual worksheet has
  no SpreadsheetML `<cols>` declaration. The raw inspector no longer allocates
  and compresses 16,384 default width/AutoFit states per ordinary columnless
  sheet.
- Preserve default dimensions, row dimensions, zero-width visibility handling,
  populated and malformed `<cols>` parsing, update budgets, and canonical
  evidence. Structural coverage makes the full-state signature raise if a
  columnless dimension fixture reaches it.

## 0.179.0 — 2026-07-28

- Reuse small, lexically validated OOXML payloads for the duration of one
  private workbook snapshot. Raw metadata readers still receive fresh parsed
  XML trees with the character-data guard, while repeated ZIP reads and lexical
  scans of the same worksheet, workbook, style, and relationship parts are
  avoided.
- Bind the cache to the materialized stable source and cap it at 128 KiB per
  payload and 4 MiB in total. Malformed, oversized, and unrelated archive
  parts are never retained. Structural coverage proves reader isolation,
  renewed character-data enforcement on cache hits, and no caching over the
  configured budget.

## 0.178.0 — 2026-07-28

- Skip five more full 16,384-column style-state expansions when a worksheet
  has no SpreadsheetML `<cols>` declaration: number formats, fonts, fills,
  alignments, and borders. The raw inspectors now keep the implicit workbook
  default for direct-cell comparisons rather than allocating and scanning an
  all-default list for every columnless sheet.
- Preserve exact private comparison evidence for direct cell styles,
  row styles, non-default workbook styles, and populated or malformed
  `<cols>` containers. Structural coverage makes each full-state signature
  raise if it is reached by a styled columnless worksheet.

## 0.177.0 — 2026-07-28

- Skip the full 16,384-column visibility-state expansion for a worksheet with
  no SpreadsheetML `<cols>` declaration. This removes four empty state arrays,
  four aggregate scans, and one full canonical-signature scan from the common
  columnless-sheet path.
- Preserve the exact canonical all-column control and count for
  `defaultColWidth="0"` without a `<cols>` declaration. Existing empty,
  malformed, and populated `<cols>` containers continue through the guarded
  control parser, update budget, and coverage-gap path unchanged.

## 0.176.0 — 2026-07-28

- Replace repeated materialized endpoint and static-reference maps while
  resolving external formula-defined names with live filtered lookup views.
  Direct and formula-derived external A1, 3-D, structured-table, and
  defined-name endpoints now share lazy alias traversal instead of copying and
  comparing growing maps for every definition.
- Preserve named-`LAMBDA` invocation boundaries, direct endpoint precedence,
  alias-chain fixed points, unresolved-name coverage, and deterministic
  external-link evidence. Incremental reverse-alias wake-ups rescan only names
  whose endpoint visibility can have changed; structural and portfolio
  regressions hold the representation and public output stable.

## 0.175.0 — 2026-07-28

- Compact propagated sensitive-call evidence for formula-defined names into one
  sparse, eleven-kind ledger per affected strongly connected component. Benign
  formula-name catalogs no longer retain direct, component, per-definition, or
  per-worksheet empty ledgers for action, DDE, custom-function, registration,
  XLM, and environment-information inspection.
- Skip the recursive-component propagation pass entirely when no
  formula-defined name has sensitive evidence, while retaining exact marker
  dependency budget accounting. Lazy ledger views preserve worksheet-local
  shadowing, qualified-name visibility, deterministic propagation order, and
  byte-identical public profiles for both benign and action-bearing catalogs.

## 0.174.0 — 2026-07-28

- Compact the private safety-marker catalog used while formula-defined names
  propagate action, DDE, custom-function, registration, XLM, and environment
  information evidence. FormulaFence now keeps its existing one
  identity-to-index catalog and generates each canonical marker only when a
  formula inspection actually resolves that name, rather than retaining eleven
  identity-to-marker maps and an eleven-kind reverse marker dictionary for a
  complete valid catalog.
- Preserve scope-aware visibility, deterministic propagated ledgers, and exact
  state-budget accounting while recovering generated markers through bounded,
  canonical prefix/index parsing. Structural and cross-release regressions hold
  the lazy representation in place and retain byte-identical public profiles.

## 0.173.0 — 2026-07-28

- Bound temporary sensitive-call propagation through formula-defined names.
  A compact acyclic name chain can legitimately preserve every inherited action,
  DDE, custom-function, registration, XLM, or environment-information call,
  but repeatedly retaining those prefixes was not covered by source, cell, or
  ordinary dependency-edge limits. FormulaFence now defaults to 1,000,000
  formula-defined-name states per workbook input and to one independent shared
  pool for each directory-portfolio side.
- Add `--max-formula-defined-name-states` to `profile`, `diff`, `check`, and
  `portfolio`, plus the matching `max-formula-defined-name-states` GitHub
  Action input. The positive-only budget reserves direct sensitive-ledger
  entries, direct name-marker dependencies, and each propagated component
  ledger before it is materialized. An overage returns status 2 before a CLI
  artifact is published; exact-boundary, no-output, portfolio-sharing, and
  Action-propagation regressions cover the guard.

## 0.172.0 — 2026-07-28

- Make formula-defined-name resolution scale with the work actually inspected
  instead of repeatedly materializing every visible workbook and worksheet-local
  name catalog. FormulaFence now reuses compact live scope overlays for static
  and resolved names, named-`LAMBDA` coverage, and all eleven sensitive-call
  marker ledgers.
- Preserve local-name shadowing, qualified-name visibility, unresolved-call
  coverage, and deterministic propagated ledgers while replacing repeated
  ready-component sorting with a priority queue. A structural regression holds
  the reusable views in place, and public-profile compatibility fixtures remain
  byte-identical to 0.171.0.

## 0.171.0 — 2026-07-28

- Bound the static local dependency records retained while a workbook snapshot
  is built. A compact formula-defined name can resolve to many local inputs at
  each caller, so source-byte and populated-cell limits alone did not cap the
  reverse/range dependency indexes used for impact review. FormulaFence now
  defaults to 2,000,000 retained local dependency-graph edges per workbook
  input, and to one independent shared pool for each side of a directory
  portfolio.
- Add `--max-dependency-edges` to `profile`, `diff`, `check`, and `portfolio`,
  plus the matching `max-dependency-edges` GitHub Action input. Every retained
  direct or range dependency and every additional legacy-CSE or observed
  dynamic-array output alias consumes the positive-only budget. An overage
  returns status 2 before a CLI artifact is published; exact-boundary,
  named-formula fanout, array-alias, portfolio-sharing, and Action propagation
  regressions cover the guard.

## 0.170.0 — 2026-07-28

- Bound `profile`'s public inventory construction before rendering. The CLI now
  defaults `--max-profile-records` to 100,000 and fails with status 2 before it
  materializes an over-limit profile object or publishes an output artifact.
  The budget counts every serialised profile-list record, including nested
  range, table-column, token/function, and dynamic-array reference entries.
- Keep this in-memory inventory boundary separate from the 32 MiB
  `--max-report-bytes` artifact boundary. Reviewers can deliberately opt up to
  a larger positive record budget for a complete known inventory; the
  programmatic `profile_snapshot` API retains its historical unlimited default.

## 0.169.0 — 2026-07-28

- Extend the 32 MiB UTF-8 rendered-artifact boundary to `profile` JSON and
  Markdown output. `profile` now accepts `--max-report-bytes`, returns status 2
  before output publication on overage, and preserves the complete artifact
  when a reviewer deliberately supplies the exact or larger budget.
- Reuse the existing incremental JSON encoder and streaming Markdown line
  writer, so profile publication has the same byte-accurate fail-closed
  behavior as `diff`, `check`, and `portfolio` without changing default
  programmatic rendering APIs.

## 0.168.0 — 2026-07-28

- Bound the rendered artifact produced by `diff`, `check`, and `portfolio` to
  32 MiB of UTF-8 text by default. This closes the gap between a safely bounded
  workbook reader and an impractically large report: a compact, repetitive
  spreadsheet can otherwise inflate into a large JSON or escaped HTML artifact.
- Expose `--max-report-bytes` and the matching public GitHub Action input.
  JSON/SARIF use incremental encoding, Markdown streams each line, and HTML
  writes each escaped review entry into one shared budget. An
  overage returns status 2 before any output path is written or replaced;
  callers can deliberately opt up for a known larger artifact.

## 0.167.0 — 2026-07-28

- Bound local semantic-change impact analysis across a complete comparison.
  `diff` and `check` now use one 100,000-state pool by default, and a directory
  portfolio shares one pool across every matched workbook. Each changed source
  and each statically reachable local dependency consumes a state, so a broad
  edit set cannot multiply the existing per-source traversal boundary into
  impractical CI work or retained report evidence.
- Expose the limit as `--max-change-analysis-states` and the matching public
  GitHub Action input. The command fails closed with status 2 on overage rather
  than emitting partial impact evidence; exact aggregate-boundary, direct CLI,
  portfolio-sharing, and Action validation/propagation regressions cover it.
- Reconstruct shortest paths lazily for the fixed serialized sample instead of
  eagerly materializing every reachable path prefix, preserving full local API
  access while avoiding quadratic evidence allocation on long dependency chains.

## 0.166.0 — 2026-07-28

- Bound the aggregate semantic state retained by a directory portfolio. Each
  baseline and candidate side now has a separate 2,000,000-populated-cell
  snapshot budget by default. FormulaFence counts the actual immutable
  snapshots it retains for nested diff evidence and candidate cross-workbook
  analysis, failing closed as soon as a new snapshot pushes that side over the
  configured total instead of continuing through later workbooks.
- Expose the control as `--max-portfolio-snapshot-cells` and the matching
  public GitHub Action input. Exact-limit, baseline/candidate overflow,
  stop-before-later-read, CLI default, and Action validation/propagation
  regressions cover the aggregate guard.

## 0.165.0 — 2026-07-28

- Add a fail-closed aggregate compressed-source-byte budget to each directory
  portfolio before FormulaFence opens any workbook snapshot. The 4 GiB default
  limits the total observed size of supported regular `.xlsx` / `.xlsm`
  sources independently on the baseline and candidate sides, rather than
  allowing the 512-workbook ceiling to multiply the 1 GiB per-workbook source
  limit into an impractical CI read set.
- Expose the limit as `--max-portfolio-source-bytes` and the matching public
  GitHub Action input. Exact-boundary, overflow-before-read, CLI default, and
  Action validation/propagation regressions keep the preflight deterministic.

## 0.164.0 — 2026-07-28

- Replace recursive globbing in portfolio discovery with a bounded direct
  `scandir` walk that propagates directory-enumeration errors instead of
  accepting a runtime globber's suppressed `OSError` as an empty subtree.
- Keep the existing raw-entry, symlink, case-collision, and file-identity
  boundaries intact; an unreadable descendant now fails the command before any
  workbook comparison can claim complete directory coverage.
- Add deterministic direct-inventory and CLI regressions for a controlled
  unreadable subdirectory, plus public-release comparison evidence.

## 0.163.0 — 2026-07-28

- Bind every later portfolio workbook read to the regular-file identity and
  state observed during inventory. The guarded read requests a no-follow final
  component where the host supports it, then checks the opened descriptor's
  device, inode, change timestamp, and size before creating its private
  inspection copy.
- A late in-place rewrite, new regular file, or symlink replacement now remains
  explicit redacted `FF078` incomplete evidence instead of silently changing
  which workbook a portfolio review inspects.
- Keep the public `discover_workbooks` path mapping stable while retaining the
  observed identity privately for comparisons, with deterministic regressions
  for all three post-inventory replacement forms.

## 0.162.0 — 2026-07-28

- Bound every recursive portfolio directory inventory before FormulaFence
  retains or sorts paths. The new default ceiling is 32,768 filesystem entries
  per supplied directory, covering ordinary non-workbook files, directories,
  Office lock files, and symlinks as well as supported workbooks.
- Add `--max-inventory-entries` to the CLI and the public GitHub Action, keeping
  this raw-entry budget independent of the existing supported-workbook and
  cross-workbook-impact limits.
- Add exact/default-limit, non-workbook, CLI, Action validation, and Action
  propagation regressions for fail-closed portfolio discovery.

## 0.161.0 — 2026-07-28

- Make ordinary `init` an atomic no-clobber publication: after writing its
  private same-directory file, it uses a new filesystem link to claim the final
  pathname. Any file, symlink, or hard link created after the initial absence
  check now remains untouched and produces the same `--force` guidance as an
  already-existing policy.
- Keep `init --force` on the existing atomic replacement path, so an explicit
  replacement still replaces a final symlink entry instead of writing through
  its target.
- Add direct ordinary-existing-file, post-check regular-file, symlink, and hard
  link regressions, including private temporary-file cleanup assertions.

## 0.160.0 — 2026-07-28

- Publish CLI reports and `init` starter policies through a private temporary
  file plus atomic replacement instead of writing through the requested final
  pathname. A final-component symlink or hard link substituted after safety
  checks is replaced as a directory entry rather than followed into an input.
- Keep the existing refusal to name an input as report output, preserve
  `init --force`, and clean up an unpublished temporary file on write or
  replacement failure.
- Add end-to-end report-path and starter-policy symlink-swap regressions, plus
  an existing-policy `init --force` compatibility regression.

## 0.159.0 — 2026-07-28

- Read a policy source through one file descriptor, requesting nonblocking mode
  where the host provides it, and verify the opened object is a regular file
  before consuming its bounded UTF-8 bytes. The 1 MiB source ceiling now also
  gates an already-known oversized file before any read.
- Close the post-check pathname-replacement case where a policy swapped for a
  FIFO or device could block a CI command before YAML validation on hosts with
  nonblocking descriptor opens. Descriptor reads use bounded blocks and retain
  the existing source-size, UTF-8, YAML, and schema diagnostics.
- Add a deterministic post-check FIFO replacement regression that asserts the
  nonblocking descriptor path, regular-file rejection, and absence of a
  pathname reopen.

## 0.158.0 — 2026-07-28

- Materialize one bounded private copy of each regular workbook source before
  archive preflight. The archive inventory, semantic-reader gate, raw OOXML
  scanners, ordinary workbook reader, and reported content hash now all use the
  same inspected bytes, while snapshots retain the caller's requested path.
- Close the source-path replacement window in which a preflighted workbook
  could be swapped before later scans. The existing 1 GiB package ceiling is
  enforced while copying; non-regular sources fail closed, and private copies
  are removed after either success or failure.
- Add deterministic source-replacement, cleanup-on-preflight-error,
  fail-before-reader, and parser-warning regressions for the stable-source
  boundary.

## 0.157.0 — 2026-07-26

- Bound policy-as-code input before YAML construction or workbook inspection:
  1 MiB source, 4,096 composed YAML nodes, 64 nesting levels, 4,096 characters
  per scalar, and 512 selectors per selector list.
- Reject duplicate mapping keys, anchors, aliases, merge keys, invalid UTF-8,
  and non-string schema keys rather than allowing YAML's last-wins or inherited
  mappings to weaken a reviewed control.
- Load a `check` policy before either workbook, preserving the CLI's stable
  policy-error status while avoiding unnecessary archive inspection. Add
  exact-limit, malformed/ambiguous YAML, selector, and fail-before-workbook
  regressions.

## 0.156.0 — 2026-07-26

- Preflight every nested Power Query `DataMashup` ZIP central directory before
  `ZipFile` can materialize an entry catalog. The shared 512-part scan budget
  now covers both logical package parts and metadata embedded-content catalogs,
  including catalogs whose content FormulaFence intentionally does not read.
- Bound raw nested member names to 1 KiB and conservatively retain coverage
  evidence for ZIP64, multi-disk, malformed, or filename-rewriting metadata
  (Unicode-path aliases, NULs, and platform separators). This keeps Python's
  post-preflight ZIP filename rewriting from bypassing the allocation boundary.
- Preserve a private opaque fingerprint and explicit `FF010`/`FF024` coverage
  evidence whenever an unsafe metadata catalog or logical package cannot be
  fully inspected. Add exact-capacity, fail-before-`ZipFile`, filename-rewrite,
  metadata-catalog, aggregate-across-mashups, and visible-diff regressions.

## 0.155.0 — 2026-07-26

- Bound each nested ZIP logical package in a Power Query `DataMashup` before
  any package member is inflated: 768 KiB package source, stored/deflated
  entries only, 512 parts, 16 MiB per part, 64 MiB aggregate expanded data
  across the Power Query scan, and a 1,000:1 maximum member ratio.
- Preserve private, hashed opaque evidence and an explicit coverage warning
  when that nested-package boundary is exceeded rather than reading the
  package's contents. A baseline-to-candidate diff consequently surfaces the
  existing `FF010`/`FF024` evidence.
- Add source-bound, compact ZIP-bomb fail-before-read, aggregate-across-
  mashups, visible-coverage, and normal Power Query regressions.

## 0.154.0 — 2026-07-26

- Bound every non-character-data XML lexical token to 128 KiB before a parser
  can retain it: comments, processing instructions, declarations, closing tags,
  and entity references. Opening tags retain their separate 128 KiB limit and
  stable start-tag diagnostic; CDATA remains under the independent 1 MiB
  decoded character-data limit.
- Explicitly forbid document-type declarations in the shared defused XML
  parser, including semantic-reader streams and bounded raw scans.
- Add exact-boundary, one-byte chunk, UTF-16/UTF-32, comment, processing-
  instruction, entity-reference, declaration, and fail-before-parser
  regressions.

## 0.153.0 — 2026-07-26

- Bound every decoded XML character-data node to 1 MiB before ElementTree's
  ordinary tree builder can join unbounded parser chunks. The shared bounded
  target covers semantic-reader streams, bounded raw-structure scans, rich-text
  and rich-data raw streams, array-formula metadata streams, and in-memory
  OOXML root parsing.
- Reject an oversized reader-visible text node through the stable
  semantic-reader safety preflight and CLI status 2. Other bounded raw readers
  retain their existing fail-closed coverage behavior rather than retaining an
  opaque text node.
- Add exact-boundary, one-byte incremental-feed, ordinary-text, CDATA,
  direct-stream, real-workbook, and fail-before-reader regressions. The limit
  resets at XML markup boundaries, so adjacent valid text and tail nodes retain
  their independent allowance.

## 0.152.0 — 2026-07-26

- Bound every XML opening tag at 128 KiB before ElementTree can construct an
  element's complete attribute map. The shared lexical gate runs before
  reader-visible streams and bounded raw XML scans; in-memory OOXML root reads
  use the same check before tree construction.
- Cover UTF-8/ASCII-compatible parts plus UTF-16 and UTF-32 fixed-width
  encodings, while correctly skipping quoted delimiters, comments, CDATA,
  processing instructions, and declarations. This keeps ordinary XML
  punctuation scanning in native byte search rather than imposing a
  character-by-character cost on large worksheets.
- Reject compact root and nested-style attribute-map payloads through the
  stable semantic-reader safety preflight and CLI status 2 before a parser
  callback runs. Add exact-boundary, one-byte-chunk, non-element-markup,
  fixed-width-encoding, root/nested attribute-map, and parser-entry
  regressions.

## 0.151.0 — 2026-07-26

- Bound compact reader-visible XML throughout `xl/styles.xml` before the
  ordinary stylesheet reader can construct its complete tree. A documented
  `extLst`, a foreign direct root subtree (or foreign root local name), and an
  ignored direct child inside a named style catalog each allow 32,768 XML
  elements; every materialized direct style record gets the same non-extension
  budget and those records share a 262,144-element budget.
- Preserve the existing named style catalog allowances while matching the
  reader's local-name and nested-sequence dispatch. Repeated ordinary-looking
  children such as `alignment`, unknown record descendants, and alternate-
  namespace extension lists now fail through the stable semantic-reader safety
  preflight and CLI status 2 rather than allocating a large complete tree.
- Add root, catalog, record, repeated-known-child, aggregate, nested,
  alternate-namespace, exact/default-limit, and fail-before-reader regressions.

## 0.150.0 — 2026-07-26

- Bound compact reader-visible XML beneath the bootstrap `xl/workbook.xml`
  part before FormulaFence's raw workbook scanners or the ordinary workbook
  reader can materialize its complete tree. A documented `extLst` payload and
  a foreign direct root subtree each allow 32,768 XML elements; the named
  workbook controls retain their existing format-aware catalog limits.
- Match the workbook parser's local-name dispatch for `workbook` and `extLst`,
  so transitional, Strict, and alternate-namespace extension containers cannot
  evade the allocation boundary. A successfully streamed overage becomes the
  stable semantic-reader safety-preflight error and CLI status 2 rather than a
  partial profile.
- Add direct/nested, nested-workbook-view, alternate-root/extension-namespace,
  exact/default-limit, normal-control, and fail-before-reader regressions.

## 0.149.0 — 2026-07-26

- Bound every relationship-selected SpreadsheetML Chartsheet and Dialogsheet
  XML tree before FormulaFence's raw protection/Custom View readers or the
  ordinary workbook reader can materialize it. These non-grid sheet grammars
  allow 32,768 XML elements per part and 65,536 across selected parts; chart
  definitions continue to use their separate DrawingML boundary.
- Cover both the documented `extLst` future-feature container and opaque
  root/descendant content without narrowing an ordinary worksheet's cell-grid
  allowance. A successfully streamed overage becomes the stable semantic-reader
  safety-preflight error and CLI status 2 instead of a partial profile.
- Add direct/nested chart-sheet and dialog-sheet, cross-type aggregate,
  exact/default-limit, normal-sheet, and fail-before-reader regressions.

## 0.148.0 — 2026-07-26

- Bound every SpreadsheetML `extLst` subtree in a reader-selected
  transitional or Strict worksheet before raw worksheet scanners or the
  ordinary workbook reader can materialize it. Each selected worksheet allows
  32,768 extension-list XML elements and the selected worksheet inventory
  allows 65,536 in aggregate.
- Keep the direct opaque-root boundary distinct: the published Worksheet
  grammar names `extLst`, but it is an arbitrary extension container. Ordinary
  `sheetData` and other named base controls retain their existing specialized
  budgets. A successfully streamed extension-list overage becomes the stable
  semantic-reader safety-preflight error and CLI status 2.
- Add direct/nested, nested-under-`sheetPr`, aggregate, exact/default-limit,
  normal-sheet, Strict-worksheet, and fail-before-reader regressions.

## 0.147.0 — 2026-07-26

- Bound complete direct worksheet-root subtrees that are absent from the
  published base SpreadsheetML Worksheet child grammar before raw worksheet
  scanners or the ordinary workbook reader can retain them. Each selected
  transitional or Strict worksheet allows 32,768 opaque-root XML elements and
  the selected worksheet inventory allows 65,536 in aggregate.
- Keep normal `sheetData` and every named base Worksheet child outside this
  narrow counter, preserving the existing populated-cell, dimension, merge,
  validation, conditional-formatting, Scenario Manager, print-break, and
  specialist metadata budgets. A successfully streamed opaque-root overage
  becomes the stable semantic-reader safety-preflight error and CLI status 2.
- Add direct/nested, aggregate, exact/default-limit, standard-root,
  Strict-worksheet, and fail-before-reader regressions.

## 0.146.0 — 2026-07-26

- Stream the raw shared-string rich-text scan one direct SpreadsheetML `si`
  item at a time instead of first constructing the complete `sst` tree. The
  scanner releases ignored foreign root children as they complete, preserving
  rich-run coverage without retaining unrelated shared-string XML.
- Add semantic-reader structural limits for shared-string shapes that either
  FormulaFence or the ordinary reader must retain: 32,768 XML elements per
  complete `si` item, 65,536 across complex/rich items, 32,768 opaque direct
  `sst`-child elements per selected table, and 65,536 in aggregate. Simple
  shared strings keep the established 500,000-entry allowance.
- Reject a successfully streamed overage with the stable semantic-reader
  safety-preflight error before raw or ordinary workbook readers start. Add
  direct/nested opaque-root, manifest/relationship-selection, per-item,
  aggregate, exact/default-limit, and fail-before-reader regressions, plus a
  streamed rich-text coverage test.

## 0.145.0 — 2026-07-26

- Add a shared semantic-reader structural gate for Table Definition XML before
  raw filter, Named Sheet View, external-data, XML Mapping, Table Style, or
  ordinary workbook readers can materialize a compact, repetitive table tree.
  Every canonical `xl/tables/*.xml` part and every safe direct internal
  worksheet `table` relationship target, including Strict and noncanonical
  targets, allows 32,768 XML elements per part and 65,536 in aggregate.
- Reject a successfully streamed table-definition structural overage through
  the stable semantic-reader safety preflight instead of emitting a partial
  report. Malformed, missing, and non-XML optional targets retain their
  established downstream coverage diagnostics; canonical orphan table parts
  stay bounded because the Table Style scanner inventories them.
- Add direct/nested opaque, Strict relationship, canonical-orphan aggregate,
  exact/default limit, fail-before-reader, and malformed-orphan regressions.

## 0.144.0 — 2026-07-26

- Bound the canonical raw `xl/metadata.xml` reader that identifies OOXML
  `XLDAPR` / `fDynamic` dynamic-array formula markers before it can materialize
  a compact, repetitive metadata tree. The dynamic-array classifier now allows
  16 MiB and 32,768 XML elements for that one named metadata part.
- Stream raw worksheet cell/formula bindings for array classification rather
  than constructing a second full worksheet XML tree after the ordinary
  workbook reader has already loaded it.
- Turn an oversized or structurally over-budget dynamic-array metadata part
  into visible `FF010` coverage evidence and fail closed by classifying affected
  array formulas as unclassified with no fixed-CSE or observed-spill aliases.
  A private streamed fallback fingerprint makes same-size opaque coverage
  changes diff-visible through `FF018` without exposing metadata, XML tags, or
  cell values; malformed XML retains its parser diagnostic.
- Add dynamic-array metadata byte/structural, exact/default, full-load,
  worksheet-streaming, same-size-fingerprint, redaction, and malformed-input
  regressions.

## 0.143.0 — 2026-07-26

- Bound raw XLM macro-sheet XML before FormulaFence, its temporary
  ordinary-workbook reader, or Custom View sanitization can materialize and
  privately canonicalize a repetitive macro-program tree. Each selected
  macro-sheet part allows 32,768 elements; the shared macro-sheet scan allows
  65,536, alongside 16 MiB per-part, 64 MiB aggregate, and 512-part limits.
- Convert a successfully streamed structural overage into explicit `FF010`
  plus `FF026`-visible opaque macro-sheet evidence, retaining only a private
  streamed content fingerprint for same-size hostile changes. Preserve the
  ordinary parser diagnostic when the structural preflight reaches malformed
  XML before an overage, and leave non-XLM raw workbook XML outside this
  scanner's scope.
- Give only selected XLM targets an empty worksheet replacement in the
  temporary ordinary reader after raw scanning, and exclude XLM targets from
  Custom View sanitization, association parsing, and generic sheet metadata.
  This includes an invalid ordinary relationship alias to the same raw target:
  the regular sheet declaration remains while a visible coverage warning
  prevents a secondary unbounded parse of macro XML.
- Add fail-before-materialization, byte/part/aggregate, exact/default,
  private-fingerprint, same-size, malformed-input, and XLM reader-isolation
  regressions for raw macro-sheet XML.

## 0.142.0 — 2026-07-26

- Bound raw query-table XML before FormulaFence can materialize and privately
  canonicalize repetitive refresh-field, sort, or extension trees. Each
  query-table relationship target allows 32,768 elements; the shared
  query-table scan allows 65,536, alongside 16 MiB per-part, 64 MiB aggregate,
  and 512-part limits.
- Share the cached boundary and a sheet-neutral private snapshot template across
  direct worksheet and table-mediated query-table references, so one part is
  neither reparsed nor recursively canonicalized once per worksheet binding. A
  successfully streamed structural overage becomes explicit `FF010` plus
  `FF023`-visible opaque query-table evidence, with a private streamed content
  fingerprint for same-size hostile changes.
- Add fail-before-materialization, nested opaque, exact/default, aggregate,
  reused-part cache, visibility, and same-size-fingerprint regressions for
  query-table XML while retaining direct and table-mediated refresh coverage.

## 0.141.0 — 2026-07-26

- Bound raw external-link package XML before FormulaFence can materialize and
  privately canonicalize a highly repetitive link definition or its direct
  relationship part. Each selected `xl/externalLinks/externalLink*.xml` or
  direct `.rels` part allows 32,768 elements; the shared external-link scan
  allows 65,536, alongside 16 MiB per-part, 64 MiB aggregate, and 512-part
  limits.
- Share the same cached boundary between external-link inventory and the
  package-indexed external-workbook resolver, so a part cannot be reparsed
  unboundedly on its second use. A successfully streamed structural overage
  becomes explicit `FF010` plus `FF025`-visible opaque package evidence, with
  a private streamed content fingerprint for same-size hostile changes.
- Add fail-before-materialization, nested opaque, exact/default, aggregate,
  reused-part cache, visibility, and same-size-fingerprint regressions for
  external-link XML while retaining package-indexed portfolio resolution.

## 0.140.0 — 2026-07-26

- Bound raw external-data `Connections` XML before FormulaFence can materialize
  and privately inspect a highly compressible connection tree. Each
  `xl/connections*.xml` part allows 32,768 elements and the complete Connections
  scan allows 65,536, alongside new 16 MiB per-part, 64 MiB aggregate, and
  512-part byte/count limits.
- Convert a successfully streamed structural overage into explicit `FF010` plus
  `FF023`-visible opaque connection coverage evidence. The fallback retains a
  private streamed content fingerprint, so same-size hostile Connections XML
  changes remain diff-visible without exposing connection or XML material.
- Add byte/part, fail-before-materialization, nested opaque, exact/default,
  aggregate, visibility, and same-size-fingerprint regressions for Connections
  XML.

## 0.139.0 — 2026-07-26

- Bound raw legacy shared-workbook `revisionHeaders` and `revisionLog` XML
  before FormulaFence can materialize and recursively canonicalize its private
  history records. Each revision XML part allows 32,768 elements and the
  complete revision scan allows 65,536, alongside the existing 16 MiB per-part,
  64 MiB aggregate, and 512-part byte/count limits.
- Convert a successfully streamed structural overage into explicit `FF010` plus
  `FF062` revision-history coverage evidence. The fallback retains a private
  streamed content fingerprint, so same-size hostile history changes remain
  diff-visible without exposing historic cells, identities, or XML.
- Add header/log fail-before-materialization, nested opaque, exact/default,
  aggregate, visibility, and same-size-fingerprint regressions.

## 0.138.0 — 2026-07-26

- Bound the decoded metadata and formula-firewall permission XML embedded in
  Power Query `DataMashup` custom XML. Each inner document allows 32,768 XML
  elements and the complete Power Query scan allows 65,536 before FormulaFence
  can materialize a private tree.
- Convert a successfully parsed inner-XML structural overage into explicit
  `FF010` plus `FF024` Power Query coverage evidence without exposing query
  material. Malformed input retains its established parser diagnostic.
- Add decoded metadata/permission fail-before-materialization, nested opaque,
  configured/default/exact capacity, aggregate, and report-visibility
  regressions.

## 0.137.0 — 2026-07-26

- Bound raw SpreadsheetML XML Maps, OPC package-signature XML, Python-in-Excel,
  and Rich Data package XML before FormulaFence can materialize and recursively
  canonicalize private trees. Each inventory allows 32,768 XML elements per
  part and 65,536 across its complete bounded XML scan, alongside the existing
  16 MiB per-part, 64 MiB aggregate, and 512-part byte/count limits.
- Include the digital-signature content-types reader in the same structural
  budget, while keeping certificate and VBA-signature binary payloads strictly
  byte-bounded rather than treating them as XML.
- Convert a successfully parsed structural overage into explicit `FF010` plus
  `FF049` XML Map, `FF050` digital-signature, `FF051` Rich Data, or `FF065`
  Python-in-Excel coverage evidence. Malformed or unreadable input retains its
  established diagnostic.
- Stream Rich Data worksheet binding attributes after the shared semantic-reader
  preflight instead of retaining a second complete worksheet XML tree.
- Add fail-before-materialization, nested opaque, exact/default, aggregate,
  binary-payload, content-types, and Rich Data streaming regressions across all
  four raw inventories.

## 0.136.0 — 2026-07-26

- Bound raw traditional Excel Note comments/VML layout XML before FormulaFence
  recursively canonicalizes private note metadata. Each part allows 32,768 XML
  elements and the complete legacy-Note scan allows 65,536, alongside the
  existing 16 MiB per-part, 64 MiB aggregate, and 512-part byte/count limits.
- Bound the shared raw worksheet embedded-control XML gateway used by worksheet
  control markup, ActiveX persistence, form-control properties, and legacy VML
  drawings with the same 32,768-per-part / 65,536-per-scan structural limits.
  A VML note drawing is independently protected at both the Note and control
  inventory paths, so either scanner declines it before tree allocation.
- Convert successfully parsed structural overages into explicit `FF010`/`FF046`
  Note or `FF010`/`FF029` embedded-control coverage evidence while malformed or
  unreadable input retains its established diagnostic.
- Add fail-before-materialization regressions for Comment, ActiveX,
  form-control-property, and shared VML paths, including nested opaque XML,
  aggregate budgets, exact capacity, default overage, and report coverage.

## 0.135.0 — 2026-07-26

- Bound relationship-selected Worksheet DrawingML XML in the shared
  semantic-reader preflight before FormulaFence's shape, native-image,
  in-content Office Web Add-in, chart, or ordinary workbook readers can
  materialize the same tree. Each internally targeted transitional or Strict
  worksheet drawing part allows 32,768 XML elements, and all unique selected
  targets allow 65,536 in aggregate.
- Stream only direct internal worksheet `drawing` relationship targets and
  fail with the stable reader-safety error for a successfully parsed structural
  overage. Missing, malformed, or non-XML optional targets keep their existing
  scanner coverage behavior, and unrelated orphan DrawingML parts are not
  broadened into the reader boundary.
- Add direct/nested opaque, Strict relationship, aggregate-target, exact
  capacity, default-capacity, orphan-scope, and fail-before-any-reader
  regressions for the shared DrawingML path.

## 0.134.0 — 2026-07-26

- Bound modern Threaded Comments and Persons XML before FormulaFence can
  materialize and recursively canonicalize private comment, reply, extension,
  or person trees. Each XML part allows 32,768 elements and the complete
  threaded-comment scan allows 65,536, alongside the existing 16 MiB per-part,
  64 MiB aggregate, and 512-part byte/count limits.
- Stream raw comment/person XML before the private parser runs, so a
  successfully parsed structural overage becomes explicit `FF010`/`FF045`
  coverage evidence while malformed or unreadable input keeps its established
  diagnostic.
- Remove threaded-comment and person relationships only from FormulaFence's
  temporary ordinary-reader copy after raw inspection. The original package and
  raw evidence remain intact, while a current or future workbook reader cannot
  re-materialize an XML tree the bounded scanner rejected.
- Add comment/person fail-before-materialization, reader-isolation,
  configured/default and exact capacity, aggregate, and nested-opaque boundary
  regressions.

## 0.133.0 — 2026-07-26

- Bound DrawingML Theme and Theme-relationship XML before FormulaFence can
  materialize and recursively canonicalize a private theme tree. Each XML part
  allows 32,768 elements and the complete Theme scan allows 65,536, alongside
  the existing 16 MiB per-part, 64 MiB aggregate, and 512-part byte/count
  limits.
- Stream raw Theme XML before the private parser runs, so a successfully parsed
  structural overage becomes explicit `FF053` coverage evidence while malformed
  or unreadable input keeps its established diagnostic. Direct Theme-image
  payloads remain byte-bounded rather than being interpreted as XML.
- Add fail-before-tree-materialization, binary-image, configured/default and
  exact capacity, aggregate, and nested-opaque Theme boundary regressions.

## 0.132.0 — 2026-07-26

- Bound generic Custom XML, Custom XML-property, Custom Data-property, custom
  document-property, and custom-state relationship XML before FormulaFence can
  materialize and recursively canonicalize private trees. Each XML part allows
  32,768 elements and the complete custom-state scan allows 65,536, alongside
  the existing 16 MiB per-part, 64 MiB aggregate, and 512-part byte/count
  limits.
- Stream raw custom-state XML before the private parser runs, so a successfully
  parsed structural overage becomes explicit `FF052` coverage evidence while
  malformed or unreadable input keeps its established diagnostic. Opaque binary
  Custom Data payloads remain byte-bounded and are never interpreted as XML.
- Hand safely classified `DataMashup` members to the Power Query scanner rather
  than reparsing every arbitrary Custom XML item. This preserves the separate
  Power Query boundary while ensuring a rejected generic or over-budget custom
  XML tree cannot be materialized by a second scanner.
- Add fail-before-all-materializing-readers, binary-payload, configured/default
  and exact capacity, aggregate, nested-opaque, and Power Query handoff
  regressions.

## 0.131.0 — 2026-07-26

- Bound Slicer and Timeline cache-definition XML before FormulaFence's private
  scanner can recursively canonicalize a compact but oversized filter tree.
  Each cache XML part allows 16,384 elements and the complete Slicer/Timeline
  package scan allows 32,768, alongside its existing 16 MiB per-part, 64 MiB
  aggregate, and 512-part byte/count limits.
- Stream the raw ZIP member before materializing its private XML tree, so a
  successfully parsed structural overage becomes explicit `FF032` coverage
  evidence while malformed or unreadable input retains its established parser
  diagnostic.
- Preserve observable coverage for configured/default and exact capacities,
  package-wide aggregation, nested opaque descendants, and the fail-before-
  materialization behavior that prevents an over-budget cache tree from being
  built by the private scanner.

## 0.130.0 — 2026-07-26

- Bound legacy PivotTable view and cache-definition XML before FormulaFence's
  private scanner can recursively canonicalize a compact but oversized tree.
  Each scanned part allows 32,768 elements and the complete PivotTable package
  scan allows 65,536, alongside its existing 16 MiB per-part, 64 MiB aggregate,
  and 512-part byte/count limits.
- Keep PivotTable refresh controls and static cell inspection available while
  preventing the ordinary workbook reader from reparsing raw pivot packages:
  FormulaFence now removes only PivotTable cache and view bindings in its
  temporary reader copy after the bounded raw inspection has retained evidence.
- Preserve observable coverage for a valid structural overage, malformed XML,
  exact/default capacity, aggregate budgets, and nested opaque descendants;
  add a reader-isolation test that rejects any underlying PivotTable cache,
  record, or view parse.

## 0.129.0 — 2026-07-26

- Bound legacy chart, ChartEx, chart-host DrawingML, and chart-overlay XML
  before FormulaFence's private scanner can recursively canonicalize a compact
  but oversized tree. Each scanned part allows 32,768 elements and the complete
  chart package scan allows 65,536, alongside its existing 16 MiB per-part,
  64 MiB aggregate, and 512-part byte/count limits.
- Preserve practical chart-cache coverage while preventing opaque unknown XML
  from bypassing the byte budget: a well-formed structural overage becomes
  explicit chart coverage evidence before its tree is materialized, while
  malformed input keeps the established parser diagnostic.
- Add fail-before-tree-materialization, configured/default/exact/aggregate, and
  opaque-nested boundary coverage, including the chart-control finding emitted
  for a new structural coverage gap.

## 0.128.0 — 2026-07-26

- Bound every OOXML package relationship (`.rels`) part in the shared semantic
  preflight before FormulaFence's raw metadata scanners or `openpyxl` can
  materialize an oversized relationship tree or catalog. Each part allows
  4,096 XML elements and all relationship parts together allow 16,384.
- Count complete relationship XML shape, including roots and opaque nested
  descendants, rather than assuming a relationship part contains only direct
  package `Relationship` records. Oversized well-formed input now gets the
  stable safety-preflight error before any specialist scanner runs.
- Preserve established coverage diagnostics for a malformed unused relationship
  part, and add fail-before-reader tests for per-part, aggregate, exact,
  opaque-nested, default-limit, and malformed compatibility cases.

## 0.127.0 — 2026-07-26

- Bound Office Web Add-in task-pane and definition XML before FormulaFence's
  private scanner can materialize their full element tree or recursively
  canonicalize opaque configuration fragments. Each part now permits 4,096
  elements and the task-pane-plus-definition package scan permits 16,384 in
  aggregate, alongside its existing 16 MiB part, 32 MiB byte, and 64-part
  limits.
- Stream the structural preflight directly from each ZIP member and fail closed
  to unrecognized Office Web Add-in coverage when its element or nesting budget
  is exceeded. Well-formed inputs at the exact boundary retain their existing
  control coverage, while malformed input preserves the established full-parser
  diagnostic.
- Add fail-before-tree-materialization coverage for both task-pane and
  definition paths, configured/default and aggregate limits, exact capacity,
  and compact direct/nested opaque XML subtrees.

## 0.126.0 — 2026-07-26

- Bound each RibbonX `customUI` XML part before FormulaFence's private
  customization scanner can materialize its full XML tree or recursively
  canonicalize opaque control fragments. The new 4,096-element-per-part
  structural budget complements the existing 16 MiB part, 32 MiB aggregate,
  and eight-part byte/count limits.
- Stream the structural preflight directly from the ZIP member and fail closed
  to unrecognized RibbonX coverage when its element or nesting budget is
  exceeded. Valid but complex customizations within the exact boundary retain
  their existing callback/control coverage and malformed inputs retain their
  established full-parser diagnostics.
- Add fail-before-tree-materialization coverage for configured/default and
  exact limits, including compact direct and nested opaque RibbonX subtrees.

## 0.125.0 — 2026-07-26

- Bound every descendant below a direct `customSheetView` in a supported legacy
  `customSheetViews` container before FormulaFence's raw Custom View scanner can
  recursively canonicalize opaque XML or materialize specialized controls. The
  new 4,096-element aggregate budget is separate from the existing 4,096-view
  declaration bound.
- Apply the bound to standard, Strict, and alternate-namespace view paths the
  raw scanner enters, while leaving foreign `customSheetViews` containers outside
  that scanner path. The published 2,052 Custom View row-plus-column page-break
  allowance remains accepted inside the new subtree capacity.
- Add fail-before-reader coverage for configured/default and exact limits,
  cross-sheet aggregation, nested opaque trees, Strict and opaque views, and
  ignored foreign containers.

## 0.124.0 — 2026-07-26

- Extend the existing page-break preflight budget into legacy Excel Custom
  Views before FormulaFence's raw Custom View scanner can materialize break
  signatures or opaque subtree evidence. The 4,096-container and 2,052-direct-
  child budgets now aggregate ordinary worksheet and supported Custom View
  paths.
- Count direct `rowBreaks`/`colBreaks` beneath a supported
  `customSheetViews` container by local name, including Strict SpreadsheetML,
  alternate-namespace `customSheetView` or break-container paths, and every
  direct child that the scanner must preserve.
- Add fail-before-reader coverage for configured/default limits, aggregation
  with ordinary worksheet breaks, exact configured limits, Strict Custom Views,
  opaque namespace paths, and fragmented Custom View break containers.

## 0.123.0 — 2026-07-26

- Bound direct worksheet page-break catalogs before FormulaFence's raw
  print-layout scanner or `openpyxl` can materialize them. The semantic-reader
  preflight permits 4,096 direct `rowBreaks`/`colBreaks` containers and 2,052
  direct child records in aggregate across selected ordinary worksheet parts.
- Preserve one worksheet's published allowance of 1,026 horizontal plus 1,026
  vertical page breaks. Every direct child of a supported break container now
  consumes the record budget, including unexpected or alternate-namespace
  children that FormulaFence must retain as raw coverage evidence.
- Add fail-before-reader coverage for configured/default limits, cross-sheet
  and cross-axis aggregation, exact published capacity, Strict SpreadsheetML,
  arbitrary direct children, and ignored foreign-namespace containers.

## 0.122.0 — 2026-07-26

- Bound reader-visible worksheet column declarations and direct containers
  before downstream raw OOXML scanners or the complete workbook reader run. The
  semantic-reader preflight now permits 16,384 `col` declarations and 4,096
  direct `cols` containers in aggregate across selected ordinary worksheet
  parts.
- Count every reader-visible SpreadsheetML `col`, including declarations that
  repeat one final column key or have unknown attributes, because `openpyxl`
  dispatches each one before FormulaFence can determine whether it is material.
- Add fail-before-reader coverage for configured/default limits, cross-sheet
  aggregation, exact limits, unknown attributes, and ignored foreign-namespace
  declaration and container cases.

## 0.121.0 — 2026-07-26

- Bound the reader-materialized formatted-row catalog before raw OOXML scanners
  or the complete workbook reader run. The semantic-reader preflight now allows
  at most 16,384 row-dimension declarations in aggregate across its selected
  ordinary worksheet parts.
- Follow `openpyxl`'s allocation trigger: a SpreadsheetML `row` counts only
  when it has an unqualified attribute other than `r` or `spans`, so coordinate-
  only rows and namespace-qualified extension attributes remain compatible.
- Add fail-before-reader regression coverage for configured/default limits,
  aggregate cross-worksheet counts, exact limits, unknown attributes, and
  ignored coordinate, namespaced-attribute, and foreign-namespace cases.

## 0.120.0 — 2026-07-26

- Bound every `openpyxl` reader-materialized stylesheet catalog before raw
  OOXML scanners or the complete workbook reader can allocate it. Number-format,
  font, fill, fill-child, gradient-stop, border, base-XF, named-style,
  differential-style, palette, table-style, table-style-element, and extension
  records each allow at most 4,096 entries; repeated known stylesheet
  containers are also capped at 4,096 in aggregate.
- Preserve the published 65,490 effective `cellXfs` style ceiling, while making
  its counter follow the reader's local-name behavior. Alternate-namespace
  `cellXfs` containers can no longer bypass the bound, and all direct children
  of `NestedSequence` style catalogs count even when they have unexpected names
  or namespaces.
- Add fail-before-reader coverage for every stylesheet catalog, repeated
  containers, unexpected nested records, alternate namespaces, an over-limit
  default font fixture, and exact configured/default font limits.

## 0.119.0 — 2026-07-26

- Bound the reader-materialized worksheet control catalogs: direct data-
  validation declarations, conditional-formatting declarations and rules, and
  Scenario Manager containers, scenarios, and input-cell records each allow at
  most 4,096 entries across reader-selected ordinary worksheet parts.
- Bound every reader-visible `sqref` list before `openpyxl` constructs a
  `CellRange` object for each whitespace-separated target: 128 KiB and 4,096
  targets per reference, with 8,192 targets in aggregate for each control
  catalog. Data-validation and conditional-formatting formula fields now follow
  the existing 8,192-character stored-formula limit as well.
- Count the same direct local-name children that the reader materializes,
  including alternate-namespace declarations, and add fail-before-reader
  coverage for individual and aggregate target budgets, nested Scenario input
  cells, formula text, default limits, and exact-limit inputs.

## 0.118.0 — 2026-07-26

- Bound merged-cell geometry before `openpyxl` expands a compact SpreadsheetML
  range into one in-memory `MergedCell` per coordinate. The semantic-reader
  preflight permits at most 4,096 direct `mergeCell` declarations across its
  selected ordinary worksheet parts, 100,000 coordinates in any one range and
  in aggregate, and 256 characters in a range reference.
- Measure the same direct child shape that the worksheet reader accepts,
  including alternate-namespace children and sheet-qualified references, so a
  repeated declaration or enormous declared geometry cannot bypass the bound.
- Add fail-before-reader regression coverage for declaration counts, alternate
  namespaces, individual and aggregate area, reference length, full-grid
  geometry, and an exact-limit sheet-qualified range.

## 0.117.0 — 2026-07-26

- Bound the remaining repeated workbook-package catalogs: direct workbook
  views, function groups, smart-tag types, and web-publish objects are limited
  before `openpyxl` can materialize them, while custom workbook views are
  limited before FormulaFence's raw Custom View scanner records them. Each
  catalog allows at most 4,096 entries.
- Bound FormulaFence's raw-scanned legacy Custom View surface as well: direct
  custom sheet-view declarations across workbook-selected worksheet, chart-
  sheet, and dialog-sheet parts are limited to 4,096 before per-view records
  are built.
- Match the reader/scanner's local-name behavior in alternate namespaces and
  add fail-before-reader tests for ordinary, alternate-namespace, exact-limit,
  and over-limit fixtures across all six catalog surfaces.

## 0.116.0 — 2026-07-26

- Bound the reader-materialized `<externalReferences>` and `<pivotCaches>`
  workbook catalogs to 4,096 direct entries each before FormulaFence starts
  raw OOXML scanners or `openpyxl`. The bound aligns with the existing bounded
  workbook relationship catalog while preventing repeated declarations from
  revisiting one safe external-link or pivot-cache target unboundedly.
- Count every direct catalog child that the reader's nested-sequence parser can
  materialize, including alternate-namespace entries, rather than only
  conventional namespace-qualified declaration tags.
- Add fail-before-reader regression coverage for ordinary, alternate-namespace,
  and real default-limit external-reference and pivot-cache fixtures.

## 0.115.0 — 2026-07-26

- Bound the reader-materialized workbook `<definedNames>` catalog to 100,000
  direct declarations before FormulaFence starts raw OOXML scanners or
  `openpyxl`, preventing a compact workbook part from creating an impractical
  name index and repeated formula-control analysis workload in CI.
- Make the workbook catalog counters follow the reader's local-name behavior:
  alternate-namespace `<definedName>` declarations count toward the new limit,
  and every direct child the reader materializes from `<sheets>` counts toward
  the existing 512-sheet ceiling.
- Add fail-before-reader regression coverage for ordinary and alternate-
  namespace sheet/name catalogs plus a real default-limit defined-name fixture.

## 0.114.0 — 2026-07-26

- Bound the OOXML reader's bootstrap catalogs before any raw scanner or
  `openpyxl` model can allocate them: 4,096 manifest `Default`/`Override`
  declarations, 4,096 workbook relationship records, and 512 workbook sheet
  declarations.
- Count every declaration in the workbook `<sheets>` catalog rather than only
  unique target parts, so repeated sheets pointing at one relationship cannot
  turn a compact package into a large in-memory workbook or repeated scanner
  workload.
- Add fail-before-reader regression coverage for all three catalog bounds and a
  real 512-declaration repeated-sheet boundary fixture.

## 0.113.0 — 2026-07-26

- Extend the fail-closed semantic-reader preflight to stream the OOXML manifest,
  workbook metadata, styles, shared-string table, and workbook-selected sheet
  parts under a 4,000,000-element-per-part and 256-level nesting bound before
  downstream scanners or `openpyxl` allocate their complete models.
- Limit reader-visible shared-string entries to 500,000, `cellXfs` styles to
  Excel's 65,490 style limit, text values to Excel's 32,767-character cell
  limit, and stored formula/defined-name text to Excel's 8,192-character
  formula limit. Shared strings follow the first manifest-selected target that
  `openpyxl` reads, then a sole workbook relationship or canonical fallback.
- Add regression coverage for structural XML, styles, cell/formula scalar
  limits, and noncanonical relationship- and manifest-selected shared-string
  parts, all proved to reject before any downstream reader starts.

## 0.112.0 — 2026-07-26

- Require `defusedxml` for FormulaFence's OOXML parsing, which also enables
  `openpyxl`'s defused XML reader in the supported installation.
- Add a second fail-closed semantic-reader preflight after the ZIP-header
  inventory and before downstream raw OOXML scanners or `openpyxl` calls. It caps an
  XML/relationship part at 64 MiB, aggregate XML material at 256 MiB, and
  follows the bounded workbook sheet relationships to stream worksheet XML and
  reject more than 500,000 populated SpreadsheetML cell records before a
  complete in-memory workbook model can be allocated.
- Stream `vbaProject.bin` into its private SHA-256 digest rather than loading
  the whole macro payload at once.
- Convert malformed cell metadata that causes `openpyxl` to raise `TypeError`
  or `IndexError` into FormulaFence's normal unreadable-workbook input error
  (exit status 2), rather than leaking a traceback into CI logs.

## 0.111.0 — 2026-07-26

- Add a fail-closed OOXML ZIP preflight before FormulaFence opens any raw
  package part or `openpyxl`. It bounds source and central-directory metadata,
  entry count, member and aggregate expanded sizes, and compression ratio;
  accepts only canonical single-disk stored/deflated members; and rejects
  duplicate/case-colliding or unsafe paths, Unicode-path aliases,
  encrypted/special members, malformed ZIP64 or local-header relationships,
  and overlapping payloads.

## 0.110.0 — 2026-07-26

- Add `--format html` for `diff`, `check`, and `portfolio`. It produces one
  self-contained review artifact with inline styles, local text/severity
  filters, expandable complete evidence, and no remote assets or network
  requests.
- Escape every workbook-derived value before it enters HTML, so evidence remains
  text rather than executable markup. The report preserves the existing
  deterministic JSON data model and visible finding/change semantics.
- Apply every existing output-only sharing boundary to HTML, including external
  workbook links, formula actions, Python-in-Excel, custom/runtime functions,
  worksheet/formula-defined registrations, XLM evaluation/action/GET.CELL/
  environment-information material, and native `CELL`/`INFO`/`SHEET`/`SHEETS`
  material. Active boundaries remain explicit in the artifact.
- Extend the composite GitHub Action's `format` input to accept `html`. HTML is
  uploaded as the configured artifact and referenced from the job summary rather
  than embedded into it.
- Add direct escaping, all-sharing-boundary CLI, `check`, portfolio, and
  composite-Action regression coverage.

## 0.109.0 — 2026-07-26

- Add opt-in `--redact-formula-environment-information` rendering for `diff`,
  `check`, and `portfolio` JSON, Markdown, and SARIF artifacts. Default local
  review output remains unchanged; the output-only boundary replaces direct
  stored `FF072` `CELL`, `INFO`, `SHEET`, and `SHEETS` material with
  `[formula environment-information material redacted]`.
- Extend the boundary to exact changed static input cells recorded by the
  private full dependency impact set, rather than the bounded impact sample
  shown in reports. Keep the `SHEET`/omitted-reference `SHEETS` raw-tab-catalog
  comparison semantics unchanged.
- Retain the private native definition-chain signature for FF072 and, when it
  changes, conservatively redact changed defined-name before/after evidence so
  a dotted workbook-defined wrapper cannot disclose a private information code
  or reference that reaches a native environment-information call deeper in the
  chain.
- Keep comparison facts, findings, policy evaluation, and exit status unchanged.
  The composite GitHub Action exposes the switch as
  `redact-formula-environment-information: 'true'`; the mode does not calculate
  a formula, determine an information type, resolve a dynamic reference,
  infer a selected cell, simulate workbook/client/workspace state, or
  reconstruct a runtime value, and it is not a general secret scrubber.
- Add direct-call, exact unsampled-static-input, resolved named-chain,
  default-evidence, policy, portfolio, JSON/Markdown/SARIF, and composite-Action
  regression coverage.

## 0.108.0 — 2026-07-26

- Add opt-in `--redact-formula-defined-xlm-environment-information-calls`
  rendering for `diff`, `check`, and `portfolio` JSON, Markdown, and SARIF
  artifacts. Default local review output remains unchanged; the output-only
  boundary replaces direct stored `FF071` GET.WORKBOOK, GET.WORKSPACE, and
  GET.DOCUMENT material with `[formula-defined XLM environment-information
  material redacted]`.
- Extend the boundary to before/after evidence for changed invoking formulas
  and exact changed static input cells recorded by the private full dependency
  impact set, rather than the bounded impact sample shown in reports.
- Retain the private formula-defined-name chain signature for FF071 and, when
  it changes, conservatively redact changed defined-name before/after evidence
  so a dotted workbook-defined wrapper cannot disclose a private information
  code or reference that reaches a selected environment-information call deeper
  in the chain.
- Keep comparison facts, findings, policy evaluation, and exit status unchanged.
  The composite GitHub Action exposes the switch as
  `redact-formula-defined-xlm-environment-information-calls: 'true'`; the mode
  does not calculate a formula, determine an information type, resolve a
  dynamic reference, simulate workbook/workspace/document state, or reconstruct
  a runtime value, and it is not a general secret scrubber.
- Add direct-call, exact unsampled-static-input, resolved named-chain,
  default-evidence, policy, portfolio, JSON/Markdown/SARIF, and composite-Action
  regression coverage.

## 0.107.0 — 2026-07-26

- Add opt-in `--redact-formula-defined-xlm-get-cell-calls` rendering for
  `diff`, `check`, and `portfolio` JSON, Markdown, and SARIF artifacts.
  Default local review output remains unchanged; the output-only boundary
  replaces direct stored `FF070` `GET.CELL` material with
  `[formula-defined XLM GET.CELL material redacted]`.
- Extend the boundary to before/after evidence for changed invoking formulas
  and exact changed static input cells recorded by the private full dependency
  impact set, rather than the bounded impact sample shown in reports.
- Retain the private formula-defined-name chain signature for FF070 and, when
  it changes, conservatively redact changed defined-name before/after evidence
  so a dotted workbook-defined wrapper cannot disclose a private information
  code or reference that reaches `GET.CELL` deeper in the chain.
- Keep comparison facts, findings, policy evaluation, and exit status unchanged.
  The composite GitHub Action exposes the switch as
  `redact-formula-defined-xlm-get-cell-calls: 'true'`; the mode does not
  calculate a formula, determine an information type, resolve a dynamic
  reference, simulate display/formatting or other Excel state, or reconstruct
  a runtime value, and it is not a general secret scrubber.
- Add direct-call, exact unsampled-static-input, resolved named-chain,
  default-evidence, policy, portfolio, JSON/Markdown/SARIF, and composite-Action
  regression coverage.

## 0.106.0 — 2026-07-26

- Add opt-in `--redact-formula-defined-xlm-actions` rendering for `diff`,
  `check`, and `portfolio` JSON, Markdown, and SARIF artifacts. Default local
  review output remains unchanged; the output-only boundary replaces direct
  stored selected `FF073` action material with `[formula-defined XLM action
  material redacted]`.
- Extend the boundary to before/after evidence for changed invoking formulas
  and exact changed static input cells recorded by the private full dependency
  impact set, rather than the bounded impact sample shown in reports.
- Retain the private formula-defined-name chain signature for FF073 and, when
  it changes, conservatively redact changed defined-name before/after evidence
  so a dotted workbook-defined wrapper cannot disclose an action target or
  handler that reaches a selected action deeper in the chain.
- Keep comparison facts, findings, policy evaluation, and exit status unchanged.
  The composite GitHub Action exposes the switch as
  `redact-formula-defined-xlm-actions: 'true'`; the mode does not calculate a
  formula, resolve an action target or event handler, load a DLL, send DDE,
  execute a macro or program, or reconstruct a dynamically assembled action,
  and it is not a general secret scrubber.
- Add direct-call, exact unsampled-static-input, resolved named-chain,
  default-evidence, policy, portfolio, JSON/Markdown/SARIF, and composite-Action
  regression coverage.

## 0.105.0 — 2026-07-26

- Add opt-in `--redact-formula-defined-xlm-evaluations` rendering for `diff`,
  `check`, and `portfolio` JSON, Markdown, and SARIF artifacts. Default local
  review output remains unchanged; the output-only boundary replaces direct
  stored `FF069` `EVALUATE` material with `[formula-defined XLM evaluation
  material redacted]`.
- Extend the boundary to before/after evidence for changed invoking formulas
  and exact changed static input cells recorded by the private full dependency
  impact set, rather than the bounded impact sample shown in reports.
- Retain the private formula-defined-name chain signature for FF069 and, when
  it changes, conservatively redact changed defined-name before/after evidence
  so a dotted workbook-defined wrapper cannot disclose expression text that
  reaches `EVALUATE` deeper in the chain.
- Keep comparison facts, findings, policy evaluation, and exit status unchanged.
  The composite GitHub Action exposes the switch as
  `redact-formula-defined-xlm-evaluations: 'true'`; the mode does not calculate
  a formula or text argument, parse a runtime-generated expression, execute a
  macro, or reconstruct dynamically assembled text, and it is not a general
  secret scrubber.
- Add direct-call, exact unsampled-static-input, resolved named-chain,
  runtime-text coverage-limit, default-evidence, policy, portfolio,
  JSON/Markdown/SARIF, and composite-Action regression coverage.

## 0.104.0 — 2026-07-26

- Add opt-in `--redact-formula-defined-xlm-registrations` rendering for
  `diff`, `check`, and `portfolio` JSON, Markdown, and SARIF artifacts. Default
  local-review output remains unchanged; the output-only boundary replaces
  direct stored `FF068` `REGISTER` material with `[formula-defined XLM
  registration material redacted]`.
- Extend the boundary to before/after evidence for changed invoking formulas
  and exact changed static input cells recorded by the private full dependency
  impact set, rather than the bounded impact sample shown in reports.
- Retain the private formula-defined-name chain signature for FF068 and, when
  it changes, conservatively redact changed defined-name before/after evidence
  so a dotted workbook-defined wrapper cannot disclose a module, procedure,
  type string, or another registration argument that reaches `REGISTER` deeper
  in the chain.
- Keep comparison facts, findings, policy evaluation, and exit status unchanged.
  The composite GitHub Action exposes the switch as
  `redact-formula-defined-xlm-registrations: 'true'`; the mode does not
  calculate a formula, execute a macro, resolve a module path, load a DLL/XLL,
  inspect host trust settings, contact a provider, or reconstruct a dynamic
  argument, and it is not a general secret scrubber.
- Add direct-call, exact unsampled-static-input, resolved named-chain,
  default-evidence, policy, portfolio, JSON/Markdown/SARIF, and composite-Action
  regression coverage.

## 0.103.0 — 2026-07-26

- Add opt-in `--redact-worksheet-code-resource-registrations` rendering for
  `diff`, `check`, and `portfolio` JSON, Markdown, and SARIF artifacts. Default
  local-review output remains unchanged; the output-only boundary replaces
  direct stored `FF067` `REGISTER.ID` material with `[worksheet code-resource
  registration material redacted]`.
- Extend the boundary to before/after evidence for changed `REGISTER.ID`
  formulas and exact changed static input cells recorded by the private full
  dependency impact set, rather than the bounded impact sample shown in reports.
- Retain the private formula-defined-name chain signature for FF067 and, when
  it changes, conservatively redact changed defined-name before/after evidence
  so a dotted workbook-defined wrapper cannot disclose a module, procedure, or
  other registration argument that reaches `REGISTER.ID` deeper in the chain.
- Keep comparison facts, findings, policy evaluation, and exit status unchanged.
  The composite GitHub Action exposes the switch as
  `redact-worksheet-code-resource-registrations: 'true'`; the mode does not
  calculate a formula, resolve a module path, load a DLL/XLL, inspect host
  trust settings, execute code, contact a provider, or reconstruct a dynamic
  argument, and it is not a general secret scrubber.
- Add direct-call, exact unsampled-static-input, resolved named-chain,
  default-evidence, policy, portfolio, JSON/Markdown/SARIF, and composite-Action
  regression coverage.

## 0.102.0 — 2026-07-26

- Add opt-in `--redact-unqualified-runtime-functions` rendering for `diff`,
  `check`, and `portfolio` JSON, Markdown, and SARIF artifacts. Default
  local-review output remains unchanged; the output-only boundary replaces
  direct stored bare `FF075` candidate material with `[unqualified
  runtime-function material redacted]`.
- Extend the boundary to before/after evidence for changed bare-call formulas
  and exact changed static input cells recorded by the private full dependency
  impact set, rather than the bounded impact sample shown in reports.
- Retain the private formula-defined-name chain signature for FF075 and, when
  it changes, conservatively redact changed defined-name before/after evidence
  so a dotted workbook-defined wrapper cannot disclose an argument that reaches
  a bare UDF deeper in the chain.
- Keep comparison facts, findings, policy evaluation, and exit status unchanged.
  The composite GitHub Action exposes the switch as
  `redact-unqualified-runtime-functions: 'true'`; the mode does not calculate
  a formula, resolve/load VBA, COM/Automation, XLL, or another provider,
  execute code, contact a runtime, reconstruct a dynamic argument, or claim to
  be a general secret scrubber.
- Add direct-call, exact unsampled-static-input, resolved named-chain,
  default-evidence, policy, portfolio, JSON/Markdown/SARIF, and composite-Action
  regression coverage.

## 0.101.0 — 2026-07-26

- Add opt-in `--redact-office-custom-functions` rendering for `diff`, `check`,
  and `portfolio` JSON, Markdown, and SARIF artifacts. Default local-review
  output remains unchanged; the output-only boundary replaces direct stored
  namespaced `FF066` formula material with `[Office custom-function material
  redacted]`.
- Extend the boundary to before/after evidence for changed custom-function
  formulas and exact changed static input cells recorded by the private full
  dependency impact set, rather than the bounded impact sample shown in reports.
- Track custom-function-relevant formula-defined-name bodies through private
  signatures and publish only an aggregate count. When that private chain
  changes, conservatively redact changed defined-name before/after evidence so
  an ordinary-looking wrapper cannot disclose a service argument.
- Keep comparison facts, findings, policy evaluation, and exit status unchanged.
  The composite GitHub Action exposes the switch as
  `redact-office-custom-functions: 'true'`; the mode does not calculate a
  formula, load an add-in or manifest, execute JavaScript, contact a runtime,
  reconstruct a dynamic argument, or claim to be a general secret scrubber.
- Add direct-call, exact unsampled-static-input, resolved named-chain,
  default-evidence, policy, portfolio, JSON/Markdown/SARIF, and composite-Action
  regression coverage.

## 0.100.0 — 2026-07-26

- Add opt-in `--redact-python-in-excel` rendering for `diff`, `check`, and
  `portfolio` JSON, Markdown, and SARIF artifacts. Default local-review output
  remains unchanged; the new output-only boundary replaces direct stored `PY`
  formula material with `[Python-in-Excel material redacted]`.
- Extend that boundary to the before/after evidence for changed PY-formula
  cells and exact changed static input cells that the private dependency graph
  recorded as reaching an inventoried PY cell. This uses the full private
  impact set rather than the bounded sample displayed to reviewers.
- Keep the pre-existing private `FF065` package/source ledger intact, while
  protecting ordinary semantic JSON evidence where Microsoft's static `PY`
  source can be visible. Comparison facts, findings, policy evaluation, and
  exit status are unchanged. The composite GitHub Action exposes the switch as
  `redact-python-in-excel: 'true'`; the mode does not parse/execute Python,
  calculate formulas, contact Microsoft Cloud, or claim to redact arbitrary
  workbook material.
- Add direct source, static-input, local-LAMBDA non-match, default-evidence,
  policy, portfolio, JSON/Markdown/SARIF, and composite-Action regression
  coverage.

## 0.99.0 — 2026-07-26

- Add opt-in `--redact-formula-external-actions` rendering for `diff`, `check`,
  and `portfolio` JSON, Markdown, and SARIF artifacts. It keeps ordinary local
  review output unchanged by default while hiding direct stored `FF064`
  formula-action/provider and `FF074` DDE formula material with a stable
  marker.
- Extend the output-only boundary to changed action/DDE cells and the exact
  statically visible input cells that the private dependency analysis recorded
  as reaching them. Conservatively hide changed defined-name before/after
  evidence when a relevant resolved action/DDE name chain changes, so a wrapper
  cannot expose an endpoint without spelling a native action itself.
- Keep comparison facts, policy evaluation, findings, and exit status unchanged.
  The composite GitHub Action exposes the switch as
  `redact-formula-external-actions: 'true'`. The mode does not evaluate
  formulas, contact a provider/DDE server, reconstruct dynamic destinations, or
  claim to be a general sensitive-data scrubber.
- Add direct-action/DDE, exact static-input beyond the bounded impact sample,
  resolved named-chain, default-evidence, policy, portfolio, JSON/Markdown/
  SARIF, and composite-Action regression coverage.

## 0.98.0 — 2026-07-26

- Add opt-in `--redact-external-workbook-links` rendering for `diff`, `check`,
  and `portfolio` JSON, Markdown, and SARIF artifacts. It keeps ordinary local
  review output unchanged by default, while replacing a whole serialized value
  that exposes a literal static external-workbook endpoint with a stable
  redaction marker. The renderer uses the existing static parser for direct
  A1, 3-D, defined-name, and book-only table spellings, plus a conservative
  visible-literal fallback; it never evaluates formulas or reconstructs
  text-built links.
- Keep redaction strictly output-only: it does not alter snapshots, comparison
  facts, policy evaluation, or exit status. The composite GitHub Action now
  exposes the same behavior as `redact-external-workbook-links: 'true'`, so a
  shared artifact and Markdown job summary can use the boundary directly.
- Add regression coverage for default local evidence, JSON/Markdown/SARIF
  direct reports, policy failures, portfolios, direct renderer use, and the
  composite Action.

## 0.97.0 — 2026-07-26

- Add `FF081` and the `no_external_workbook_link_surface_changes` policy guard
  (`FFP081`). The new private ledger detects material static external-workbook
  source or target changes even when an existing external formula remains at
  the same worksheet cell—an intentional boundary not covered by
  `no_new_external_links` / `FF004`.
- Inventory literal static endpoints persisted in worksheet formulas,
  workbook/sheet-local defined names, data-validation criteria, and standard
  DrawingML or ChartEx chart formula elements. Preserve only a count-only
  public ledger profile; source paths, workbook/sheet/name identities,
  formulas, validation ranges, chart part identities, and endpoint spellings
  stay inside the private ledger signature.
- Treat unreadable or otherwise unrecognized chart formula material as opaque
  coverage evidence for the new guard, so a material chart change fails closed
  rather than claiming link-surface coverage. Do not evaluate formulas, open,
  resolve, refresh, or trust an external source/cache. Conditional formatting
  remains outside this ledger because its spreadsheet formula grammar forbids
  external-cell references.
- Add same-cell source-swap, named-definition, data-validation, standard chart,
  opaque-chart, policy, generated-policy, profile/Markdown, and SARIF-redaction
  coverage.

## 0.96.0 — 2026-07-26

- Extend the candidate-only `FF079` portfolio graph through safely eligible,
  workbook-scoped named `LAMBDA` definitions. A caller such as
  `=ExternalMetric(7)` can retain each static direct or package-validated
  external A1, 3-D, workbook/sheet-local-name, or book-only table-selector
  endpoint in that LAMBDA body. The bridge also supports a nested eligible
  named LAMBDA and an eligible global formula-defined name that calls one.
- Preserve all fixed internal input edges alongside those private endpoint
  models. A definition such as
  `=LAMBDA(value,SUM(value,Inputs!$B$2,ExternalInput))` now makes a real
  invocation depend on both `Inputs!$B$2` and the validated external input;
  the same preservation fixes an eligible non-`LAMBDA` formula-defined wrapper
  that combines static internal and external inputs. Multiple endpoint kinds in
  one definition are retained rather than silently stopping at the first kind.
- Treat named-LAMBDA handling as static input-edge extraction, not evaluation:
  retain a LAMBDA endpoint only at a function call, never at a bare name. Keep
  dynamic, relative, recursive, local/shadowed, broken, unresolved, local-3-D,
  spill, explicit-intersection, tokenizer-failed, and unsupported definitions
  outside the graph. Keep names, paths, source identities/selectors, and
  consumer bridge identities private in `FF079` evidence. Add direct/package,
  nested/wrapper, mixed-endpoint, local-input, unsafe-form, scope, and
  JSON/Markdown/SARIF-redaction regression coverage.

## 0.95.0 — 2026-07-26

- Extend the candidate-only `FF079` portfolio graph through safely eligible,
  workbook-scoped non-`LAMBDA` formula-defined names. A definition such as
  `=SUM(ExternalInput)` or `=SUM('..\\inputs\\[Inputs.xlsx]Data'!$B$2:$B$4)`
  may retain each static direct or package-validated external A1, 3-D,
  workbook/sheet-local-name, or book-only table-selector endpoint to a calling
  worksheet formula. Eligible formula names can call another eligible global
  formula name without repeatedly rescanning unrelated definitions.
- Treat this as static input-edge extraction rather than formula evaluation.
  Require every external token to map to an existing validated endpoint and
  reject definitions with broken, unresolved, dynamic, relative, local-3-D,
  spill, explicit-intersection, tokenization-failed, sheet-local, or named
  `LAMBDA` semantics. Keep raw paths, source identities/selectors, and
  consumer bridge identities private in `FF079` evidence, while ordinary
  defined-name declarations remain normal review context.
- Add direct/package, A1/name/table/3-D, nested-formula, duplicate-endpoint,
  dynamic/relative/local/unresolved, scope-shadowing, and JSON/Markdown/SARIF
  redaction coverage.

## 0.94.0 — 2026-07-26

- Extend the candidate-only `FF079` portfolio graph through exact static
  external structured-table selectors: direct book-only spellings such as
  `='..\\inputs\\source.xlsx'!Sales[Amount]` and validated package-indexed
  spellings such as `=[1]!Sales[#Data]`. Finite, acyclic workbook-scoped
  consumer alias chains may terminate in either table endpoint form.
- Resolve a selector only after its private source spelling identifies an
  already inspected relative candidate and that candidate has exactly one
  case-insensitive matching table. Reuse FormulaFence's existing static table
  bounds for columns, contiguous column ranges, and `#All`, `#Data`,
  `#Headers`, or `#Totals`; source cells, not synthetic table nodes, remain
  `FF079` roots.
- Fail closed for bare table names (ambiguous with external names),
  sheet-qualified table spellings, `@`/`#This Row`, unsupported selectors,
  missing or colliding source tables, unsafe paths, local/shadowed consumer
  aliases, formula wrappers, and every unresolved package declaration. Keep
  raw source paths, table identities/selectors, and endpoint aliases private
  in JSON, Markdown, and SARIF evidence. Add parser, selector-bound,
  direct/package, alias-chain, collision, and redaction coverage.

## 0.93.0 — 2026-07-26

- Extend the candidate-only `FF079` portfolio graph through exact static
  external 3-D A1 spans: direct spellings such as
  `=[Inputs.xlsx]Jan:Mar!$B$2` and validated package-indexed spellings such as
  `=[1]Jan:Mar!$B$2`. The same finite, acyclic workbook-scoped consumer alias
  chains added in 0.92.0 can terminate in either span form.
- Expand a span only after its private source spelling resolves to an already
  inspected relative candidate. FormulaFence requires that candidate's complete
  raw OOXML tab catalog to agree with its inspected worksheet order, then adds
  the same static A1 bounds for every worksheet from the first endpoint through
  the last. It does not open, fetch, refresh, calculate, or trust an external
  target or cache.
- Fail closed for incomplete or inconsistent source tab metadata, reversed,
  missing, ambiguous, or non-worksheet endpoints; malformed paths, quoting,
  indexes, or A1 payloads; local/shadowed consumer aliases; formula wrappers,
  dynamic forms, and every unresolved package declaration. Keep source paths,
  endpoint identities, package targets, and consumer aliases out of JSON,
  Markdown, and SARIF portfolio evidence. Add parser, tab-catalog, direct,
  package, alias-chain, span-boundary, and redaction coverage.

## 0.92.0 — 2026-07-26

- Extend the candidate-only `FF079` portfolio graph through finite, acyclic
  chains of exact workbook-scoped consumer-name aliases. Every bridge must be
  one unqualified non-A1 name identity (with the usual optional leading `=`),
  and must terminate in an already-supported direct or validated
  package-indexed external A1, workbook-name, or sheet-local-name endpoint.
  The resolver retains only that prevalidated endpoint; it never evaluates the
  alias formula, resolves a path during parsing, or invents a source name.
- Keep the graph fail closed for sheet-local consumer aliases, local shadowing
  at a formula use site, functions, operators, constants, ranges, structured
  references, sheet-qualified targets, missing links, and cyclic chains.
  Source local-name scope and the existing direct/package validation boundaries
  remain unchanged.
- Add parser and portfolio coverage for direct and package A1/global/local-name
  alias chains, canonical no-leading-`=` name text, cycles, unresolved bridges,
  formula wrappers, local aliases, and JSON/Markdown/SARIF redaction. Validate
  the boundary against an independently maintained external-data workbook.

## 0.91.0 — 2026-07-26

- Extend the candidate-only `FF079` portfolio graph through exact,
  workbook-scoped consumer aliases whose stored definition is a direct external
  A1 cell/range or direct workbook-scoped external name, for example
  `'..\\inputs\\[Inputs.xlsx]Data'!$B$2:$B$4` or
  `'..\\inputs\\[Inputs.xlsx]InputRange'`. Canonical literal forms with one
  leading `=` are handled too. Direct formulas and the existing package-indexed
  and sheet-local paths retain their previous behavior.
- Preserve Excel name scope: any sheet-local consumer name shadows a
  same-named workbook alias, including aliases backed by an external package
  link. Formula-valued and sheet-local aliases never become an inferred
  cross-workbook edge.
- Tighten direct external literal parsing and fail closed for formula wrappers
  and operators, malformed quoting, invalid A1/name payloads, numeric package
  indexes presented as filenames, unsafe paths, and every unresolved target.
  Add parser, alias, scope-shadowing, and JSON/Markdown/SARIF-redaction
  coverage, plus an independently maintained external-workbook validation.

## 0.90.0 — 2026-07-26

- Extend the candidate-only `FF079` portfolio graph to static external
  sheet-local defined names: direct forms such as
  `=[Inputs.xlsx]Data!LocalInput` and documented package-indexed forms such as
  `=[N]Data!LocalInput`. Direct or package-indexed workbook-scoped consumer
  aliases can retain the same exact static spelling. The source sheet selects
  only that candidate workbook's matching local-name scope; a safely expanded
  local formula alias may resolve to fixed internal A1 destinations, but no
  global-name fallback is inferred.
- Preserve the existing no-following boundary. Indexed `N` is resolved only
  through the validated one-based document order of `externalReferences`, one
  declared `externalLink` part, one `externalBook`, and one external
  `externalLinkPath` relationship; a target must still normalize to an
  already-inspected relative candidate workbook. Raw source paths, indexes,
  sheet/name identities, and consumer aliases remain private in portfolio
  evidence; no target is opened, fetched, refreshed, cached, or evaluated.
- Fail closed for malformed/ambiguous quoting or indexes, A1/whole-row/
  whole-column/structured payloads in the name path, 3-D forms, unknown source
  sheets, source global fallbacks, different-sheet locals, dynamic/relative/
  cyclic/external/non-static source definitions, consumer sheet-local or
  formula aliases, unsafe paths, and every unvalidated package shape. Add
  parser, direct/package-alias, static-local-alias, scope-collision, dynamic,
  and JSON/Markdown/SARIF-redaction coverage.

## 0.89.0 — 2026-07-26

- Extend the candidate-only `FF079` portfolio graph to Excel's documented
  package-indexed external-A1 syntax, such as `=[N]Data!$B$2:$B$4`, including
  a direct workbook-scoped consumer alias whose stored definition is exactly
  `[N]Data!$B$2:$B$4`. `N` is resolved only through the validated one-based
  document order of `externalReferences`, one declared `externalLink` part,
  one `externalBook`, and one external `externalLinkPath` relationship; its
  target must still normalize to an already-inspected relative candidate
  workbook. Static cells, ranges, whole rows, and whole columns are supported;
  no target is opened, fetched, or evaluated.
- Keep package relationship targets, indexed source spellings, and consumer
  alias identities out of portfolio evidence. Decimal index zero, malformed or
  ambiguously quoted syntax, 3-D spans, names/structured references, sheet-
  scoped aliases, formula-defined aliases, invalid/ambiguous package shapes,
  unsafe paths, caches, DDE/OLE/non-workbook parts, and every other non-static
  form remain unresolved rather than creating a guessed edge.
- Add parser, source-alias, declaration-order, relative-range, whole-column,
  absolute-path, local/formula-alias, package-ambiguity, and JSON/Markdown/
  SARIF-redaction coverage. Validate the boundary against an independently
  maintained OpenPyExcel workbook containing a real `[1]Sheet1!$A$1` alias.

## 0.88.0 — 2026-07-26

- Extend the candidate-only `FF079` portfolio graph to Excel's package-indexed
  external-name spelling, `=[N]!InputRange`, and to a direct
  workbook-scoped consumer alias whose stored definition is exactly
  `[N]!InputRange`. `N` is resolved only through the one-based document order
  of `externalReferences`, one declared `externalLink` part, one
  `externalBook`, and one external `externalLinkPath` relationship; its target
  must still normalize to an already-inspected relative candidate workbook.
  The candidate source name must fully expand to static internal A1
  destinations without evaluation.
- Keep package relationship targets, indexed external source spellings, and any
  cached `externalLink` material out of portfolio evidence. Ordinary source or
  consumer defined-name declarations remain normal workbook-review items.
  FormulaFence neither trusts cached values nor opens a target.
  Absolute/UNC/URI/escaping paths, malformed
  or ambiguous declarations, non-workbook/DDE/OLE parts, package external-A1
  forms, sheet-scoped aliases, consumer formula-name aliases, dynamic/missing
  source names, and every other non-static form remain unresolved rather than
  creating a guessed edge.
- Add declaration-order, direct/indexed-name, source-alias, dynamic-name,
  absolute-path, malformed-package, and JSON/Markdown/SARIF redaction
  coverage, including an independently maintained public external-link
  workbook fixture.

## 0.87.0 — 2026-07-26

- Extend the candidate-only `FF079` portfolio graph to the documented direct
  external workbook-scoped name form such as `=[Inputs.xlsx]InputRange`. The
  exact relative source path must still identify an already-inspected candidate
  workbook, and its source name must expand completely to static internal A1
  destinations. Safe global formula-defined aliases work through the existing
  non-evaluating name resolver.
- Keep external source paths and defined-name identities private. Sheet-scoped
  names, direct external structured references, absolute/UNC/URI/escaping
  paths, missing names, and names whose source definitions are dynamic,
  relative, cyclic, external, 3-D, malformed, or otherwise not fully static
  remain unresolved rather than creating a guessed edge. A defined-name
  declaration change remains its ordinary `FF008` review event; `FF079` roots
  are changed candidate cells.
- Add parser, snapshot, graph, policy, JSON/Markdown/SARIF-redaction, static
  formula-name alias, dynamic-name rejection, and unsafe-path coverage for the
  expanded boundary.

## 0.86.0 — 2026-07-26

- Add `FF079`, a high-severity candidate-only static cross-workbook impact
  boundary for directory portfolios. It follows local formula dependencies and
  only direct external A1 cells/ranges whose exact relative workbook spelling
  resolves to an already-inspected candidate path; evidence contains relative
  workbook identities, Excel logical cells, counts, and deterministic shortest
  paths. The raw external source spelling stays private.
- Add `FFP079` through `no_cross_workbook_impacts`, including the generated
  starter policy. FormulaFence does not open, download, evaluate, refresh, or
  otherwise follow an external link; it does not guess basename matches or
  resolve absolute, UNC, URI, escaping, named, table, 3-D, or dynamic forms.
- Add a global `--max-link-impact` bound (100,000 source-to-node states by
  default) to the portfolio CLI and composite GitHub Action. Reaching the bound
  emits redacted critical `FF080`, retains partial evidence, and exits 2 rather
  than claiming complete impact coverage.
- Extend portfolio JSON, Markdown, and SARIF with safe cross-workbook evidence,
  and add fixture coverage for relative/case-insensitive links, lazy ranges,
  transitive local/cross-workbook paths, unresolved-path privacy, policy
  enforcement, and fail-closed traversal bounds.

## 0.85.0 — 2026-07-26

- Add a bounded, local-first `formulafence portfolio` command for recursive
  `.xlsx` / `.xlsm` directory comparison. It uses a relative path as the
  workbook identity, produces consolidated JSON, Markdown, and SARIF evidence,
  and deliberately reports a move as a removal plus an addition rather than
  guessing a rename.
- Add high-severity `FF077` for a supported workbook added to or removed from a
  portfolio and `FFP077` through the new
  `no_portfolio_membership_changes` policy. Existing policy controls apply
  independently to each matched workbook.
- Fail closed for unsupported spreadsheet formats, control-character or
  case-colliding paths, symlinks, inventory limits, and empty dual portfolios.
  Transient Office lock files are ignored. An unreadable supported workbook
  yields redacted critical `FF078` evidence plus exit status 2 while the report
  preserves all other portfolio results; a newly added or removed unreadable
  workbook also retains its known `FF077` / `FFP077` membership evidence.
- Extend the public composite GitHub Action to detect matching directory inputs
  automatically and expose a validated `max-workbooks` bound (512 by default).
  Portfolio SARIF carries relative workbook artifact URIs and per-cell logical
  locations without emitting an absolute runner path. The CLI and Action refuse
  report paths that would overwrite an input or write inside a portfolio root.

## 0.84.0 — 2026-07-26

- Extend the existing `FF065` / `FFP065` Python-in-Excel boundary to recognize
  the separately stored 2022 `pythonScripts.xml` package contract—its root,
  script records, workbook relationship, and content type—alongside the
  documented 2023 `python.xml` contract.
- Fingerprint every stored Python package part independently when both
  representations coexist. Public profiles continue to expose only aggregate
  physical-part, formula-call, script, environment, initialization, and
  coverage counts; Python source, IDs, script indexes, formulas, locations,
  and raw XML remain private.
- Treat conflicting package declarations, malformed roots, missing/unbound
  parts, and read-budget limits as explicit coverage evidence rather than
  guessing which contract Excel will run. Relationship-ID-only rewrites still
  normalize and neither Python representation is loaded, evaluated, or sent to
  a runtime.

## 0.83.0 — 2026-07-26

- Add `FF076`, a high-severity private boundary for Excel 4.0 / XLM
  automatic-macro routing. It recognizes only workbook-scoped `Auto_Open`,
  `Auto_Close`, `Auto_Activate`, and `Auto_Deactivate` names (including an
  optional `_xlnm.` prefix) whose direct internal single-cell A1 definitions
  target a raw declared XLM macro sheet.
- Preserve private detection of same-count target or stored-definition changes,
  while public profiles, findings, policy results, and SARIF expose only total
  and per-event counts. The general defined-name diff deliberately remains
  readable for normal workbook review.
- Add the fail-closed `no_xlm_automatic_macro_binding_changes` policy rule
  (`FFP076`). FormulaFence does not evaluate or resolve names, rely on the
  reserved/unused `definedName@xlm` attribute, parse or execute XLM commands,
  inspect Excel trust settings, or claim that a binding will run.

## 0.82.0 — 2026-07-26

- Add `FF075`, a private ledger for unknown unqualified worksheet-call
  candidates that could resolve through VBA, a COM/Automation add-in, an XLL,
  or another registered runtime. The public profile and finding details expose
  only formula-cell, call, and relevant formula-defined-name counts; candidate
  names, formulas, arguments, locations, provider identities, and host details
  remain private.
- Use a stable native Excel function catalogue derived from Microsoft's current
  alphabetical reference, with explicit modern/runtime compatibility additions,
  so ordinary calls such as `XLOOKUP`, `VSTACK`, `FIELDVALUE`, and `PY` do not
  become candidates because a third-party parser changes. Workbook-defined
  names, local `LET`/`LAMBDA` bindings, qualified calls, and dedicated legacy
  XLM spellings remain outside this generic boundary.
- Propagate candidates through formula-defined names, named LAMBDAs, recursive
  groups, and sheet-local definitions; retain private signatures for same-count
  callable and definition changes; inventory dormant stored definitions; and
  guard ordinary edits that statically reach a candidate formula.
- Add the fail-closed `no_unqualified_runtime_function_changes` policy rule
  (`FFP075`). FormulaFence does not evaluate formulas, resolve/load VBA,
  COM/Automation, XLL, or any registered provider, inspect host trust settings,
  or execute code.

## 0.81.0 — 2026-07-26

- Add a root composite GitHub Action for FormulaFence policy checks and
  semantic diffs. It installs the selected action source, accepts baseline,
  candidate, optional policy, report-format, output, and severity inputs, and
  exposes deterministic report-path and exit-code outputs.
- Keep report generation, Markdown job-summary evidence, and optional artifact
  upload alive through a FormulaFence policy failure, then re-emit the original
  exit code. The action confines reports and inputs to the workspace and
  refuses to overwrite an input workbook or policy.
- Add action metadata, isolated action-script contract tests, and a GitHub CI
  workflow smoke check that runs the public Action itself.

## 0.80.0 — 2026-07-26

- Add `FF074`, a private ledger for direct DDE-style formula syntax using the
  documented `application|topic!item` form in worksheet formulas,
  formula-defined names, and named `LAMBDA` bodies. The lexical detector
  deliberately skips pipes inside ordinary quoted sheet names and string
  literals, and it never evaluates a formula, resolves an endpoint, looks up
  or launches a DDE server, or sends a command.
- Propagate stored DDE definitions through nested, recursive, and sheet-local
  names to their invoking formulas, including a conservative raw `LAMBDA`
  wrapper fallback when direct DDE syntax defeats the underlying tokenizer.
  Public profiles expose only formula-cell, DDE-link, and defined-name counts;
  services, topics, items, formulas, locations, and identities stay inside
  private comparison signatures. Raw `externalLink` DDE/OLE packages remain a
  distinct `FF025` package boundary.
- Add the fail-closed `no_formula_dde_link_changes` policy rule (`FFP074`) and
  static-input coverage for invoking named LAMBDAs.

## 0.79.0 — 2026-07-26

- Add `FF073`, a private ledger for selected legacy XLM action and
  event-dispatch calls—`CALL`, `EXEC`, `EXECUTE`, `RUN`, `SEND.KEYS`, and
  `ON.*`—when stored in formula-defined names or named `LAMBDA` bodies. The
  ledger follows nested, recursive, and sheet-local names to their invoking
  formula cells without evaluating a formula or resolving an action target,
  handler, DLL, DDE command, macro, or program.
- Preserve private review of same-count definition/invocation changes,
  uninvoked stored names, and ordinary static input changes. Safe profiles and
  Markdown expose only invoking-cell, selected-action, and
  formula-defined-name counts; targets, handlers, formulas, arguments,
  locations, and name identities stay private. Workbook-defined callables
  shadow these spellings rather than being asserted to be legacy XLM actions.
- Add the fail-closed `no_formula_defined_xlm_action_changes` policy rule
  (`FFP073`). Direct worksheet action calls and raw XLM macro-sheet parts
  remain outside this stored-definition boundary; FormulaFence does not claim
  to interpret arbitrary XLM commands or execute any action.

## 0.78.0 — 2026-07-26

- Extend `FF072` / `FFP072` from native `CELL` and `INFO` to native `SHEET`
  and `SHEETS` calls in worksheet formulas, formula-defined names, and named
  `LAMBDA` bodies. The existing private definition-chain, sheet-local-name,
  recursion, static-input, redaction, and fail-closed policy path now protects
  workbook-structure information without evaluating a formula.
- Read the private raw OOXML workbook tab catalog so `SHEET` and omitted-
  reference `SHEETS()` calls are compared against all declared tabs, including
  hidden, very-hidden, chart, macro, and dialog sheets. A catalog membership,
  order, or name change emits `FF072`; a visibility-only change does not,
  matching the documented inclusion of hidden tabs.
- Add safe `SHEET`, `SHEETS`, and omitted-`SHEETS` aggregate counts to profiles
  and Markdown. The raw tab-catalog comparison material, formulas, arguments,
  locations, defined-name identities, and calculated values remain private in
  the dedicated ledger; ordinary sheet inventory remains normal reviewer
  context. Explicit `SHEETS(reference)` arguments are inventoried but not
  evaluated or guessed as single-sheet versus 3-D references.

## 0.77.0 — 2026-07-26

- Extend the existing `FF064` formula external-action ledger to stored
  provider-backed `STOCKHISTORY` and every documented Cube function:
  `CUBEKPIMEMBER`, `CUBEMEMBER`, `CUBEMEMBERPROPERTY`,
  `CUBERANKEDMEMBER`, `CUBESET`, `CUBESETCOUNT`, and `CUBEVALUE`.
  FormulaFence now detects same-count market-provider, connection, query, and
  Cube-set changes without evaluating a formula, contacting a provider, or
  exposing formula arguments.
- Preserve the existing named-formula, named-`LAMBDA`, recursive-chain,
  sheet-local-resolution, static-input, private-signature, and fail-closed
  `no_formula_external_action_changes` (`FFP064`) semantics for the expanded
  provider family. Workbook-defined callables shadow the native spelling rather
  than being misclassified.
- Extend safe profiles and Markdown reports with separate `STOCKHISTORY` and
  aggregate Cube-function counts. Connection names, MDX expressions, market
  symbols, field/property names, formula cells, and results remain private.

## 0.76.0 — 2026-07-26

- Add FF072, a private ledger for native Excel CELL and INFO calls in worksheet
  formulas, formula-defined names, and named LAMBDA bodies. The calls can
  observe file, folder, location, client, calculation, and workbook state
  beyond ordinary visible precedents; FormulaFence captures stored syntax and
  definition chains without evaluating a call or simulating Excel state.
- Propagate stored calls through nested and sheet-local formula names to
  invoking cells. Profiles aggregate only safe formula-cell, call,
  formula-defined-name, and CELL-without-explicit-reference counts; same-count
  definition/invocation changes, uninvoked names, and ordinary static input
  edits remain reviewable through private signatures.
- Add the fail-closed no_formula_environment_information_changes policy rule
  (FFP072). FormulaFence does not infer the selected cell, determine an
  information type, resolve dynamic arguments, inspect file/folder/client or
  workspace state, or claim that a state-only workbook change alters a call.

## 0.75.0 — 2026-07-26

- Add FF071, a private ledger for selected legacy XLM environment-information
  calls—GET.WORKBOOK, GET.WORKSPACE, and GET.DOCUMENT—stored in
  formula-defined names and named LAMBDA bodies. The calls can observe
  workbook, workspace/client, or document state; FormulaFence captures their
  stored surface without evaluating a call, inferring its information type, or
  simulating Excel state.
- Propagate stored calls through nested and sheet-local formula names to
  invoking cells. Same-count invocation or definition changes, uninvoked stored
  names, and ordinary cell edits that statically reach a stored argument edge
  remain reviewable through private signatures while profiles expose only safe
  formula-cell, call, and formula-defined-name counts.
- Add the fail-closed no_formula_defined_xlm_environment_information_changes
  policy rule (FFP071). FormulaFence does not evaluate a formula, resolve a
  dynamic reference, inspect document/workspace/client state, enumerate loaded
  add-ins, inspect printers, or execute an XLM macro. Direct worksheet calls
  and raw XLM macro sheets remain outside this narrow stored-definition
  boundary.

## 0.74.0 — 2026-07-26

- Add `FF070`, a private ledger for legacy XLM `GET.CELL` information calls
  stored in formula-defined names and named `LAMBDA` bodies. Microsoft
  identifies `GET.CELL` / `xlfGetCell` as an XLM information function; this
  captures the stored formula surface without claiming FormulaFence evaluates
  a call, identifies its requested information type, or simulates Excel state.
- Propagate stored calls through nested and sheet-local formula names to
  invoking cells. Same-count invocation or definition changes, uninvoked stored
  names, and ordinary cell edits that statically reach a stored argument edge
  remain reviewable through private signatures while profiles expose only safe
  formula-cell, call, and formula-defined-name counts.
- Add the fail-closed `no_formula_defined_xlm_get_cell_changes` policy rule
  (`FFP070`). FormulaFence does not evaluate a formula, resolve a dynamic
  reference, render formatting or display text, inspect comments/protection,
  or execute an XLM macro. Direct worksheet `GET.CELL` calls and raw XLM macro
  sheets remain outside this narrow stored-definition boundary.

## 0.73.0 — 2026-07-26

- Add `FF069`, a private ledger for legacy XLM `EVALUATE` calls stored in
  formula-defined names and named `LAMBDA` bodies. Microsoft documents
  `EVALUATE` as reducing a valid character string to a worksheet value; this
  captures a stored dynamic-expression surface without claiming that
  FormulaFence calculates or parses that expression.
- Propagate stored calls through nested and sheet-local formula names to
  invoking cells. Same-count invocation or definition changes, uninvoked stored
  names, and ordinary cell edits that statically reach a stored argument edge
  remain reviewable through private signatures while profiles expose only safe
  formula-cell, call, and formula-defined-name counts.
- Add the fail-closed `no_formula_defined_xlm_evaluation_changes` policy rule
  (`FFP069`). FormulaFence does not evaluate formula text, infer dependencies
  inside a runtime-generated expression, execute an XLM macro, or inspect host
  trust settings. Direct worksheet `EVALUATE` calls and raw XLM macro sheets
  remain outside this narrow stored-definition boundary.

## 0.72.0 — 2026-07-26

- Add `FF068`, a private ledger for legacy XLM `REGISTER` calls stored in
  formula-defined names and named `LAMBDA` bodies. Microsoft documents that
  `REGISTER` can register DLL functions/commands or load an XLL, and that its
  macro types can be called from a defined-name definition. This closes the
  stored-definition gap without treating direct worksheet formulas or raw XLM
  macro-sheet XML as the same surface.
- Propagate registrations through nested and sheet-local formula names to
  invoking cells. Same-count invocation or name-definition changes, uninvoked
  stored names, and ordinary cell edits that statically reach an invocation
  remain reviewable through private signatures while profiles expose only safe
  formula-cell, call, and formula-defined-name counts.
- Add the fail-closed `no_formula_defined_xlm_registration_changes` policy rule
  (`FFP068`). FormulaFence does not evaluate a formula, execute a macro,
  resolve a module path, load a DLL/XLL, or inspect host trust settings.

## 0.71.0 — 2026-07-26

- Extend the existing `FF064` formula external-action boundary through
  formula-defined names and named `LAMBDA` bodies. Stored `HYPERLINK`,
  `WEBSERVICE`, `IMAGE`, and `RTD` calls can no longer be hidden behind a
  formula name or nested named function.
- Track relevant named definitions in a separate private signature, including
  an uninvoked stored definition. Same-count name-definition changes and
  ordinary cell edits that statically reach an invoking action formula emit
  `FF064`; profiles expose only formula-cell, call, function, and
  formula-defined-name counts.
- `no_formula_external_action_changes` continues to make this fail closed as
  `FFP064`. FormulaFence does not evaluate a formula, resolve a destination,
  fetch content, follow a link, or start an RTD provider.

## 0.70.0 — 2026-07-26

- Add `FF067`, a private ledger for stored worksheet and formula-defined
  `REGISTER.ID` calls.
  Microsoft documents this legacy worksheet function as returning a DLL or
  code-resource registration ID and registering the resource when required.
  Public profiles and finding details expose only formula-cell, call, and
  relevant formula-defined-name counts; module paths, procedure names, type
  strings, formulas, arguments, locations, and name identities remain private.
- Propagate registrations held in formula-defined names and named `LAMBDA`
  bodies to their invoking worksheet formulas. Same-count formula or
  named-definition changes remain visible via private signatures, and ordinary
  cell edits that statically reach a registration emit `FF067`. Recursive and
  sheet-local named definitions remain cycle-safe and scope-aware.
- Add the fail-closed `no_worksheet_code_resource_registration_changes` policy
  rule (`FFP067`). FormulaFence does not evaluate a formula, resolve or load a
  DLL/XLL, inspect host trust settings, or determine whether registration
  succeeds. The separate XLM macro-sheet boundary continues to cover raw
  macro-sheet `CALL`/`REGISTER` program material.

## 0.69.0 — 2026-07-26

- Add `FF066`, a private ledger for namespaced Office custom-function call
  candidates, such as the documented `CONTOSO.ADD` form. The public profile and
  finding details expose only formula-cell, call, and namespace counts; call
  names, namespaces, locations, formulas, and arguments remain private.
- Flag same-count callable or argument changes and ordinary cell edits that
  statically reach a candidate formula. The ledger excludes known native dotted
  Excel functions, direct workbook-defined callables, and `_xlfn.` / `_xlws.`
  compatibility forms. Candidates stored in formula-defined names and named
  `LAMBDA` bodies are instead propagated to their invoking worksheet formulas.
  It deliberately does not classify unqualified VBA, COM, or XLL UDFs, and a
  candidate is not treated as proof that an Office Add-in is installed or runnable.
- Add the fail-closed `no_office_custom_function_changes` policy rule
  (`FFP066`). FormulaFence does not resolve a candidate to an add-in, read a
  manifest, load or execute an add-in, evaluate a formula, or contact a
  custom-function runtime.

## 0.68.0 — 2026-07-26

- Add `FF065`, a private Python-in-Excel ledger for the documented workbook
  Python part and stored `PY` formulas. FormulaFence fingerprints bounded raw
  Python XML—including code, environment definitions, script order, and
  extensions—privately while profiles and finding details expose only safe
  package, formula-cell, function, script, environment, initialization, and
  coverage counts.
- Flag source-code/environment/package changes, stored PY formula-binding
  changes, and ordinary cell edits that statically reach a PY formula. This
  catches static source changes such as `=_xlfn._xlws.PY(0,0,A1)` without
  parsing or evaluating Python, decoding a script index, or contacting the
  Microsoft Cloud runtime. Relationship-ID-only rewrites normalize; malformed,
  missing, unbound, oversized, unreadable, and over-budget metadata stays
  visible as coverage evidence.
- Add the fail-closed `no_python_in_excel_changes` policy rule (`FFP065`).
  FormulaFence never loads Python source as Python, runs code, evaluates a PY
  formula, resolves its result, contacts Microsoft Cloud, or validates runtime
  package support.

## 0.67.0 — 2026-07-26

- Add `FF064`, a private ledger for stored `HYPERLINK`, `WEBSERVICE`, `IMAGE`,
  and `RTD` formula calls, including `_xlfn.` compatibility spellings. Public
  profiles and finding details expose only safe action-cell and per-function
  counts; formulas, arguments, destinations, provider names, results, and
  locations remain private action evidence.
- Flag same-count call/argument/provider changes and ordinary edits that reach
  an action formula through FormulaFence's static dependency graph, covering
  source-cell retargeting such as `=HYPERLINK(A1, ...)` without evaluating it.
  Dynamic or unresolved argument sources remain explicit parser-coverage
  boundaries rather than ungrounded reachability claims.
- Add the fail-closed `no_formula_external_action_changes` policy rule
  (`FFP064`). FormulaFence never calculates, resolves, fetches, opens, clicks,
  follows, authenticates to, or executes an action function or RTD provider.

## 0.66.0 — 2026-07-26

- Add `FF063`, a bounded package-wide ledger for every canonical OPC root or
  part-level relationship with `TargetMode="External"`. It catches hyperlink,
  image, and opaque remote targets even when their source is outside a known
  workbook-feature binding. Source parts, types, identifiers, targets, unknown
  metadata, and raw XML remain in private signatures; public output exposes
  only safe aggregate counts.
- Add the fail-closed `no_external_relationship_changes` policy rule (`FFP063`)
  and normalize writer-selected relationship-ID rewrites. Material target,
  type, source, or coverage changes remain high-severity findings, while known
  feature-specific boundaries retain their own focused findings.
- Bound relationship XML inspection to 16 MiB per part, 64 MiB per workbook,
  and 512 parts. Duplicate, orphaned, malformed, unsafe, unreadable,
  oversized, or over-budget relationship metadata produces explicit coverage
  evidence; FormulaFence does not resolve, open, fetch, execute, or establish
  trust for any target.

## 0.65.0 — 2026-07-26

- Extend the existing `FF028` Office Web Add-in boundary beyond document task
  panes to documented worksheet `x15:webExtensions` bindings and in-content
  DrawingML `we:webextensionref` frames. FormulaFence validates worksheet
  `appRef` values against direct web-extension definition bindings, privately
  fingerprints local formulas, and follows active `mc:Choice` frame
  relationships while exposing only safe structural counts.
- Handle the native-picture fallback of an in-content frame as a normal bounded
  worksheet image (`FF059`) rather than double-counting it as a shape or
  silently dropping its preview bytes. Stable nonvisual and relationship-ID
  rewrites normalize; material worksheet-binding or active-frame changes emit
  `FF028` and remain subject to `no_office_web_addin_changes` (`FFP028`).
- Keep supported worksheet binding extensions out of the ordinary reader only
  after raw evidence is collected, avoiding its lossy unsupported-extension
  warning. Missing, malformed, unbound, duplicate, unsafe, oversized, or
  over-budget metadata remains explicit coverage evidence. Worksheet-binding
  and in-content DrawingML scans are bounded to 16 MiB per part, 64 MiB per
  workbook, and 512 parts; FormulaFence does not install, fetch, execute, or
  render an add-in or follow external targets.

## 0.64.0 — 2026-07-26

- Extend the existing `FF044` Worksheet DrawingML SmartArt boundary to
  fingerprint bounded direct internal Image targets from each Diagram Data
  part. Image bytes, names, targets, relationship IDs, and raw XML remain
  private; profiles, Markdown, JSON, and SARIF expose only safe image,
  fingerprinted-image, and uninspected-image counts.
- Support both transitional and Strict Image relationship forms. Coordinated
  Diagram Data image relationship-ID rewrites remain quiet, while byte-only
  image changes emit the existing high-severity `FF044` finding and can be
  blocked by `no_worksheet_drawing_shape_changes` (`FFP044`).
- Bound direct image hashing to 32 MiB per image, 64 MiB per workbook, and 512
  images. Missing, duplicate, external, unsafe, unreadable, oversized,
  over-budget, or unsupported component-side relationships become visible
  coverage evidence. FormulaFence never decodes or renders media, fetches a
  target, or follows hyperlinks, second-hop targets, or relationships from
  other SmartArt component kinds.

## 0.63.0 — 2026-07-26

- Inspect Office 2016+ DrawingML ChartEx (`cx:chart`) workbook controls under
  the existing high-severity `FF030` / `FFP030` chart boundary. FormulaFence
  recognizes ChartEx graphic frames inside Excel's `mc:AlternateContent`
  fallback form, follows the Office 2014 `chartEx` relationship to bounded
  `chartEx*.xml` parts, and compares their private XML and normalized binding
  semantics without exposing formulas, labels, titles, target paths, XML, or
  payload bytes.
- Fingerprint bounded direct ChartEx style, colour-style, drawing, image,
  theme-override, and embedded-package targets. Profiles and Markdown now
  expose safe ChartEx references, parts, series, titles, data-reference, and
  direct-payload counts alongside legacy chart counts. Relationship-ID rewrites
  and equivalent internal target spellings remain quiet.
- Treat malformed, missing, duplicate, orphaned, external, unsafe, unsupported,
  oversized, and over-budget ChartEx material as visible coverage evidence.
  FormulaFence still does not calculate formulas, render or assess ChartEx
  visuals, infer chart-to-cell impact, parse media/package formats, resolve
  second-hop relationships, or interpret ChartEx-specific visualization or
  nested-chart semantics.

## 0.62.0 — 2026-07-26

- Extend the existing `FF044` Worksheet DrawingML control boundary to inspect
  non-chart `xdr:graphicFrame` SmartArt diagrams. FormulaFence follows the
  graphic frame's `dgm:relIds` bindings for diagram data, layout, quick-style,
  and colour parts, plus direct worksheet-drawing `diagramDrawing` rendering
  parts; private signatures retain the frame, anchor, relationship semantics,
  and component material while profiles, Markdown, JSON, and SARIF expose only
  structural counts.
- Keep chart graphic frames in `FF030` and native pictures in `FF059`. Unknown
  non-chart graphic-frame URI types, malformed/missing/duplicate/unsafe
  SmartArt bindings, unreadable/oversized/over-budget component material, and
  component-side relationships outside the bounded scan become visible coverage
  evidence rather than silent omissions. Transitional and Strict DrawingML are
  supported.
- A SmartArt component or frame change emits the existing high-severity
  `FF044` finding and can be blocked by the existing fail-closed
  `no_worksheet_drawing_shape_changes` policy (`FFP044`). Relationship-ID and
  non-visual-ID rewrites normalize when their supported semantics are unchanged.
  FormulaFence compares stored declarations only: it does not render SmartArt,
  calculate final layout or visibility, resolve themes, follow component-side
  media/hyperlink targets, or parse/hash media.

## 0.61.0 — 2026-07-26

- Inspect legacy shared-workbook revision history directly from raw OOXML:
  workbook-bound `revisionHeaders` parts, header-to-log relationships, bounded
  `revisionLog` records, and shared/tracking/history-retention/protection
  controls. Profiles, Markdown, JSON, and SARIF expose only structural
  aggregates; historic values, locations, author identities, timestamps,
  comments, GUIDs, relationship identifiers, and XML remain private.
- Emit `FF062` for a material revision header/log, audit record, relationship,
  control, or coverage change and add the fail-closed
  `no_shared_workbook_revision_changes` policy rule (`FFP062`). Equivalent
  Boolean/integer spelling, coordinated relationship-ID rewrites, and
  transitional/Strict relationship types normalize without suppressing material
  history changes.
- Support header-only and header-plus-log packages, including Strict
  SpreadsheetML. Missing, duplicate, malformed, unsupported, unresolved,
  unsafe, oversized, or over-budget revision metadata is visible coverage
  evidence. FormulaFence fingerprints stored declarations rather than applying
  revisions, reconstructing historic state, resolving conflicts, validating
  author/timestamp claims, rendering Excel, or interpreting arbitrary future
  extensions.

## 0.60.0 — 2026-07-26

- Inspect Excel Table presentation controls directly from raw OOXML: applied
  `tableStyleInfo` bindings/toggles, applicable custom `tableStyle` /
  `tableStyleElement` definitions, semantic Dxf material, and direct
  Table/TableColumn Dxf and named-cell-style references. Profiles, Markdown,
  JSON, and SARIF retain only structural counts; table/style names, formatting,
  colours, identifiers, and XML remain private.
- Emit `FF061` for a material Table Style control or coverage change and add
  the fail-closed `no_table_style_control_changes` policy rule (`FFP061`).
  Boolean/default spelling, case-only style names, Excel `xr9:uid` revision
  provenance, and coordinated Dxf reordering/ID rewrites normalize without
  suppressing a material presentation change.
- Isolate Table Style XML and direct table-local presentation references from
  the ordinary reader only after collecting raw evidence. Transitional and
  Strict SpreadsheetML are supported; missing, duplicate, malformed,
  unresolved, unsupported, oversized, or over-budget controls become explicit
  coverage evidence. FormulaFence does not render final Table appearance,
  resolve themes, apply conditional formatting, or cover PivotTable-only style
  regions; `defaultTableStyle` remains a new-table preference rather than an
  existing-table binding.

## 0.59.0 — 2026-07-26

- Inspect legacy Excel Custom Views directly from raw workbook and sheet OOXML:
  GUID-linked `customWorkbookView` / `customSheetView` declarations across
  transitional and Strict SpreadsheetML worksheets, dialog sheets, and chart
  sheets. Private signatures retain alternate display, hidden/filter, print,
  pane, comment/object, and supported per-sheet state while profiles, Markdown,
  JSON, and SARIF expose structural aggregates only.
- Emit `FF060` for a material Custom View declaration, linked per-sheet state,
  or coverage change, and add the fail-closed
  `no_custom_workbook_view_changes` policy rule (`FFP060`). Coordinated GUID and
  sheet-ID/active-sheet-ID rewrites, Boolean/default spelling, and unsigned
  integer spelling normalize without hiding material changes.
- Keep Custom View XML out of the ordinary workbook-reader copy only after raw
  evidence is collected, so readers that reject otherwise valid legacy views
  cannot silently erase coverage. Missing, duplicate, malformed, unsupported,
  unsafe, oversized, over-budget, or incompletely linked declarations become
  visible coverage evidence. FormulaFence does not activate/render views,
  calculate filtered results or final print output, interpret future
  extensions, or support Custom Views on unsupported sheet types.

## 0.58.0 — 2026-07-26

- Extend `FF044` Worksheet DrawingML controls to inspect transitional and
  strict `xdr:cxnSp` connectors, including connectors nested in `xdr:grpSp`.
  Private signatures retain connector anchor, geometry/style, nonvisual
  declarations, and `stCxn`/`endCxn` endpoint attachment semantics alongside
  the existing regular-shape/group controls; profiles, Markdown, JSON, and
  SARIF expose connector and attachment aggregates only.
- Resolve connector endpoint IDs against supported DrawingML object identities
  without leaking those IDs. Consistent rewrites of nonvisual IDs and matching
  connector endpoints normalize; reattachment, endpoint-site, geometry, style,
  or anchor changes produce the existing high-severity `FF044` finding and the
  existing fail-closed `no_worksheet_drawing_shape_changes` policy (`FFP044`).
- Treat malformed, duplicate, missing, or unsupported connector endpoint
  metadata as visible coverage evidence. Free connectors remain supported.
  FormulaFence still does not render DrawingML, evaluate final routing or
  visibility, resolve themes, fetch external targets, parse media, or inspect
  `xdr:graphicFrame`, SmartArt, and other unsupported drawing objects; native
  `xdr:pic` images remain in `FF059`.

## 0.57.0 — 2026-07-26

- Inspect native worksheet image controls before ordinary readers discard their
  package bindings: anchored transitional and strict DrawingML `xdr:pic`
  objects (including pictures inside groups), direct worksheet backgrounds, and
  VML-backed header/footer watermark images. Private signatures retain anchors,
  visual declarations, relationship semantics, and bounded direct image-payload
  hashes; profiles, Markdown, JSON, and SARIF expose aggregate counts only.
- Emit `FF059` for a material native worksheet-image change and add the
  fail-closed `no_worksheet_image_changes` policy rule (`FFP059`). This catches
  a changed floating picture, sheet background, or printed watermark while
  formulas and ordinary cells remain unchanged.
- Normalize writer-selected non-visual and VML IDs plus consistent
  relationship-ID rewrites. Keep chart drawings, rich-data/in-cell images,
  Themes, ActiveX/OLE image controls, text/group shapes, and header/footer text
  in their existing dedicated boundaries. Missing, duplicate, malformed,
  unsafe, oversized, or over-budget XML/payload material is explicit coverage
  evidence; FormulaFence neither renders nor decodes images, follows an
  external target, or calculates final print/layout behavior.

## 0.56.0 — 2026-07-26

- Inspect material worksheet-dimension controls directly from raw transitional
  and strict SpreadsheetML before ordinary readers flatten row/column state:
  `sheetFormatPr` default row/column/base sizing, positive direct row heights,
  layered effective positive column widths, `bestFit`, `customHeight`, Office
  2010 `x14ac:dyDescent` baseline adjustments, and active thick-border automatic
  row-height adjustments. Profiles, Markdown, JSON, and SARIF expose aggregate
  counts only; dimension values, sheet names, row/column targets, writer hints,
  and raw XML remain private.
- Emit `FF058` for a material worksheet-dimension control change and add the
  fail-closed `no_worksheet_dimension_changes` policy rule (`FFP058`). This
  closes the review gap where wrapped content, report framing, baseline layout,
  or automatic pagination can change without a formula or ordinary cell edit.
- Keep zero/hidden dimensions in the existing `FF036` visibility boundary while
  allowing a positive override to be compared as an ordinary sizing change.
  Normalize decimal/Boolean spelling, ordinary baseline defaults, the
  `customWidth` writer hint, inert thick-border flags under fixed custom heights,
  and equivalent layered column-range splitting. Malformed, duplicate,
  unsupported, or budget-exhausted dimension metadata is explicit coverage
  evidence; malformed dimensions are isolated only in the temporary ordinary
  reader copy so the original raw package remains auditable fail-closed.

## 0.55.0 — 2026-07-26

- Inspect material effective cell-border controls directly from raw
  transitional and strict SpreadsheetML before ordinary readers flatten style
  inheritance: reusable `<borders>/<border>` definitions, `borderId`, base
  `cellStyleXfs`, `xfId`/`applyBorder`, direct-cell, custom-row, and column
  assignments. The boundary covers ordinary edge sides, Office 2010 logical
  start/end sides, diagonals, outline, styles, and stored colours. Profiles,
  Markdown, JSON, and SARIF expose aggregate counts only; border definitions,
  colours, style IDs, and cell/row/column targets remain private.
- Emit `FF057` for a material effective cell-border control change and add the
  fail-closed `no_cell_border_changes` policy rule (`FFP057`). This closes the
  review gap where a report boundary, total, exception box, or warning cue can
  change while ordinary cell values and formulas stay fixed.
- Normalize omitted/`none` side declarations, Boolean and colour spellings,
  unused diagonal payload, ineffective empty `outline="false"`, base-XF
  inheritance, `applyBorder`, and equivalent effective column-range splitting.
  Missing, duplicate, malformed, or unsupported metadata is an explicit
  coverage warning. FormulaFence compares stored declarations only: it does
  not render Excel, resolve theme/palette colours, choose adjacent-cell border
  precedence, apply conditional-format/table/differential-style borders,
  calculate print layout, or infer client behavior.

## 0.54.0 — 2026-07-26

- Inspect material saved worksheet print-layout controls directly from raw
  transitional and strict SpreadsheetML before ordinary readers normalize them:
  `_xlnm.Print_Area` / `_xlnm.Print_Titles` definitions, print options,
  margins, page setup/fit-to-page, headers and footers, and manual row/column
  page breaks. Profiles, Markdown, JSON, and SARIF expose structural counts
  only; print ranges, header/footer text, page values, printer-setting
  references, and raw XML remain private.
- Emit `FF056` for a material worksheet print-layout control change and add the
  fail-closed `no_worksheet_print_layout_changes` policy rule (`FFP056`). This
  closes the review gap where an unchanged workbook can print a smaller,
  reordered, reframed, or differently labelled report.
- Normalize omitted/default, Boolean, integer, and decimal spellings; paired
  print-gridline flags; inactive first/even header-footer content; disabled
  first-page numbers; the fit-to-page versus percentage-scale selection; and
  automatic-break display noise. Missing, duplicate, malformed, or unsupported
  metadata is an explicit coverage warning. FormulaFence compares stored
  declarations only: it does not render or preview Excel, calculate page
  geometry/counts or automatic pagination, resolve printer/client defaults or
  `devMode` settings, or cover custom/legacy sheet-view and extension print
  controls.

## 0.53.0 — 2026-07-26

- Inspect material raw transitional and strict SpreadsheetML worksheet-display
  controls before ordinary readers normalize them: hidden zeroes, formula display,
  gridlines/custom gridline colour, row/column headers, outline symbols, rulers,
  page whitespace/margins, right-to-left layout, page-oriented view modes, and
  split/frozen panes. Profiles, Markdown, JSON, and SARIF expose
  structural counts only; sheet names, targets, and raw view XML remain private.
- Emit `FF055` for a material worksheet-display control change and add
  the fail-closed `no_worksheet_display_control_changes` policy rule
  (`FFP055`). This closes the review gap where unchanged values can
  appear blank, controls can be obscured, or the saved workbook surface can be
  materially reframed without an ordinary cell diff.
- Normalize omitted/default controls, Boolean and active custom-gridline-colour
  spellings, finite pane-split decimals, and ordinary selection/top-left/zoom
  navigation churn. Malformed or
  unsupported display metadata produces a visible coverage warning rather than a
  silent omission. FormulaFence compares stored declarations only: it does not
  render Excel, resolve an effective palette colour, calculate viewport geometry,
  decide final visibility, inspect print settings, or compose view controls with
  styles, objects, or client state.

## 0.52.0 — 2026-07-26

- Inspect effective raw SpreadsheetML cell-alignment controls before ordinary
  readers can flatten their style inheritance: horizontal/vertical placement,
  text rotation, wrapping, shrinking, indentation, relative indentation,
  justification, and reading order across default, direct-cell, row, and
  column-style assignments. Profiles, Markdown, JSON, and SARIF expose
  structural counts only; alignment values, style IDs, and target locations
  remain private.
- Emit `FF054` for a material effective cell-alignment control change and add
  the fail-closed `no_cell_alignment_changes` policy rule (`FFP054`). This
  closes the review gap where an unchanged value, warning, or classification can
  be moved, rotated, wrapped, shrunk, or indented without a normal cell diff.
- Normalize equivalent explicit defaults, Boolean and integer spellings,
  semantically inert `mergeCell` compatibility material, base-XF inheritance,
  `applyAlignment` semantics, and effective column-range splitting.
  Missing, duplicate, malformed, or unsupported alignment metadata produces a
  visible coverage warning rather than a silent omission. FormulaFence compares
  stored declarations only: it does not compute widths/heights, merge layout,
  overflow, final text visibility, font/fill/conditional-format composition,
  or Excel client rendering.

## 0.51.0 — 2026-07-26

- Inspect the raw workbook-level DrawingML Theme before ordinary workbook
  readers can reduce it to local style references: workbook-to-Theme bindings,
  transitional and strict Theme XML colour/font/format schemes, direct
  Theme-image relationships, and bounded direct image payloads. Profiles,
  Markdown, JSON, and SARIF expose aggregate counts only; Theme XML, scheme
  names, colour values, font names, image bytes, relationship IDs, and targets
  remain private.
- Emit `FF053` for a material stored workbook-Theme control change and
  add the fail-closed `no_workbook_theme_changes` policy rule
  (`FFP053`). This closes the review gap where a colour, font,
  effect, or direct Theme-image control can change a themed cell, chart, or
  drawing appearance while ordinary cells and local style references stay
  fixed.
- Normalize writer-selected Theme relationship IDs/order and equivalent
  internal target spelling. Missing, duplicate, malformed, unsafe, unbound,
  unreadable, oversized, or over-budget metadata emits a visible coverage
  warning; reads are bounded to 16 MiB per part, 64 MiB per workbook, and 512
  parts. FormulaFence does not resolve effective styles, render a workbook,
  calculate contrast, decode an image, fetch a target, calculate formulas, or
  infer Excel client behavior.

## 0.50.0 — 2026-07-26

- Inspect raw custom workbook data stores before ordinary workbook readers can
  omit them: generic Custom XML data and property/schema declarations,
  package/item relationships, workbook-bound Custom Data Properties and opaque
  binary Custom Data payloads, and custom document properties. Power Query
  `DataMashup` Custom XML remains exclusively under the existing Power Query
  controls. Profiles, Markdown, JSON, and SARIF expose aggregate counts only;
  custom XML, property names and values, storage IDs, binary payloads,
  relationship IDs, and targets remain private.
- Emit `FF052` for a material persisted custom-data-store change and add the
  fail-closed `no_custom_data_store_changes` policy rule (`FFP052`). This
  closes the review gap where add-in state, opaque binary data, or custom
  document properties can change while ordinary cells, formulas, and Power
  Query controls remain fixed.
- Normalize writer-selected relationship IDs/order and document-property
  `pid` values. Custom XML `itemID` and Custom Data `id` storage identities are
  compared privately because add-ins can bind state to them. Missing, duplicate,
  malformed, unsafe, unbound, unreadable, oversized, or over-budget metadata
  emits a visible coverage warning; reads are bounded to 16 MiB per part, 64
  MiB per workbook, and 512 parts. FormulaFence does not execute an add-in,
  resolve a property, follow or fetch a target, interpret a binary payload,
  calculate formulas, or infer Excel client behavior.

## 0.49.0 — 2026-07-24

- Inspect raw Excel rich-data controls before ordinary workbook readers can omit
  or normalize them: Rich Value Data, structures, types, arrays, supporting
  property bags/structures, styles, web images, rich-value relationships,
  workbook/package relationships, and `XLRICHVALUE` metadata/cell bindings.
  Profiles, Markdown, JSON, and SARIF expose aggregate counts only; entity
  values, provider data, field names, identifiers, URLs, image references,
  relationship IDs, and bound-cell locations remain private.
- Emit `FF051` for a material rich-data control change and add the fail-closed
  `no_rich_data_changes` policy rule (`FFP051`). This closes the review gap
  where provider-linked entities, attached data, or external image
  associations can change while ordinary cell values and formulas stay fixed.
- Normalize writer-selected relationship IDs/order and equivalent internal
  target spelling. Missing, duplicate, malformed, unsafe, unreadable,
  oversized, or over-budget metadata emits a visible coverage warning; reads
  are bounded to 16 MiB per XML part, 64 MiB per workbook, and 512 parts.
  FormulaFence does not contact providers, refresh data, calculate formulas,
  fetch endpoints, validate target content, or infer Excel client behavior.

## 0.48.0 — 2026-07-24

- Inspect raw OPC package-signature controls before ordinary workbook readers
  can omit or normalize them: package-root signature origins, origin-to-XML
  signature relationships, XMLDSIG envelope/reference material, embedded
  certificate values, certificate-part relationships/payloads, and conventional
  VBA project signature payloads/relationships
  (`vbaProjectSignature.bin`, Agile, and V3). Profiles, Markdown, JSON, and
  SARIF expose aggregate counts only; signature XML, reference URIs,
  certificate identities/contents, binary payloads, relationship IDs, and
  relationship targets remain private.
- Emit `FF050` for material package- or VBA-signature envelope changes and add
  the fail-closed `no_digital_signature_changes` policy rule (`FFP050`).
  This closes the review gap where provenance/integrity-assurance metadata can
  be added, removed, or altered while ordinary cell values, formulas, and
  `xl/vbaProject.bin` bytes remain unchanged.
- Normalize writer-selected relationship IDs/order, equivalent internal target
  spelling, and XMLDSIG base64 whitespace. Missing, duplicate, malformed,
  unsafe, unbound, unreadable, oversized, or over-budget metadata emits a
  visible coverage warning; reads are bounded to 16 MiB per part, 64 MiB per
  workbook, and 512 parts. FormulaFence inventories envelopes only: it does
  not validate cryptography, signed-reference coverage, certificate identity or
  trust, expiry, revocation, timestamps, signed contents, or VBA-code validity.

## 0.47.0 — 2026-07-24

- Inspect raw SpreadsheetML XML Maps, XML-table column properties, and
  single-cell XML table parts before ordinary workbook readers can discard or
  normalize the mapping surface. FormulaFence privately compares embedded
  schemas, map and data-binding refresh/export behavior, mapped XPath/table/
  cell declarations, and related workbook/worksheet relationship targets while
  profiles, Markdown, JSON, and SARIF expose aggregate counts only—never
  schemas, map names, XPath expressions, table identities, target cells,
  connection identities, or relationship targets.
- Emit FF049 for a material XML-mapped workbook control change and add the
  fail-closed `no_xml_mapping_changes` policy rule (FFP049). This closes the
  review gap where a data import/export template can be redirected or have its
  refresh behavior changed without an ordinary worksheet-cell diff.
- Normalize equivalent Boolean and unsigned-integer spelling,
  writer-selected relationship IDs/order, and equivalent internal target
  spelling. Missing, duplicate, malformed, unsafe, unbound, unreadable,
  oversized, or over-budget metadata produces a visible coverage warning;
  bounded raw reads use 16 MiB per part, 64 MiB per workbook, and 512 parts.
  FormulaFence does not import/export XML, validate data against schemas, open
  map bindings, fetch data, calculate refresh results, or infer client
  behavior.

## 0.46.0 — 2026-07-24

- Inspect raw Office 2010 `x14:sparklineGroups` worksheet extensions before
  ordinary workbook readers discard them. FormulaFence privately compares
  sparkline source/date-axis formulas, destination cells, group membership,
  type/axis/display/marker controls, line weight, and colour definitions while
  profiles, Markdown, JSON, and SARIF expose aggregate counts only—never
  source formulas, locations, control values, or colour definitions.
- Emit `FF048` for a material worksheet-sparkline control change and add the
  fail-closed `no_worksheet_sparkline_changes` policy rule (`FFP048`). This
  closes the review gap where a compact trend can be retargeted or restyled
  without changing ordinary worksheet values.
- Normalize equivalent direct local-range, Boolean/numeric, colour-case, and
  declaration-order spelling. Missing, duplicate, malformed, unreadable,
  oversized, or over-budget metadata produces a visible coverage warning. A
  Sparkline Group-removed temporary reader copy is made only after raw
  inspection, so reader loss cannot suppress evidence; FormulaFence does not
  calculate, render, resolve, or fetch sparkline sources or assess visual
  accessibility.

## 0.45.1 — 2026-07-24

- Stabilize the cross-version redaction test for an unknown custom number
  format by using a fixture-specific sentinel instead of a generic numeric
  substring. This does not change FormulaFence's workbook inspection or report
  surface.

## 0.45.0 — 2026-07-24

- Inspect raw standard SpreadsheetML and Office 2016 revision worksheet-cell
  hyperlink declarations before ordinary workbook readers can normalize their
  target bindings. FormulaFence privately compares cell/range binding, external
  and internal relationship semantics, location, display override, and
  ScreenTip material while profiles, Markdown, JSON, and SARIF expose aggregate
  counts only—never targets, cell references, locations, display strings,
  ScreenTips, relationship IDs, or revision UIDs.
- Emit `FF047` for a material worksheet-cell hyperlink control change and add
  the fail-closed `no_cell_hyperlink_changes` policy rule (`FFP047`). This closes
  the review gap where a familiar cell label can redirect a reviewer to a
  different URL, file, or in-workbook destination without changing the ordinary
  cell value.
- Normalize writer-chosen relationship IDs/revision UIDs, relationship ordering,
  and equivalent internal target spelling. Missing, duplicate, malformed,
  unsafe, unbound, unreadable, oversized, or over-budget metadata produces a
  visible coverage warning. The ordinary workbook reader receives a
  hyperlink-removed temporary copy only after raw inspection, so malformed
  package metadata cannot suppress evidence; FormulaFence does not render,
  resolve, fetch, follow, reputation-check, or execute a link, inspect linked
  content, or interpret `HYPERLINK()` formulas beyond the ordinary formula
  diff.

## 0.44.0 — 2026-07-24

- Inspect raw legacy Excel Note comments parts and their worksheet-bound VML
  Note shapes before workbook readers can omit author/text or display
  declarations. FormulaFence privately compares note text and rich-text
  presentation, author association, cell binding, comment properties, Note
  visibility/layout, and relationship semantics while profiles, Markdown, JSON,
  and SARIF expose aggregate counts only—never note text, authors, locations,
  VML, targets, relationship IDs, or GUIDs.
- Recognize the documented legacy Note placeholder that Excel can retain beside
  a modern threaded comment. A consistent placeholder GUID/author rekey stays
  quiet, while a changed placeholder reconciliation declaration remains guarded
  independently of the modern thread.
- Emit FF046 for a material legacy Note or threaded-placeholder change and add
  the fail-closed no_legacy_comment_changes policy rule (FFP046). This closes
  the review gap where an instruction, review context, or its visibility/layout
  can change outside ordinary worksheet cells.
- Bound raw comments/VML XML reads to 16 MiB per part, 64 MiB per workbook, and
  512 parts. Missing, duplicate, malformed, unsafe, external, unbound, or
  oversized metadata produces a visible coverage warning. The ordinary workbook
  reader receives a temporary Note-quarantined copy after raw inspection, so
  parser tolerance cannot suppress Note evidence; FormulaFence does not render
  Notes, resolve authors, fetch targets, execute linked content, or infer
  client/cloud state.

## 0.43.0 — 2026-07-24

- Inspect raw modern Excel threaded-comment and person package parts before
  workbook readers can omit their review annotations. FormulaFence compares the
  private comment/reply graph, stored text, cell binding, timestamp, resolution
  state, mention range/person binding, extension material, and author/mentioned
  person definitions while profiles, Markdown, JSON, and SARIF expose only
  aggregate counts—never comment text, locations, timestamps, names, user IDs,
  provider IDs, relationship IDs, or GUIDs.
- Emit `FF045` for a material threaded-comment control change and add the
  fail-closed `no_threaded_comment_changes` policy rule (`FFP045`). This closes
  the review gap where an assumption, instruction, or approval reply can change
  without changing any ordinary worksheet cell.
- Rebuild comment trees and person/mention links from their private identities
  so consistent writer-chosen comment, parent, person, mention, and package
  relationship-ID rewrites stay quiet. Missing, duplicate, unsafe, unbound,
  malformed, unreadable, oversized, or over-budget metadata becomes a visible
  coverage warning rather than a silent omission. FormulaFence compares stored
  package declarations only: it does not render comments, validate mention text
  offsets, send notifications, resolve accounts, fetch targets, or inspect
  legacy note/placeholder content.

## 0.42.0 — 2026-07-24

- Inspect raw non-chart Worksheet DrawingML regular shapes (`xdr:sp`) and group
  shapes (`xdr:grpSp`) before workbook readers can discard their text-box
  presentation. FormulaFence compares anchor/layout declarations, text and
  visual XML, group nesting, macro assignments, text links, and click/hover
  relationship semantics privately; profiles, Markdown, JSON, and SARIF expose
  structural counts only, never text, formatting, anchors, formulas, macro
  names, relationship IDs, or targets.
- Emit `FF044` for a material Worksheet DrawingML shape-control change and add
  the fail-closed `no_worksheet_drawing_shape_changes` policy rule (`FFP044`).
  This catches a text-box warning whose stored cell values and concatenated
  text remain unchanged while its presentation becomes less visible.
- Normalize writer-chosen non-visual shape IDs, relationship-ID rewrites, colour
  case, and relationship target spelling while retaining meaningful z-order and
  shape/group declarations. Missing, malformed, unsupported, oversized, or
  over-budget shape metadata becomes a visible parser-coverage warning rather
  than a silent omission. FormulaFence compares stored declarations only: it
  does not render DrawingML, resolve themes, evaluate text links, execute macro
  assignments, fetch targets, inspect arbitrary media, or claim coverage for
  pictures, connectors, graphic frames, SmartArt, or other non-`xdr:sp`
  drawing objects.

## 0.41.0 — 2026-07-24

- Inspect raw shared-string and inline-string character presentation that normal
  workbook readers reduce to concatenated text. FormulaFence compares rich
  `<r>/<rPr>` property sequences, styled character boundaries, and phonetic
  presentation material privately; profiles, Markdown, JSON, and SARIF expose
  structural counts only, never text, colours, fonts, indexes, or locations.
- Emit `FF043` for a material rich-text run control change and add the
  fail-closed `no_rich_text_run_changes` policy rule (`FFP043`).
  A formatting-only change such as making a warning phrase white is detected
  even when the normal cell value remains unchanged; an ordinary text-only edit
  within the same run-property sequence remains a normal semantic cell diff.
- Normalize rich-run property ordering, colour case, explicit false Boolean
  properties, and equivalent shared-versus-inline storage. Malformed,
  unsupported, missing, or unreadable rich-text metadata becomes a visible
  parser-coverage warning rather than a silent omission. FormulaFence compares
  stored declarations only: it does not render cells, resolve theme colours,
  calculate contrast, decide visibility, or guarantee Excel rendering.

## 0.40.0 — 2026-07-24

- Inspect raw SpreadsheetML formula-result caches alongside formula text. Cache
  values, error text, per-cell digests, and formula-cell locations remain only
  in private comparison entries; profiles, Markdown, JSON, and SARIF expose
  aggregate formula/cached/missing/result-type/malformed counts only.
- Emit `FF042` when a stored formula result changes without a changed formula at
  that cell or an ordinary changed cell that reaches it through the static
  dependency graph. Add the fail-closed
  `no_formula_cached_result_changes` policy rule (`FFP042`).
- Normalize equivalent finite numeric and Boolean result spellings, keep absent
  or blank caches visible as missing rather than inventing a result, and make
  malformed or unsupported cache metadata an explicit parser-coverage warning.
  FormulaFence does not calculate or validate results, distinguish a stale
  result from a tampered one, or model volatile, dynamic, external, or
  calculation-engine dependencies; a legitimate recalculation without a static
  visible precedent can therefore require review.

## 0.39.0 — 2026-07-24

- Extend the existing filter, sort, and row/column-visibility boundary to the
  documented zero-sized states that hide worksheet content without changing a
  cell: direct row `ht="0"`, column `width="0"`, and worksheet-default
  `defaultRowHeight="0"` / `defaultColWidth="0"` controls. Retain dimensions,
  row/column targets, and raw declarations only in private signatures; profiles,
  `FF036`, and SARIF expose structural counts only.
- Resolve worksheet-default zero dimensions before direct row and layered column
  declarations, so a later positive height or width is compared as an effective
  override while an equivalent inherited zero stays quiet. Positive ordinary
  resizes remain outside this narrow concealment boundary.
- Extend `FF036` / `no_filter_visibility_changes` (`FFP036`) without adding a
  new policy switch. Invalid, negative, non-finite, or application-out-of-range
  dimensions become explicit parser-coverage warnings rather than silent
  omissions. FormulaFence does not render widths/heights, infer overflow, or
  track arbitrary nonzero layout changes.

## 0.38.0 — 2026-07-24

- Inspect raw workbook cell-fill controls that can change what a reviewer sees
  without changing a stored cell value or formula: `<fills>` definitions,
  including patterned and gradient fills, base `<cellStyleXfs>`, effective
  `<cellXfs>`, direct cell `s`, `customFormat=1` row styles, and worksheet
  `<cols>/<col style>` defaults. Retain fill colours, pattern/gradient material,
  style IDs, and targets only in private signatures; profiles, `FF041`, and
  SARIF expose structural counts only.
- Resolve fill-ID remapping, base-XF inheritance, `applyFill`, valid
  pattern-colour child ordering, semantically inert no-fill/solid-background
  declarations, and equivalent column-range splitting. Record a column fill only
  as an OOXML default for unallocated/new cells rather than claiming to
  re-render allocated cells.
- Emit `FF041` for a material cell-fill-control change and add the fail-closed
  `no_cell_fill_changes` policy rule (`FFP041`). Invalid or missing fill/style
  references, malformed definitions, invalid targets, and bounded parser
  failures become explicit coverage warnings rather than silent omissions.
  FormulaFence does not resolve theme colours or rendering, calculate
  text/background contrast, apply conditional-format differential styles or
  table styles, or model borders, alignment, rich text, width/overflow, or
  arbitrary visual formatting.

## 0.37.0 — 2026-07-24

- Inspect raw workbook cell-font controls that can change what a reviewer sees
  without changing a stored cell value or formula: `<fonts>` definitions, base
  `<cellStyleXfs>`, effective `<cellXfs>`, direct cell `s`, `customFormat=1`
  row styles, and worksheet `<cols>/<col style>` defaults. Retain font names,
  colours, effects, style IDs, and targets only in private signatures;
  profiles, `FF040`, and SARIF expose structural counts only.
- Resolve font-ID remapping, base-XF inheritance, `applyFont`, equivalent font
  child ordering, common Boolean spellings, and equivalent column-range
  splitting. Record a column font only as an OOXML default for unallocated/new
  cells rather than claiming to re-render allocated cells.
- Emit `FF040` for a material cell-font-control change and add the fail-closed
  `no_cell_font_changes` policy rule (`FFP040`). Invalid or missing font/style
  references, malformed definitions, invalid targets, and bounded parser
  failures become explicit coverage warnings rather than silent omissions.
  FormulaFence does not resolve theme colours or rendering, model fills,
  borders, alignment, rich-text runs, table styles, width/overflow, or arbitrary
  visual formatting.

## 0.36.0 — 2026-07-24

- Inspect raw workbook number-format controls that can change what a reviewer
  sees without changing a stored cell value or formula: custom `<numFmt>`
  definitions, base `<cellStyleXfs>`, effective `<cellXfs>`, direct cell `s`,
  `customFormat=1` row styles, and worksheet `<cols>/<col style>` defaults.
  Retain format codes, style IDs, and targets only in private signatures;
  profiles, `FF039`, and SARIF expose structural counts only.
- Resolve custom-format ID remapping, base-XF inheritance, and
  `applyNumberFormat`; normalize equivalent custom-ID allocation, Boolean
  spelling, and effective column-range splitting. Record column styles as
  defaults for unallocated/new cells rather than claiming to re-render allocated
  cells.
- Emit `FF039` for a material number-format-control change and add the
  fail-closed `no_number_format_changes` policy rule (`FFP039`). Invalid or
  missing format/style references, conflicting custom definitions, invalid
  targets, and bounded parser failures become explicit coverage warnings rather
  than silent omissions. FormulaFence does not render locale-specific output,
  validate format syntax, calculate values, model width/overflow, or track
  arbitrary non-number-format visual styling.

## 0.35.0 — 2026-07-25

- Inspect raw worksheet `<cols>/<col>` visibility declarations for hidden,
  outlined, and collapsed columns without relying on a workbook reader that can
  flatten or lose compressed column ranges. Retain column positions and effective
  state only in private signatures; profiles, `FF036`, and SARIF expose safe
  counts only.
- Normalize Boolean/default and unsigned-integer spellings, equivalent range
  segmentation, and layered column declarations by applying later *present*
  visibility attributes in OOXML file order. A later width/style-only record
  does not erase an existing visibility state.
- Extend `FF036` / `no_filter_visibility_changes` (`FFP036`) to block effective
  hidden, outlined, or collapsed column changes alongside existing filter, sort,
  and row-visibility controls. Make malformed column bounds, attributes, child
  markup, and bounded-update exhaustion visible coverage warnings rather than
  silently omitting the affected controls. FormulaFence does not render
  outlines, apply filters, calculate results, track widths/styles, or interpret
  outline-display settings.

## 0.34.0 — 2026-07-25

- Inspect modern Excel Named Sheet Views from the documented relationship-backed
  worksheet parts, retaining view names, IDs, alternate filter criteria, target
  ranges, table bindings, table-column IDs, and sort keys only in private
  signatures—not profiles, `FF038`, or SARIF.
- Reconcile each stored filter to its base AutoFilter using Excel's documented
  UID, table-ID, then worksheet-owned fallback sequence. Normalize equivalent
  GUID, local A1 case/absolute-reference, Boolean/default, and unsigned-integer
  spellings while making a resolved target rebinding material to the diff.
- Emit `FF038` for a Named Sheet View definition, alternate filter/sort rule, or
  binding change; add the fail-closed `no_named_sheet_view_changes` policy rule
  (`FFP038`).
- Make missing, ambiguous, mismatched, malformed, unsupported, oversized, and
  unsafe relationship parts or filter bindings visible parser-coverage warnings
  rather than silently dropping them. FormulaFence does not activate/render a
  saved view, calculate a filtered result, infer formula visibility sensitivity,
  repair metadata, or interpret full differential-format, future extension, or
  rich-sort semantics.

## 0.33.0 — 2026-07-25

- Inspect standard worksheet `ignoredErrors` declarations and Office 2010
  `x14:ignoredErrors` extension declarations directly from raw OOXML. Private
  signatures retain target ranges and enabled warning types without serializing
  them into profiles, `FF037`, or SARIF.
- Emit `FF037` when suppressed Excel evaluation, inconsistent-formula,
  omitted-range, unlocked-formula, empty-reference, list-validation,
  calculated-column, text-number, or two-digit-year warning controls change;
  add the fail-closed `no_ignored_error_changes` policy rule (`FFP037`).
  Equivalent local A1 case/absolute-reference, Boolean, and target-order
  spellings are normalized.
- Make malformed or unsupported containers, extension material, attributes,
  flags, targets, and child markup visible parser-coverage warnings rather than
  silently dropping them. FormulaFence does not determine whether Excel would
  show a warning, calculate a formula, repair an error, or change application-
  level error-checking options.

## 0.32.0 — 2026-07-25

- Inspect worksheet and Table Definition-part AutoFilters directly from raw
  OOXML, including private filter criteria, selected values, filter-button
  state, AutoFilter sort state, and sort conditions. Also inspect explicit row
  `hidden` / outline state and the `sheetFormatPr@zeroHeight` hidden-by-default
  optimization without serializing criteria, sort keys/lists, table names, or
  row/range references into profiles, `FF036`, or SARIF.
- Emit `FF036` for a material filter, sort, or row-visibility control change,
  and add the fail-closed `no_filter_visibility_changes` policy rule (`FFP036`).
  Equivalent local A1 case/absolute-reference, Boolean/default, and unsigned
  integer spellings are normalized; filter-member ordering is canonicalized.
- Make malformed or unsupported declarations, extensions, and unsafe/missing
  table relationships visible parser-coverage warnings rather than silently
  dropping them. FormulaFence does not apply filters, calculate results,
  determine formula visibility sensitivity, render a report, or track hidden
  columns.

## 0.31.0 — 2026-07-25

- Inspect Excel Scenario Manager declarations directly from worksheet OOXML
  (`scenarios` / `scenario` / `inputCells`). Private signatures retain selected
  and shown state, summary references, names, protection flags, comments/users,
  changing-cell references, stored input values, deleted/undone state, and
  display number formats without serializing that material into profiles,
  `FF035`, or SARIF.
- Emit `FF035` for a material Scenario Manager definition or stored-input
  change, and add the fail-closed `no_scenario_manager_changes` policy rule
  (`FFP035`). Equivalent local A1 case/absolute-reference, Boolean, and
  unsigned-integer spellings plus schema-default false flags are normalized.
  Missing, malformed, duplicate-within-worksheet, or unsupported declarations
  are visible parser-coverage warnings rather than silently ignored.
- Treat Scenario Manager as worksheet-scoped: duplicate scenario names on
  different worksheets remain valid. Do not show/apply scenarios, calculate
  results, infer scenario-to-formula dependencies, or expose scenario names,
  comments, users, stored values, references, or raw XML.

## 0.30.0 — 2026-07-25

- Inspect Excel What-If Data Table masters directly from worksheet OOXML
  (`f t="dataTable"`). Private signatures retain the declared output range,
  one-/two-variable mode, orientation, input references, deleted-input flags,
  recalculation request, and supported generic formula metadata without
  serializing those controls into profiles, `FF034`, or SARIF.
- Stabilize the workbook reader's `DataTableFormula` representation, eliminating
  false self-diffs caused by process-local object addresses while preserving
  ordinary formula add/remove guards with a safe `=TABLE()` placeholder.
- Emit `FF034` for a material What-If Data Table definition or control change,
  and add the fail-closed `no_what_if_data_table_changes` policy rule
  (`FFP034`). Equivalent A1 case/absolute-reference and Boolean spellings are
  normalized. Missing, malformed, overlapping, or unsupported declarations are
  visible parser-coverage warnings rather than silently ignored.
- Do not calculate scenarios, infer a Data Table's output formula, predict
  recalculation results, or add Data Table inputs to the ordinary dependency
  graph. Cached scenario-output cells remain under the normal cell-diff
  boundary.

## 0.29.0 — 2026-07-24

- Inspect embedded Power Pivot/Data Model packages from the workbook's explicit
  `powerPivotData` relationship and `x15:dataModel` declaration. Private
  fingerprints retain complete declaration material, normalized relationship
  semantics, and bounded raw model payload hashes without serializing table,
  column, relationship, connection, DAX, stored-value, target, XML, or payload
  content.
- Emit `FF033` for a Data Model binding, declaration, direct model-part
  relationship, or bounded raw payload change, and add the fail-closed
  `no_power_pivot_data_model_changes` policy rule (`FFP033`). Relationship IDs,
  equivalent internal target spellings, and writer-generated Data Model GUIDs
  are normalized away. Missing, malformed, orphaned, unbound, externally
  targeted, unexpected directly related, oversized, over-budget, or
  unrecognized material remains a visible coverage warning. Raw model payload
  reads are bounded to 512 MiB per part, 512 MiB per workbook, and 16 parts.
- Never deserialize the embedded Analysis Services payload, evaluate DAX,
  refresh a model, calculate/render a report, infer model-to-cell impact, or
  fetch an external target.

## 0.28.0 — 2026-07-24

- Inspect Slicer and Timeline cache definitions directly from documented
  workbook extension declarations and explicit workbook relationships. Private
  fingerprints retain Slicer item selections, Timeline state/filter material,
  PivotTable/table source bindings, filtered-PivotTable bindings, normalized
  relationships, and complete cache definitions without serializing cache
  names, source fields, selected values, date ranges, PivotTable names,
  relationship targets, or XML.
- Emit `FF032` for a Slicer/Timeline workbook binding, filter state, cache
  definition, source binding, filtered-PivotTable binding, or unexpected direct
  cache-part relationship change, and add the fail-closed
  `no_slicer_timeline_cache_changes` policy rule (`FFP032`). Relationship IDs,
  equivalent internal target spellings, coordinated Slicer/Timeline PivotCache
  extension-ID renumbering,
  known optional Slicer defaults, Boolean spellings, and Timeline GUIDs are
  normalized away. Malformed, orphaned, unbound, externally targeted,
  oversized, over-budget, or unrecognized material remains a visible coverage
  warning. Cache XML reads are bounded to 16 MiB per part, 64 MiB per workbook,
  and 512 parts.
- Treat the documented 2010 Timeline-cache relationship and the widely emitted
  2011 compatibility relationship as one equivalent workbook binding.
- Never apply a Slicer or Timeline filter, calculate/render a PivotTable or
  table, infer downstream cell impact, fetch an external target, or model
  worksheet/drawing view geometry and styles.

## 0.27.0 — 2026-07-24

- Inspect PivotTable view definitions, cache schemas, shared cache items, and
  bounded raw cache-record payloads directly from the documented workbook-cache
  and worksheet-PivotTable OOXML relationships. Private fingerprints retain
  layouts, cache material, normalized relationships, and record hashes without
  serializing names, source ranges, field/item values, formulas, cache records,
  relationship targets, XML, or payload bytes.
- Emit `FF031` for PivotTable bindings/layouts, cache definitions, shared items,
  cache-record relationships, or cache-record payload changes, and add the
  fail-closed `no_pivot_table_definition_changes` policy rule (`FFP031`). Cache
  source and refresh settings deliberately remain with `FF023`. Relationship
  IDs, equivalent internal target spellings, and cache-ID renumbering are
  normalized away; malformed, orphaned, unbound, oversized, over-budget, or
  unrecognized material remains a visible coverage warning. PivotTable and
  cache-definition XML reads are bounded to 16 MiB per part, 64 MiB per
  workbook, and 512 parts; raw cache-record hashes are bounded to 32 MiB per
  part, 64 MiB per workbook, and 512 parts.
- Detach cache-record relationships only in a temporary reader copy before the
  underlying workbook library loads cells, so it does not eagerly materialize
  unbounded record streams. The original workbook is never modified.
- Never refresh or calculate a PivotTable, render a report, infer
  PivotTable-to-cell impact, fetch an external target, parse cached record
  values, or interpret OLAP, extension-list, or slicer semantics.

## 0.26.0 — 2026-07-24

- Inspect DrawingML chart definitions, cached series presentation data, and
  chart `userShapes` overlays directly through standard worksheet/chartsheet
  drawing relationships. Private fingerprints retain chart definition and cache
  material separately, normalized relationships, overlays, and bounded direct
  related-part payload hashes without serializing formulas, cached values,
  titles, shape text, relationship targets, XML, or payload bytes.
- Emit `FF030` for chart bindings, definitions, cached data, overlays,
  relationships, or direct related-payload changes, and add the fail-closed
  `no_chart_definition_changes` policy rule (`FFP030`). Writer-chosen
  relationship IDs and equivalent internal target spellings are normalized
  away; malformed, orphaned, unbound, oversized, over-budget, or unrecognized
  material remains a visible coverage warning. Chart and overlay XML reads are
  bounded to 16 MiB per part, 64 MiB per workbook, and 512 parts; direct related
  payload hashes are bounded to 32 MiB per part, 64 MiB per workbook, and 512
  parts.
- Never calculate a series formula, render a chart, infer chart-to-cell impact,
  follow an external target, parse media or embedded-package formats, or
  interpret modern `chartEx`/nested-chart semantics.

## 0.25.0 — 2026-07-24

- Extend the existing worksheet-control guardrail to legacy VML form controls.
  FormulaFence now follows standard worksheet `vmlDrawing` relationships and
  privately fingerprints non-`Note` VML `ClientData` control material, including
  macro assignments, cell/range bindings, camera source ranges, and directly
  referenced VML-part relationship semantics. Ordinary VML comment notes are
  deliberately excluded from the control inventory.
- Keep the `FF029` / `FFP029` contract while making it cover modern worksheet
  controls, legacy VML controls, and OLE objects together. Relationship IDs and
  equivalent internal target spellings remain normalized; malformed, orphaned,
  missing, oversized, and over-budget VML material remains a visible coverage
  warning. VML XML shares the existing 16 MiB-per-part, 64 MiB-per-workbook,
  512-part control XML budget.
- Never render a VML drawing, read comment text into the control profile, execute
  a macro, evaluate a binding, or open a relationship target. Macro names,
  formulas/ranges, captions, relationship targets, and VML XML remain private.

## 0.24.0 — 2026-07-24

- Inspect relationship-backed worksheet ActiveX, form-control, and OLE-object
  chains directly from raw OOXML before the workbook reader can omit them.
  Private fingerprints retain worksheet declarations, control configuration,
  ActiveX persistence XML, form-control properties, relationships, and bounded
  direct ActiveX/OLE/package payload hashes without serializing control names,
  class IDs, licenses, macro assignments, formulas/ranges, OLE identities,
  relationship targets, XML, or payload bytes.
- Emit `FF029` for worksheet-control bindings, definitions, ActiveX/form-control
  material, OLE configuration, related relationships, or direct payload changes,
  and add the fail-closed `no_worksheet_embedded_control_changes` policy rule
  (`FFP029`). Writer-chosen relationship IDs and equivalent internal target
  spellings are normalized away. Malformed, orphaned, unbound, oversized, or
  over-budget material remains a visible coverage warning. XML reads are bounded
  to 16 MiB per part, 64 MiB per workbook, and 512 parts; raw direct payload
  hashes are bounded to 32 MiB per part, 64 MiB per workbook, and 512 parts.
  FormulaFence never initializes an ActiveX control, deserializes or opens an OLE
  object/package, renders its drawing surface, follows an external target, or
  infers event dispatch.

## 0.23.0 — 2026-07-24

- Inspect document-linked Office Web Add-in task-pane packages directly from
  their workbook declarations, through task-pane bindings and direct
  web-extension definitions. Private fingerprints retain task-pane
  configuration, add-in references, auto-show properties, bindings, snapshots,
  and relationship material without serializing add-in IDs, store references,
  property/binding values, XML, snapshot data, or relationship targets.
- Emit `FF028` for Office Web Add-in task-pane workbook bindings,
  configuration, definitions, or relationships and add the fail-closed
  `no_office_web_addin_changes` policy rule (`FFP028`). Writer-chosen
  relationship IDs and equivalent internal target spellings are normalized
  away, while malformed, oversized, unbound, or otherwise unrecognized parts
  remain visible coverage warnings. Task-pane and web-extension XML reads are
  bounded to 16 MiB per part, 32 MiB per workbook, and 64 parts. FormulaFence
  never installs, loads, executes, or fetches an add-in or manifest, follows an
  external target, or models worksheet-scoped Web Add-in markup outside this
  task-pane chain.

## 0.22.0 — 2026-07-24

- Inspect Office RibbonX custom-UI package parts directly from their root
  declarations, including the documented 2006 and Office 2010-era package
  forms. Private fingerprints retain complete custom-UI XML and direct
  relationship material without serializing control IDs, labels, callback
  names, image targets, or XML content.
- Emit `FF027` for RibbonX package, callback/control, or relationship changes
  and add the fail-closed `no_ribbon_customization_changes` policy rule
  (`FFP027`). Writer-chosen relationship IDs and equivalent internal target
  spellings are normalized away, while malformed, oversized, unbound,
  version-mismatched, or otherwise unrecognized parts remain visible coverage
  warnings. Reads are bounded to 16 MiB per part, 32 MiB per workbook, and
  eight parts. FormulaFence never invokes RibbonX callbacks, follows external
  targets, or parses image payloads.

## 0.21.0 — 2026-07-24

- Extend the XLM macro-sheet control boundary to direct internal related parts,
  including embedded OLE objects and packages. FormulaFence streams those raw
  bytes into private fingerprints without parsing or serializing payload
  contents; a payload-only change now remains a critical `FF026` finding.
- Bound that work to 32 MiB per related part, 64 MiB across a workbook, and
  256 parts. Missing, unreadable, oversized, or over-budget targets surface as
  safe inventory counts and parser-coverage warnings rather than silent gaps.

## 0.20.0 — 2026-07-24

- Inspect Excel 4.0 / XLM Macro Sheet and International Macro Sheet package
  parts directly, before the workbook reader can omit their executable cells.
  Private fingerprints retain complete macro XML, workbook bindings, and
  related package relationships without serializing commands, cell values,
  targets, identifiers, or embedded-object payloads.
- Emit `FF026` for XLM macro-sheet changes and add the fail-closed
  `no_xlm_macro_sheet_changes` policy rule (`FFP026`). Relationship-id-only
  rewrites are normalized away, while malformed, unbound, oversized, or
  unrecognized parts remain visible coverage warnings. FormulaFence does not
  execute or emulate XLM commands, resolve targets, or load embedded objects.

## 0.19.0 — 2026-07-24

- Inspect raw `xl/externalLinks/externalLink*.xml` parts for external-workbook,
  DDE, and OLE link definitions. Private fingerprints retain declaration-to-part
  bindings, endpoint relationships, definition material, cached values,
  item behavior, and unmodelled XML without serializing targets, names, source
  data, or extension payloads.
- Emit `FF025` for external-link package changes and add the fail-closed
  `no_external_link_package_changes` policy rule (`FFP025`). FormulaFence does
  not follow or execute external-workbook, DDE, or OLE links, establish source
  trust, or infer returned data.

## 0.18.0 — 2026-07-24

- Inspect Power Query Data Mashup custom XML parts directly: private fingerprints
  cover the embedded `Section1.m` formula document, logical package material,
  stable query metadata, and formula-firewall permissions without serializing
  M text, query/source names, metadata values, embedded content, telemetry IDs,
  or user-bound permission bindings.
- Ignore documented refresh-result metadata and `sqmid` telemetry noise while
  preserving high-signal query-definition and execution-control changes. Query
  tables linked through normal Excel table relationships are now inventoried in
  addition to directly worksheet-linked tables.
- Emit `FF024` for changed Power Query formulas or semantic controls and add
  the fail-closed `no_power_query_changes` policy rule (`FFP024`). FormulaFence
  does not execute M, refresh connections, assess sources, or inspect DDE/OLE
  links or full PivotTable layout semantics.

## 0.17.0 — 2026-07-24

- Inventory external-data refresh controls directly from OOXML: workbook-wide
  external-link/refresh flags, data connections, linked query tables, and
  pivot-cache sources and refresh behavior.
- Normalize schema defaults and preserve source paths, connection strings,
  query material, identifiers, names, descriptions, parameter values, cached
  records, and opaque extension XML as private comparison fingerprints rather
  than report content.
- Emit `FF023` for changed external-data connection or refresh controls and add
  the fail-closed `no_external_data_connection_changes` policy rule (`FFP023`).
  FormulaFence does not execute connections, refresh data, or parse Power Query
  M, DDE/OLE links, or full PivotTable layout semantics.

## 0.16.0 — 2026-07-24

- Inventory operational protection controls directly from OOXML: workbook
  structure/windows/revision locks; worksheet, dialog-sheet, and chart-sheet
  permissions; protected range targets; and compact direct cell/row/column
  locked/hidden assignments on active protected sheets.
- Normalize worksheet action defaults so omitted and explicit OOXML spellings
  compare equal. Preserve unmodelled protection metadata through private
  fingerprints, and compare legacy/modern verifier material, protected-range
  names, and security descriptors without serializing any of their values.
- Emit `FF022` for changed protection controls and add the fail-closed
  `no_protection_changes` policy rule (`FFP022`). Protection remains an
  operational review boundary, not file encryption, authentication, or a claim
  to reproduce Excel's complete style cascade.

## 0.15.0 — 2026-07-24

- Inventory worksheet conditional-formatting controls directly from OOXML:
  compact target ranges, global precedence, criteria, rule flags, differential
  styles, color scales, data bars, icon sets, and retained extension fragments.
- Resolve differential styles rather than comparing unstable `dxfId` values;
  normalize schema boolean defaults, leading `=` criteria, priority-number
  gaps, and extension GUID links to avoid writer-only control diffs. Profiles
  redact criteria, text rules, and raw style/extension XML.
- Emit `FF021` for changed conditional-formatting controls and add the
  fail-closed `no_conditional_formatting_changes` policy rule (`FFP021`). The
  tool records display-control semantics but does not calculate Excel's final
  conditional formatting result.

## 0.14.0 — 2026-07-24

- Inventory worksheet data-validation controls as compact target ranges and
  compare their effective type, operator, criteria, blank/dropdown behavior,
  prompts, error alert, IME mode, and worksheet-level prompt-disable setting.
  This does not expand full-column validations into cells.
- Normalize omitted OOXML defaults (`none`, `between`, `stop`, and `noControl`)
  and an optional leading `=` in criteria, so equivalent Excel-compatible
  writers do not create a control-only diff. Identical controls are also joined
  when a writer splits their target groups. Profiles redact criteria and
  prompt/error text; local reports retain full before/after evidence.
- Emit `FF020` for changed validation controls and add the fail-closed
  `no_data_validation_changes` policy rule (`FFP020`). Validation expressions
  remain inspectable controls, not calculations FormulaFence evaluates.

## 0.13.0 — 2026-07-24

- Trace formulas that read a non-anchor member of a dynamic array's current
  OOXML output range. The compact anchor-to-consumer edge makes an input of a
  dynamic array reach its current direct and range consumers without expanding
  the spill into virtual cells.
- Preserve the safety boundary: dynamic output ranges are profiled as observed,
  not fixed, because recalculation can grow, shrink, or block a spill. Profiles
  list the observed range and every linked output-member consumer.
- Emit `FF019` when a formula newly intersects an observed non-anchor dynamic
  output member, including when a changed observed extent reaches an unchanged
  formula. Add `no_new_dynamic_array_output_references` (`FFP019`) as a
  fail-closed policy control.

## 0.12.0 — 2026-07-24

- Trace fixed legacy CSE array output members without expanding their declared
  ranges: an input of an array anchor now reaches ordinary formulas that read
  non-anchor result cells, including cross-sheet and range consumers.
- Inspect raw OOXML dynamic-array metadata to keep the boundary safe. Dynamic
  anchors are inventoried but never receive aliases for a current spill extent;
  unrecognized array metadata becomes a visible coverage warning instead of a
  guessed fixed CSE graph.
- Compare array-formula execution mode independently of formula text and emit
  `FF018` when a legacy-CSE or dynamic formula is added, removed, or changes
  mode, or when a legacy CSE fixed output range changes. Add the fail-closed
  `no_array_formula_semantics_changes` policy rule (`FFP018`).

## 0.11.0 — 2026-07-24

- Trace the exact selected cell for direct static A1 implicit intersection,
  including literal `@A1:A3` and persisted OOXML `_xlfn.SINGLE(A1:A3)` forms.
  Other explicit intersection expressions retain conservative static input
  edges instead of being evaluated.
- Inventory explicit implicit intersection in profiles, emit `FF017` for new
  uses, and add the fail-closed `no_new_implicit_intersections` policy rule
  (`FFP017`). Formula-defined names containing this context-dependent behavior
  remain unresolved at call sites.
- Normalize direct display and OOXML spellings of `#`/`ANCHORARRAY` and
  `@`/`SINGLE` in formula fingerprints to avoid a serialization-only formula
  diff.

## 0.10.0 — 2026-07-24

- Trace the static anchor behind direct internal `A1#` spilled-array references
  and OOXML-style `ANCHORARRAY(A1)` calls without evaluating Excel.
- Inventory spill-reference consumers in profiles and report new instances as
  `FF015`; add `no_new_spill_references` (`FFP015`) for a fail-closed CI
  boundary. Dynamic spill extent and blocking cells remain explicit limits.
- Surface formula-tokenization failures at workbook level instead of silently
  omitting their graph. New failures emit `FF016` and can be blocked with
  `no_new_tokenization_failures` (`FFP016`).
- Keep formula-defined names containing a spill reference unexpanded, so a
  named formula cannot hide the dynamic boundary behind inferred dependencies.

## 0.9.0 — 2026-07-24

- Expand calls to workbook- and worksheet-local defined names whose complete
  definitions are statically resolvable `LAMBDA` expressions, including nested
  named-LAMBDA calls and formula-defined names that call them.
- Preserve Excel name scope and local-name precedence for callable definitions.
  Recognize OOXML formula definitions without a leading `=` and serialized
  `_xlfn.LAMBDA`, `_xlpm.`, and `_xlop.` local-name forms.
- Leave dynamic, relative, cyclic, external, 3-D, tokenizer-unsupported, and
  otherwise non-static named LAMBDAs visible as unresolved references at their
  call sites rather than creating guessed graph edges.

## 0.4.0 — 2026-07-24

- Resolve the conservative, fully qualified Excel-table subset: table names,
  single or contiguous column ranges, and `#All`/`#Data`/`#Headers`/`#Totals`
  regions.
- Inventory table metadata in profiles and flag table additions, removals, and
  definition changes as `FF013`; add the `no_table_definition_changes` policy
  control (`FFP013`).
- Keep this-row (`@`) and complex table syntax as explicit unresolved coverage
  rather than inferring a dependency that static inspection cannot justify.

## 0.3.0 — 2026-07-24

- Resolve ordinary workbook and sheet-local defined names with static A1
  destinations into the dependency graph, including explicit references to
  sheet-local names and names that resolve to external workbooks.
- Make static-analysis coverage visible: profiles identify unresolved range
  tokens and dynamic `INDIRECT`/`OFFSET` formulas; diffs report new instances
  as `FF011` and `FF012`.
- Add opt-in `no_new_unresolved_references` and `no_new_dynamic_references`
  policy controls (`FFP011`, `FFP012`) for teams that need to fail closed on
  new dependency-coverage gaps.

## 0.2.0 — 2026-07-24

- Add deterministic shortest dependency-path samples to every changed cell's
  impact record, FormulaFence hazard finding metadata, Markdown reports, and
  SARIF properties.
- Include the same path evidence when an impact-limit policy fails.

## 0.1.1 — 2026-07-24

- Capture workbook-parser warnings as structured profile coverage notes instead
  of writing raw dependency warnings to the console.
- Flag newly introduced parser coverage warnings in diffs (`FF010`) and support
  the `no_new_parser_warnings` policy control.
- Validate the profile path against a public 18-sheet financial cap-table model;
  see [validation notes](docs/validation.md).

## 0.1.0 — 2026-07-24

- Initial public release: formula-aware semantic diffing, explicit dependency
  impact, workbook-control checks, policy-as-code, and Markdown/JSON/SARIF output.
