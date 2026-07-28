# External validation notes

FormulaFence's test suite builds small fixtures that isolate individual risks.
Those tests are necessary but insufficient for confidence in an Office-file
reader, so each release should also be exercised on independently maintained
workbooks without copying their contents into this repository.

## Dependency-graph edge budget — 2026-07-28

Version 0.171.0 closes a retained-graph amplification gap after the bounded
reader has finished. FormulaFence stores compact local direct and range
dependency indexes for downstream impact review. A source can remain small and
stay within the populated-cell ceiling while one formula-defined name resolves
to many local inputs at every caller, multiplying those retained indexes.

`profile`, `diff`, and `check` now give each workbook input a positive-only
`--max-dependency-edges` budget of 2,000,000 by default. `portfolio` uses one
independent pool for all successfully retained baseline snapshots and another
for candidate snapshots. A direct local dependency, a compact local range
dependency, and each additional fixed-CSE or observed dynamic-array output
alias consumes one record. The boundary intentionally does not expand a range
into cells and does not replace the candidate-only cross-workbook graph or
local impact-analysis limits. An overage returns status 2 before a CLI artifact
can be rendered or published.

For an independent compact stress input, a 57,243-byte valid `.xlsx` defined
`Fanout` as a 100-input local `SUM` name and used `=Fanout` in 10,000 caller
cells. This creates exactly 1,000,000 retained reverse-dependency edges without
any range or array-alias records. The public 0.170.0 wheel wrote its normal
39,844-byte JSON profile in 6.081 seconds. The 0.171.0 candidate rejected a
999,999 edge budget in 6.334 seconds with exit 2 and no output path. At exactly
1,000,000 it completed in 6.254 seconds and produced byte-identical JSON
(SHA-256 `9ceade979ecbf37a36d17a4b18693fc0daa988da1b5dfcbb6bbef2454a8abdc6`).
Focused regressions also prove the 12-edge named-formula boundary, direct CLI
no-output behavior, legacy-CSE and dynamic-array alias accounting, independent
portfolio-side sharing with stop-before-later-read behavior, nonpositive API
and Action-input rejection, and Action propagation. The final source suite
passed 1,282 tests in 211.56 seconds; Ruff, bytecode compilation,
Action-shell syntax, and whitespace checks were clean.

## Profile inventory-record budget — 2026-07-28

Version 0.170.0 closes the retained-profile-state gap left by the artifact byte
boundary. A profile deliberately emits safe location and coverage evidence, but
the 32 MiB `--max-report-bytes` check occurs only after FormulaFence has built
that public Python object graph. A compact valid source can therefore force a
large profile inventory even when no artifact is ultimately published.

`profile` now defaults `--max-profile-records` to 100,000. Before constructing
the profile, it counts every serialised list item: top-level sheet/control
inventories, nested table columns and ranges, parser warnings, per-location
token/function lists, and dynamic-array reference entries. The command returns
status 2 before profile construction and before any output file exists when the
aggregate exceeds the budget. It is separate from the reader's source/snapshot
limits and from rendered bytes; a reviewer can deliberately choose a larger
positive count for a complete known inventory.

On the valid 2,590,768-byte workbook with 500,000 repeated `INDIRECT` formulas
used for the 0.169 artifact boundary, version 0.170.0 rejected the default
100,000-record budget in 418.596 seconds with no output. The known complete
profile contains exactly 1,000,007 list records: 500,000 location entries,
500,000 function entries, six parser warnings, and one sheet record. Supplying
that exact record count together with the known 65,429,579-byte render budget
completed in 434.598 seconds and produced a byte-for-byte match to the public
0.168.0 JSON (SHA-256
`862c1870e9307164749e9b6f4a24a82f47dc606f537f971a997a98439bdc7597`).
Focused regression coverage proves exact capacity, nonpositive rejection,
preflight before any location rendering, CLI default propagation, and no output
publication on overage. The final source suite passed 1,274 tests in 389.76
seconds with Ruff, bytecode compilation, Action-shell syntax, and whitespace
checks clean. The final wheel and source distribution passed `twine check`; both
fresh isolated installs passed `pip check`, wrote a normal JSON profile, and
rejected a one-record JSON profile without publishing an output path.

## Profile artifact byte budget — 2026-07-28

Version 0.169.0 completes the rendered-artifact boundary for the remaining CLI
path. A profile intentionally publishes safe structural and coverage metadata,
including one location/function record for each dynamically referenced formula.
Those records remain useful review evidence, but a compressed source at the
allowed reader cell count can otherwise serialize into a much larger JSON or
Markdown file after bounded workbook inspection has already completed.

`profile` now accepts the same positive `--max-report-bytes` control as the
comparison commands and defaults to 33,554,432 UTF-8 bytes (32 MiB). JSON uses
the incremental encoder; Markdown streams every line through the shared byte
counter. An overage returns status 2 before FormulaFence writes or replaces the
requested output. This is an artifact-publication boundary, not a reduction of
the separate bounded profile/snapshot state retained while inspection runs; a
reviewer can deliberately opt up for a complete known-size artifact.

For a valid 2,590,768-byte `.xlsx` containing exactly 500,000 repeated
`INDIRECT` formulas on a 28-character sheet name, the public 0.168.0 wheel
completed in 6m44.010s and wrote a 65,429,579-byte JSON profile. The 0.169.0
candidate default failed closed in 7m06.956s with no report at the 32 MiB
ceiling. Giving it exactly 65,429,579 bytes completed in 6m59.999s and wrote a
byte-for-byte identical profile (SHA-256
`862c1870e9307164749e9b6f4a24a82f47dc606f537f971a997a98439bdc7597`).
The focused profile/rendering/CLI regression set passed 12 tests in 3.31
seconds. The final source suite passed 1,272 tests in 356.55 seconds, with
Ruff, bytecode compilation, Action-shell syntax, and whitespace checks clean.
The final wheel and source distribution passed `twine check`; fresh isolated
installs of both passed `pip check`, wrote a normal JSON profile, and rejected a
one-byte Markdown profile budget without publishing an output path.

## Rendered report byte budget — 2026-07-28

Version 0.168.0 closes the remaining artifact-amplification gap after workbook
reading and semantic comparison have been bounded. A small compressed OOXML
source can legitimately contain many repeated long values; a complete JSON
report records before/after evidence and can be vastly larger than the package
that caused it. Building one enormous JSON, Markdown, HTML, or SARIF string
also risks converting a bounded review into an impractical CI-memory and
artifact-upload task.

`diff`, `check`, and `portfolio` now cap the final UTF-8 review artifact at
33,554,432 bytes (32 MiB) by default with `--max-report-bytes` / Action
`max-report-bytes`. JSON and SARIF count incremental encoder chunks; Markdown
streams each line; HTML writes each escaped review entry through one
shared budget. An overage returns status 2 before FormulaFence writes or
replaces the requested output path. The control is intentionally independent
from workbook input, local impact, portfolio snapshot, and cross-workbook graph
budgets; a reviewer can opt up with a deliberate positive byte value.

Direct regressions prove exact UTF-8 JSON accounting and fail-closed rendering
for JSON, Markdown, HTML, and SARIF. CLI and portfolio commands prove no output
file appears after an overage, and the public Action metadata/default,
invalid-input, and propagation contracts are covered.

For a compact control of two 78,679-byte / 78,683-byte `.xlsx` files containing
1,000 changed 32,767-character text cells, the public 0.167.0 wheel completed
in 32.719226 seconds and wrote a 66,287,419-byte JSON report. The 0.168.0
candidate default failed closed in 32.398541 seconds with no report at its
33,554,432-byte ceiling. Giving it exactly 66,287,419 bytes completed in
32.726499 seconds and wrote exactly 66,287,419 bytes. The focused output,
CLI, portfolio, and Action contract set passed 10 tests in 2.65 seconds. The
final source suite passed 1,269 tests in 369.61 seconds, with Ruff, bytecode
compilation, Action-shell syntax, and whitespace checks clean.

## Aggregate local impact-analysis budget — 2026-07-28

Version 0.167.0 closes an algorithmic multiplication gap in semantic diffing.
FormulaFence already capped one changed cell's local downstream walk, but a
workbook can contain many independently changed sources. Applying the same
per-source ceiling to every edit could still multiply CPU and retained impact
evidence into an impractical CI run, even when the workbook reader and each
individual traversal stayed within their own limits.

`diff` and `check` now share a 100,000-state local-analysis pool by default;
`portfolio` shares one pool across every matched workbook. A changed source and
each statically reachable local dependent consume one state. The
`--max-change-analysis-states` / Action `max-change-analysis-states` control is
positive-only, and an overage returns status 2 before a partial local-impact
artifact can be written. The pre-existing per-source walk remains an independent
coverage limit, and the candidate-only cross-workbook graph retains its separate
`--max-link-impact` pool. Shortest paths are now reconstructed lazily only for
the fixed serialized sample, avoiding eager quadratic materialization of every
prefix in a long chain while retaining normal mapping access for callers that
need a particular path.

Direct regressions prove that two independent two-state source/dependent paths
pass at an exact four-state bound and fail at three; a portfolio proves that its
single budget is shared across separate matched workbooks. CLI/default and
Action metadata, invalid-input, and propagation cases cover the public
contract.

For a compact fan-out control with 250 changed inputs and an 800-formula chain,
the public 0.166.0 wheel completed all 200,250 source-to-reachable states,
produced 250 changes in a 2,310,669-byte JSON report, and took 12.662560
seconds. The 0.167.0 candidate with a 10,000-state budget failed closed in
2.436415 seconds without publishing a report; its default 100,000-state budget
also failed closed in 3.358328 seconds. Five no-change controls averaged
2.379241 seconds in the public wheel and 2.455888 seconds in the candidate.
The final source suite passed 1,260 tests in 382.03 seconds, with Ruff,
bytecode compilation, Action-shell syntax, and whitespace checks clean.

## Aggregate portfolio snapshot-cell budget — 2026-07-28

Version 0.166.0 closes the retained-semantic-state gap left after the source-byte
preflight. A portfolio report retains immutable workbook snapshots for its
nested diff evidence, and candidate snapshots also remain available for the
cross-workbook dependency graph. The individual 500,000-cell reader ceiling
therefore still allowed a 512-workbook directory to accumulate an impractical
number of populated cell records even when every source package was valid and
the aggregate compressed input stayed within its own limit.

Each baseline and candidate side now has an independent 2,000,000-populated-cell
budget, configurable with `--max-portfolio-snapshot-cells` and the matching
Action input. Actual populated cells are knowable only after one source is read,
so FormulaFence records the new immutable snapshot immediately and fails closed
before it opens a later workbook if that side has crossed its total. This keeps
the per-source reader boundary intact while bounding the state retained across
the portfolio; it is deliberately separate from raw-entry, source-byte,
supported-workbook, and cross-workbook graph limits.

Direct regressions prove an exact two-workbook total succeeds, a baseline
overflow stops before the corresponding later candidate source is opened, and a
larger candidate snapshot also fails before a later source can be read. The CLI
and public Action reject nonpositive inputs, default to the documented limit,
and propagate an aggregate overage as status 2.

For a controlled six-workbook portfolio of compact 20,000-populated-cell files
(945,826 total source bytes), the public 0.165.0 wheel completed every snapshot
in 89.803903 seconds at 105,888 KiB maximum RSS. The 0.166.0 candidate with a
20,000-cell per-side limit stopped on the second baseline snapshot at 40,000
cells, returned status 2 in 47.647901 seconds at 82,872 KiB, and did not open a
later source. The final source suite passed 1,252 tests in 369.50 seconds, with
Ruff, bytecode compilation, Action-shell syntax, and whitespace checks clean.

## Aggregate portfolio source budget — 2026-07-28

Version 0.165.0 closes the next portfolio-scale resource gap. FormulaFence
already limited each source package to 1 GiB, but the 512-workbook directory
ceiling still allowed one supplied side to schedule up to 512 GiB of source
bytes for private copies and semantic readers. A broad generated-model
portfolio could therefore remain impractical despite every individual workbook
passing its own archive checks.

Each side now sums the size captured with every supported regular-file identity
during inventory. The default 4 GiB ceiling is checked after supported-file
filtering but before any `load_snapshot` call. It neither reopens a pathname
nor mixes baseline and candidate budgets: a later replacement still hits the
existing identity guard, while each directory gets its own configurable
`--max-portfolio-source-bytes` allowance.

Direct regressions prove an exact two-workbook boundary succeeds and that a
one-byte overflow reaches no snapshot loader at all. A separate larger-candidate
fixture proves the candidate-side preflight occurs before any source read too.
CLI defaults and a controlled low-budget command return status 2, while the
public composite Action validates the new positive integer input and propagates
the same fail-closed result.

Against the public 0.164.0 wheel, 20 complete comparisons of a 16-workbook
baseline/candidate portfolio took 73.563222 seconds at 40,632 KiB maximum RSS;
the 0.165.0 candidate wheel took 73.085537 seconds at 40,444 KiB. Fresh isolated
wheel and sdist installations reported 0.165.0, passed `pip check`, completed a
normal exact-per-side-budget portfolio, and rejected a one-byte aggregate
overflow before an intentionally malformed `.xlsx` source could be inspected.
The final source suite passed 1,246 tests in 387.82 seconds, with Ruff, bytecode
compilation, Action-shell syntax, and whitespace checks clean.

## Fail-closed portfolio traversal — 2026-07-28

Version 0.164.0 closes a coverage gap in recursive directory portfolios. The
previous implementation relied on `Path.rglob("*")`; contemporary Python glob
implementations can suppress an `OSError` raised while scanning a descendant,
turning an unreadable subtree into an apparent empty branch. That makes an
otherwise successful review artifact falsely imply complete directory coverage.

Portfolio discovery now uses an iterative, bounded `scandir` walk. It counts
every entry before retaining or sorting it, does not recurse through symlinked
directories, and converts any directory-read or entry-type error into a
redacted portfolio failure before workbook comparison. The existing 32,768 raw
entry ceiling, supported-workbook ceiling, and later source-identity check are
unchanged.

With a controlled `PermissionError` on a nested directory, the public 0.163.0
wheel returned success and retained `model.xlsx`; the candidate returned
`Could not inventory baseline portfolio directory.` and CLI exit code 2. Five
ordinary 10,000-entry inventories took 0.827227 seconds at 43,780 KiB RSS in
0.163 and 0.820001 seconds at 46,476 KiB in 0.164. The final source suite
passed 1,239 tests in 203.75 seconds, with Ruff, bytecode compilation,
Action-shell syntax, and whitespace checks clean. The final wheel and source
distribution passed `twine check`. Fresh isolated installs of both passed
`pip check`, version, normal/hostile-policy, exact/overflow-inventory, and
blocked-subtree controls.

## Portfolio source-identity binding — 2026-07-28

Version 0.163.0 closes the remaining inventory-to-read path window in directory
portfolios. Earlier releases rejected symlinks during recursive inventory, but
retained a pathname and opened it later while comparing workbooks. A concurrent
producer could replace that file after inventory with a new regular file or a
symlink, causing the report to inspect a different workbook than the one whose
relative membership had been reviewed.

Each retained regular workbook now carries the device, inode, change timestamp,
and size observed during inventory. The later read requests a no-follow final
component where the host supports it, verifies the opened descriptor still has
that observed state, then makes the existing bounded private inspection copy.
Any in-place rewrite, new regular file, or symlink substitution becomes redacted
`FF078` incomplete evidence and exit code `2`; the report remains available for
the rest of the portfolio.

On one ordinary baseline/candidate workbook pair, 20 complete portfolio
comparisons took 2.100858 seconds at 37,504 KiB RSS with the public 0.162.0
wheel and 2.148245 seconds at 37,540 KiB with the 0.163.0 candidate. In the
controlled late-replacement path, 0.162 reported
`incomplete=False, status=changed`; 0.163 reported
`incomplete=True, status=unreadable, findings=FF078`. The final source suite
passed 1,238 tests in 201.81 seconds, with Ruff, bytecode compilation,
Action-shell syntax, and whitespace checks clean. The final wheel and source
distribution passed `twine check`. Fresh isolated installs of both passed
`pip check`, version, normal/hostile-policy, exact/overflow-inventory, and
direct/CLI late-source-change controls.

## Bounded portfolio filesystem inventories — 2026-07-28

Version 0.162.0 closes a portfolio discovery allocation gap. The prior
supported-workbook ceiling did not constrain arbitrary directory entries:
FormulaFence recursively collected and sorted every path before it decided
whether the path was a supported workbook. A broad CI directory full of source
files, generated assets, or other ordinary non-workbook material could therefore
consume unbounded inventory memory and sorting time even when it contained no
workbook at all.

Each baseline and candidate directory now has a separate 32,768-entry default
budget, configurable with `--max-inventory-entries` and the matching Action
input. The count applies before filtering, retaining, or sorting paths, so it
includes non-workbook files, directories, transient lock files, and symlinks.
Exact-boundary, overflow, zero-limit, CLI, Action validation, and Action
propagation controls cover the fail-closed behavior without changing the
existing supported-workbook or cross-workbook-impact budgets.

A directory containing 10,000 ordinary `.txt` entries and no workbook took
0.203777 seconds at 41,620 KiB RSS for the public 0.161.0 inventory path. The
candidate rejected the same directory at a deliberately reduced 128-entry
budget in 0.004794 seconds at 38,636 KiB RSS, before collecting the remaining
paths for sorting. The final source suite passed 1,234 tests in 203.63 seconds,
with Ruff, bytecode compilation, Action-shell syntax, and whitespace checks
clean. The locally built wheel and source distribution passed `twine check`;
fresh installs of both passed `pip check`, normal and hostile policy controls,
the exact/overflow portfolio-entry controls, and the retained report-output
swap control.

## Atomic no-clobber starter-policy publication — 2026-07-28

Version 0.161.0 closes the remaining `init` no-overwrite time-of-check/time-of-use
gap. Version 0.160.0 checked whether the requested starter-policy pathname
existed and then atomically replaced that pathname. That kept a symlink target
safe, but a normal file, symlink, or hard link created after the initial check
could still have its final directory entry replaced even though the caller had
not supplied `--force`.

The new default path writes the complete starter policy privately beside its
destination, then uses the filesystem's no-replace link creation to claim the
final name. The operation fails rather than overwriting when an entry has
appeared. Deterministic controls substitute a normal policy, a symlink to a
protected file, and a hard link to a protected file immediately before that
publication step: each must return status 2, preserve the competing entry and
protected bytes, and remove the unpublished private file. The explicit
`init --force` path still atomically replaces a final symlink entry without
modifying its target. The boundary remains scoped to final entry publication;
a hostile parent directory is still outside the caller's workspace-permission
trust boundary.

Against the public 0.160.0 wheel, a normal policy created in the gap produced
exit 0 and `concurrent_preserved=False`; the 0.161.0 source returned exit 2
with `concurrent_preserved=True`. In fresh processes, 1,000 normal `init`
publications took 1.423195 seconds at 38,764 KiB RSS with 0.160.0 and 1.439989
seconds at 38,292 KiB RSS with 0.161.0. The small no-clobber operation therefore
preserves the existing startup-scale behavior while making the default contract
filesystem-enforced. The final source suite passed 1,229 tests in 200.92
seconds, with Ruff, bytecode compilation, and whitespace checks clean. The
locally built wheel and source distribution passed `twine check`; fresh installs
of both passed `pip check`, normal and hostile policy controls, the retained
report-path swap control, the new no-clobber race, and forced-final-symlink
publication.

## Atomic report and policy publication — 2026-07-28

Version 0.160.0 closes the final output-path race after FormulaFence has
validated that a report will not overwrite an inspected input. Earlier releases
performed that check, then wrote directly through the report pathname. A
concurrent final-component symlink replacement could therefore redirect the
write into a workbook or policy input after it had been inspected.

A deterministic `profile` control began with a normal workbook and an absent
report path, then substituted a report-path symlink to the workbook immediately
after the public 0.159.0 output check. The command returned success, left the
report path as a symlink, and changed the workbook into Markdown. Version
0.160.0 writes the rendered text to a private same-directory temporary file and
atomically replaces the final directory entry. Under the identical swap, it
returned success with the workbook byte-identical and the report path replaced
by an ordinary report file. Equivalent `init` coverage confirms a swapped
starter-policy path cannot overwrite a target and that `--force` still replaces
an existing policy.

In fresh processes, ten JSON profiles of a 5,301-byte normal workbook took
0.744942 seconds at 38,172 KiB RSS with the public 0.159.0 wheel and 0.744002
seconds at 38,860 KiB RSS with the 0.160.0 source. The final-entry guard does
not protect a hostile output parent directory; its permissions remain part of
the caller's workspace boundary. End-to-end report-swap, starter-policy-swap,
`init --force`, input-overwrite, and ordinary CLI regressions cover the change.
The complete source suite passed 1,225 tests in 198.66 seconds. Release
artifacts are separately built from the tagged commit, checked with `twine`,
installed from both wheel and source distribution into fresh environments, and
compared against the public release bytes.

## Policy source descriptor boundary — 2026-07-28

Version 0.159.0 closes a post-check pathname boundary in policy loading. The
previous loader checked that a policy pathname was a file, then opened that
pathname with ordinary blocking I/O. A concurrent replacement between those two
steps could substitute a FIFO or device and stall a `check` or policy-backed
`portfolio` command before YAML validation.

On a host with POSIX nonblocking descriptor opens, a deterministic control
started with a valid policy then replaced its pathname with a FIFO immediately
before the public 0.158.0 wheel opened it. An alarm had to interrupt the old
pathname open after 1.000191 seconds; its eventual diagnostic was the generic
policy-read error. The 0.159.0 source opens one descriptor with `O_NONBLOCK`,
checks it is regular with `fstat`, and reads the bounded source from that same
descriptor. The equivalent replacement returned `Policy source is not a regular
file.` in 0.000094 seconds, without reaching YAML construction. Hosts without a
nonblocking descriptor-open flag retain their platform behavior; the regular
descriptor check and bounded read still apply.

In fresh processes, 250 parses of a 46-byte normal policy took 0.068762 seconds
at 31,872 KiB RSS with the public 0.158.0 wheel and 0.060042 seconds at 32,364
KiB RSS with the 0.159.0 source. Exact source-limit, post-check FIFO,
regular-file, source-size, UTF-8, YAML-ambiguity, and fail-before-workbook
regressions cover the policy boundary. The complete source suite passed 1,222
tests in 201.42 seconds. Release artifacts are separately built from the tagged
commit, checked with `twine`, installed from both wheel and source distribution
into fresh environments, and compared against the public release bytes.

## Stable workbook source snapshots — 2026-07-28

Version 0.158.0 closes a source-path time-of-check/time-of-use boundary in
workbook inspection. Earlier releases preflighted a pathname, then reopened
that pathname for semantic preflight, raw OOXML scanners, the ordinary reader,
and its final content hash. A concurrent replacement could therefore make the
preflight describe one package while the resulting snapshot described another.

A deterministic controlled pair used an original workbook whose `Model!B2`
formula was `=Inputs!B2*2` and a replacement whose formula was
`=Inputs!B3*2`. A shim that copied the replacement immediately after the public
0.157.0 wheel's archive preflight produced a snapshot hash matching the
replacement and the replacement formula. For the 0.158.0 source, a shim copied
the same replacement to the public pathname immediately after FormulaFence had
made its private source copy. Its snapshot hash instead matched the original
and `Model!B2` retained `=Inputs!B2*2`, while the public pathname's final hash
matched the replacement. The snapshot is consequently internally coherent
after materialization; it is not a lock against an in-place producer write or a
same-identity process that can alter the private temporary file.

In fresh processes, ten complete snapshots of a 5,301-byte, two-sheet normal
control took 0.710387 seconds at 35,432 KiB RSS with the public 0.157.0 wheel
and 0.716368 seconds at 34,720 KiB RSS with the 0.158.0 source. The source adds
one bounded 1 MiB-block copy before the existing reader work, with no material
normal-control memory or elapsed-time regression in this control. Exact
source-replacement, cleanup-on-preflight-error, source-size fail-before-reader,
and parser-warning regressions cover the boundary; the complete source suite
passed 1,221 tests in 200.04 seconds. Release artifacts are separately built
from the tagged commit, checked with `twine`, installed from both wheel and
source distribution into fresh environments, and compared against the public
release bytes.

## Policy-as-code YAML bounds — 2026-07-26

Version 0.157.0 closes the policy-input boundary used by `check` and
policy-backed `portfolio` runs. Earlier releases read an entire YAML file before
parsing it, accepted duplicate mappings with PyYAML's last-wins behavior, and
loaded a `check` policy only after both workbook snapshots. A pull request could
therefore hide a disabled rule behind a repeated key or make a CI job allocate
for a large policy before FormulaFence reached the schema error.

FormulaFence now reads at most 1 MiB of strict UTF-8 source and composes one
bounded document before it inspects a workbook: 4,096 YAML nodes, 64 nesting
levels, 4,096 characters per scalar, and 512 selectors in either selector
list. Anchors, aliases, merge keys, and duplicate mapping keys are rejected,
as are non-string schema keys. This deliberately small configuration subset
makes the reviewed policy unambiguous rather than attempting to support YAML
inheritance or graph semantics.

A controlled pair of 4,815-byte workbooks used a 20,971,553-byte policy with
an unknown top-level field containing a 20 MiB scalar. In fresh processes, the
public 0.156.0 wheel read and parsed that policy before emitting its ordinary
unknown-field error in 7.252205 seconds at 98,768 KiB RSS. The 0.157.0 source
rejected it before YAML construction or workbook inspection in 0.341406 seconds
at 38,704 KiB RSS, with the explicit 1 MiB policy-source diagnostic. The same
46-byte normal policy completed in 0.427354 seconds at 38,176 KiB RSS with
0.156.0 and 0.446255 seconds at 38,188 KiB RSS with the new source. Exact
source, node, nesting, scalar, selector, UTF-8, ambiguous-YAML, schema-key,
and fail-before-workbook regressions cover the boundary; the complete source
suite passed 1,219 tests in 202.43 seconds. The final wheel and source
distribution passed `twine check`; fresh isolated installs of each reported
FormulaFence 0.157.0, passed `pip check`, accepted the normal control, and
rejected the hostile policy before workbook inspection.

## Power Query nested ZIP catalog preflight — 2026-07-26

Version 0.156.0 closes the remaining nested-ZIP catalog path in Power Query
`DataMashup` custom XML. The previous release bounded a logical package before
it inflated members, but metadata-side embedded content only called
`ZipFile(...).infolist()` to count parts. Python creates one `ZipInfo` object
per central-directory entry before returning that count, so a compact embedded
ZIP could still allocate a large catalog even though FormulaFence never read
its content.

FormulaFence now scans bounded central-directory bytes before Python can create
that catalog. It applies the existing 768 KiB nested-source limit and a 512-part
budget shared across logical packages and metadata embedded-content ZIPs, caps
raw member names at 1 KiB, and treats ZIP64, multi-disk, malformed, and
filename-rewriting metadata (Unicode-path aliases, NULs, or platform
separators) as explicit coverage gaps. Logical-package content still receives
the independent stored/deflated, member-size, aggregate-size, and
compression-ratio checks before any member read. Metadata embedded content is
never opened or decompressed merely to count it.

A controlled workbook used a 769,622-byte metadata embedded-content ZIP with
7,400 empty entries. Its Custom XML part was 1,028,589 bytes and the outer
workbook was 78,646 bytes. In fresh processes, the public 0.155.0 wheel
accepted the catalog in 0.179804 seconds at 42,596 KiB RSS, reported 7,401
embedded-content parts, and emitted no warning. The 0.156.0 source stopped at
the shared catalog boundary in 0.159670 seconds at 38,968 KiB RSS, retained the
ordinary logical-package count of one, and emitted the explicit opaque coverage
warning. The corresponding normal control completed in 0.123762 seconds at
35,564 KiB RSS with 0.155.0 and 0.123570 seconds at 35,916 KiB RSS with the new
source, without warnings. Exact-capacity, fail-before-`ZipFile`, filename-
rewrite, metadata-catalog, aggregate-across-mashups, report-visibility, and
ordinary Power Query regressions cover the boundary; the complete source suite
passed 1,207 tests in 200.16 seconds, with Ruff, bytecode compilation, and
diff-whitespace checks also passing. The final wheel and source distribution
passed `twine check`; fresh isolated installs of each reported FormulaFence
0.156.0, passed `pip check`, accepted the normal control, and retained the
hostile catalog's opaque warning.

## Power Query nested ZIP package bounds — 2026-07-26

Version 0.155.0 closes a nested-archive expansion path in Power Query
`DataMashup` custom XML. The outer custom XML and decoded metadata/permission
XML already have independent boundaries, but the logical package is itself a
ZIP stream. Earlier releases listed and read every member to hash it, so a
small inner archive could make the inspection repeatedly inflate a much larger
declared payload.

FormulaFence now checks the package source before opening it and checks every
listed entry before any member read: 768 KiB source, stored/deflated entries,
512 parts, 16 MiB per part, 64 MiB aggregate expanded data across the Power
Query scan, and a 1,000:1 maximum member compression ratio. The boundary is a
CI allocation and coverage limit rather than a Power Query validity rule. An
overage retains a private opaque fingerprint and produces a parser coverage
warning; a baseline-to-candidate diff therefore reports `FF010` and `FF024`.

A controlled workbook used a 537,768-byte inner ZIP with `Section1.m` plus
128 zero-filled 4 MiB configuration entries. It declared 536,870,947 expanded
bytes while the Base64 `DataMashup` Custom XML remained 718,485 bytes and the
outer workbook was 12,385 bytes. The public 0.154.0 wheel accepted that package
in 2.738963 seconds. The 0.155.0 source returned the explicit coverage warning
without inflating a nested member in 0.491038 seconds; its normal Power Query
control completed in 0.470811 seconds. Source-bound, fail-before-open,
fail-before-member-read, aggregate-across-mashups, report-visibility, and
normal-control regressions cover the boundary. The complete 0.155.0 source
suite passed 1,197 tests in 199.85 seconds; Ruff, bytecode compilation, and
diff-whitespace checks also passed.

## XML non-character-data lexical bounds — 2026-07-26

Version 0.154.0 closes the parser-frontier gap left by bounded opening tags
and decoded text. An XML parser can need a complete comment, processing
instruction, declaration, closing tag, or entity reference before an ordinary
start/end stream event is available. FormulaFence now streams lexical
punctuation before parser construction and allows at most 128 KiB of physical
bytes for each of those non-character-data tokens. Opening tags retain their
separate 128 KiB limit and stable diagnostic. The scanner covers
UTF-8/ASCII-compatible XML plus UTF-16 and UTF-32 punctuation, preserves
quoted delimiters, and does not count CDATA content as lexical markup: the
existing 1 MiB incremental character-data target bounds that content instead.
The shared defused parser now explicitly forbids document-type declarations,
including paths that do not begin with an ASCII byte sequence. Python's
[ElementTree parsing documentation](https://docs.python.org/3/library/xml.etree.elementtree.html)
describes the incremental parser/target boundary that this guard precedes.

Three controlled raw-ZIP fixtures began with the normal 4,835-byte workbook:
a 20,000,000-character random comment and processing instruction appended at
the Styles root, plus a 20,000,000-character comment in a document-type
internal subset. Their `xl/styles.xml` members were 20,002,638, 20,002,638,
and 20,002,663 bytes; their packages were 15,154,584, 15,154,668, and
15,154,367 bytes. The random payloads avoid the ZIP ratio gate, so the parser
frontier remains the measured path. In fresh isolated CLI processes, the exact
public 0.153.0 wheel accepted the normal control in 0.404297 seconds at 37,900
KiB RSS, then accepted the comment, processing-instruction, and document-type
fixtures in 2.644092 seconds / 130,792 KiB, 2.529314 seconds / 149,520 KiB,
and 1.728552 seconds / 130,920 KiB RSS respectively.

The 0.154.0 source accepted the normal control in 0.372747 seconds at 38,144
KiB RSS. It rejected those three hostile packages before parser entry in
0.323810 seconds / 38,008 KiB, 0.323149 seconds / 38,120 KiB, and 0.335219
seconds / 37,300 KiB RSS, each with CLI status 2 and the stable
non-character-data-markup safety-preflight error. Exact boundary, one-byte
chunk, UTF-16/UTF-32, comment, processing-instruction, CDATA, end-tag,
entity-reference, declaration, document-type, direct-stream, real-workbook,
and fail-before-parser regressions cover the gate. The archive-safety suite
passed 325 tests in 14.61 seconds, and the complete versioned source suite
passed 1,193 tests in 199.38 seconds. Ruff, bytecode compilation, and
`git diff --check` passed. The release wheel and sdist passed `twine check`;
fresh isolated installs of each reported FormulaFence 0.154.0, accepted the
normal control, and rejected all three hostile workbooks through the CLI with
status 2 and the non-character-data-markup safety-preflight error.

## XML character-data bounds — 2026-07-26

Version 0.153.0 closes the remaining parser-frontier gap after XML start tags
have passed their lexical limit. ElementTree supplies parsed character data to
its tree-builder target while parsing, but the ordinary builder joins the
successive chunks for one text or tail node until an XML markup boundary. That
means an end-event semantic check is too late for otherwise ignored opaque text.
FormulaFence now uses a bounded tree-builder target with its defused XML parser:
each decoded character-data node can retain at most 1 MiB before the next chunk
is handed to the ordinary builder. The same target covers semantic-reader
streams, raw XML structure scans, rich-text/rich-data/array-formula streams,
and in-memory OOXML root reads. The public
[ElementTree parsing documentation](https://docs.python.org/3/library/xml.etree.elementtree.html)
describes parser targets and incremental parsing; the limit is an allocation
boundary, not an OOXML validity rule.

A controlled raw-ZIP fixture started from the normal 4,835-byte workbook and
appended one ignored `ff:opaque` Styles-root child with 20,000,000 base64
characters. Its `xl/styles.xml` expanded to 20,002,693 bytes and the package
to 15,154,612 bytes, avoiding the ZIP ratio gate so the parser path could be
measured. In fresh processes, the exact released 0.152.0 wheel accepted the
fixture in 1.500564 seconds at 88,792 KiB RSS. The 0.153.0 source rejected it
in 0.087663 seconds at 35,492 KiB RSS with the stable XML-character-data
safety-preflight error. The normal control remained accepted in 0.053025
seconds / 35,224 KiB RSS, versus 0.051307 / 35,160 for the released wheel.

Exact-boundary, one-byte incremental-feed, ordinary-text, CDATA, independent
text/tail-node, direct-stream, real-workbook, and fail-before-reader
regressions cover the target. The archive-safety suite passed 284 tests, and
the complete source suite passed 1,153 tests in 191.38 seconds.
Twine metadata checks passed for the wheel and sdist. Fresh isolated wheel and
sdist installs each accepted the normal control and rejected the hostile
workbook through the CLI with status 2 and the XML-character-data
safety-preflight error.

## XML opening-tag attribute bounds — 2026-07-26

Version 0.152.0 closes the parser-frontier allocation gap left by structural
element counts. Python documents that an ElementTree iterparse start event has
already seen the closing delimiter and has defined attributes; its
[ElementTree documentation](https://docs.python.org/3/library/xml.etree.elementtree.html)
and [pyexpat callback reference](https://docs.python.org/3/library/pyexpat.html)
therefore place a complete attribute mapping on the parser side of a start
callback. The existing defused XML parser blocks document-type and entity
paths, but a compact ordinary element can still carry a very large number of
distinct attributes before an element-count callback runs.

Controlled raw-ZIP fixtures began with the normal 4,835-byte workbook used by
the stylesheet audit. One placed 500,000 distinct attributes on the Styles
root; another put them on one direct cell-format record. Each Styles part was
6,502,631 bytes, while the compressed workbooks were 1,184,083 and 1,184,103
bytes respectively. The exact released 0.151.0 wheel completed the root case
in 7.326036 seconds at 212,576 KiB RSS. It reached the nested-record case in
6.654143 seconds and 211,676 KiB RSS before the ordinary style constructor
reported its unexpected attributes. A Python expat start callback with a
post-callback attribute-count check still reached 126,628 KiB, confirming that
the limit must be lexical and precede parser construction.

The 0.152.0 source lexically streams each XML opening tag before ElementTree
starts, allowing 128 KiB of physical tag bytes. It uses byte-search fast paths
for UTF-8/ASCII-compatible parts and fixed-width handling for UTF-16/UTF-32;
quoted delimiters, comments, CDATA, processing instructions, and declarations
do not impersonate elements. A fresh process accepted the normal control in
0.051743 seconds / 35,260 KiB RSS, compared with 0.043672 / 35,412 for the
exact released 0.151.0 wheel. It rejected the root and nested-record fixtures
before parser entry in 0.009795 seconds / 34,852 KiB and 0.009885 / 34,236 KiB
respectively, with the stable XML start-tag safety-preflight error. Exact
boundary, one-byte chunk, quoted/non-element-markup, UTF-16/UTF-32,
root/nested-map, and fail-before-parser regressions cover the lexical gate. The
archive-safety suite passed 279 tests in 13.10 seconds, and the complete source
suite passed 1,147 tests in 189 seconds with zero failures or errors. Twine
metadata checks passed for the 0.152.0 wheel and sdist. Fresh isolated wheel
and sdist installs each accepted the normal control and rejected both
500,000-attribute fixtures through the CLI with status 2 before report output.

## Stylesheet XML structural bounds — 2026-07-26

Version 0.151.0 closes the remaining compact-allocation paths in
`xl/styles.xml`. The Open XML SDK's
[Stylesheet reference](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.stylesheet?view=openxml-3.0.1)
lists the named root catalogs and its
[ExtensionList reference](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.extensionlist?view=openxml-3.0.1)
documents the extensible `ext` payload. The ordinary stylesheet reader reads
the complete part into bytes and constructs a complete XML tree before it
materializes its catalog objects. It dispatches named controls by local name;
some nested sequences turn every direct child into a style object, while other
descriptors ignore an unknown child only after that tree is already allocated.

The semantic-reader preflight now streams the Styles part before that reader
starts. Every local-name `extLst` subtree allows 32,768 elements. A foreign
direct `styleSheet` child, a non-`styleSheet` root local name, and an ignored
direct child inside a named root catalog each separately allow 32,768 elements.
Each materialized direct style record allows 32,768 non-extension descendants,
with 262,144 such descendants across the Styles part. Existing named catalog
limits—including the 65,490 effective-`cellXfs` ceiling—remain in force. A
successfully streamed overage produces the stable safety-preflight error and
CLI status 2 rather than a partial profile. These are CI allocation limits, not
SpreadsheetML validity rules.

Controlled raw-ZIP fixtures began with a normal 4,835-byte workbook whose
Styles part was 2,631 bytes. Seven hostile variants placed 500,000 elements in
the documented root extension list (16,684-byte package / 6,002,830-byte
Styles part), a foreign direct root child (16,632 / 6,002,705), an ignored
`cellXfs` child (16,663 / 6,002,705), a foreign child inside one `xf` (16,668 /
6,002,709), repeated `alignment` children inside one `xf` (16,618 /
6,002,635), repeated `name` children inside one font (10,027 / 3,502,631), and
a foreign Styles-root local name (16,624 / 6,002,748). The exact released
0.150.0 wheel successfully profiled the normal, extension, root, catalog,
record, repeated-alignment, repeated-font, and foreign-root fixtures in
0.396208 seconds / 37,504 KiB, 6.613773 / 146,628 KiB, 6.257524 / 146,584 KiB,
6.290636 / 147,260 KiB, 6.207826 / 146,740 KiB, 12.517622 / 267,672 KiB,
9.636139 / 416,496 KiB, and 6.198354 / 147,236 KiB RSS respectively.

The 0.151.0 source keeps the normal fixture accepted in 0.437032 seconds /
37,444 KiB. It rejects the seven hostile fixtures in 0.508010–0.589207 seconds
and 37,900–38,624 KiB RSS before profile output, with the corresponding
extension-list, opaque-root, opaque-catalog, or style-record safety-preflight
error. Root, catalog, record, repeated-known-child, aggregate, nested,
alternate-namespace, exact/default-limit, and fail-before-reader regressions
accompany the fixtures. The archive-safety suite passed 265 tests, and the
complete source suite passed 1,133 tests in 162.48 seconds with zero failures
or errors. `twine check` passed for the wheel and sdist. Fresh wheel and sdist
installs each accepted the normal workbook and rejected all seven hostile
fixtures before profile output with the corresponding stable stylesheet-XML
safety-preflight error.

## Workbook XML structural bounds — 2026-07-26

Version 0.150.0 closes the same compact-allocation path in the bootstrap
`xl/workbook.xml` part. The Open XML SDK's
[Workbook reference](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.workbook?view=openxml-3.0.1)
lists `extLst` among the root's named children, and Microsoft's
[XLSX workbook specification](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/e29966e2-5baa-4fcf-84c9-082025e1be13)
documents that container's `ext` payloads. The ordinary workbook reader first
reads this entire part into bytes and builds a complete package tree; FormulaFence
also has raw workbook scanners for tab and legacy Custom View metadata. Neither
path needs an unbounded extension payload or foreign root subtree to produce a
profile.

The semantic-reader preflight now streams the bootstrap workbook part before
either reader begins. It allows 32,768 elements beneath every local-name
`extLst` subtree, including lists nested in a workbook view or using an
alternate namespace, and separately allows 32,768 elements beneath a foreign
direct workbook-root child. The named Workbook controls retain their existing
format-aware catalog limits, including 512 sheet declarations and 100,000
defined names. A successfully parsed overage produces the stable
safety-preflight error and CLI status 2 rather than a partial profile. These
are CI allocation limits, not SpreadsheetML validity rules.

Controlled raw-ZIP fixtures began with a normal 4,835-byte workbook. One
appended a documented `extLst`/`ext` container with 500,000 foreign children to
`xl/workbook.xml`; its package was 16,653 bytes and the workbook XML expanded
to 6,000,749 bytes. A second appended a foreign direct root container with the
same child count; its package was 16,591 bytes and workbook XML expanded to
6,000,624 bytes. The exact released 0.149.0 wheel completed successful profiles
for the normal, extension, and opaque-root fixtures in 0.396157 seconds /
37,988 KiB, 28.430383 seconds / 139,648 KiB, and 28.208687 seconds /
139,608 KiB RSS respectively. The 0.150.0 source rejects the extension fixture
in 0.436918 seconds / 37,100 KiB and the opaque fixture in 0.426571 seconds /
38,012 KiB before profile output; the normal workbook remains accepted in
0.395590 seconds / 38,080 KiB. Direct/nested, nested-view, alternate-root and
alternate-extension-namespace, exact/default-limit, normal-control, and
fail-before-reader regressions accompany the fixtures. The archive-safety suite
passed 244 tests, and the complete source suite passed 1,112 tests in 160.48
seconds with zero failures or errors. `twine check` passed for the wheel and
sdist. Fresh wheel and sdist installs each accepted the normal workbook and
rejected both hostile fixtures before profile output with the stable
workbook-XML safety-preflight error.

## Non-grid sheet structural bounds — 2026-07-26

Version 0.149.0 closes the adjacent compact-allocation path in relationship-
selected chart-sheet and dialog-sheet XML. The Open XML SDK's
[Chartsheet](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.chartsheet?view=openxml-3.0.1)
reference identifies `extLst` as a direct child of the chart-sheet root. The
published [Dialogsheet grammar](https://c-rex.net/samples/ooxml/e1/part4/OOXML_P4_DOCX_dialogsheet_topic_ID0ETIX4.html)
likewise names `extLst` and contains only sheet-level control elements rather
than a cell grid. FormulaFence's raw protection and Custom View scans parse
both kinds of selected sheet; the ordinary reader loads a chart sheet as one
complete XML tree and routes a dialog sheet through its worksheet path. Neither
path needs an unbounded extension or opaque XML subtree to produce a profile.

The semantic-reader preflight now streams each relationship-selected Chartsheet
and Dialogsheet before either raw or ordinary reader starts. It permits 32,768
XML elements in each non-grid sheet part and 65,536 across the selected
inventory. That whole-part boundary covers direct `extLst` content, nested
foreign extension descendants, and opaque root trees without imposing a small
tree limit on ordinary worksheet cell grids; relationship-selected DrawingML
continues to use its separate structural budget. A successfully parsed overage
returns the stable safety-preflight error and CLI status 2 rather than a
partial profile. These are CI allocation limits, not SpreadsheetML validity
rules.

Raw-ZIP fixtures began with a normal workbook containing one ordinary worksheet
and a chart sheet. One fixture appended an `extLst`/`ext` container with
500,000 foreign children to the chart sheet. A second changed the workbook
relationship and content type to a `xl/dialogsheets/sheet1.xml` Dialogsheet,
then added the same valid extension container. The chart-sheet package was
18,858 bytes and its selected XML expanded to 6,000,521 bytes; the dialog-sheet
package was 18,576 bytes and expanded to 6,000,306 bytes. The exact released
0.148.0 wheel completed successful profiles in 3.550517 seconds / 90,328 KiB
and 3.682121 seconds / 90,700 KiB respectively. The 0.149.0 source rejects the
chart fixture in 0.386294 seconds / 38,000 KiB and the dialog fixture in
0.406587 seconds / 37,580 KiB, before a profile is produced. A normal chart
workbook remains accepted in 0.406379 seconds / 37,592 KiB. Direct/nested
extension-list, direct opaque-root, cross-type aggregate, exact/default-limit,
normal-sheet, and fail-before-reader regressions accompany the fixtures. The
archive-safety suite passed 231 tests,
and the complete source suite passed 1,099 tests in 158.62 seconds with zero
failures or errors. `twine check` passed for the wheel and sdist. Fresh wheel
and sdist installs each accepted the normal chart workbook and rejected both
hostile packages before profile output with the stable non-grid-sheet
safety-preflight error.

## Worksheet extension-list structural bounds — 2026-07-26

Version 0.148.0 closes the remaining compact-allocation path inside a named
Worksheet extension container. The Open XML SDK's
[Worksheet](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.worksheet?view=openxml-3.0.1)
reference lists `WorksheetExtensionList` alongside the ordinary root children.
That correctly leaves `extLst` outside the 0.147 opaque-root counter, but an
extension list can itself carry arbitrary XML. FormulaFence's Office Web
Add-in worksheet scan parses every selected worksheet to locate that extension
location, and the ordinary worksheet reader dispatches it as an extension list.
Neither path needs an unbounded opaque extension tree to produce a profile.

The semantic-reader preflight now tracks every element below a SpreadsheetML
`extLst` in each selected transitional or Strict worksheet before either reader
starts. It permits 32,768 extension-list elements per worksheet and 65,536 in
aggregate. The scan detects a list nested beneath another named Worksheet
control as well as a direct root list, without double-counting an element if a
nested extension list occurs. Ordinary `sheetData` and other named base
controls retain their existing specialist limits. A successfully parsed
overage returns the stable safety-preflight error and CLI status 2 rather than
a partial profile. These are CI allocation limits for extension content, not a
SpreadsheetML validity rule.

A raw-ZIP fixture generated without FormulaFence retained a four-sheet model
and appended one `extLst`/`ext` container with 500,000 foreign children to one
worksheet. The workbook was 18,225 bytes; its selected worksheet XML expanded
to 6,000,799 bytes. The exact released 0.147.0 wheel completed a successful
profile in 16.513316 seconds at 126,176 KiB RSS. The 0.148.0 source rejects the
same input before the traced Office Web Add-in scan in 0.477682 seconds at
37,436 KiB RSS, with the stable worksheet extension-list safety-preflight error
and no profile JSON. A normal 6,408-byte model profile completed in 0.467634
seconds at 37,996 KiB RSS. Direct/nested, nested-under-`sheetPr`, aggregate,
exact/default-limit, normal-sheet, Strict-worksheet, and fail-before-reader
regressions accompany the fixture. The full archive-safety suite passed 220
tests with zero failures, and the complete source suite passed 1,088 tests in
158.73 seconds with zero failures or errors.

## Worksheet opaque-root structural bounds — 2026-07-26

Version 0.147.0 closes a compact-allocation path shared by FormulaFence's raw
worksheet scanners and the ordinary worksheet reader. The Open XML SDK's
[Worksheet](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.worksheet?view=openxml-3.0.1)
reference enumerates the base root children, including `sheetData`, controls,
relationships, and the `extLst` extension location. The semantic-reader
preflight intentionally permits a large `sheetData` sequence under its existing
cell and dimension budgets. It previously had no compact structural boundary
for a complete direct root subtree whose tag is outside that base grammar,
though raw worksheet metadata readers can build a complete XML tree and the
ordinary reader can retain an unrecognized root child until parsing completes.

The preflight now matches the documented direct child grammar in both
transitional and Strict Worksheet parts before either reader starts. It counts
each complete subtree rooted at any other direct child, allowing 32,768 XML
elements per selected worksheet and 65,536 in aggregate. Named base children,
including `sheetData`, remain outside this narrow counter and retain their
existing specialist limits. A successfully parsed overage returns the stable
safety-preflight error and CLI status 2 rather than a partial profile. These
are CI allocation limits for opaque root content, not a SpreadsheetML validity
rule.

A raw-ZIP fixture generated without FormulaFence retained a four-sheet model
and appended 500,000 ignored foreign direct children to one worksheet root.
The workbook was 87,143 bytes; its selected worksheet XML expanded to
28,001,080 bytes. The exact released unguarded 0.146.0 wheel completed a
successful profile in 18.774197 seconds at 133,860 KiB RSS, emitting no warning
or failure. The 0.147.0 source rejects the same input before any workbook
reader starts in 0.459208 seconds at 37,588 KiB RSS, with the stable worksheet
opaque-root safety-preflight error and no profile JSON. A normal 6,408-byte
model profile completed in 0.443235 seconds at 37,508 KiB RSS. Direct/nested,
aggregate, exact/default-limit, standard-root, Strict-worksheet, and
fail-before-reader regressions accompany the fixture. The complete source suite
passed 1,081 tests in 158.31 seconds with zero failures or errors.

## Shared-string rich-text structural bounds — 2026-07-26

Version 0.146.0 closes a compact-allocation path in the raw rich-text scanner
and the ordinary shared-string reader. SpreadsheetML uses the
[Shared String Table](https://learn.microsoft.com/en-us/office/open-xml/spreadsheet/working-with-the-shared-string-table)
to hold ordinary values and rich-text runs in `si` items. A table may
legitimately have many simple entries, but FormulaFence previously constructed
the entire table before inspecting rich runs, while the standard reader could
retain ignored direct root children until parsing completed. A compact foreign
root subtree therefore had no shared-string-specific structural boundary.

The semantic-reader preflight now streams every shared-string table selected by
either the manifest-backed ordinary reader or the raw rich-text relationship
before either reader starts. It limits each complete `si` item to 32,768 XML
elements, complex/rich items to 65,536 in aggregate, and opaque direct `sst`
children to 32,768 elements per table and 65,536 in aggregate. Simple values
retain the established 500,000-entry allowance. The raw rich-text scanner now
keeps only one direct `si` tree at a time and removes completed irrelevant root
children; the generic structural stream likewise detaches completed elements
instead of leaving cleared siblings attached to their parent. A successfully
parsed overage returns the stable safety-preflight error and CLI status 2,
rather than producing a partial profile. These are CI allocation limits, not a
SpreadsheetML validity rule.

A raw-ZIP fixture generated without FormulaFence retained one ordinary rich
shared string and appended 500,000 ignored foreign direct children to
`xl/sharedStrings.xml`. The workbook was 85,611 bytes; the shared-string XML
expanded to 27,500,347 bytes. The exact released unguarded 0.145.0 wheel
completed a successful profile in 2.928342 seconds at 110,200 KiB RSS, emitted
no warning, and reported no failure. The 0.146.0 source rejects the same input
before either reader starts in 0.424907 seconds at 38,000 KiB RSS, with the
stable shared-string opaque-XML safety-preflight error and no profile JSON. A
normal rich-string profile completed in 0.390174 seconds at 37,616 KiB RSS.
Direct/nested opaque-root, manifest/relationship-selection, per-item,
aggregate, exact/default-limit, fail-before-reader, and streamed-rich-text
regressions accompany the fixture.
The complete source suite passed 1,073 tests in 157.84 seconds with zero
failures or errors.

## Table Definition XML semantic-preflight bounds — 2026-07-26

Version 0.145.0 closes a compact-allocation path shared by the raw readers and
ordinary workbook reader that can reach an Excel Table Definition part. A
worksheet's [`tableParts`](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.tableparts?view=openxml-3.0.1)
declaration identifies each table through a relationship ID. FormulaFence's
raw filter, Named Sheet View, external-data, XML Mapping, and Table Style
scanners can inspect the same definition before or alongside `openpyxl`; the
generic semantic-reader XML ceiling was intentionally too high to protect a
small, repetitive table tree.

The semantic-reader preflight now streams every canonical `xl/tables/*.xml`
part and every safe direct internal worksheet `table` relationship target before
any of those readers can construct it. Canonical orphan parts are included
because the Table Style scanner inventories them; standard transitional and
Strict relationships plus noncanonical safe targets are included because the
other readers and ordinary workbook reader can follow them. Each target allows
32,768 elements and the complete Table Definition inventory allows 65,536. A
successfully parsed overage returns the stable safety-preflight error and CLI
status 2 rather than a partial profile. Missing, malformed, and non-XML
optional targets keep their existing downstream coverage diagnostics. These are
CI allocation limits, not an OOXML validity rule.

A raw-ZIP fixture generated without FormulaFence retained one ordinary table
and added 1,000,000 ignored foreign direct children to
`xl/tables/table1.xml`. The workbook was 26,997 bytes; its Table Definition
part was 11,000,385 bytes expanded and 21,613 bytes compressed. The exact
released unguarded 0.144.0 wheel completed a profile in 5.727898 seconds at
261,232 KiB RSS, reported one table, and emitted no warning. The 0.145.0 source
rejects the same input before any workbook reader starts in 0.059982 seconds at
37,540 KiB RSS, with the stable table-definition safety-preflight error and no
profile JSON. A normal baseline profile completed in 0.047602 seconds at
35,152 KiB RSS. Direct/nested opaque, Strict relationship, canonical-orphan
aggregate, exact/default-capacity, fail-before-reader, and malformed-orphan
regressions accompany the fixture. The complete source suite passed 1,065
tests in 157.507 seconds with zero failures or errors.

## Dynamic-array metadata XML structural bounds — 2026-07-26

Version 0.144.0 closes a compact-allocation path in the raw
`xl/metadata.xml` reader used to distinguish legacy CSE formulas from dynamic
arrays. The metadata definition links a `futureMetadata` record to a metadata
type, including the `XLDAPR` dynamic-array properties used by FormulaFence's
classification ([MS-XLSX metadata](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/3dd44d53-847b-402f-a8c7-41a85024caf7)). A byte-permitted metadata part could
therefore first materialize a repetitive tree even when all added children were
foreign and irrelevant to the mapping. The classifier also previously built a
second complete worksheet tree only to recover direct `c` / `f` bindings.

FormulaFence now streams the canonical metadata part before parsing it, allowing
16 MiB and 32,768 XML elements. A successful structural or byte overage stays
out of the metadata tree parser, marks raw array-formula classification
incomplete, and leaves every observed array formula unclassified with no
fixed-CSE or observed-spill alias. It emits `FF010`; a private streamed SHA-256
fallback fingerprint makes same-size opaque coverage changes visible through
`FF018`, while JSON, Markdown, and SARIF expose neither metadata nor raw XML.
The direct worksheet binding scan is now streaming as well. Malformed metadata
retains its parser diagnostic. These are allocation boundaries for this named
dynamic-array classifier, not a general SpreadsheetML validity limit.

A controlled dynamic-array fixture retained a real `cm=1` / `XLDAPR` binding
at `Model!B1` and appended 1,000,000 ignored foreign direct metadata children.
Its `xl/metadata.xml` part was 6,000,781 bytes uncompressed and 9,243 bytes
compressed; the whole workbook was 15,291 bytes. The released unguarded 0.143.0
reader completed the hostile comparison in 1.301145 seconds / 134,580 KiB with
no findings, no warning, and no diff. The 0.144.0 source completed the same
comparison in 0.284295 seconds / 47,360 KiB without materializing the hostile
tree, emitting `FF010` and `FF018` with
`parser_coverage_warning_added` and `array_formula_metadata_coverage_changed`.
A normal self-comparison completed in 0.160159 seconds / 38,316 KiB with no
findings or changes. The hostile output contained neither the injected namespace
nor the foreign element tag.

## XLM macro-sheet XML structural bounds — 2026-07-26

Version 0.143.0 closes a compact-allocation path in raw Excel 4.0 / XLM
macro-sheet XML. FormulaFence already scanned macro-sheet formulas before the
ordinary workbook library could omit them, but a byte-permitted macro part
could first materialize a repetitive private tree and recursively canonicalize
unsupported children. The ordinary workbook reader and Custom View
sanitization also had secondary paths that could reopen that same raw macro
tree.

FormulaFence now streams each selected raw macro-sheet XML part before private
parsing. Each part allows 32,768 elements and the complete shared macro-sheet
scan allows 65,536, together with 16 MiB per-part, 64 MiB aggregate, and
512-part limits. A successfully streamed structural overage becomes explicit
`FF010` plus `FF026`-visible opaque macro-sheet evidence. Its private streamed
SHA-256 content fingerprint keeps same-size hostile XML diff-visible without
retaining macro commands, cell values, relationships, embedded payloads, or
raw XML. After the raw scan, the temporary ordinary-workbook reader gets an
empty worksheet replacement for selected XLM targets, while Custom View
sanitization and association parsing exclude them. An invalid ordinary-sheet
relationship alias to the same raw target is likewise excluded from generic
sheet metadata and receives a visible coverage warning. Malformed macro XML
reached before a structural overage retains its ordinary parser diagnostic.
These are allocation limits for the named raw macro-sheet readers, not general
legacy-workbook package validity rules.

A controlled XLM fixture added 1,000,000 foreign-namespace direct children to
`xl/macrosheets/sheet1.xml`. The affected macro part was 6,000,635 bytes
uncompressed, 9,146 bytes compressed, and the complete workbook was 16,685
bytes. In isolated source measurements, the released unguarded 0.142.0 reader
completed the hostile comparison in 7.090131 seconds / 528,944 KiB and emitted
only `FF026`. The 0.143.0 source completed the same comparison in 0.262575
seconds / 44,808 KiB without materializing the hostile tree, emitting `FF010`
and `FF026` with `parser_coverage_warning_added` and
`xlm_macro_sheets_changed` records. A normal baseline self-comparison completed
in 0.198028 seconds / 38,128 KiB with no findings or changes. The hostile JSON
contained neither the injected namespace nor the raw element tag. The full
source suite passed 1,051 tests in 156.63 seconds; Ruff and bytecode
compilation were clean.

## Query-table XML structural bounds — 2026-07-26

Version 0.142.0 closes a compact-allocation path in raw query-table XML.
FormulaFence already inspected query-table refresh controls reached directly
from worksheets or through Excel tables, but a byte-permitted query-table part
could first materialize a repetitive private tree and recursively canonicalize
unsupported children. One target can be bound from multiple worksheets, so the
boundary caches both the parsed root and a private sheet-neutral control
template instead of repeating the parse or opaque traversal per binding.

FormulaFence now streams each selected query-table relationship target before
private parsing. Each part allows 32,768 elements and the complete shared
query-table scan allows 65,536, together with 16 MiB per-part, 64 MiB aggregate,
and 512-part limits. A successfully streamed structural overage becomes
explicit `FF010` plus `FF023`-visible opaque query-table evidence. Its private
streamed SHA-256 content fingerprint keeps same-size hostile XML diff-visible
without retaining query-table names, connection metadata, field data, sort
state, extension material, or raw XML. Malformed input retains its ordinary
parser diagnostic. These are allocation limits for the raw query-table reader,
not general table or external-data package validity rules.

A controlled external-data fixture added 1,000,000 foreign-namespace direct
children to `xl/queryTables/queryTable1.xml` through a direct worksheet
relationship. The affected part was 6,000,506 bytes uncompressed, 9,118 bytes
compressed, and the complete workbook was 17,431 bytes. In isolated source
measurements, the released unguarded 0.141.0 reader completed the hostile
comparison in 6.685531 seconds / 536,832 KiB and emitted only `FF023`. The
0.142.0 source completed the same comparison in 0.260500 seconds / 44,292 KiB
without materializing the hostile tree, emitting `FF010` and `FF023` with
`parser_coverage_warning_added` and
`query_table_refresh_controls_changed` records. A normal baseline
self-comparison had no findings or changes. The hostile JSON contained neither
the injected namespace nor raw element tag. The full source suite passed 1,037
tests in 150.53 seconds; Ruff and bytecode compilation were clean.

## External-link package XML structural bounds — 2026-07-26

Version 0.141.0 closes a compact-allocation path in raw external-link package
XML. FormulaFence already inventories external-workbook, DDE, and OLE link
definitions, but a byte-permitted `externalLink` part could first materialize a
highly repetitive private tree and recursively canonicalize opaque children.
The inventory and package-indexed external-workbook resolver can both read the
same link part, so the boundary is shared and cached rather than allowing a
second unbounded parse.

FormulaFence now streams selected
`xl/externalLinks/externalLink*.xml` parts and their direct relationship parts
before private parsing. Each part allows 32,768 elements and the complete
shared external-link scan allows 65,536, together with 16 MiB per-part, 64 MiB
aggregate, and 512-part limits. A successfully streamed structural overage
becomes explicit `FF010` plus `FF025`-visible opaque package evidence. Its
private streamed SHA-256 content fingerprint keeps same-size hostile XML
diff-visible without retaining package targets, names, DDE/OLE metadata, cache
values, or raw XML. Malformed input retains its ordinary parser diagnostic.
These are allocation limits for the named raw external-link readers, not
general package-relationship validity rules.

A controlled three-part external-link fixture added 1,000,000 foreign-namespace
direct children to `xl/externalLinks/externalLink3.xml`. The affected part was
6,000,645 bytes uncompressed, 9,134 bytes compressed, and the complete workbook
was 17,577 bytes. In isolated source measurements, the released unguarded
0.140.0 reader completed the hostile comparison in 8.561901 seconds / 714,764
KiB and emitted only `FF025`. The 0.141.0 source completed the same comparison
in 0.255000 seconds / 42,308 KiB without materializing the hostile tree,
emitting `FF010` and `FF025` with
`external_link_packages_changed` and `parser_coverage_warning_added` records.
A normal baseline self-comparison had no findings or changes. The hostile JSON
contained neither the injected namespace nor raw element tag. The full source
suite passed 1,026 tests in 146.22 seconds; Ruff and bytecode compilation were
clean.

## External-data Connections XML structural bounds — 2026-07-26

Version 0.140.0 closes a compact-allocation path in raw external-data
Connections XML. FormulaFence already inspected connection refresh controls,
but its raw `xl/connections*.xml` reader could first expand a highly repetitive
yet byte-permitted XML tree before recording its opaque metadata. This release
adds 16 MiB per-part, 64 MiB aggregate, and 512-part read limits as well as the
structural element boundary.

FormulaFence now streams each raw Connections XML part before private parsing,
with 32,768 elements allowed per part and 65,536 across a Connections scan. A
successfully streamed structural overage becomes explicit `FF010` plus
`FF023`-visible opaque connection evidence; malformed input retains its
ordinary parser diagnostic. The fallback includes a private streamed SHA-256
content fingerprint, so same-size hostile parts remain diff-visible without
retaining connection names, paths, strings, commands, or raw XML. These are
allocation boundaries for raw Connections XML, not general external-data
package validity rules.

A controlled fixture added 1,000,000 foreign-namespace direct children to a
normal Connections root. Its `xl/connections.xml` expanded to 13,001,347 bytes,
compressed to 25,951 bytes, and contained 1,000,009 XML elements. In the same
source environment, the unguarded 0.139.0 reader completed it in 1.126806
seconds / 145,232 KiB; the 0.140 source completed it in 0.181229 seconds /
40,708 KiB without materializing the hostile tree. A normal baseline comparison
had no findings or changes. Comparing the normal baseline with the hostile
candidate produced `FF010` and `FF023`, an
`external_data_connections_changed` record, and no hostile tag or raw content
in JSON, Markdown, or SARIF. The source suite passed 1,015 tests in 143.70
seconds with Ruff, bytecode compilation, and diff whitespace checks clean.
The exact final wheel completed the normal and hostile fixtures in 0.132709
seconds / 33,224 KiB and 0.226601 seconds / 39,280 KiB. A fresh Python 3.12
environment installed it from `site-packages`, passed `pip check`, confirmed
`openpyxl.DEFUSEDXML`, emitted no normal findings, and emitted only
`FF010`/`FF023` for the hostile comparison. Its SHA-256 is
`7225517d66715a7706c96e53fc1944e1442e4f7590b6836e7b6348e3d7147ebd`.

## Legacy shared-workbook revision XML structural bounds — 2026-07-26

Version 0.139.0 closes a compact-allocation path in raw legacy
shared-workbook revision history. FormulaFence already limited `revisionHeaders`
and `revisionLog` parts to 16 MiB each, 64 MiB per scan, and 512 parts, but the
private revision scanner parsed and recursively canonicalized each permitted
XML tree. A highly repetitive history therefore could consume substantial CI
memory despite a very small ZIP member.

FormulaFence now streams every revision header, log, and relationship XML part
before private parsing. Each part allows 32,768 elements and the complete
revision scan allows 65,536. A successfully streamed overage becomes explicit
`FF010` plus `FF062` coverage evidence; malformed input retains its established
parser diagnostic. The fallback includes a private streamed SHA-256 content
fingerprint, so same-size hostile revisions remain diff-visible without
retaining history, identity, cell, or XML content. These are allocation
boundaries, not shared-workbook validity rules.

Two controlled fixtures inserted 1,000,000 direct valid records into the
`revisionLog` and `revisionHeaders` roots. The log was 6,000,127 bytes,
compressed to 8,879 bytes; the header was 9,000,123 bytes, compressed to
17,605 bytes. In the same source environment, the unguarded 0.138.0 reader
completed them in 4.211653 seconds / 449,600 KiB and 5.087386 seconds /
454,444 KiB. The 0.139 source completed them in 0.158966 seconds / 41,268 KiB
and 0.166300 seconds / 40,688 KiB without materializing either hostile tree.
Normal baseline-to-baseline diffs had no findings or changes; each hostile
baseline comparison produced `FF010`/`FF062`, a
`shared_workbook_revisions_changed` record, and no exposed revision material.
The source suite passed 1,005 tests in 143.94 seconds with Ruff, bytecode
compilation, and diff whitespace checks clean.
The exact final wheel completed the log and header fixtures in 0.204736 seconds
/ 40,052 KiB and 0.247325 seconds / 40,164 KiB; its SHA-256 is
`3e94b11538e31b5d3cea0614e9468a7634941602b466fa0aa4b8306357c6bfa4`.
A fresh Python 3.12 environment installed that wheel, passed `pip check`,
imported FormulaFence 0.139.0 from `site-packages`, confirmed
`openpyxl.DEFUSEDXML`, produced no baseline findings, and emitted only the
expected `FF010`/`FF062` pair for each hostile fixture.

## Power Query nested XML structural bounds — 2026-07-26

Version 0.138.0 closes an allocation path inside Power Query `DataMashup`
Custom XML. FormulaFence already bounds the outer Custom XML item, but the
documented length-prefixed stream can carry independent metadata and
formula-firewall permission XML after Base64 decoding. Those payloads were
previously tree parsed privately without an element ceiling.

FormulaFence now streams each decoded metadata or permission document before
private parsing, with a 32,768-element per-document limit and a 65,536-element
aggregate across the Power Query scan. A successfully parsed overage becomes
visible `FF010` plus `FF024` coverage evidence; malformed input keeps its
existing parser diagnostic. The limits are CI allocation boundaries, not Power
Query validity rules.

Two controlled fixtures inserted 1,000,000 opaque direct children in the
decoded `LocalPackageMetadataFile` and `PermissionList` documents. The nested
XML payloads expanded to 5,000,641 and 5,000,272 bytes; each resulting
`customXml/item1.xml` member was 6,669,121 bytes, compressed to 17,795 and
17,852 bytes. In the same source environment, the unguarded 0.137.0 reader
completed them in 13.966945 seconds / 415,808 KiB and 11.693012 seconds /
402,960 KiB. The 0.138 source completed them in 0.260228 seconds / 63,524 KiB
and 0.258095 seconds / 58,508 KiB without materializing the hostile tree.
Normal baseline-to-baseline diffs had no findings or changes; each hostile
report contained `FF010` and `FF024`. The source suite passed 995 tests in
142.23 seconds with Ruff, bytecode compilation, and diff whitespace checks
clean. The exact final wheel completed the metadata and permission cases in
0.314384 seconds / 62,312 KiB and 0.310034 seconds / 57,084 KiB; its SHA-256
is `2ecbaa81180dfb388f3268eef5c32375f1b95c13bf32d42362b37ce82015b036`.
A fresh Python 3.12 environment installed that wheel, passed `pip check`,
imported FormulaFence 0.138.0 from `site-packages`, confirmed
`openpyxl.DEFUSEDXML`, and produced no normal findings plus `FF010`/`FF024`
for each hostile fixture.

## XML Maps, signatures, Python, and Rich Data structural bounds — 2026-07-26

Version 0.137.0 closes four compact-allocation paths in raw OOXML inventories:
SpreadsheetML XML Maps/table bindings, OPC package-signature envelopes,
Python-in-Excel package XML, and Excel Rich Data package XML. Each inventory
already imposed a 16 MiB per-part, 64 MiB aggregate, and 512-part byte/count
limit, but a highly compressible XML part could still contain enough elements
to allocate a large private tree before its scanner compared it.

The release streams each materialized package XML part before tree parsing,
with a 32,768-element per-part limit and a 65,536-element aggregate limit for
each inventory. A successfully parsed overage becomes explicit `FF010` plus
the relevant `FF049`, `FF050`, `FF051`, or `FF065` coverage evidence; malformed
or unreadable input keeps its established parser diagnostic. Package-signature
certificate and VBA-signature binaries remain byte-bounded, never XML parsed.
Rich Data now streams only `vm` and `r` attributes from worksheet cells after
the shared reader preflight, avoiding a second full worksheet tree.

Four controlled fixtures each inserted 1,000,000 opaque direct children into a
selected raw XML root. The expanded parts were 5,000,776 bytes for XML Maps,
5,000,726 for a package XMLDSIG envelope, 5,000,562 for Python-in-Excel, and
5,000,280 for Rich Data; their compressed members were 7,772, 7,691, 7,601,
and 7,516 bytes respectively. In the same Python 3.12 environment, the prior
0.136 source completed those loads in 5.992546 seconds / 418,768 KiB, 3.262024
seconds / 382,864 KiB, 7.227243 seconds / 660,168 KiB, and 5.476918 seconds /
503,312 KiB. The 0.137 source completed them in 0.425491 seconds / 37,992 KiB,
0.458029 seconds / 38,540 KiB, 0.408754 seconds / 38,128 KiB, and 0.466614
seconds / 38,640 KiB without materializing the hostile tree.

The exact final wheel completed the same four loads in 0.488242 seconds /
36,316 KiB, 0.534549 seconds / 36,708 KiB, 0.431343 seconds / 36,716 KiB, and
0.513357 seconds / 36,320 KiB. Normal baseline-to-baseline diffs had no
findings or changes; the hostile reports contained `FF010`/`FF049`,
`FF010`/`FF050`, `FF010`/`FF065`, and `FF010`/`FF051`, respectively. The final
wheel SHA-256 is
`ee5317b6ea93a374c6a84715b7040ad5dde752819f17782ee3897117f5c42d16`.
The source suite passed 985 tests in 138.92 seconds with Ruff, bytecode
compilation, and artifact metadata checks clean. A fresh Python 3.12
environment installed that exact wheel, passed `pip check`, imported
FormulaFence 0.137.0 from `site-packages`, and confirmed `openpyxl.DEFUSEDXML`.

## Legacy Note and embedded-control XML structural bounds — 2026-07-26

Version 0.136.0 closes compact-allocation paths in two raw OOXML inventories.
Traditional Excel Notes require a SpreadsheetML comments part and can carry
visibility/layout in a VML `legacyDrawing`; FormulaFence recursively
canonicalizes both privately before ordinary cell inspection. The
embedded-control inventory independently reads worksheet control markup, ActiveX
`ocx` persistence XML, form-control properties, and all relationship-selected
legacy VML drawings so it can exclude `Note` `ClientData` while retaining form
controls. Microsoft's [Comment](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.comment?view=openxml-3.0.1),
[LegacyDrawing](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.legacydrawing?view=openxml-3.0.1),
and [`ocx` persistence](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oe376/b30a660a-95eb-4716-b201-a46aae788610)
definitions establish those persisted surfaces; they do not make FormulaFence's
allocation ceilings Excel-file-validity rules.

Both raw gateways now stream complete XML structure before calling their private
tree parsers. Each part permits 32,768 elements and each complete scan permits
65,536; the pre-existing 16 MiB per-part, 64 MiB aggregate, and 512-part
byte/count limits remain in force. A successfully parsed structural overage is
not silently skipped: the Note path emits visible `FF010`/`FF046` coverage
evidence and the embedded-control path emits `FF010`/`FF029`. The VML path has
both independent gates because the two inventories each previously materialized
the same tree; malformed or unreadable input keeps its established diagnostic.

Raw-ZIP fixtures generated without FormulaFence's readers added 1,000,000
opaque direct children to otherwise ordinary Note package parts. The Comments
fixture was a 15,086-byte workbook whose `xl/comments/comment1.xml` expanded to
5,000,277 bytes (7,509 compressed); the VML fixture was 15,131 bytes with
`xl/drawings/commentsDrawing1.vml` at 5,001,035 expanded bytes (7,877
compressed). In the same Python 3.12 environment, the exact 0.135.0 wheel
completed normal diffs in 8.585594 seconds at 739,200 KiB for Comments and
3.738134 seconds at 209,108 KiB for VML. The candidate source completed the
corresponding fail-closed coverage diffs in 0.588982 seconds at 41,604 KiB and
0.634418 seconds at 45,216 KiB, respectively, without materializing either
opaque tree. The exact 0.136.0 wheel completed them in 0.706144 seconds at
40,576 KiB and 0.737067 seconds at 41,696 KiB, respectively. It passed a
normal baseline-to-baseline CLI diff; the hostile Comment report contained
`FF010`/`FF046`, while the shared VML report contained `FF010`/`FF029`/`FF046`
and both structural coverage warnings. Focused regressions prove
fail-before-tree-materialization for Comments, ActiveX, form-control properties,
and shared VML; they also cover nested opaque descendants, aggregate budgets,
exact capacity, default overage, and report findings. The complete 966-test
suite passed in 134.27 seconds with Ruff and bytecode compilation clean. A
fresh Python 3.12 environment installed the exact wheel (SHA-256
`3598ee5dfd436a0baea40a7b45c3bd59224ed88cbe0ecf31c2b94bc6146b7781`), passed
`pip check`, imported FormulaFence 0.136.0 from `site-packages`, and confirmed
`openpyxl.DEFUSEDXML`.

## Shared Worksheet DrawingML semantic-preflight bounds — 2026-07-26

Version 0.135.0 closes a shared compact-allocation path across Worksheet
DrawingML readers. FormulaFence's shape, native-image, in-content Office Web
Add-in, and worksheet-chart boundaries can all reach an internal DrawingML
target before the ordinary workbook reader starts. The later chart scanner had
its own structural guard, but the earlier shape and image paths could already
materialize the same opaque XML tree. Microsoft's Open XML SDK documents the
Worksheet Drawing root as `xdr:wsDr` and its shape surface as `xdr:sp` in the
[Drawing.Spreadsheet namespace](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.spreadsheet?view=openxml-3.0.1);
that makes a relationship-selected drawing worth inspecting, but does not turn
FormulaFence's allocation ceiling into an Excel-file-validity rule.

The semantic-reader preflight now follows direct internal transitional or
Strict worksheet `drawing` relationships, deduplicates their XML targets, and
streams complete structure before any materializing DrawingML or workbook
reader can run. It permits 32,768 elements per target and 65,536 across those
selected targets. A successfully parsed overage gets the stable safety-preflight
error and CLI status 2 rather than a partial report. Missing, malformed, or
non-XML optional targets retain their existing scanner coverage diagnostics,
and orphan DrawingML parts are deliberately outside the relationship-selected
boundary.

A raw-ZIP fixture generated without FormulaFence's reader contained a
35,274-byte workbook with 1,000,000 opaque direct children beneath a worksheet
DrawingML root (`xl/drawings/drawing1.xml`: 14,001,804 expanded bytes and
27,894 compressed bytes). In the same Python 3.12 environment, the exact
0.134.0 wheel completed a normal diff in 9.913578 seconds at 143,444 KiB,
eventually emitting a late chart coverage warning after the shape/image paths
had already handled the tree. The exact 0.135.0 wheel rejects the same workbook
in 0.584261 seconds at 39,464 KiB before any metadata reader runs; it emits no
JSON report because the command exits 2 with the stable Worksheet-DrawingML
safety error. Focused archive-safety regressions cover direct and nested opaque
descendants, Strict relationships, aggregate targets, exact/default capacities,
orphan scope, and fail-before-reader behavior. The complete 955-test source
suite passed in 132.53 seconds, with Ruff, bytecode compilation, and `git diff
--check` clean. A fresh Python 3.12 environment installed the exact wheel
(SHA-256 `621e1e31ecc41d8c2015af0416626fc35247e3435a9ca1ae7f6da1304974de5e`),
passed `pip check`, imported FormulaFence 0.135.0 from `site-packages`,
confirmed `openpyxl.DEFUSEDXML`, completed a normal CLI diff, and retained the
expected hostile status-2/empty-report outcome.

## Threaded-comment and Persons XML structural bounds — 2026-07-26

Version 0.134.0 closes the compact allocation path in FormulaFence's modern
comment scanner. It follows worksheet-associated Threaded Comments and the
workbook-associated Persons part, then recursively canonicalizes private
comment/reply, extension, and person material before ordinary cell inspection.
Its existing 16 MiB per-part, 64 MiB aggregate, and 512-part byte/count limits
did not bound the XML objects created by those private trees.

Microsoft specifies that a [Threaded Comments part](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/66e1875d-c60a-48eb-bf88-41066d45fea8)
is XML associated with one worksheet through a threaded-comment relationship,
while a [Persons part](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/1a170d26-42a2-46f0-b2b6-0ff1dec1c344)
is XML associated with the workbook through a person relationship. That makes
their stored review and identity material worth comparing, but does not make
FormulaFence's element count an Excel-file-validity rule. FormulaFence now
streams each comment/person XML part before tree materialization, allowing
32,768 elements per part and 65,536 across the complete scan. A successfully
streamed overage becomes explicit `FF010`/`FF045` coverage evidence; malformed
or unreadable input keeps its established parser diagnostic. After raw
inspection, FormulaFence removes only these relationships from its temporary
ordinary-reader copy, so the workbook reader cannot re-materialize a rejected
tree while the original package and private raw evidence remain intact.

A raw-ZIP fixture generated without FormulaFence's reader contained a
10,565-byte workbook with 100,000 opaque direct Threaded Comments children
(1,400,899 bytes of `xl/threadedComments/threadedComment1.xml`, 3,152 bytes
compressed). In the same Python 3.12.3 environment, the exact 0.133.0 wheel
completed in 1.263552 seconds at 115,312 KiB despite already producing an
unsupported-metadata warning. The exact 0.134.0 wheel completed in 0.671988
seconds at 39,892 KiB before materializing that tree, recorded one unrecognized
threaded-comment coverage entry, and emitted the structural coverage warning.
The normal installed CLI diff exited 0; the baseline-to-fixture diff emitted
`FF010` and `FF045` and exited 1 with `--fail-on medium`. The complete 947-test
source suite passed in 131.19 seconds, with Ruff, bytecode compilation,
`git diff --check`, the 35-test action-contract suite, and `twine check` clean.
A fresh Python 3.12 environment installed the exact wheel (SHA-256
`1590eaf9a1becb7b3d1778e527b815a95e8502a1faa9ec4f3f6b43e2b895b0ec`),
confirmed FormulaFence 0.134.0 with `openpyxl.DEFUSEDXML` enabled, and retained
the expected normal and hostile CLI outcomes.

## Workbook Theme XML structural bounds — 2026-07-26

Version 0.133.0 closes the compact allocation path in FormulaFence's
workbook-level DrawingML Theme scanner. The scanner follows the raw workbook
Theme binding, privately canonicalizes Theme XML, and follows direct
Theme-image relationships before ordinary cell inspection. Its existing 16 MiB
per-part, 64 MiB aggregate, and 512-part byte/count limits did not bound the
XML objects created by the Theme tree or its recursive canonicalizer.

Microsoft's Open XML documentation says that a SpreadsheetML package can have
zero or one [Theme part](https://learn.microsoft.com/en-us/office/open-xml/presentation/structure-of-a-presentationml-document),
bound from the Workbook, and that the part contains the colour, font, and
format schemes that affect spreadsheet cell contents and charts. That makes the
stored control material worth comparing, but does not make FormulaFence's
element count an Excel-file-validity rule. FormulaFence now streams Theme and
Theme-relationship XML before tree materialization, allowing 32,768 elements
per XML part and 65,536 across the complete Theme scan. A successfully streamed
overage becomes explicit `FF053` coverage evidence; malformed or unreadable
input keeps its established parser diagnostic. Direct Theme-image payloads stay
byte-bounded rather than being interpreted as XML.

A raw-ZIP fixture generated without FormulaFence's reader contained a
9,591-byte workbook with 100,000 opaque direct Theme XML children
(1,210,322 bytes of `xl/theme/theme1.xml`). In the same Python 3.12.3
environment, the exact 0.132.0 wheel completed in 1.090010 seconds at
112,240 KiB and reported no unrecognized Theme coverage. The exact 0.133.0
wheel completed in 0.683019 seconds at 41,512 KiB before materializing that
tree, recorded explicit unrecognized Theme coverage, and emitted the structural
coverage warning. The normal installed CLI diff exited 0; the
baseline-to-fixture diff emitted `FF010` and `FF053` and exited 1 with
`--fail-on medium`. The complete 937-test source suite passed in 129.68
seconds, with Ruff, bytecode compilation, `git diff --check`, and `twine check`
clean. A fresh Python 3.12 environment installed the exact wheel (SHA-256
`206ae032803648b08b8c1df4ea23d7d06f94ac4866d8937e3e2bb8ac70726829`),
confirmed FormulaFence 0.133.0 with `openpyxl.DEFUSEDXML` enabled, and retained
the expected normal and hostile CLI outcomes.

## Custom workbook data-store XML structural bounds — 2026-07-26

Version 0.132.0 closes the compact allocation path in FormulaFence's generic
Custom XML and persisted add-in-state scanner. The scanner reads generic Custom
XML data/property/schema parts, Custom Data property parts, custom document
properties, and their package relationships before an ordinary workbook reader
can omit them. Its existing 16 MiB per-part, 64 MiB aggregate, and 512-part
byte/count limits did not bound the XML objects created by its private recursive
canonicalizer. A later Power Query discovery pass also used to revisit arbitrary
Custom XML items while looking for `DataMashup` definitions.

Microsoft describes Custom XML parts as a place to store [arbitrary XML data in
Office documents](https://learn.microsoft.com/en-us/visualstudio/vsto/custom-xml-parts-overview?view=visualstudio),
and its [Office Scripts API](https://learn.microsoft.com/en-us/javascript/api/office-scripts/excelscript/excelscript.customxmlpart?view=office-scripts)
exposes whole-part `getXml` / `setXml` operations.
That means an element count is a FormulaFence reader-allocation and coverage
boundary, not an Excel-file-validity rule. FormulaFence now streams every
custom-state XML member before tree materialization, allowing 32,768 elements
per part and 65,536 across the complete custom-state scan. A successfully
streamed overage becomes visible `FF052` coverage evidence; malformed or
unreadable input retains its established parser diagnostic. Opaque binary
Custom Data remains byte-bounded rather than being parsed as XML. The Power
Query scanner consumes only `DataMashup` items safely classified by this same
bounded pass, so a rejected generic tree cannot be materialized a second time.

A raw-ZIP fixture generated without FormulaFence's reader contained an
11,493-byte workbook with 100,000 opaque direct Custom XML children
(1,400,203 bytes of Custom XML). In the same Python 3.12.3 environment, the
exact 0.131.0 wheel completed in 0.8529 seconds at 94,776 KiB and reported no
unrecognized custom-data-store coverage. The exact 0.132.0 wheel completed in
0.1860 seconds at 36,024 KiB before materializing that tree, recorded explicit
unrecognized custom-data-store coverage, and emitted one structural warning.
The normal installed CLI diff exited 0; the baseline-to-fixture diff emitted
`FF010` and `FF052` and exited 1 with `--fail-on medium`. The complete
929-test source suite passed in 128.02 seconds, with Ruff, bytecode
compilation, `git diff --check`, the 35-test action-contract suite, and
`twine check` clean. A fresh Python 3.12 environment installed the exact wheel
(SHA-256 `c34d9c5f0caba6d7ad32a3f377c4fac0e8f5ac9bb5812d1e7304c1c8284d162e`),
confirmed FormulaFence 0.132.0 with `openpyxl.DEFUSEDXML` enabled, and retained
the expected normal and hostile CLI outcomes.

## Slicer and Timeline cache XML structural bounds — 2026-07-26

Version 0.131.0 closes the corresponding compact allocation path in
FormulaFence's Slicer and Timeline cache scanner. The raw scanner compares
private filter state, source bindings, and cache definitions by recursively
canonicalizing each cache XML tree. Its existing 16 MiB per-part, 64 MiB
aggregate, and 512-part byte/count limits did not bound the number of XML
objects that a compact cache could create.

FormulaFence now streams every Slicer and Timeline cache member before reading
it into the private tree, allowing 16,384 XML elements per part and 32,768
across a complete cache scan. [Excel documents 10,000 items displayed in a
filter drop-down list](https://support.microsoft.com/en-US/Excel/excel-specifications-and-limits),
so the structural capacity deliberately leaves room above that display limit
without treating an overage as invalid. A successfully streamed overage instead
becomes explicit `FF032` filter-cache coverage evidence before the recursive
scanner runs; malformed or unreadable input keeps its established full-parser
diagnostic.

A raw-ZIP fixture generated without FormulaFence's reader contained a
13,633-byte workbook with 100,000 opaque direct Slicer-cache children
(1,400,536 bytes of cache XML). In the same Python 3.12.3 environment, the
exact 0.130.0 wheel completed the fixture in 0.7380 seconds at 85,536 KiB and
reported no unrecognized Slicer/Timeline coverage. The 0.131.0 source candidate
completed in 0.1613 seconds at 35,304 KiB, before the complete cache tree was
materialized, and recorded one structural coverage warning. A normal candidate
CLI diff completed successfully; the baseline-to-fixture diff emitted `FF032`
and returned status 1 with `--fail-on medium`. The complete 920-test source
suite passed in 126.68 seconds, with Ruff, bytecode compilation, and
`git diff --check` clean. A fresh Python 3.12 environment installed the exact
candidate wheel, confirmed FormulaFence 0.131.0 with `openpyxl.DEFUSEDXML`
enabled, and completed a normal CLI diff; its fixture diff retained one
structural coverage warning, emitted `FF032`, and returned the expected status
1 at `--fail-on medium`.

## PivotTable XML structural bounds — 2026-07-26

Version 0.130.0 closes two compact allocation paths in FormulaFence's
PivotTable handling. The raw scanner recursively canonicalizes PivotTable view
and cache-definition XML to compare layout, cache schema, and private shared
items without exposing them. The ordinary workbook reader can then follow the
same cache and view bindings a second time. Existing 16 MiB per-part, 64 MiB
aggregate, and 512-part byte/count limits did not bound either materialized
tree.

FormulaFence now streams every PivotTable view and cache-definition part before
reading it into a tree, with 32,768 elements per part and 65,536 across a
package scan. [Microsoft documents up to 1,048,576 unique items per PivotTable
field](https://support.microsoft.com/en-US/Excel/excel-specifications-and-limits),
so a CI reader cannot assume a small catalog is invalid; a successfully
streamed overage instead becomes explicit PivotTable coverage evidence before
the private tree is built. Malformed or unreadable input retains the established
full-parser diagnostic. After raw inspection, the temporary source for
`openpyxl` removes only PivotTable cache and view bindings, retaining ordinary
cell analysis without letting that reader reparse the package graph.

Independent raw-ZIP fixtures measured the exact 0.129.0 wheel and the 0.130.0
source candidate in the same Python 3.12 environment. An 11,968-byte workbook
with 100,000 opaque PivotTable-view children (1,400,733 bytes of view XML)
completed in 1.1149 seconds at 88,576 KiB through the prior wheel with no
unrecognized PivotTable coverage; the candidate records a structural coverage
gap in 0.5067 seconds at 41,344 KiB. An 11,766-byte workbook containing
100,000 valid cached shared items (1,300,855 bytes of cache-definition XML)
took 7.1374 seconds at 207,188 KiB through the prior wheel with no structural
coverage marker; the candidate completes in 0.5825 seconds at 44,944 KiB with
the coverage gap visible. The complete 913-test source suite passed in bounded
runner batches (127.17 seconds aggregate), with Ruff, bytecode compilation,
and `git diff --check` clean. A fresh Python 3.12 environment installed the
exact 0.130.0 wheel, confirmed `openpyxl.DEFUSEDXML`, and completed a normal
CLI diff. The hostile PivotTable view retained one structural coverage warning;
its installed diff returned `FF031` and the expected status 1 with
`--fail-on medium`. The installed reader-isolation check rejected every
underlying PivotTable cache, record, and view parser while raw PivotTable
evidence remained available.

## Chart XML structural bounds — 2026-07-26

Version 0.129.0 closes a compact allocation path in FormulaFence's private
chart scanner. The scanner follows legacy chart, Office 2016+ ChartEx,
chart-host DrawingML, and chart-overlay parts, then recursively canonicalizes
private XML so relationship IDs and cached series data can be compared without
being exposed. Its existing 16 MiB per-part, 64 MiB aggregate, and 512-part
byte/count limits did not bound a broad but compact XML tree.

Before a chart XML part is read into a tree, FormulaFence now streams it with a
32,768-element per-part limit and a 65,536-element package budget. The larger
capacity is deliberate: [Microsoft documents chart data-point capacity as
memory-limited](https://support.microsoft.com/en-US/Excel/excel-specifications-and-limits)
rather than as a small fixed catalog. A successfully streamed
overage becomes chart coverage evidence before the recursive scanner starts;
malformed or unreadable input retains the established full-parser diagnostic.

Independent raw-ZIP fixtures measured the exact 0.128.0 wheel and the 0.129.0
source candidate in the same Python 3.12 environment. A 12,453-byte workbook
with 100,000 opaque chart children (1,401,828 bytes of chart XML) completed in
0.9359 seconds at 84,948 KiB through the prior wheel with no unrecognized chart
coverage. The candidate reports one structural chart coverage gap in 0.3739
seconds at 46,756 KiB, before the chart XML tree is materialized. The complete
905-test source suite passed in bounded runner batches (129.41 seconds
aggregate), with Ruff, bytecode compilation, and `git diff --check` clean. A
fresh Python 3.12 environment installed the exact 0.129.0 wheel, confirmed
`openpyxl.DEFUSEDXML`, and completed a normal CLI diff. The hostile chart
profile retained one structural coverage warning; its installed diff returned
`FF030` and the expected status 1 with `--fail-on medium`.

## Package relationship-part structural bounds — 2026-07-26

Version 0.128.0 closes a shared allocation path across FormulaFence's raw OOXML
metadata scanners. Version 0.127.0 already bounded the workbook relationship
catalog, but many scanners also parse optional relationship parts beside
task-pane, drawing, control, and other package members. A compact workbook can
therefore contain a very large valid relationship tree that has no reader-visible
cell and is not represented as an unrecognized control.

The semantic-reader preflight now streams every package member ending in
`.rels` before any raw scanner or the complete `openpyxl` reader runs. It allows
4,096 XML elements per relationship part and 16,384 across all relationship
parts, counting every element including the root and opaque nested descendants.
An over-budget part returns the stable input safety error. A malformed optional
part still reaches its established coverage warning after the small bounded
stream fails to parse, so valid-but-unrecognized metadata remains observable.

Independent raw-ZIP fixtures measured the exact 0.127.0 wheel and the 0.128.0
source candidate in the same Python 3.12 environment. The 13,334-byte Office
Web Add-in definition fixture and 13,325-byte task-pane fixture each contain
100,000 empty relationship records (about 2.0 MiB of relationship XML). The
previous wheel completed in 1.094 seconds at 83,960 KiB and 1.127 seconds at
82,016 KiB respectively, with no unrecognized relationship coverage. The
candidate rejects the definition fixture in 0.0089 seconds at 34,788 KiB and
the task-pane fixture in 0.0092 seconds at 34,784 KiB, before its private
Office Web Add-in scanner begins. The archive suite passed 182 tests, and all
898 source tests passed in bounded runner batches (130.08 seconds aggregate),
with Ruff and `git diff --check` clean. A fresh Python 3.12 environment
installed the exact 0.128.0 wheel, confirmed `openpyxl.DEFUSEDXML`, and
completed a normal CLI diff. Both relationship fixtures returned the expected
input-error status (2) through the installed `profile` command.

## Office Web Add-in task-pane and definition XML bounds — 2026-07-26

Version 0.127.0 closes two compact allocation paths in FormulaFence's private
Office Web Add-in scanner. Task-pane and `webextension` definition parts have
16 MiB per-part, 32 MiB aggregate, and 64-part byte/count limits, but their raw
XML is then recursively canonicalized to preserve unsupported configuration
without exposing it in reports. A small compressed part can therefore still
make a CI worker construct a broad or deeply nested XML tree even when no
worksheet cell is populated.

Before either scanner reads the XML payload, it now streams the ZIP member and
counts elements under the existing nesting limit. Each task-pane or definition
part permits 4,096 elements, including its root, and their package scan permits
16,384 elements in aggregate. An over-budget part becomes unrecognized Office
Web Add-in coverage with an explicit warning; malformed or unreadable input
still reaches the established parser so its normal diagnostic remains visible.
Exact per-part capacity remains covered, and the aggregate counter prevents a
many-small-parts bypass.

Independent raw-ZIP fixtures measured the exact 0.126.0 wheel and the 0.127.0
source candidate in the same Python 3.12 environment. An 11,021-byte workbook
with 100,000 direct opaque task-pane children (1,300,374 bytes of XML) loaded
through 0.126.0 in 0.586 seconds at 97,840 KiB resident; the candidate marks it
unrecognized in 0.135 seconds at 34,116 KiB. An 11,041-byte workbook with
100,000 direct opaque `webextension` children (1,300,917 bytes of XML) loaded
in 0.579 seconds at 98,104 KiB and now stops in 0.135 seconds at 34,124 KiB.
The nested equivalents likewise moved from 0.597/97,592 KiB and 0.573/97,704
KiB to about 0.145 seconds and 34 MiB. The completed Office Web Add-in diff
suite passed **496 tests in 79.87 seconds**, with Ruff and `git diff --check`
clean. The completed source suite passed **892 tests in 123.09 seconds**. A
fresh Python 3.12 environment installed the exact candidate wheel, confirmed
FormulaFence 0.127.0 with `openpyxl.DEFUSEDXML` enabled, and completed a normal
CLI comparison. All four independently generated direct/nested task-pane and
definition fixtures produced exactly one structural warning in the installed
reader and the expected `FF010` coverage finding; with `--fail-on medium`, each
installed CLI comparison returned status 1.

## RibbonX structural XML bounds — 2026-07-26

Version 0.126.0 closes a compact allocation path in FormulaFence's private
RibbonX customization scanner. The existing 16 MiB per-part, 32 MiB aggregate,
and eight-part limits bound raw `customUI` payload size and count, but a small
compressed XML part can still expand into a broad or deeply nested element tree.
The scanner then creates a recursive canonical fragment for unknown controls,
so a workbook can consume substantial memory without a populated worksheet.

Before the scanner reads a bounded RibbonX part into memory, it now streams the
ZIP member and counts every XML element while enforcing the existing nesting
limit. Each part permits 4,096 elements, including its root. A successfully
streamed part that exceeds either structural bound becomes unrecognized RibbonX
coverage with an explicit warning; malformed or unreadable input still reaches
the established parser so its normal diagnostic is preserved. An exact
4,096-element customization remains fully covered, and both direct opaque
controls and one opaque nested subtree consume the same boundary.

Independent raw-ZIP fixtures measured the exact 0.125.0 wheel and the 0.126.0
source candidate in the same Python 3.12 environment. A 10,158-byte workbook
with 100,000 direct opaque RibbonX children (1,400,497 bytes of `customUI` XML)
loaded through 0.125.0 in 0.559 seconds at 93,568 KiB resident; the candidate
marks it unrecognized in 0.137 seconds at 33,352 KiB. A 9,976-byte workbook
with one opaque RibbonX child containing 100,000 nested entries (1,300,522
bytes of XML) loaded in 0.578 seconds at 92,872 KiB and now stops in 0.137
seconds at 34,092 KiB. The completed source suite passed **885 tests in 124.65
seconds**, with Ruff and `git diff --check` clean. A fresh Python 3.12
environment installed the exact candidate wheel and confirmed its FormulaFence
0.126.0 version with `openpyxl.DEFUSEDXML` enabled, then completed a normal CLI
comparison. Both independently generated RibbonX fixtures produced exactly one
structural warning in the installed reader and the expected `FF010` coverage
finding; with `--fail-on medium`, each installed CLI comparison returned status
1.

## Legacy Custom View descendant bounds — 2026-07-26

Version 0.125.0 closes the remaining compact subtree path in the legacy Custom
View scanner. Version 0.124.0 bounded Custom View page-break records and
containers, but a supported `<customSheetViews>/<customSheetView>` hierarchy can
also carry an unknown direct child or a single opaque nested subtree. FormulaFence
preserves that unsupported XML privately through a recursive canonical signature;
one view declaration could therefore still make a small package allocate a long
tuple tree even though its page-break catalog was within bounds.

The semantic-reader preflight now counts every descendant below each direct
`customSheetView` that the raw scanner will enter from a transitional or Strict
`customSheetViews` container. It permits 4,096 descendants in aggregate, in
addition to the existing 4,096 direct-view declaration boundary. This covers
standard views, Strict views, and alternate-namespace views that reach the
opaque signature path, while foreign `customSheetViews` containers remain
outside that raw scanner path. The bound is large enough for the published
2,052 row-plus-column page-break allowance plus its two containers; an exact
4,096-descendant opaque fixture remains accepted.

Independent raw-ZIP fixtures measured the exact 0.124.0 wheel and the 0.125.0
source candidate in the same Python 3.12 environment. An 8,978-byte workbook
with 100,000 direct opaque Custom View children (1,401,837 bytes of sheet XML)
loaded through 0.124.0 in 4.197 seconds at 111,704 KiB resident; the candidate
rejects it in 0.018 seconds at 33,700 KiB. An 8,798-byte workbook with one
opaque Custom View child containing 100,000 nested entries (1,301,862 bytes of
sheet XML) loaded in 3.773 seconds at 100,188 KiB and now rejects in 0.019
seconds at 33,660 KiB. Focused archive-safety and version coverage passed
**177 tests in 6.26 seconds**; the completed source suite passed **879 tests in
123.22 seconds**, with Ruff, bytecode compilation, and `git diff --check`
clean. A fresh Python 3.12 environment installed the exact candidate wheel,
confirmed `FormulaFence 0.125.0` and `openpyxl.DEFUSEDXML`, completed a normal
CLI comparison, and returned the normal input-error exit status (2) for both
independently generated 100,000-entry Custom View subtree fixtures.

## Legacy Custom View page-break catalog bounds — 2026-07-26

Version 0.124.0 closes the equivalent allocation paths inside legacy Excel
Custom Views. Version 0.123.0 bounded direct worksheet
`<rowBreaks>`/`<colBreaks>` catalogs, but FormulaFence's separate Custom View
scanner also parses direct break containers beneath a supported
`<customSheetViews>/<customSheetView>` hierarchy. It constructs break
signatures for standard and Strict SpreadsheetML views, and hashes the same
subtrees when an alternate-namespace view or break container must be retained
as opaque coverage evidence. That left a small valid workbook able to bypass
the ordinary worksheet break counter without a populated cell.

The semantic-reader preflight now shares its 2,052-direct-child and
4,096-container aggregate budgets with that supported Custom View path. It
recognizes transitional and Strict `<customSheetViews>` containers, then counts
the local `customSheetView`, `rowBreaks`, and `colBreaks` shapes the raw scanner
will handle. Every direct break-container child consumes the record budget,
including unexpected names and alternate namespaces. This is deliberately
scoped to the raw scanner's supported container hierarchy; a foreign
`customSheetViews` container remains outside that scanner path.

Independent raw-ZIP fixtures measured the exact 0.123.0 wheel and the 0.124.0
source candidate in the same fresh Python 3.12 environment. A 19,396-byte
workbook with 100,000 Custom View row-break records (4,501,791 bytes of sheet
XML) loaded through 0.123.0 in 9.043 seconds at 109,628 KiB resident; the
candidate rejects it in 0.012 seconds at 33,272 KiB. A 9,138-byte workbook with
100,000 empty Custom View row-break containers (1,501,764 bytes of sheet XML)
loaded in 3.981 seconds at 74,332 KiB and now rejects in 0.023 seconds at
33,852 KiB. Focused standard, Strict, opaque-path, aggregation, and exact-limit
regressions passed alongside the completed source suite: **869 tests in 121.67
seconds**. A fresh Python 3.12 environment installed the exact candidate wheel,
confirmed `FormulaFence 0.124.0` and `openpyxl.DEFUSEDXML`, completed a normal
CLI comparison, and returned the normal input-error exit status (2) for both
independently generated 100,000-entry Custom View fixtures.

## Worksheet page-break catalog bounds — 2026-07-26

Version 0.123.0 closes two compact allocation paths in worksheet print
metadata. Excel's
[published limits](https://support.microsoft.com/en-US/Excel/excel-specifications-and-limits)
allow 1,026 horizontal and 1,026 vertical page breaks. `openpyxl` turns each
direct transitional SpreadsheetML `<brk>` record into a page-break object, while
FormulaFence's raw print-layout scanner retains every direct child of
`<rowBreaks>` or `<colBreaks>` as either a modeled break or coverage evidence.
Repeated empty containers are also retained by the raw scanner. Neither path
needs a populated cell, so the ordinary worksheet-cell budget cannot constrain
them.

The semantic-reader preflight now permits 2,052 direct break-container children
in aggregate across selected ordinary worksheet parts: one worksheet's complete
published row-plus-column allowance. It separately permits 4,096 direct
`rowBreaks`/`colBreaks` containers, blocking container fragmentation before raw
XML scanners can build a large direct-child list. The record counter follows
both allocation paths: every direct child of an ordinary or Strict SpreadsheetML
container consumes it, including unexpected or alternate-namespace children
that FormulaFence must retain as raw coverage evidence. Foreign-namespace
containers remain outside the supported reader/scanner paths.

Independent raw-ZIP fixtures established both costs using the exact 0.122.0
wheel in Python 3.12, then the 0.123.0 source candidate in that same
environment. A 17,127-byte workbook with 100,000 valid row-break records
(4,200,530 bytes of worksheet XML) loaded in 9.762 seconds at 70,900 KiB
resident and now rejects in 0.012 seconds at 33,676 KiB. A 7,166-byte workbook
with 100,000 empty direct row-break containers (1,200,466 bytes of worksheet
XML) loaded in 4.077 seconds at 39,036 KiB and now rejects in 0.015 seconds at
33,796 KiB. Focused archive-safety and version coverage passed **157 tests in
5.50 seconds**; the completed source tree passed **859 tests in 122.86
seconds**, with Ruff, bytecode compilation, and `git diff --check` clean. A
fresh Python 3.12 environment installed the exact release-candidate wheel,
confirmed `FormulaFence 0.123.0` and `openpyxl.DEFUSEDXML`, completed a normal
CLI comparison, and returned the normal input-error exit status (2) for both
independently generated 100,000-entry page-break fixtures.

## Column-dimension declaration and container bounds — 2026-07-26

Version 0.122.0 closes two compact allocation paths around worksheet columns.
Excel allows up to
[16,384 columns](https://support.microsoft.com/en-US/Excel/excel-specifications-and-limits),
yet a valid package can repeat the same column declaration without expanding
the grid. `openpyxl` dispatches every transitional SpreadsheetML `<col>` into
its column-dimension parser before FormulaFence can decide whether the width,
style, or visibility is relevant; repeated entries can therefore create costly
parser work even when they overwrite the same final column key. FormulaFence's
raw dimension scanners also retain direct `<cols>` container lists.

The semantic-reader preflight now limits reader-visible `col` declarations to
16,384 and direct `cols` containers to 4,096, each in aggregate across selected
ordinary worksheet parts. The declaration count follows the reader's actual tag
dispatch, so unknown unqualified attributes do not evade it. The container count
follows the raw scanner's direct worksheet-child behavior; foreign namespace
declarations stay outside both supported reader paths. These are separate
CI-oriented cardinality limits, not a claim that a valid Excel grid at its
documented column maximum is unsafe.

Independent raw-ZIP fixtures established both costs. A 20,680-byte workbook
with 100,000 repeated valid `col` declarations (5,400,663 bytes of worksheet
XML) loaded before the gate in 9.028 seconds at 82,776 KiB despite retaining one
final column key; it now rejects in 0.074 seconds at 35,828 KiB. A 7,246-byte
workbook with 100,000 empty direct `cols` containers (1,200,642 bytes of
worksheet XML) loaded in 4.226 seconds at 72,280 KiB and now rejects in 0.014
seconds at 34,744 KiB. A 10,000-declaration fixture remains accepted in 0.848
seconds at 39,680 KiB; exact 16,384-declaration and 4,096-container fixtures
completed in 1.382 seconds and 0.203 seconds, respectively. The completed
source tree passed **847 tests in 120.57 seconds**; focused archive-safety and
version coverage passed **145 tests in 4.59 seconds**, with Ruff, bytecode
compilation, and `git diff --check` clean. A fresh Python 3.12 environment
installed the exact release-candidate wheel, confirmed `FormulaFence 0.122.0`
and `openpyxl.DEFUSEDXML`, completed a normal CLI comparison, and returned the
normal input-error exit status (2) for independently generated 100,000-entry
declaration and container fixtures.

## Empty formatted-row dimension bound — 2026-07-26

Version 0.121.0 closes a compact allocation path in ordinary worksheet loading.
Excel permits a worksheet grid of up to
[1,048,576 rows](https://support.microsoft.com/en-US/Excel/excel-specifications-and-limits),
but populated-cell cardinality does not bound empty row formatting. In
`openpyxl`, parsing a transitional SpreadsheetML `<row>` stores a declaration
for later `RowDimension` construction whenever it has an unqualified attribute
other than `r` or `spans`. A compact run of empty rows with `ht` and
`customHeight` therefore materializes one Python dimension object per row;
FormulaFence's raw worksheet display scanners can revisit the same XML later.

The semantic-reader preflight now allows at most 16,384 such declarations in
aggregate across selected ordinary worksheet parts before either downstream
scanner or the complete reader starts. It streams both supported SpreadsheetML
variants for the raw display boundary, but uses the reader's exact attribute
rule: coordinate-only rows and namespace-qualified extension attributes do not
consume the budget, while an unknown unqualified attribute does. This is a
separate CI-oriented cardinality boundary, not a claim that Excel's grid limit
is unsafe.

Independent raw-ZIP fixtures established the impact. A normal workbook loaded
in 0.041 seconds at 35,152 KiB. A 259,597-byte workbook containing 100,000
empty rows with only height metadata (4,089,411 bytes of worksheet XML) loaded
before the gate in 8.580 seconds at 142,092 KiB. With the gate it rejects before
the workbook reader in 0.074 seconds at 36,076 KiB. A 10,000-row formatted
fixture remains accepted in 0.809 seconds at 45,624 KiB. Regression coverage
exercises configured/default limits, exact-limit acceptance, cross-sheet
aggregation, unknown attributes, and ignored coordinate, namespaced-attribute,
and foreign-namespace cases. The completed source tree passed **836 tests in
122.58 seconds**; focused archive-safety and version coverage passed **134
tests in 4.06 seconds**, with Ruff, bytecode compilation, and `git diff --check`
clean. A fresh Python 3.12 environment installed the exact release-candidate
wheel, confirmed `FormulaFence 0.121.0` and `openpyxl.DEFUSEDXML`, completed a
normal CLI comparison, and returned the normal input-error exit status (2) for
the independently generated 100,000-row fixture.

## Stylesheet catalog bounds — 2026-07-26

Version 0.120.0 closes compact allocation paths in `styles.xml`. Excel's
[published workbook limits](https://support.microsoft.com/en-US/Excel/excel-specifications-and-limits)
list 512 fonts per workbook, 256 fill styles, roughly 200–250 number formats,
and 65,490 unique cell styles. `openpyxl`, however, dispatches stylesheet
containers by local name and builds Python records for every direct child of
its `NestedSequence` containers, including unexpected or alternate-namespace
children. Before this gate, the existing preflight counted only conventional
namespace `cellXfs` children, leaving the rest of the stylesheet catalog and an
alternate-namespace `cellXfs` container outside that cardinality boundary.

FormulaFence now streams the stylesheet before any raw scanner or complete
reader runs. It permits 4,096 repeated known stylesheet containers in aggregate
and 4,096 records in each number-format, font, fill, fill-child, gradient-stop,
border, base-XF, named-style, differential-style, palette, table-style,
table-style-element, and extension catalog. The existing 65,490 effective
`cellXfs` ceiling remains intact. Counts match the reader's local-name and
nested-sequence behavior, so namespace decoration and unexpected direct records
cannot evade the gate.

Independent raw-ZIP fixtures established the impact. A normal workbook loaded
in 0.041 seconds at 35,092 KiB. Adding 100,000 valid font records produced a
285,934-byte package with 13,802,886 bytes of `styles.xml`; before the gate it
loaded in 18.321 seconds at 336,052 KiB. With the gate it rejects before any
scanner or workbook reader in 0.068 seconds at 34,868 KiB. A 10,000-font
fixture previously reached 65,464 KiB in 1.663 seconds. The regression suite
covers every stylesheet list, repeated containers, exact configured and default
font limits, unexpected nested children, and alternate-namespace `cellXfs`.
The completed source tree passed **828 tests in 122.30 seconds**; focused
archive-safety and version coverage passed **126 tests in 3.54 seconds** with
Ruff, bytecode compilation, and `git diff --check` clean. A clean Python 3.12
environment installed the exact release-candidate wheel, confirmed
`FormulaFence 0.120.0` and `openpyxl.DEFUSEDXML`, completed a normal comparison,
and returned the normal input-error exit status (2) for the independently
generated 100,000-font workbook.

## Worksheet control catalog and `sqref` bounds — 2026-07-26

Version 0.119.0 closes three related allocation paths in ordinary worksheet
loading. A [data-validation collection can contain up to 65,534 entries](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/4ec60c5e-de69-4d16-944d-7fab3e45fdff),
while the [interoperability notes](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oi29500/0114969a-4859-443e-ae3e-0c1fdebfcc64)
permit up to 32,767 references in one data-validation `sqref`; the standard
conditional-formatting collection and
each of its `cfRule` sequences are explicitly [unbounded](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/ecd8b657-48f7-410b-b856-c4ad9c2fe127)
and [unbounded](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/98f4ba7a-b8f8-4969-92a0-d671d5d8ca8a),
respectively. Scenario Manager similarly stores an `sqref` sequence plus a
collection of scenarios and per-scenario input-cell records, as shown by the
Open XML [Scenarios](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.scenarios?view=openxml-3.0.1)
and [InputCells](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.inputcells?view=openxml-3.0.1)
definitions.

`openpyxl` converts each whitespace-separated `sqref` token into a `CellRange`
inside a `MultiCellRange`; it also builds one object for each data validation,
conditional-formatting rule, scenario, and scenario input cell. FormulaFence's
raw conditional-formatting and Scenario Manager scanners retain corresponding
private semantic records. Neither ZIP-member counts nor populated-cell counts
bound those allocations.

The semantic-reader preflight now streams direct local-name children in
reader-selected ordinary worksheet parts before raw scanners or the complete
reader begin. It permits at most 4,096 data-validation declarations,
conditional-formatting declarations, conditional-formatting rules, Scenario
Manager containers, scenarios, and scenario input cells. Each `sqref` is capped
at 128 KiB and 4,096 whitespace-separated targets; each control catalog is
capped at 8,192 targets in aggregate. The stored text of a data-validation
`formula1`/`formula2` or conditional-formatting `formula` now follows the
existing 8,192-character formula ceiling. The counters follow `openpyxl`'s
direct local-child behavior, including alternate-namespace children, so a
namespace rewrite cannot sidestep the preflight.

Independent fixtures establish the allocation path. Before the gate, an
18,378-byte workbook with 100,000 repeated data validations (4,600,734 bytes
of worksheet XML) took 9.640 seconds and 172,260 KiB despite collapsing to one
semantic rule. A single 218,389-byte validation `sqref` with 100,000 targets
(688,894 characters) took 1.018 seconds and 77,708 KiB. Ten thousand
conditional-formatting rules in a 34,777-byte package took 2.143 seconds and
64,552 KiB; a 100,000-target conditional-formatting `sqref` took 1.147 seconds
and 86,228 KiB. Ten thousand Scenario Manager declarations took 1.612 seconds
and 69,172 KiB, while 10,000 input-cell records in one scenario took 0.673
seconds and 51,920 KiB.

With the normal gate enabled, those generated inputs rejected before downstream
work in 0.008–0.044 seconds at 34,748–36,820 KiB. Exact-limit fixtures remained
accepted: 4,096 validations, 4,096 conditional-formatting rules, 4,096
scenarios with input cells, and a 4,096-target validation `sqref` completed in
0.061–0.665 seconds at 35,632–48,060 KiB. The completed source tree passed
**807 tests in 120.15 seconds**. Focused archive-safety and version coverage
also passed **105 tests in 2.82 seconds**, along with Ruff, bytecode
compilation, and `git diff --check`. The release wheel and source distribution
passed `twine check`; a clean Python 3.12 environment installed the exact wheel,
confirmed `FormulaFence 0.119.0` and `openpyxl.DEFUSEDXML`, compared a normal
workbook successfully, and returned the normal input-error exit status (2) for
independently generated 4,097-entry data-validation, 4,097-target
conditional-formatting, and 4,097-input-cell Scenario Manager inputs.

## Reader merged-cell geometry bound — 2026-07-26

Version 0.118.0 closes a compact allocation path in ordinary worksheet loading.
The SpreadsheetML [mergeCells interoperability notes](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oi29500/8a2a99c9-bfc5-4d44-8d00-3046b75af83c)
permit an unbounded number of `mergeCell` elements (with Excel allowing up to
4,294,967,294 occurrences). When `openpyxl` binds a worksheet, it does not just
retain a merge reference: it creates an in-memory `MergedCell` for every
coordinate in the range. A tiny sheet part can therefore request an enormous
allocation even when it contains almost no populated cells.

The semantic-reader preflight now streams direct `mergeCell` declarations in
reader-selected ordinary worksheet parts before raw OOXML scanners or the
complete reader begin. It limits the declaration count to 4,096, every range
and their aggregate expanded coordinate area to 100,000 cells, and every
reference attribute to 256 characters. The geometry check accepts the same
sheet-qualified range grammar as the reader and counts matching direct local
children, including an alternate-namespace child, so neither range spelling nor
namespace decoration bypasses the limit.

Independent generated fixtures establish the boundary. A 4,889-byte package
with `A1:ALL100` (100,000 coordinates) remained accepted, taking 0.862 seconds
and 72,524 KiB. A 4,890-byte `A1:ALL101` package (101,000 coordinates) failed
in 0.002 seconds at 34,672 KiB with the range-specific safety error; a
4,895-byte full-worksheet `A1:XFD1048576` package failed in 0.002 seconds at
34,804 KiB before `openpyxl` could expand the grid. The completed source tree
passed **779 tests in 127.54 seconds**. Focused archive-safety and version
coverage also passed **77 tests in 2.32 seconds**, along with Ruff, bytecode
compilation, and `git diff --check`. The release wheel and source distribution
passed `twine check`; a clean Python 3.12 environment installed the exact wheel,
confirmed `FormulaFence 0.118.0` and `openpyxl.DEFUSEDXML`, compared a normal
workbook successfully, and returned the normal input-error exit status (2) for
an independently generated full-grid `A1:XFD1048576` merge declaration.

## Reader view and auxiliary catalog bounds — 2026-07-26

Version 0.117.0 closes the remaining repeated-child paths in `openpyxl`'s
workbook package and FormulaFence's legacy Custom View scanner. The SpreadsheetML
[BookViews format reference](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.bookviews?view=openxml-3.0.1)
explicitly permits an unlimited number of workbook views. `openpyxl` builds
objects for every direct `bookViews` child and each local-named
`functionGroup`, `smartTagType`, and `webPublishObject` entry. Its package
parser has the same repeated-child behavior for `customWorkbookViews`, which
FormulaFence removes from its complete-reader copy; FormulaFence's raw Custom
View scanner then retains a per-view record for every direct custom workbook or
custom sheet view it inspects. None of those lists is bounded by a ZIP-part
count, relationship count, or ordinary worksheet-cell count.

The semantic-reader preflight now limits direct workbook book-view,
custom-workbook-view, function-group, smart-tag-type, and web-publish-object
catalog entries to 4,096 each. It also limits direct legacy custom sheet-view
declarations to 4,096 in aggregate across workbook-selected worksheet,
chart-sheet, and dialog-sheet parts. The nested-sequence catalogs count every
direct child that the reader would materialize, while the typed catalogs count
the reader's matching local name. The raw Custom View boundary similarly counts
every direct child it would inspect, so an alternate-namespace entry cannot
bypass the safety limit.

Controlled 10,000-entry fixtures establish the allocation path: 160,591 bytes
of workbook XML for 10,001 book views took 0.678 seconds and 40,376 KiB;
10,000 custom workbook views in a 34,797-byte package took 2.369 seconds and
87,112 KiB; the function-group, smart-tag, and web-publish-object fixtures took
0.978, 1.446, and 1.498 seconds; and a 7,170-byte package with 10,000 custom
sheet views took 0.867 seconds and 67,232 KiB. With the normal gate enabled,
all six fixtures rejected before downstream work in 0.013–0.022 seconds with
stable catalog-specific input errors. Exact 4,096-entry fixtures remained
accepted for every catalog; the 4,097th declaration was rejected.

Focused archive-safety and version coverage passed **67 tests in 2.02 seconds**;
the completed source tree passed **769 tests in 130.23 seconds**, plus Ruff,
bytecode compilation, and `git diff --check`. The final wheel and source
distribution passed `twine check`; a clean Python 3.12 environment installed
the wheel, confirmed `FormulaFence 0.117.0` and `openpyxl.DEFUSEDXML`, profiled
an ordinary workbook, and returned the normal input-error exit status (2) for
independently generated 4,097-entry book-view, custom-workbook-view,
function-group, smart-tag-type, web-publish-object, and custom-sheet-view
fixtures.

## Reader relationship-backed catalog bounds — 2026-07-26

Version 0.116.0 closes the next workbook-level repetition path. An
`externalReference` and a `pivotCache` each select a workbook relationship, but
the relationship inventory alone cannot prevent a compact workbook XML part
from declaring the same relationship thousands of times. `openpyxl` first
materializes both direct catalogs; FormulaFence then inventories external-link
declarations and rereads each declared pivot-cache definition while collecting
refresh controls.

The semantic-reader preflight now rejects the 4,097th direct external-reference
or pivot-cache declaration before any raw OOXML scanner or complete reader
starts. The limits preserve every distinct relationship allowed by FormulaFence's
existing 4,096-record relationship catalog, while making duplicate declarations
finite. As with sheets and defined names, direct alternate-namespace catalog
entries count because the reader's nested-sequence parser materializes them.

Before the gate, a 10,000-declaration repeated-pivot-cache fixture had a
491,026-byte `xl/workbook.xml` and a 10,130-byte package, yet FormulaFence
spent 7.350 seconds and retained 10,000 cache controls. A corresponding
10,000-declaration repeated-external-reference fixture was a 10,452-byte
package and took 1.277 seconds despite pointing at only three external-link
parts. The new gate rejected those inputs in 0.018 and 0.016 seconds,
respectively, with stable catalog-specific errors before downstream reader work.
Focused archive-safety and version coverage passed **49 tests in 1.95 seconds**;
the completed source tree passed **751 tests in 131.91 seconds**, plus Ruff,
bytecode compilation, and `git diff --check` before final package verification.
Both final distribution artifacts passed `twine check`; a fresh Python 3.12
environment installed the wheel, returned `FormulaFence 0.116.0`, confirmed
`openpyxl.DEFUSEDXML`, profiled an ordinary workbook, and returned the normal
input-error exit status (2) for both 4,097-declaration catalog fixtures.

## Reader defined-name catalog bound — 2026-07-26

Version 0.115.0 closes a workbook-level allocation path that remained after
the sheet catalog limit. The SpreadsheetML schema permits an unbounded sequence
of `<definedName>` children, and Microsoft's interoperability notes say Excel
permits up to 2,147,483,647 occurrences of that element. FormulaFence has to
make a narrower CI-worker choice because `openpyxl` reads the complete workbook
part into a `WorkbookPackage` before FormulaFence can decide whether any name is
relevant to a policy or formula-control ledger.

The semantic-reader preflight now streams direct workbook defined-name
declarations and rejects the 100,001st before any raw OOXML scanner or complete
reader starts. Its catalog checks deliberately use the same local-name behavior
as the reader: an alternate-namespace `definedName` still materializes and is
therefore counted. The existing 512-sheet counter receives the same correction,
so an alternate-namespace direct sheet entry cannot bypass it.

A generated fixture with 100,000 direct names had 7,300,628 bytes of
`xl/workbook.xml` but only a 274,754-byte package. It remained accepted; a
FormulaFence snapshot took 22.616 seconds and peaked at 192,288 KiB in the
isolated Python process. Adding one declaration made the package fail in 0.312
seconds with the stable defined-name preflight error, before any downstream
reader ran. Focused archive-safety and version coverage passed **43 tests in
2.08 seconds**; the completed source tree passed **745 tests in 116.02
seconds**, plus Ruff, bytecode compilation, and `git diff --check`. The schema
and Excel compatibility detail are documented in Microsoft's
[definedNames interoperability notes](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oi29500/1aea93a7-b208-44b2-a2e5-83dc2b635b41).
Both final distribution artifacts passed `twine check`; a fresh Python 3.12
environment installed the wheel, returned `FormulaFence 0.115.0`, confirmed
`openpyxl.DEFUSEDXML`, profiled an ordinary workbook, and returned the normal
input-error exit status (2) for the 100,001-name fixture.

## Reader bootstrap catalog bounds — 2026-07-26

Version 0.114.0 closes a separate allocation path before worksheet data is
considered. A SpreadsheetML workbook catalogs `<sheet>` declarations that each
point to a part through a relationship ID, as shown in Microsoft's
[SpreadsheetML package structure](https://learn.microsoft.com/en-us/office/open-xml/spreadsheet/structure-of-a-spreadsheetml-document).
`openpyxl` materializes the package manifest's `Default` and `Override`
sequences before it finds the workbook part, so a ZIP-member bound alone does
not bound the reader's bootstrap object catalogs.

The semantic-reader preflight now streams these catalogs without retaining XML
trees and rejects more than 4,096 manifest declarations, 4,096 workbook
relationships, or 512 workbook sheet declarations. The sheet count is by
declaration rather than unique target, so a compact package cannot repeat one
safe relationship hundreds of times and make the reader or FormulaFence's raw
sheet boundaries revisit it as a large workbook.

Focused fixtures prove every new bound fails before the first downstream
scanner. A direct boundary reproduction appended 512 copies of one valid
`<sheet>` declaration to a four-sheet workbook: stock `openpyxl` accepted the
result as 516 sheets, while FormulaFence returned the stable semantic-reader
preflight error before loading the workbook. The completed source tree passed
**741 tests in 121.54 seconds**, plus Ruff, bytecode compilation, and `git diff
--check`. Both final distribution artifacts passed `twine check`; a fresh Python
3.12 environment installed the wheel, returned `FormulaFence 0.114.0`, confirmed
`openpyxl.DEFUSEDXML`, profiled a normal workbook, and returned the normal
input-error exit status (2) for the repeated-sheet boundary fixture.

## Reader XML cardinality and scalar limits — 2026-07-26

Version 0.113.0 extends the 0.112.0 semantic-reader preflight around the
specific XML surfaces that ordinary workbook loading materializes: the package
manifest, workbook metadata and relationship catalog, canonical styles,
shared-string table, and workbook-selected sheets. This matters because
`openpyxl`'s [shared-string reader](https://openpyxl.readthedocs.io/en/3.1/_modules/openpyxl/reader/strings.html)
appends every `<si>` item to a Python list even when few worksheet cells refer
to it; a worksheet-cell count alone cannot bound that allocation.

The preflight streams those reader-visible parts under 4,000,000 elements per
part and 256 nesting levels, without retaining their trees. It rejects more
than 500,000 shared-string entries or populated worksheet cells, more than
65,490 `cellXfs` entries, text values above 32,767 characters, and stored
formula/defined-name text above 8,192 characters. The text, formula, and style
ceilings align with [Excel's published specifications and limits](https://support.microsoft.com/en-US/Excel/excel-specifications-and-limits).
For shared strings it follows the first matching manifest Override—the same
selection path `openpyxl` uses—then a sole workbook relationship or canonical
fallback. Malformed unrelated extension XML retains its existing coverage-
warning route rather than becoming a broad input rejection.

Focused fixtures prove every new bound fails before a downstream scanner,
including relationship- and manifest-selected noncanonical shared-string
parts. A direct boundary probe built a 27,245-byte source workbook containing
500,001 shared-string items and confirmed the stable preflight rejection before
the first raw workbook scanner could run. The completed source tree passed
**737 tests in 115.64 seconds**, plus Ruff, bytecode compilation, and `git diff
--check`. Both final distribution artifacts passed `twine check`; a fresh
Python 3.12 environment installed the wheel, returned `FormulaFence 0.113.0`,
confirmed `openpyxl.DEFUSEDXML`, profiled a normal workbook, and returned the
normal input-error exit status (2) for the actual 500,001-item boundary fixture.

## Semantic-reader resource boundary and defused XML — 2026-07-26

The archive-header inventory from 0.111.0 bounds ZIP expansion, but FormulaFence
uses a complete non-streaming `openpyxl` workbook model to compare semantic cell
and control state. Version 0.112.0 adds a second, fail-closed reader preflight
between that structural ZIP check and downstream raw OOXML scanners or
`openpyxl` calls. It caps each XML/relationship part at 64 MiB and aggregate XML material at
256 MiB, follows bounded workbook sheet relationships, and streams those
selected worksheet parts without retaining cells or values to reject more than
500,000 populated SpreadsheetML cell records. A nonstandard
relationship-selected worksheet target is counted too; unrelated malformed
extension parts retain their existing explicit coverage-warning path rather
than becoming a broad false-positive input rejection.

`defusedxml>=0.7,<1` is now a declared runtime dependency. FormulaFence uses it
for raw OOXML parsing, and a clean supported installation confirms that it also
enables `openpyxl.DEFUSEDXML`. Focused fixtures prove normal and ZIP64 workbooks
still load; per-part, aggregate, and populated-cell bounds fail before any
downstream scanner; an entity-bearing selected worksheet fails at the reader
preflight; and `vbaProject.bin` is SHA-256 hashed through a stream rather than
one whole-payload `ZipFile.read`. Malformed cell/package metadata that causes
`openpyxl` to raise `TypeError` or `IndexError` now emits FormulaFence's normal
unreadable-workbook input error (exit 2), rather than a traceback.

The completed source tree passed **727 tests in 126.37 seconds**, plus Ruff,
bytecode compilation, and `git diff --check`. Both distribution artifacts passed
`twine check`; their final SHA-256 values are published with the GitHub Release
assets rather than copied into this self-included validation note. A fresh
Python 3.12 environment installed the final wheel and declared dependencies,
returned `FormulaFence 0.112.0`, profiled a normal formula workbook, confirmed
the defused parser path, and rejected the entity-bearing workbook before a
reader ran.

## Bounded OOXML archive preflight — 2026-07-26

Workbook files are untrusted CI inputs before FormulaFence can inspect their
formula semantics. The reader now performs a bounded ZIP inventory before it
opens any OOXML part or calls `openpyxl`: a one-GiB source cap, 32-MiB central
directory cap, 4,096-entry cap, canonical single-disk member inventory, 512-MiB
per-member expansion cap, 768-MiB aggregate expansion cap, and a 1,000:1
compression-ratio cap. It accepts only stored or deflated members and verifies
their local headers without extracting data. Duplicate or case-colliding paths,
unsafe paths, encrypted or special-file members, malformed ZIP64 metadata, and
overlapping payloads fail before downstream scanners run.

Focused regression fixtures prove that ordinary and valid ZIP64 workbooks still
load, while source-size and central-directory gates run before a general ZIP
reader; oversized members and aggregate expansion, compression bombs,
traversal and case-colliding paths, duplicate members, encrypted records, ZIP
Unicode-path aliases, malformed ZIP64 metadata, local-header disagreement,
symbolic links, and the CLI error path all fail closed. The completed source
tree passed **719 tests in 105.67 seconds**. This follows Python's warning about
ZIP decompression resource pitfalls and Microsoft's archive-validation guidance
on entry counts and expanded-size limits: [Python `zipfile` documentation](https://docs.python.org/3/library/zipfile.html#decompression-pitfalls)
and [Microsoft archive best practices](https://learn.microsoft.com/en-us/dotnet/standard/io/zip-tar-best-practices).

## Self-contained HTML review artifacts — 2026-07-26

FormulaFence 0.110.0 adds `--format html` for `diff`, `check`, and `portfolio`.
It is a deliberately local review artifact rather than a hosted viewer: the
page contains inline CSS, a fixed text/severity filtering script, and escaped
expandable evidence. It has no remote assets or report-initiated network
requests. This keeps an uploaded CI artifact useful to reviewers without
requiring a service that receives workbook data.

The renderer starts from the same deterministic report payload as Markdown and
applies every existing output-only sharing boundary before inserting anything
into the page. A hostile workbook-derived string is escaped in both the summary
and expanded JSON evidence, so it cannot become a report tag or script. The
native environment-information fixture also caught an incidental static
stylesheet word matching a controlled marker; that token was removed so exact
redaction scans remain meaningful instead of merely hiding workbook evidence.

The completed source tree passed **703 tests in 104.00 seconds**, including
direct escaped-evidence tests, every sharing-redaction CLI path, HTML `check`
policy evidence, HTML portfolios, and the composite Action contract. Ruff,
`git diff --check`, Python bytecode compilation, and Action shell syntax checks
were clean. A release-candidate wheel and source distribution passed `twine
check`. The wheel was installed with declared dependencies into a fresh
environment, reported `FormulaFence 0.110.0`, retained controlled markers in a
default HTML diff, removed them from redacted `check` and portfolio HTML, and
preserved `FF072` plus `FFP072` with a policy exit code of `1`. Final artifact
digests are published with the GitHub release.

## Shared native environment-information report redaction — 2026-07-26

FormulaFence 0.109.0 adds the separate, opt-in
`--redact-formula-environment-information` rendering boundary for generic
reports. It is deliberately separate from the count-only `FF072` ledger and
from external-workbook-link, formula-action, Python-in-Excel, Office
custom-function, unqualified-runtime-function, worksheet-code-resource
registration, formula-defined XLM registration, evaluation, action, GET.CELL,
and environment-information redaction modes. Default local-review output
remains unchanged. When enabled, JSON, Markdown, and SARIF hide direct stored
native `CELL`, `INFO`, `SHEET`, and `SHEETS` material and exact changed static
input evidence the private dependency analysis recorded as reaching an
inventoried call. A formula-defined-name body can pass a private information
code or reference through a dotted workbook-defined wrapper to a native call
deeper in the chain, so FormulaFence privately compares the resolved definition
chain and conservatively hides changed defined-name before/after evidence when
that signature changes.

Microsoft's [CELL function
documentation](https://support.microsoft.com/en-us/office/cell-function-51bd39a5-f338-4dbe-a33f-955d67c2b2cf)
documents file, location, formatting, and content information, while its
[INFO function documentation](https://support.microsoft.com/en-au/office/info-function-725f259a-0e4b-49b3-8b52-58815c69acae)
lists operating-environment data. The [SHEET function
documentation](https://support.microsoft.com/en-us/excel/functions/sheet-function)
and [SHEETS function
documentation](https://support.microsoft.com/en-us/excel/functions/sheets-function)
cover tab-number/count behavior. Those documented stored calls explain why
generic report evidence can disclose private material without proving that a
call will calculate or what information Excel would return. FormulaFence does
not calculate a formula or information call, determine an information type,
resolve a dynamic reference, infer a selected cell, simulate workbook/client/
workspace state, or reconstruct a runtime value. Its existing raw-tab-catalog
comparison semantics for stored `SHEET` and omitted-reference `SHEETS` calls
remain unchanged.

Focused fixtures cover a direct stored native call plus a changed static input,
an input whose native-call consumer falls beyond the report's bounded impact
sample, and a dotted named wrapper whose private reference can only be
associated with a native call through the private fixed-point definition
analysis. The redacted JSON, Markdown, SARIF, policy, portfolio, and
composite-Action contracts retain `FF072` / `FFP072` and their exit behavior
while omitting controlled information-code, reference, input, and nested-name
markers. The final source tree passed **701 tests in 101.90 seconds**, plus a
clean Ruff check, `git diff --check`, and shell syntax check for the composite
Action. A wheel built from the release candidate passed `twine check`, was
installed with its declared dependencies into a fresh environment, reported
`FormulaFence 0.109.0`, retained controlled environment/input markers in
default JSON, removed them from redacted JSON/Markdown/SARIF/policy/portfolio
output, and still returned `1` with `FF072` and `FFP072` for the redacted policy
check. Final artifact digests are published with the GitHub release.

## Shared formula-defined XLM environment-information report redaction — 2026-07-26

FormulaFence 0.108.0 adds the separate, opt-in
`--redact-formula-defined-xlm-environment-information-calls` rendering boundary
for generic reports. It is deliberately separate from the count-only `FF071`
ledger and from the external-workbook-link, formula-action, Python-in-Excel,
Office custom-function, unqualified-runtime-function, worksheet-code-resource
registration, formula-defined XLM registration, formula-defined XLM evaluation,
formula-defined XLM action, and formula-defined XLM GET.CELL redaction modes.
Default local-review output remains unchanged. When enabled, JSON, Markdown,
and SARIF hide direct stored selected environment-information material, changed
invoking-formula evidence, and exact changed static input evidence the private
dependency analysis recorded as reaching an inventoried call. A
formula-defined-name body can pass a private information code or reference
through a dotted workbook-defined wrapper to GET.WORKBOOK, GET.WORKSPACE, or
GET.DOCUMENT deeper in the chain, so FormulaFence privately compares the
resolved definition chain and conservatively hides changed defined-name
before/after evidence when that signature changes.

Microsoft's [Excel C API
reference](https://learn.microsoft.com/en-us/office/client-developer/excel/programming-with-the-c-api-in-excel)
identifies workspace information functions such as GET.CELL and GET.WORKBOOK;
Microsoft's [xlfFree example](https://learn.microsoft.com/en-us/office/client-developer/excel/xlfree)
demonstrates GET.WORKSPACE returning platform information, and its [expression
evaluation reference](https://learn.microsoft.com/en-us/office/client-developer/excel/excel-worksheet-and-expression-evaluation)
identifies GET.DOCUMENT as an XLM information function. Those documented stored
arguments explain why generic report evidence can disclose private material
without proving that a call will calculate or what information Excel would
return. FormulaFence does not calculate a formula or information call,
determine an information type, resolve a dynamic reference, simulate
workbook/workspace/document state, or reconstruct a runtime value.

Focused fixtures cover a direct stored selected call plus a changed static
input, an input whose environment-information consumer falls beyond the
report's bounded impact sample, and a dotted named wrapper whose private
reference can only be associated with a selected call through the private
fixed-point definition analysis. The redacted JSON, Markdown, SARIF, policy,
portfolio, and composite-Action contracts retain `FF071` / `FFP071` and their
exit behavior while omitting controlled information-code, reference, input,
and nested-name markers. The final source tree passed **695 tests in 100.98
seconds** (deterministic command-limited chunks), plus a clean Ruff check,
`git diff --check`, and shell syntax check for the composite Action. The exact
release wheel was installed in a fresh environment; its CLI reported
`FormulaFence 0.108.0`, retained controlled environment/input markers in
default JSON, removed them from redacted JSON/Markdown/SARIF/policy/portfolio
output, and still returned `1` with `FF071` and `FFP071` for the redacted policy
check. The wheel `formulafence-0.108.0-py3-none-any.whl` passed `twine check`
with SHA-256
`62c2163c2f3f2eecc1978daf05790573173fe88621fe6610b25bc3cc16034524`.
The source distribution also passed `twine check`; its digest is intentionally
omitted because this validation note is included in that archive.

## Shared formula-defined XLM GET.CELL report redaction — 2026-07-26

FormulaFence 0.107.0 adds the separate, opt-in
`--redact-formula-defined-xlm-get-cell-calls` rendering boundary for generic
reports. It is deliberately separate from the count-only `FF070` ledger and
from the external-workbook-link, formula-action, Python-in-Excel, Office
custom-function, unqualified-runtime-function, worksheet-code-resource
registration, formula-defined XLM registration, formula-defined XLM
evaluation, and formula-defined XLM action redaction modes. Default
local-review output remains unchanged. When enabled, JSON, Markdown, and SARIF
hide direct stored GET.CELL material, changed invoking-formula evidence, and
exact changed static input evidence the private dependency analysis recorded as
reaching an inventoried call. A formula-defined-name body can pass a private
information code or reference through a dotted workbook-defined wrapper to
GET.CELL deeper in the chain, so FormulaFence privately compares the resolved
definition chain and conservatively hides changed defined-name before/after
evidence when that signature changes.

Microsoft's [Excel C API
reference](https://learn.microsoft.com/en-us/office/client-developer/excel/programming-with-the-c-api-in-excel)
identifies GET.CELL as `xlfGetCell`. That documented stored material explains
why generic report evidence can disclose private call arguments without proving
that a call will calculate or what information Excel would return. FormulaFence
does not calculate a formula or information call, determine an information
type, resolve a dynamic reference, simulate display/formatting or other Excel
state, or reconstruct a runtime value.

Focused fixtures cover a direct stored GET.CELL call plus a changed static
input, an input whose GET.CELL consumer falls beyond the report's bounded
impact sample, and a dotted named wrapper whose private reference can only be
associated with GET.CELL through the private fixed-point definition analysis.
The redacted JSON, Markdown, SARIF, policy, portfolio, and composite-Action
contracts retain `FF070` / `FFP070` and their exit behavior while omitting
controlled information-code, reference, input, and nested-name markers. The
final source tree passed **689 tests in 100.15 seconds** (deterministic
command-limited chunks), plus a clean Ruff check, `git diff --check`, and shell
syntax check for the composite Action. The exact release wheel was installed in
a fresh environment; its CLI reported `FormulaFence 0.107.0`, retained
controlled GET.CELL/input markers in default JSON, removed them from redacted
JSON/Markdown/SARIF/policy/portfolio output, and still returned `1` with
`FF070` and `FFP070` for the redacted policy check. The wheel
`formulafence-0.107.0-py3-none-any.whl` passed `twine check` with SHA-256
`14c23132397029643ba78f305200cfe03702d8d5911ef6909edbe51cdcc77c74`.
The source distribution also passed `twine check`; its digest is intentionally
omitted because this validation note is included in that archive.

## Shared formula-defined XLM action report redaction — 2026-07-26

FormulaFence 0.106.0 adds the separate, opt-in
`--redact-formula-defined-xlm-actions` rendering boundary for generic reports.
It is deliberately separate from the count-only `FF073` ledger and from the
external-workbook-link, formula-action, Python-in-Excel, Office custom-function,
unqualified-runtime-function, worksheet-code-resource registration,
formula-defined XLM registration, and formula-defined XLM evaluation redaction
modes. Default local-review output remains unchanged. When enabled, JSON,
Markdown, and SARIF hide direct stored selected-action material, changed
invoking-formula evidence, and exact changed static input evidence the private
dependency analysis recorded as reaching an inventoried action. A
formula-defined-name body can pass a private target or handler through a dotted
workbook-defined wrapper to a selected XLM action deeper in the chain, so
FormulaFence privately compares the resolved definition chain and
conservatively hides changed defined-name before/after evidence when that
signature changes.

Microsoft's [Excel C API
reference](https://learn.microsoft.com/en-us/office/client-developer/excel/programming-with-the-c-api-in-excel)
describes XLM command-equivalent functions and event traps, while its
[DLL-access guidance](https://learn.microsoft.com/en-us/office/client-developer/excel/how-to-access-dlls-in-excel)
documents `CALL` as an XLM route to a DLL function or command. Those documented
stored arguments explain why generic report evidence can disclose sensitive
action material without proving that an action is available, trusted, or will
succeed. FormulaFence does not calculate a formula, resolve a target or
handler, load a DLL, send DDE, execute a macro or program, or reconstruct a
dynamically assembled action.

Focused fixtures cover a direct stored selected action plus a changed static
input, an input whose action consumer falls beyond the report's bounded impact
sample, and a dotted named wrapper whose private target can only be associated
with the action through the private fixed-point definition analysis. The
redacted JSON, Markdown, SARIF, policy, portfolio, and composite-Action
contracts retain `FF073` / `FFP073` and their exit behavior while omitting
controlled target, handler, input, and nested-name markers. The final source
tree passed **683 tests in 98.49 seconds** (deterministic command-limited
chunks), plus a clean Ruff check, `git diff --check`, and shell syntax check for
the composite Action. The exact release wheel was installed in a fresh
environment; its CLI reported `FormulaFence 0.106.0`, retained controlled
action/input markers in default JSON, removed them from redacted
JSON/Markdown/SARIF/policy/portfolio output, and still returned `1` with
`FF073` and `FFP073` for the redacted policy check. The wheel
`formulafence-0.106.0-py3-none-any.whl` passed `twine check` with SHA-256
`d47f1f28d1baa2e96286ce570408e0bb207aa9e99b84261d2928a8c6dc727bb3`.
The source distribution also passed `twine check`; its digest is intentionally
omitted because this validation note is included in that archive.

## Shared formula-defined XLM evaluation report redaction — 2026-07-26

FormulaFence 0.105.0 adds the separate, opt-in
`--redact-formula-defined-xlm-evaluations` rendering boundary for generic
reports. It is deliberately separate from the count-only `FF069` ledger and
from the external-workbook-link, formula-action, Python-in-Excel, Office
custom-function, unqualified-runtime-function, worksheet-code-resource
registration, and formula-defined XLM registration redaction modes. Default
local-review output remains unchanged. When enabled, JSON, Markdown, and SARIF
hide direct stored `EVALUATE` material, changed invoking-formula evidence, and
exact changed static input evidence the private dependency analysis recorded as
reaching an inventoried evaluation. A formula-defined-name body can pass a
private expression through a dotted workbook-defined wrapper to `EVALUATE`
deeper in the chain, so FormulaFence privately compares the resolved definition
chain and conservatively hides changed defined-name before/after evidence when
that signature changes.

Microsoft's [Excel expression-evaluation
reference](https://learn.microsoft.com/en-us/office/client-developer/excel/excel-worksheet-and-expression-evaluation)
identifies `EVALUATE` as an XLM function that reduces a valid character string
to a worksheet value. That documented behavior explains why stored expression
text can be sensitive without proving that the text is valid or will be
evaluated. FormulaFence does not calculate a formula or text argument, parse a
runtime-generated expression, execute a macro, or reconstruct dynamically
assembled text.

Focused fixtures cover a direct stored `EVALUATE` expression plus a changed
static input, an input whose evaluation consumer falls beyond the report's
bounded impact sample, and a dotted named wrapper whose private expression can
only be associated with the evaluation through the private fixed-point
definition analysis. A separate fixture proves that FormulaFence neither
re-tokenizes nor redacts an ordinary change referenced only inside runtime
expression text. The redacted JSON, Markdown, SARIF, policy, portfolio, and
composite-Action contracts retain `FF069` / `FFP069` and their exit behavior
while omitting controlled expression, literal, input, and nested-name markers.
The final source tree passed **677 tests in 96.49 seconds** (deterministic
command-limited chunks), plus a clean Ruff check, `git diff --check`, and shell
syntax check for the composite Action. The exact release wheel was installed in
a fresh environment; its CLI reported `FormulaFence 0.105.0`, retained
controlled evaluation/input markers in default JSON, removed them from redacted
JSON/Markdown/SARIF/policy/portfolio output, and still returned `1` with
`FF069` and `FFP069` for the redacted policy check. The wheel
`formulafence-0.105.0-py3-none-any.whl` passed `twine check` with SHA-256
`65f7adc49f9165739c4763d313b5d874c72b3120f724eae9b2ffebbb8461d059`.
The source distribution also passed `twine check`; its digest is intentionally
omitted because this validation note is included in that archive.

## Shared formula-defined XLM registration report redaction — 2026-07-26

FormulaFence 0.104.0 adds the separate, opt-in
`--redact-formula-defined-xlm-registrations` rendering boundary for generic
reports. It is deliberately separate from the count-only `FF068` ledger and
from the external-workbook-link, formula-action, Python-in-Excel, Office
custom-function, unqualified-runtime-function, and worksheet-code-resource
registration redaction modes. Default local-review output remains unchanged.
When enabled, JSON, Markdown, and SARIF hide direct stored `REGISTER` material,
changed invoking-formula evidence, and exact changed static input evidence the
private dependency analysis recorded as reaching an inventoried registration. A
formula-defined-name body can send a private module, procedure, type string, or
other argument through a dotted workbook-defined wrapper to `REGISTER` deeper
in the chain, so FormulaFence privately compares the resolved definition chain
and conservatively hides changed defined-name before/after evidence when that
signature changes.

Microsoft's [`xlfRegister` Form 1 reference](https://learn.microsoft.com/en-au/office/client-developer/excel/xlfregister-form-1)
documents the XLM `REGISTER` equivalent for DLL-function or command
registration and identifies macro types callable from a defined-name
definition. Its [`Form 2` reference](https://learn.microsoft.com/en-au/office/client-developer/excel/xlfregister-form-2)
documents XLL loading and activation. Those documented stored arguments explain
why a generic report can expose sensitive implementation material without
proving that a DLL/XLL is available, trusted, loaded, or successfully
registered. FormulaFence does not evaluate a formula, execute a macro, resolve
a module path, load a DLL/XLL, inspect host trust settings, contact a provider,
or reconstruct a dynamic argument.

Focused fixtures cover a direct stored `REGISTER` definition plus a changed
static module input, an input whose invoking registration falls beyond the
report's bounded impact sample, and a dotted named wrapper whose private module
literal can only be associated with the registration through the private
fixed-point definition analysis. The redacted JSON, Markdown, SARIF, policy,
portfolio, and composite-Action contracts retain `FF068` / `FFP068` and their
exit behavior while omitting controlled module, procedure, type, input, and
nested-name markers. The final source tree passed **671 tests in 95.24 seconds**
(deterministic command-limited chunks), plus a clean Ruff check,
`git diff --check`, and shell syntax check for the composite Action. The exact
release wheel was installed in a fresh environment; its CLI retained controlled
registration/input markers in default JSON, removed them from redacted
JSON/Markdown/SARIF/policy/portfolio output, and still returned `1` with
`FF068` and `FFP068` for the redacted policy check. The wheel
`formulafence-0.104.0-py3-none-any.whl` passed `twine check` with SHA-256
`31b41bacfe2a5299b4c07926cd815b32c2a2990c9616747f2696aaf40f3a505f`.
The source distribution also passed `twine check`; its digest is intentionally
omitted because this validation note is included in that archive.

## Shared worksheet code-resource registration report redaction — 2026-07-26

FormulaFence 0.103.0 adds the separate, opt-in
`--redact-worksheet-code-resource-registrations` rendering boundary for generic
reports. It is deliberately separate from the count-only `FF067` ledger and
from the external-workbook-link, formula-action, Python-in-Excel, Office
custom-function, and unqualified-runtime-function redaction modes. Default
local-review output remains unchanged. When enabled, JSON, Markdown, and SARIF
hide direct stored `REGISTER.ID` material, changed registration-cell evidence,
and exact changed static input evidence the private dependency analysis recorded
as reaching an inventoried registration. A formula-defined-name body can send a
private module, procedure, or other argument through a dotted workbook-defined
wrapper to `REGISTER.ID` deeper in the chain, so FormulaFence privately compares
the resolved definition chain and conservatively hides changed defined-name
before/after evidence when that signature changes.

Microsoft's [`REGISTER.ID` reference](https://support.microsoft.com/en-us/office/register-id-function-f8f0af0f-fd66-4704-a0f2-87b27b175b50)
documents `REGISTER.ID(module_text, procedure, [type_text])`, says it returns a
DLL/code-resource registration ID, and says it registers an unregistered DLL or
code resource. It also distinguishes the worksheet-capable function from
`REGISTER`. Those documented stored arguments establish why a generic report
can expose sensitive implementation material without proving that a module is
available, trusted, or successfully registered. FormulaFence does not evaluate
a formula, resolve a module path, load a DLL/XLL, inspect host trust settings,
execute code, contact a provider, or reconstruct a dynamic argument.

Focused fixtures cover a direct `REGISTER.ID` expression plus a changed static
module input, an input whose registration consumer falls beyond the report's
bounded impact sample, and a dotted named wrapper whose private module literal
can only be associated with the registration through the private fixed-point
definition analysis. The redacted JSON, Markdown, SARIF, policy, portfolio, and
composite-Action contracts retain `FF067` / `FFP067` and their exit behavior
while omitting controlled module, procedure, type, input, and nested-name
markers. The final source tree passed **665 tests in 95.22 seconds**
(deterministic command-limited chunks), plus a clean Ruff check,
`git diff --check`, and shell syntax check for the composite Action. The exact
release wheel was installed in a fresh environment; its CLI reported
`FormulaFence 0.103.0`, retained controlled registration/input markers in
default JSON, removed them from redacted JSON/Markdown/SARIF/policy/portfolio
output, and still returned `1` with `FF067` and `FFP067` for the redacted policy
check. The wheel `formulafence-0.103.0-py3-none-any.whl` passed `twine check`
with SHA-256 `8b45bf02dd631ca142a849d95ec624311b886af02d8848b53ae75a93df37c5f5`.
The source distribution also passed `twine check`; its digest is intentionally
omitted because this validation note is included in that archive.

## Shared unqualified runtime-function report redaction — 2026-07-26

FormulaFence 0.102.0 adds the separate, opt-in
`--redact-unqualified-runtime-functions` rendering boundary for generic
reports. It is deliberately separate from the count-only `FF075` ledger and
from the external-workbook-link, formula-action, Python-in-Excel, and Office
custom-function redaction modes. Default local-review output remains unchanged.
When enabled, JSON, Markdown, and SARIF hide direct stored bare runtime-call
material, changed runtime-call-cell evidence, and exact changed static input
evidence the private dependency analysis recorded as reaching an inventoried
call. A formula-defined-name body can send a private argument through a dotted
workbook-defined wrapper to a bare UDF deeper in the chain, so FormulaFence
privately compares the resolved definition chain and conservatively hides
changed defined-name before/after evidence when that signature changes.

Microsoft's [installed UDF guidance](https://support.microsoft.com/en-us/excel/user-defined-functions-that-are-installed-with-add-ins-reference)
documents add-in and Automation functions, its [VBA custom-function guidance](https://support.microsoft.com/en-us/excel/create-custom-functions-in-excel)
shows a bare worksheet call such as `=DISCOUNT(D7,E7)`, and the
[XLL registration/call guidance](https://learn.microsoft.com/en-us/office/client-developer/excel/accessing-xll-code-in-excel)
documents XLL code exposed to Excel. Those documents establish why a stored
bare call and its arguments can be sensitive without proving a provider is
installed or runnable. FormulaFence does not evaluate a formula, resolve/load
VBA, COM/Automation, XLL, or another provider, execute code, contact a runtime,
or reconstruct a dynamic argument.

Focused fixtures cover a direct `PRIVATEUDF` formula plus a changed static
input, an input whose bare-runtime-function consumer falls beyond the report's
bounded impact sample, and a dotted named wrapper whose private argument can
only be associated with the bare call through the private fixed-point definition
analysis. The redacted JSON, Markdown, SARIF, policy, portfolio, and
composite-Action contracts retain `FF075` / `FFP075` and their exit behavior
while omitting the controlled UDF, argument, input, and nested-name markers.
The final source tree passed **659 tests in 92.70 seconds** (deterministic
command-limited chunks), plus a clean Ruff check, `git diff --check`, and shell
syntax check for the composite Action. The exact release wheel was installed in
a fresh environment; its CLI reported `FormulaFence 0.102.0`, retained the
controlled UDF/input markers in default JSON, removed them from redacted
JSON/Markdown/SARIF/policy/portfolio output, and still returned `1` with
`FF075` and `FFP075` for the redacted policy check. The wheel
`formulafence-0.102.0-py3-none-any.whl` passed `twine check` with SHA-256
`2e73319cc362eeb2900c66b466458b6028b1601c8db719d3dfbd39c6ae5088bf`.
The source distribution also passed `twine check`; its digest is intentionally
omitted because this validation note is included in that archive.

## Shared Office custom-function report redaction — 2026-07-26

FormulaFence 0.101.0 adds the separate, opt-in
`--redact-office-custom-functions` rendering boundary for generic reports. It
is deliberately separate from the count-only `FF066` ledger and from the
external-workbook-link, formula-action, and Python-in-Excel redaction modes.
Default local-review output remains unchanged. When enabled, JSON, Markdown,
and SARIF hide direct stored namespaced custom-function material, changed
custom-function-cell evidence, and exact changed static input evidence the
private dependency analysis recorded as reaching an inventoried call. A changed
formula-defined-name body can pass a private argument to a namespaced call
through an ordinary-looking named `LAMBDA`, so FormulaFence privately compares
the relevant resolved definition chain and conservatively hides changed
defined-name before/after evidence when that signature changes.

Microsoft's [custom-functions overview](https://learn.microsoft.com/en-us/office/dev/add-ins/excel/custom-functions-overview)
documents JavaScript/TypeScript functions exposed through a manifest namespace,
and its [web-data guidance](https://learn.microsoft.com/en-us/office/dev/add-ins/excel/custom-functions-web-reqs)
documents requests to APIs, WebSockets, and streaming calls. Those documents
establish why a stored namespaced formula and its static arguments can be
sensitive even though the normal workbook does not carry the manifest or
runtime code. FormulaFence does not evaluate a formula, load an add-in,
execute JavaScript, contact a runtime, or reconstruct a dynamic argument.

Focused fixtures cover a direct `CONTOSO` formula plus a changed static input,
an input whose custom-function consumer falls beyond the report's bounded
impact sample, and an unqualified nested named-LAMBDA wrapper whose private
argument can only be associated with a custom function through the private
fixed-point definition analysis. The redacted JSON, Markdown, SARIF, policy,
portfolio, and composite-Action contracts retain `FF066` / `FFP066` and their
exit behavior while omitting the controlled call, query, input, and nested-name
markers. The final source tree passed **653 tests in 90.63 seconds**, plus a
clean Ruff check, `git diff --check`, and shell syntax check for the composite
Action. The exact release wheel was installed in a fresh environment; its CLI
reported `FormulaFence 0.101.0`, retained every controlled marker in default
JSON, removed them from redacted JSON/Markdown/SARIF/policy/portfolio output,
and still returned `1` with `FF066` and `FFP066` for the redacted policy check.
The wheel `formulafence-0.101.0-py3-none-any.whl` passed `twine check` with
SHA-256 `4916176df4ee0edab14a821d895427913282a65a92266c6475fa956f2a4b2d65`.
The source distribution also passed `twine check`; its digest is intentionally
omitted because this validation note is included in that archive.

## Shared Python-in-Excel report redaction — 2026-07-26

FormulaFence 0.100.0 adds the separate, opt-in
`--redact-python-in-excel` rendering boundary for generic reports. It is
deliberately separate from the count-only `FF065` ledger and from the
external-workbook-link / formula-action redaction modes. Default local-review
output remains unchanged. When enabled, JSON, Markdown, and SARIF hide direct
inventoried `PY` formula material, changed PY-cell evidence, and exact changed
static input evidence that the private dependency analysis recorded as reaching
an inventoried PY cell. It does not parse Python, evaluate a formula, contact
Microsoft Cloud, change comparison/policy results, or reconstruct a runtime
value.

Microsoft's [PY function reference](https://support.microsoft.com/en-us/excel/functions/py-function)
documents static Python source in `PY(python_code, return_type)`; the OOXML
[Python in Excel definition](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/151e4bcd-90a0-4d82-8b98-f16bf273e4ff)
also describes the associated Python script contract. A disposable
baseline/candidate pair outside this repository changed both direct PY source
and a static input reaching a separate stored PY binding. The exact wheel's
unredacted JSON retained every controlled marker, confirming the established
local-review contract. Redacted JSON, Markdown, SARIF, and a one-workbook
portfolio JSON artifact omitted every marker. A redacted policy check still
exited `1` with `FF065` and `FFP065`, proving the rendering boundary did not
weaken the policy control.

The isolated environment reported `FormulaFence 0.100.0`. The built wheel
`formulafence-0.100.0-py3-none-any.whl` passed `twine check` with SHA-256
`25d951f45de5520f1052abf51c812b529009721692a1fc377c69cc3070bcc427`.
The source distribution passed `twine check`; its digest is intentionally
omitted because this validation note is included in that source archive. The
final source tree passed **647 tests in 87.73 seconds**, plus a clean Ruff
check, `git diff --check`, and shell syntax check for the composite Action.

## Shared formula external-action / DDE report redaction — 2026-07-26

FormulaFence 0.99.0 adds the separate, opt-in
`--redact-formula-external-actions` rendering boundary for generic reports. It
is deliberately separate from the count-only `FF064` and `FF074` ledgers. The
default local-review output remains unchanged. When enabled, JSON, Markdown,
and SARIF hide direct inventoried action/provider or DDE formula material,
changed action/DDE cell evidence, and exact changed static input evidence that
the private dependency analysis recorded as reaching one. A changed resolved
formula-name chain can carry an endpoint without spelling the native action, so
the mode conservatively hides changed defined-name before/after values in that
case. It does not evaluate a formula, contact a provider or DDE server, change
a comparison/policy result, or reconstruct a dynamically assembled endpoint.

A disposable release-validation pair outside this repository contained a direct
`HYPERLINK`, a `HYPERLINK(A9, ...)` static input, and a formula-defined direct
DDE link. Candidate-only action, input, and DDE markers changed. The exact
wheel's unredacted JSON retained all six controlled markers, confirming the
normal local-review contract. With the redaction option, wheel-produced JSON,
Markdown, SARIF, and a one-workbook portfolio JSON artifact omitted every
marker. The redacted policy check still exited `1` with `FF064`, `FF074`,
`FFP064`, and `FFP074`, proving output redaction did not weaken the policy
boundary.

The isolated environment reported `FormulaFence 0.99.0`. The built wheel
`formulafence-0.99.0-py3-none-any.whl` passed `twine check` with SHA-256
`a72c646d3de630b230078f39eadb5c0958f905a4b21ce2788a457551b4917eaf`.
The source distribution passed `twine check`; its digest is intentionally
omitted because this validation note is included in that source archive. The
final source tree passed **642 tests in 86.63 seconds**, plus a clean Ruff
check, `git diff --check`, and shell syntax check for the composite Action.

## Shared external-workbook-link report redaction — 2026-07-26

FormulaFence 0.98.0 adds an explicit rendering boundary for generic reports:
`--redact-external-workbook-links`. It is deliberately separate from the
count-only `FF081` ledger. Existing local review output remains unchanged when
the option is absent; when present, JSON, Markdown, and SARIF replace an entire
serialized value containing a literal static external-workbook endpoint with a
stable redaction marker. It does not evaluate a formula, change a comparison or
policy result, or attempt to reconstruct a dynamic reference assembled from
text fragments.

A fresh disposable baseline/candidate pair was generated with XlsxWriter 3.2.9
outside this repository. Each workbook contained a direct external A1 formula,
a direct external formula-defined name, and an external data-validation
criterion. Only a controlled source marker changed. The unredacted JSON diff
contained both markers, confirming that the normal local-review contract stayed
intact. The exact release wheel then produced redacted JSON, Markdown, and
SARIF artifacts with neither marker present. Its narrow link-surface policy
returned exit `1` with `FF008`, `FF020`, `FF081`, and `FFP081`, proving the
rendering switch did not weaken policy evidence. An exact-wheel one-workbook
portfolio JSON artifact likewise retained its report structure while omitting
both markers.

The isolated environment reported `FormulaFence 0.98.0`. The built wheel
`formulafence-0.98.0-py3-none-any.whl` passed `twine check` with SHA-256
`06d177296bcd5430e15b2b64c3e4237143af86cb5105b645e1fda6a7c9dc9c35`.
The source distribution passed `twine check`; its digest is intentionally
omitted because this validation note is included in that source archive. The
final source tree passed **636 tests in 85.73 seconds**, plus a clean Ruff
check, `git diff --check`, and shell syntax check for the composite Action.

## Static external-workbook link surfaces — 2026-07-26

Microsoft's [workbook-link guidance](https://support.microsoft.com/en-us/office/create-workbook-links-c98d1803-dd75-4668-ac6a-d7cca2a9b95f)
documents formulas that reference another workbook. Its
[Name Manager guidance](https://support.microsoft.com/en-us/excel/use-the-name-manager-in-excel)
distinguishes formula-bearing names, and the SpreadsheetML
[data-validation definition](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/20ed0abd-113f-4b8a-8de3-c68e733a300a)
permits an external cell reference in a validation criterion. FormulaFence
0.97.0 therefore compares the literal persisted endpoints privately across
worksheet formulas, defined names, data-validation criteria, and parsed chart
formula parts; it does not evaluate a formula, open a source, refresh data, or
trust cached results.

A separate disposable pair was written with XlsxWriter 3.2.9 outside this
repository. Each workbook contained three static surfaces: `Model!D2` held an
external A1 formula, `ExternalLimit` held a direct external workbook-name
formula, and `Model!E2` held an external data-validation criterion. The
candidate changed only the controlled source marker; the existing external
formula remained at `Model!D2`, so the old new-location rule had no reason to
emit `FF004`. This is a controlled integration fixture, not a claim that the
writer evaluated or resolved the non-existent external source.

An isolated environment installed the exact built
`formulafence-0.97.0-py3-none-any.whl` and ran `formulafence check` with only
`no_external_workbook_link_surface_changes: true`. It returned exit `1` with
`FF081` and `FFP081`, no `FF004`, one unchanged external-reference cell, and
three surfaces / three endpoints. The candidate profile, plus the `FF081` and
`FFP081` JSON and SARIF results, omitted both controlled source markers. The
ordinary defined-name and data-validation findings retain their pre-existing
local review evidence; this validation asserts the new ledger's distinct
count-only evidence contract rather than claiming those unrelated reports are
redacted.

The 0.97.0 source tree passed **632 tests in 84.03 seconds**, a clean Ruff
check, and `git diff --check`. Fresh source and wheel distributions passed
`twine check`.

## Static external endpoints inside named LAMBDAs — 2026-07-26

Microsoft documents that a named [LAMBDA function](https://support.microsoft.com/en-us/excel/functions/lambda-function)
is reusable throughout a workbook and called like a native Excel function; its
[workbook-link guidance](https://support.microsoft.com/en-us/excel/create-workbook-links)
also permits formulas to reference another workbook. FormulaFence 0.96.0
therefore extracts only a narrow static input edge from a **workbook-scoped**
named LAMBDA at a real function call. It can preserve both fixed internal
inputs and every already validated direct/package external endpoint in the
body, including through a nested named LAMBDA or a formula-defined wrapper. A
bare LAMBDA name does not create an edge. FormulaFence does not calculate the
definition: dynamic, relative, recursive, local/shadowed, broken, unresolved,
local-3-D, spill, explicit-intersection, and tokenizer-failed forms remain
outside the portfolio graph.

The independently maintained [XlsxWriter table comparison fixture](https://github.com/jmcnamara/XlsxWriter/blob/main/xlsxwriter/test/comparison/xlsx_files/table09.xlsx)
was downloaded to a disposable directory outside this repository (SHA-256
`bf30d9a6b8b94cd5f75c15316a41d54c4063a5745e32ff2f89eb39d252605a04`). It
supplies `Table1` on `Sheet1!B3:K6`. A separately created consumer declared
`ExternalMetric = =LAMBDA(value,SUM(value,Inputs!$B$2,'..\\inputs\\source.xlsx'!Table1[Column2]))`,
a nested `ExternalMetricNested`, and a formula-defined
`ExternalMetricWrapper`; it called those from `Summary!D2`, `Summary!E2`, and
`Summary!F2`, while `Summary!G2` used the bare LAMBDA name as a control. The
unchanged consumer SHA-256 was
`0f94ad37f5403339c19ab5bc64759b456f23a92a07d6251797f6ad4f4b793b07`.
Changing only the disposable source copy's `Sheet1!C4` produced source SHA-256
`4fcb3407f69e90c8494d32b26bd1f577d6984c30727ae512d7b718733188a9d0` and
exactly three `FF079` impacts: `Summary!D2`, `Summary!E2`, and `Summary!F2`.
The `G2` control did not become an impact. JSON omitted the controlled names,
the raw table selector, and the relative external path. The upstream workbook
was never executed, refreshed, modified in place, or copied into this
repository.

The 0.96.0 source tree passed **626 tests in 83.02 seconds**, a clean Ruff
check, and `git diff --check`. Fresh source and wheel distributions passed
`twine check`. An isolated environment installed the exact release wheel and
reran the temporary portfolio through its CLI; it returned policy exit `1`
with `FF079` and `FFP079`, while the controlled names, selector, and path
remained absent from JSON.

## Static external endpoints inside formula-defined names — 2026-07-26

Excel permits a defined name to contain a formula, and Microsoft's
[workbook-link guidance](https://support.microsoft.com/en-us/office/create-workbook-links-c98d1803-dd75-4668-ac6a-d7cca2a9b95f)
shows that a formula can use a reference in another workbook. Its
[defined-name guidance](https://support.microsoft.com/en-US/Excel/names-in-formulas)
also describes names as formula references. FormulaFence 0.95.0 therefore
extracts only a narrow static input edge from a workbook-scoped non-`LAMBDA`
formula-defined name such as
`=SUM('..\\inputs\\source.xlsx'!Table1[Column2])` or a second name that calls
it. It does not calculate either formula: every external token must already
be a parsed static direct or package-validated endpoint, every other reference
must be static, and dynamic, relative, local-3-D, spill, explicit-intersection,
broken, unresolved, tokenizer-failed, local, and `LAMBDA` definitions remain
outside the portfolio graph.

The independently maintained [XlsxWriter table comparison fixture](https://github.com/jmcnamara/XlsxWriter/blob/main/xlsxwriter/test/comparison/xlsx_files/table09.xlsx)
was again downloaded to a disposable directory outside this repository
(SHA-256 `bf30d9a6b8b94cd5f75c15316a41d54c4063a5745e32ff2f89eb39d252605a04`).
It supplies `Table1` on `Sheet1!B3:K6`. A separate consumer used
`ExternalTableFormula = =SUM('..\\inputs\\source.xlsx'!Table1[Column2])` and
`ExternalTableFormulaSecond = =SUM(ExternalTableFormula)`, then called those
names from `Summary!D2` and `Summary!E2`. Changing only the disposable source
copy's `Sheet1!C4` produced source SHA-256
`1ed8b78dc667c0690ce4554d1a6d1c5f4f9742af228e1cec9e2145a04156df90`
and exactly two `FF079` impacts plus `FFP079` at those two cells. The consumer
was unchanged (SHA-256
`b70c8b5b431e408fa3872990ef2b6f5584aaec68ad6a1de173758931fe96b05f`).
The JSON report omitted both controlled formula-name identities, the raw table
selector, and the relative external path. The upstream workbook was never
executed, refreshed, modified in place, or copied into this repository.

The 0.95.0 source tree passed **624 tests in 82.23 seconds**, a clean Ruff
check, and `git diff --check`. Fresh source and wheel distributions passed
`twine check`. An isolated environment installed the release wheel and reran
the temporary portfolio through its CLI; it returned policy exit `1` with
`FF079` and `FFP079`, while the controlled name and selector material remained
absent from JSON.

## Candidate-only external structured-table selectors — 2026-07-26

Microsoft's [structured-reference grammar](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oi29500/089fbdef-ed49-4a14-9509-794c95651b17)
defines a table identifier as a workbook prefix plus a table name, followed by
the table selector; it does not add a source worksheet to that identifier. Its
[structured-reference guidance](https://support.microsoft.com/en-US/Excel/using-structured-references-with-excel-tables)
also documents the data/header/total selector semantics and external-workbook
constraint. FormulaFence 0.94.0 therefore accepts only book-only,
selector-bearing direct forms such as `='..\\inputs\\source.xlsx'!Table1[Column2]`
and package-validated forms such as `=[1]!Table1[Column2]`. It resolves the
selector only after finding exactly one case-insensitive matching table in the
already inspected candidate source; it does not open, fetch, refresh, evaluate,
or use cached external data.

The independently maintained [XlsxWriter table comparison fixture](https://github.com/jmcnamara/XlsxWriter/blob/main/xlsxwriter/test/comparison/xlsx_files/table09.xlsx)
was downloaded to a temporary directory outside this repository (SHA-256
`bf30d9a6b8b94cd5f75c15316a41d54c4063a5745e32ff2f89eb39d252605a04`). It
contains `Table1` over `Sheet1!B3:K6`. A disposable source copy was paired with
a separate consumer using direct one-column and two-column `#Data` selectors,
plus a direct workbook-name alias and one-hop alias. Source-sheet-qualified and
`@` controls remained outside the graph. Changing only the copy's
`Sheet1!C4` produced exactly four `FF079` impacts and `FFP079` at
`Summary!D2`, `Summary!E2`, `Summary!F2`, and `Summary!G2`. JSON, Markdown,
and SARIF omitted the raw selectors and controlled consumer aliases; the safe
profile also omitted the selector text. The upstream artifact was never
executed, refreshed, changed in place, or copied into this repository.

The 0.94.0 source tree passed **624 tests in 82.34 seconds**, a clean Ruff
check, and `git diff --check` before packaging. Fresh 0.94.0 source and wheel
distributions passed `twine check`. An isolated environment installed the exact
final wheel and reran the temporary public portfolio through its CLI; it
returned policy exit `1` with `FF079` and `FFP079`, while the raw selectors and
controlled aliases remained absent from its JSON output.

## Candidate-only external 3-D A1 spans — 2026-07-26

Microsoft's [cell-reference grammar notes](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oi29500/3c420ebb-6ef1-4b0d-959d-76e88c841c3e)
define an external cell reference whose sheet prefix can be a `sheet-range`,
including a workbook index. Its [3-D-reference guidance](https://support.microsoft.com/en-us/excel/create-a-reference-to-the-same-cell-range-on-multiple-worksheets)
states that a span uses every worksheet between its endpoints. FormulaFence
0.93.0 applies that only after a direct or package-validated source spelling
resolves to an already-inspected candidate: the candidate must expose a
complete raw OOXML tab catalog consistent with the inspected ordinary-worksheet
order, and both endpoint tabs must be unique, present, and forward.

The independently maintained [MullinsLab external-data workbook](https://github.com/MullinsLab/excel-external-data/blob/5b4d55319c2eab3ad25408a85de025bdffa35e8b/external-data-blank.xlsx)
was downloaded at commit `5b4d55319c2eab3ad25408a85de025bdffa35e8b` to a
temporary directory outside this repository (SHA-256
`b194aa281d64f1b5cf7f953a328adca211d67245c6b2d0fe64b5245c352a7b68`). A
temporary source copy received controlled `FenceJan`, `FenceFeb`, `FenceMar`,
and out-of-span `FenceOutside` worksheets. A separate consumer used one direct
external `FenceJan:FenceMar` A1 span, a one-hop workbook-name alias to the same
span, a leading-`=` alias, and a reversed-span control. Changing only
`FenceFeb!B3` yielded exactly three `FF079` impacts and `FFP079`; the reversed
control did not add an edge. JSON, Markdown, SARIF, and the source CLI's JSON
omitted the uppercase external-source spelling, all controlled alias names, and
the span endpoint identities. The upstream workbook was neither executed,
refreshed, modified in place, nor copied into this repository.

The release-versioned source tree passed **620 tests in 81.76 seconds**, a
clean Ruff check, and `git diff --check`. Fresh 0.93.0 source and wheel
distributions passed `twine check`. An isolated environment installed the exact
final wheel and reran the temporary public portfolio through its CLI; it
returned policy exit `1` with `FF079` and `FFP079`, while the controlled aliases,
uppercase source spelling, and endpoint identities remained absent from JSON.

## Exact workbook-name alias chains — 2026-07-26

Microsoft's [name-formula grammar](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oi29500/da7f42ad-0083-451a-98cf-b475b578d91d)
permits a `name-reference`, and its [name grammar notes](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oi29500/1399f3cb-927c-4611-96ee-666143a8be35)
define that reference separately from an external name. FormulaFence 0.92.0
uses that narrow grammar only to bridge a finite, acyclic chain of
workbook-scoped aliases: each intermediate stored definition is exactly one
unqualified, non-A1 name identity, with or without the optional leading `=`.
The chain must end at an already-supported direct or package-validated external
A1/global-name/sheet-local-name endpoint. Sheet-local consumer aliases,
operators, functions, ranges, structured references, missing targets, and
cycles stay unresolved; no formula is evaluated.

The independently maintained [MullinsLab external-data workbook](https://github.com/MullinsLab/excel-external-data/blob/5b4d55319c2eab3ad25408a85de025bdffa35e8b/external-data-blank.xlsx)
was downloaded at commit `5b4d55319c2eab3ad25408a85de025bdffa35e8b` to a
temporary directory outside this repository. A temporary source copy gained a
controlled static global name, while a separate consumer used a direct external
A1 alias plus a one-hop bridge, and a direct external global-name alias plus a
two-hop bridge whose final definition omitted `=`. Changing the candidate
copy's `External data!A1` produced exactly four `FF079` impacts and policy
`FFP079`. JSON, Markdown, and SARIF omitted every controlled source and
consumer alias identity and the raw uppercase external-workbook spelling. The
upstream workbook was never executed, refreshed, changed in place, or copied
into this repository.

The release-versioned source tree passed **616 tests in 80.66 seconds**, a
clean Ruff check, and `git diff --check`. Fresh 0.92.0 source and wheel
distributions passed `twine check`. An isolated environment installed the
exact final wheel and reran the temporary public portfolio through both the
library and CLI. The CLI returned policy exit `1` with `FF079` and `FFP079`;
the controlled aliases and uppercase external source spelling remained absent
from its JSON output.

## Direct external-alias portfolio boundary — 2026-07-26

Microsoft's [cell-reference grammar notes](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oe376/e531ebe0-a152-4978-a876-28e2a68f746e)
separate an external cell reference from an A1 reference, and its
[name grammar notes](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oe376/a469e5f5-a102-49bd-9642-8a8e8aaf1623)
distinguish an external name. FormulaFence 0.91.0 therefore follows a
workbook-scoped consumer alias only when its stored text is one exact static
direct external A1 or workbook-scoped-name literal. It accepts the canonical
optional leading `=` spelling only when the remaining text is still exactly
that literal. Sheet-local consumer aliases and formula wrappers remain outside
the graph; a sheet-local name with the same identity shadows the global alias.
The source is still resolved only as an already-inspected relative candidate,
and FormulaFence never opens, fetches, refreshes, caches, or evaluates a link.

An independently maintained
[MullinsLab external-data workbook](https://github.com/MullinsLab/excel-external-data/blob/5b4d55319c2eab3ad25408a85de025bdffa35e8b/external-data-blank.xlsx)
was downloaded at commit `5b4d55319c2eab3ad25408a85de025bdffa35e8b` into a
temporary directory outside this repository. A temporary source copy gained
one controlled static workbook name, and a separate temporary consumer used
direct A1, direct workbook-name, and leading-`=` A1 aliases to that source.
Changing the original workbook's `External data!A1` in the candidate copy
produced exactly three impacts with `FF079` and policy `FFP079`. JSON,
Markdown, and SARIF omitted each controlled source/consumer alias identity,
the upstream filename, and the repository identity. The upstream workbook was
never executed, refreshed, changed in place, or copied into this repository.

The release-versioned source tree passed **614 tests in 80.20 seconds**, a
clean Ruff check, `git diff --check`, and GitHub Action shell syntax
validation. Fresh 0.91.0 source and wheel distributions passed `twine check`.
An isolated environment installed the exact final wheel and reran the same
temporary portfolio both through the library and the CLI. The CLI returned
policy exit `1` with `FF079` and `FFP079`; the controlled aliases, upstream
filename, and repository identity remained absent from its JSON output.

## External sheet-local-name portfolio boundary — 2026-07-26

Microsoft's [name grammar notes](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oe376/a469e5f5-a102-49bd-9642-8a8e8aaf1623)
define an external name as a `single-sheet-prefix` plus a name, while its
[cell-reference grammar](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oe376/e531ebe0-a152-4978-a876-28e2a68f746e)
defines that prefix and the nonzero package-index meaning. Excel's
[Name Manager guidance](https://support.microsoft.com/en-us/excel/use-the-name-manager-in-excel)
also distinguishes worksheet and workbook scope. FormulaFence 0.90.0 therefore
recognizes only static direct `[Book.xlsx]Sheet!LocalName` and validated
package-indexed `[N]Sheet!LocalName` forms. The named source must resolve only
inside the explicitly named source sheet's local scope to fixed internal A1
destinations; it never substitutes a same-named global or different-sheet
local name. A direct or package-indexed workbook-scoped consumer alias can
retain that exact static spelling, but no target is opened, fetched, refreshed,
cached, or evaluated.

An independently maintained
[MullinsLab external-data workbook](https://github.com/MullinsLab/excel-external-data/blob/5b4d55319c2eab3ad25408a85de025bdffa35e8b/external-data-blank.xlsx)
was downloaded at commit `5b4d55319c2eab3ad25408a85de025bdffa35e8b` into a
temporary directory outside this repository. Its OOXML has a real static
worksheet-local `External_data` name scoped to `External data` and covering
`A1:N9592`. A temporary baseline/candidate copy gained one controlled `A1`
value change; a separate temporary consumer used the quoted direct local-name
form. FormulaFence emitted `FF079` and policy `FFP079` for the one reachable
consumer formula, with exit status `1`. JSON contained only relative workbook
paths and logical cells: the source-name identity, upstream filename, and
repository identity were absent. The upstream workbook was neither executed,
refreshed, changed in place, nor copied into this repository.

The final source checkout passed **613 tests in 80.13 seconds**, a clean Ruff
check, `git diff --check`, and GitHub Action shell syntax validation. Fresh
0.90.0 source and wheel distributions passed `twine check`. An isolated virtual
environment installed the exact final wheel and reran the temporary independent
portfolio with `no_cross_workbook_impacts`; it returned policy exit `1` with
both `FF079` and `FFP079`, while `External_data`, the upstream filename, and
the repository identity remained absent from JSON.

## Package-indexed external-A1 portfolio boundary — 2026-07-26

Microsoft's [cell-reference grammar notes](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oe376/e531ebe0-a152-4978-a876-28e2a68f746e)
describe the package cell form `[N]Sheet!A1` and state that a nonzero index
identifies an external book in the external-link collection. FormulaFence
0.89.0 treats that integer as package metadata, never as a filename. It accepts
only one static A1 cell/range/whole-row/whole-column form, and only after the
same declaration-order, single-`externalLink`, single-`externalBook`, and
external-`externalLinkPath` validation used by the indexed-name boundary. A
direct workbook-scoped consumer alias may store that exact static spelling;
sheet-local aliases, formula aliases, 3-D forms, names, structured references,
caches, malformed syntax, and unsafe targets remain unresolved.

A controlled two-link source/consumer portfolio placed the source workbook in
the second `externalReference` position while workbook relationships were
written in reverse order. One direct `[2]Data!$B$2:$B$4` formula and one direct
workbook-scoped alias both produced `FF079` and `FFP079` when the covered source
cell changed. Whole-column parsing, absolute target rejection, local/formula
alias rejection, malformed package links, and JSON/Markdown/SARIF redaction
are covered in the suite. No package target, index spelling, or consumer alias
identity reaches portfolio evidence.

The independently maintained public
[openpyexcel external-link fixtures](https://github.com/sciris/openpyexcel/tree/1fde667a1adc2f4988279fd73a2ac2660706b5ce/openpyexcel/workbook/external_link/tests/data)
were downloaded at commit `1fde667a1adc2f4988279fd73a2ac2660706b5ce` into a
temporary directory outside this repository. Its `workbook_external_range.xml`
contains a real workbook-scoped `[1]Sheet1!$A$1` alias. A temporary consumer
copy was changed only to call that existing alias; changing the paired
`book2.xlsx` source cell yielded a complete `FF079` graph with paths to the
consumer formulas and no raw external relationship target in JSON, Markdown,
or SARIF. Neither upstream workbook was executed, refreshed, changed in place,
or copied into this repository.

The final source checkout passed **610 tests in 81.62 seconds**, a clean Ruff
check, `git diff --check`, and GitHub Action shell syntax validation. Fresh
0.89.0 source and wheel distributions passed `twine check`. An isolated virtual
environment installed the exact wheel and ran the temporary public-fixture
portfolio with `no_cross_workbook_impacts`; it returned policy exit `1` with
both `FF079` and `FFP079`, while the existing external-alias identity remained
absent from the JSON report.

## Package-indexed external-name portfolio boundary — 2026-07-26

Microsoft's [name grammar notes](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oe376/a469e5f5-a102-49bd-9642-8a8e8aaf1623)
state that Office prefixes an external name with the index of its associated
relationship, while its [external-name formula grammar](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oe376/75328f70-50a7-43af-a4da-3abade67f5f9)
limits the stored definition to references in the same external book.
FormulaFence 0.88.0 uses that evidence only as a bounded candidate-portfolio
bridge: `[N]!SourceName` must select one document-order `externalReference`,
one declared `externalLink` package part, one `externalBook`, and one external
`externalLinkPath` relationship. The private target must then match an
already-inspected relative candidate, and the source's workbook-scoped name
must fully expand to static internal A1 destinations. No cached external-link
value or name definition is trusted. Microsoft's [external-workbook base-path
notes](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-xlsb/ea3aa674-0411-498d-9eb0-3c1f99a9c0c0)
describe `externalLinkPath` as relative to the package containing that
relationship; FormulaFence deliberately rejects Excel's alternate
startup/library/missing-path relationship forms rather than inferring one.

A controlled two-link source/consumer portfolio deliberately wrote workbook
relationships in reverse order, while the consumer's `[2]!PrivateName` direct
formula and workbook-scoped alias referred to the second declaration. Changing
one static source cell produced `FF079` and `FFP079` for exactly those two
consumer formulas. A dynamic source name, an absolute package target, and an
external-link part containing two `externalBook` definitions produced no edge.
Rebinding two indexed declarations to one external-link part also produced no
edge. Portfolio JSON, Markdown, and SARIF omitted the controlled indexed
source spelling and private target sentinel. As intended, source and consumer
defined-name declarations remained ordinary items in their respective profile
inventories; no profile exposed the private package relationship target or a
source-to-consumer mapping.

The independently maintained public
[openpyexcel external-link fixtures](https://github.com/sciris/openpyexcel/tree/1fde667a1adc2f4988279fd73a2ac2660706b5ce/openpyexcel/workbook/external_link/tests/data)
were rechecked at commit `1fde667a1adc2f4988279fd73a2ac2660706b5ce` in a
temporary directory outside this repository. Its consumer uses a
workbook-local name backed by `[1]!B2range` and `externalLink1.xml`; changing
`book2.xlsx` `Sheet1!A2` yielded `FF079` and `FFP079` only at
`book1.xlsx` `Sheet1!C1`. The fixture was not executed, refreshed, or changed
in place; temporary copies and report output contained only relative workbook
identities and logical cells.

The final source checkout passed **608 tests in 79.80 seconds**, a clean Ruff
check, and `git diff --check`. A fresh isolated virtual environment then
installed the exact built 0.88.0 wheel outside the checkout. Its package-indexed
source/consumer pair produced `FF079` and `FFP079` for one direct `[1]!Name`
formula and one direct workbook-scoped consumer alias, with two reachable
formulas total. The JSON omitted the temporary source path and every controlled
indexed source-name spelling; ordinary source and consumer names remained
normal defined-name profile context. Both the wheel and source distribution passed
`twine check`.

## Cross-workbook defined-name portfolio boundary — 2026-07-26

Microsoft's [workbook-link guidance](https://support.microsoft.com/en-us/excel/create-workbook-links)
documents that another workbook can be referenced by a cell/range or by a
defined name. FormulaFence 0.87.0 adds only the direct workbook-scoped syntax
such as `=[source.xlsx]InputRange`, and only when the exact relative path names
an already-inspected candidate workbook whose source name fully expands to
static internal A1 destinations. This is static reachability evidence, not an
attempt to calculate Excel or resolve a link package.

A controlled source/summary portfolio declared a private direct global range,
a formula-defined global alias of that range, and an `OFFSET`-based name. A
summary formula used the exact-relative external alias. Changing one source
cell produced `FF079` and `FFP079` for the one reachable summary formula. The
dynamic name, sheet-qualified name form, absolute-path form, and
portfolio-escaping form produced no invented edge. JSON, Markdown, and SARIF
contained only relative workbook identities and logical cells: the controlled
source path and all defined-name identities remained absent.

The independently maintained public
[openpyexcel external-link fixtures](https://github.com/sciris/openpyexcel/tree/1fde667a1adc2f4988279fd73a2ac2660706b5ce/openpyexcel/workbook/external_link/tests/data)
were rechecked at commit `1fde667a1adc2f4988279fd73a2ac2660706b5ce`. Their
consumer formulas use workbook-local names backed by external-link-package
metadata, not a direct external workbook-scoped formula token. FormulaFence
0.87.0 continued to record no portfolio edge rather than infer a filename, package
target, or name identity from that metadata. Neither public workbook was
executed, refreshed, altered, copied into this repository, nor emitted in a
report.

The final source checkout passed **605 tests in 87.89 seconds**, a clean Ruff
check, and `git diff --check`. A fresh isolated virtual environment then
installed the built 0.87.0 wheel outside the checkout. Its source/summary pair
produced `FF079` and `FFP079` only for the static alias destination; the dynamic
and absolute forms did not add an impact. The resulting JSON omitted the
temporary root, external path sentinel, and every controlled name identity.
Both wheel and source distribution passed `twine check`.

## Cross-workbook portfolio impact boundary — 2026-07-26

FormulaFence 0.86.0 implements the direct external-A1 forms documented in
Microsoft's [workbook-link guidance](https://support.microsoft.com/en-us/excel/create-workbook-links)
only when an exact relative source path maps to an already-inspected candidate
workbook. It never opens, refreshes, calculates, or downloads a link target.

The independently maintained public
[openpyexcel external-link fixtures](https://github.com/sciris/openpyexcel/tree/1fde667a1adc2f4988279fd73a2ac2660706b5ce/openpyexcel/workbook/external_link/tests/data)
were inspected at commit `1fde667a1adc2f4988279fd73a2ac2660706b5ce`.
`book1.xlsx` and `book2.xlsx` are real OOXML workbooks: the first contains one
explicit external-reference formula and one external-link package, but no
direct static external A1 cell/range. FormulaFence recorded zero resolvable
portfolio edges rather than guessing that its external name/package target was
the paired filename. Neither workbook was executed, refreshed, altered, copied
into this repository, nor emitted in a report.

A clean virtual environment then installed the built 0.86.0 wheel outside the
source checkout. A disposable source/summary pair used the
documented direct form `=[source.xlsx]Data!A1`; changing the source produced
`FF079` and `FFP079`, with two reachable summary formulas and exit code 1. The
same summary also contained an absolute-link sentinel. JSON, Markdown, and
SARIF output retained only relative workbook identities and logical cells: the
temporary root and sentinel did not appear. The source distribution and wheel
both passed `twine check`.

The final source checkout passed **602 tests in 87.67 seconds**, a clean Ruff
check, shell syntax validation, `git diff --check`, and another isolated
wheel/source-distribution build.

## Portfolio change-control boundary — 2026-07-26

FormulaFence 0.85.0 was exercised against the independently maintained public
[Python in Excel course workbooks](https://github.com/LinkedInLearning/python-in-excel-quick-start-4551222)
at commit `cbf3219e864a1816adb8019d9ca4683de7429dae`. The portfolio contract is
also motivated by the spreadsheet change-review workflow described in
[Governance of Spreadsheets through Spreadsheet Change Reviews](https://arxiv.org/abs/1211.7100):
inventory and an auditable structural-change boundary are practical controls
for a group of critical spreadsheets. No workbook content was executed,
uploaded, or copied into this repository.

A temporary baseline contained four unmodified public workbooks. The candidate
kept two files byte-identical, placed the course's corresponding finished
workbook at the same relative path as one start workbook, removed one public
workbook, and added another public workbook. `formulafence portfolio` reported
five relative paths: three matched, one semantically changed, two unchanged,
one added, one removed, and zero unreadable. The changed real workbook produced
13 semantic changes and nine stored-control findings; FormulaFence did not
calculate any formula or Python code. The added and removed paths each emitted
exactly one `FF077` record, while a narrow
`no_portfolio_membership_changes` policy added `FFP077` for each and exited 1.

The same real-corpus run with `--fail-on high --format sarif` exited 1 and
produced 11 SARIF results. Every physical artifact URI was a relative workbook
path (never the temporary directory), including the added/removed records.
The external run therefore verifies actual recursive inventory, matching,
single-workbook comparison, membership policy, threshold enforcement, and
SARIF attribution in one local-only workflow. It intentionally does not infer
a rename or compare any workbook that cannot be read.

A freshly created virtual environment installed the built 0.85.0 wheel with
its declared dependencies, returned `FormulaFence 0.85.0`, and reproduced the
five-workbook public portfolio report plus the expected membership-policy exit
code of 1 without importing the source checkout. Both wheel and source
distribution passed `twine check` before this smoke run.

Controlled fixtures separately cover a same-path formula-to-value edit with
per-workbook policy enforcement, additions/removals, an unreadable archive that
still yields redacted `FF078` evidence and exit 2, membership evidence retained
when an unreadable file is newly added or removed, safe JSON/Markdown/SARIF
output for a new workbook containing a private sentinel, default inventory
limits, unsupported Excel formats, Office lock-file skipping, case-portable
identity, symlink refusal, and report-output protection. The full suite passed
**594 tests in 82.20 seconds**, followed by a clean Ruff check.

## PythonScripts compatibility boundary — 2026-07-26

FormulaFence 0.84.0 was checked against Microsoft's [Python in Excel
introduction](https://support.microsoft.com/en-US/Excel/python/introduction-to-python-in-excel)
and the OOXML [Python part](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/151e4bcd-90a0-4d82-8b98-f16bf273e4ff)
definition. The newer standard part is not the only package shape found in
real workbooks, so the scanner was also exercised against independently
maintained public [Python in Excel course workbooks](https://github.com/LinkedInLearning/python-in-excel-quick-start-4551222).
No workbook formulas or Python source were executed, copied into this
repository, or emitted in a report.

The public `Debug.xlsx` workbook carries only the separate 2022
`pythonScripts.xml` contract: its `pythonscripts` content type, `PythonScripts`
workbook relationship, and one script record are all present. Before this
release it appeared as zero inspected Python parts/scripts with a coverage gap;
FormulaFence 0.84.0 reports one physical Python part, one stored script, one
`PY` formula cell/call, and no unsupported-Python metadata. The public
`DFexplore_finished.xlsx` workbook carries both the 2023 `python.xml` and 2022
`pythonScripts.xml` contracts. It reports two physical parts and thirty stored
script records (fifteen in each stored representation), one environment
definition/initialization, fifteen `PY` formula cells/calls, and no Python
coverage gap. This verifies package inventory only; it does not decide which
representation a particular Excel runtime will execute.

A temporary copy of the legacy-only public workbook then received a code-only
change inside its `pythonScripts.xml` part. A fresh virtual environment
installed the release wheel, reported `FF065` with a private definition change,
and the report JSON contained neither the original nor replacement source
string. This confirms that the compatibility format is both detected and
redacted in an independently maintained workbook, rather than only in a
synthetic fixture.

Controlled fixtures additionally changed only a private 2022 script, changed
that script while a 2023 part coexisted, renumbered only its workbook
relationship ID, and malformed its XML. The first two emit `FF065` and
`FFP065` under the existing policy; the ID rewrite normalizes; malformed XML
remains explicit coverage evidence. Profile JSON, Markdown, report JSON, and
SARIF were checked for every controlled source sentinel and exposed none. The
full suite passed **580 tests in 88.97 seconds**, followed by a clean Ruff
check.

## XLM automatic-macro binding boundary — 2026-07-26

FormulaFence 0.83.0 was checked against Microsoft's
[RunAutoMacros documentation](https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.runautomacros),
which identifies the workbook-attached automatic macro surface, and its
[XlRunAutoMacro enumeration](https://learn.microsoft.com/en-us/office/vba/api/excel.xlrunautomacro),
which lists `Auto_Open`, `Auto_Close`, `Auto_Activate`, and
`Auto_Deactivate`. The OOXML defined-name compatibility notes independently
confirm that `_xlnm.` is a reserved built-in-name prefix, `localSheetId` marks a
sheet-local name, and `definedName@xlm` is reserved and unused. Microsoft also
continues to document that Excel supports XLM macros under its macro-security
settings; FormulaFence deliberately does not infer those settings or runtime
execution.

Controlled macro-enabled OOXML fixtures were generated and mutated without
opening Excel. Each fixture had a raw declared XLM macro sheet and harmless
private test cells. The boundary recognized all four documented event names,
including `_xlnm.Auto_Open` and mixed-case forms. Retargeting `_xlnm.Auto_Open`
from one macro-sheet A1 cell to another kept every public automatic-binding
count at `1 / 1 / 0 / 0 / 0` while emitting `FF076`; it did not emit `FF026`,
because the macro-sheet declaration and program remained unchanged. The policy
fixture converted the result to high-severity `FFP076`.

Counterexamples added a workbook `Auto_Open` name pointing to an ordinary
worksheet, a sheet-local `Auto_Close` name pointing to the macro sheet, and an
`Auto_Activate` multi-cell macro-sheet range. None was classified as an XLM
automatic binding. This validates the narrow static contract, not any claim
about a dynamic name formula, external target, range, or local name being
executable. The specialized profile section, Markdown,
`FF076` change details, `FFP076`, and SARIF result were checked for the
controlled name and target formulas and contained none. The ordinary
defined-name diff intentionally retains reviewer-visible name/formula context
outside this redacted ledger. A duplicate raw macro-sheet workbook relationship
made the sheet binding ambiguous, produced an explicit automatic-binding
coverage warning, and removed it from the specialized inventory; comparing it
with the valid fixture still emitted `FF076` rather than silently accepting a
guessed target. The full suite passed **574 tests in 77.52 seconds**.

## Unqualified runtime-function candidate boundary — 2026-07-26

FormulaFence 0.82.0 was checked against Microsoft's [alphabetical Excel
function catalogue](https://support.microsoft.com/en-us/office/excel-functions-alphabetical-b3944572-255d-4efb-bb96-c6d90033e188),
[installed UDF reference](https://support.microsoft.com/en-us/excel/user-defined-functions-that-are-installed-with-add-ins-reference),
[VBA custom-function guidance](https://support.microsoft.com/en-us/excel/create-custom-functions-in-excel),
and [XLL registration/call guidance](https://learn.microsoft.com/en-us/office/client-developer/excel/accessing-xll-code-in-excel).
Those sources establish that a bare worksheet function can be supplied by an
installed add-in, VBA, or a registered XLL, while the stored formula alone does
not identify its provider or prove it can run.

Controlled `.xlsx` fixtures were generated without opening Excel. The direct
fixture had three candidate formula cells and four bare calls, alongside native
`SUM`, `XLOOKUP`, `VSTACK`, `FIELDVALUE`, `PY`, a workbook-defined `LAMBDA`, and
a local `LET`/`LAMBDA` binding as exclusion controls. Replacing one private
candidate name preserved the public `3 / 4 / 0` formula-cell/call/definition
counts while emitting `FF075`; changing only a static input emitted `FF075`
with a static-input count of one. The named fixture exercised a formula-defined
value, a named `LAMBDA`, and a nested named `LAMBDA`, producing three invoking
formula cells, five calls, and three relevant definitions. Same-count hidden
definition changes and static inputs were separately guarded. Recursive and
sheet-local named-LAMBDA tests verified cycle safety and local precedence. An
uninvoked stored definition remained independently reviewable as `0 / 0 / 1`
without exposing its candidate name.

The formula parser tests also verify that namespaced calls stay under `FF066`,
documented native calls do not become generic candidates, and a current
openpyxl native-function catalogue is a subset of FormulaFence's pinned
allowlist; FormulaFence does not import that mutable catalogue at runtime. The
public profile, Markdown, `FF075` details, `FFP075` policy result, and SARIF
result were checked for all controlled candidate names and private input/query
values and contained none. The full suite passed **569 tests in 80.99 seconds**.

## GitHub Action execution boundary — 2026-07-26

FormulaFence 0.81.0 was exercised as the root composite GitHub Action, rather
than only as a CLI invocation. The contract tests parse the public action
metadata, generate a fresh baseline/candidate `.xlsx` pair, and run the Action
shell entry point in an isolated workspace. A `no_formula_to_value` policy
failure generated a Markdown report and job-summary evidence, wrote exit code
`1` to `GITHUB_OUTPUT`, and deliberately left the entry point successful so
the composite Action can upload the report before its final step re-emits the
failure. A valid pair returned exit code `0`.

The same contract rejects a report path or workbook/policy input that escapes
`GITHUB_WORKSPACE`; it also refuses a report path that resolves to any supplied
input. The workflow file passed `actionlint`, and a clean Python environment
installed the Action source with `install: true`, returned `FormulaFence
0.81.0`, and produced a report. This boundary does not execute formulas or
macros, comment on a pull request, or send workbook contents to a FormulaFence
service.

## Direct DDE-style formula-link boundary — 2026-07-26

FormulaFence 0.80.0 was checked against the Windows [Dynamic Data Exchange
overview](https://learn.microsoft.com/en-us/windows/win32/dataxchg/about-dynamic-data-exchange),
which documents Excel's `='Quote'|'NYSE'!ZAXX` application/topic/item formula
shape, and Excel's [DDE security settings](https://learn.microsoft.com/en-us/troubleshoot/microsoft-365-apps/excel/security-settings),
which distinguishes server lookup from the not-recommended server-launch
option. The scope is lexical and static: FormulaFence recognizes only a pipe
outside quoted text followed by a topic and an `!item` boundary; it skips ordinary quoted
sheet names and double-quoted strings. It never evaluates a formula, resolves
an endpoint, looks up or starts a DDE server, sends a command, or attempts to
reproduce Trust Center behavior.

Four fresh controlled `.xlsx` files were generated in a standalone validation
directory without opening Excel. The baseline uses one direct worksheet link,
one direct formula-defined name, a nested name, and a named `LAMBDA`
invocation. Its dedicated public ledger reports four invoking formula cells,
four links, and three relevant formula-defined names. Changing only a private
named-definition topic kept every public count fixed while emitting `FF074`;
the narrow policy exited 1 with `FFP074`. Changing only the visible argument to
the invoking named `LAMBDA` also kept the ledger equal and emitted `FF074` with
a static-input count of one.

The state-only candidate adds an unrelated worksheet while retaining every
direct-DDE formula and name. It emitted no `FF074` or `FFP074` and passed the
narrow policy with exit 0. The suite additionally covers quoted-sheet and
string-literal non-matches, quoted/unquoted/embedded/missing-item syntax,
same-count definition changes, static inputs, uninvoked definitions, named
`LAMBDA` tokenizer fallback, and sheet-local name precedence. The full suite
passed with 552 tests.

The dedicated `formula_dde_links` profile object and the `FF074` / `FFP074`
SARIF results excluded controlled services, topics, items, formulas, inputs,
and name identities. Ordinary semantic and defined-name differences remain
normal reviewer context, so this privacy claim is deliberately limited to the
dedicated ledger and policy-facing results. The staged
`formulafence-0.80.0-py3-none-any.whl` (SHA-256
`be9f57844e0e7f519b7809e771cc5f716473e65de339cd0e63766b569626aeb5`)
and its source distribution passed `twine check`. The source-distribution
digest is intentionally omitted here because this validation note is itself
included in the source archive. The wheel installed with declared dependencies
into a fresh environment, returned `FormulaFence 0.80.0`, and reproduced the
direct DDE validation results above.

## Formula-defined XLM action and event-dispatch boundary — 2026-07-26

FormulaFence 0.79.0 was checked against Microsoft's [Excel C API
reference](https://learn.microsoft.com/en-us/office/client-developer/excel/programming-with-the-c-api-in-excel),
which describes XLM command-equivalent functions and event traps including
`ON.ENTRY` and `ON.TIME`, and its [DLL-access
guidance](https://learn.microsoft.com/en-us/office/client-developer/excel/how-to-access-dlls-in-excel),
which documents `CALL` and `REGISTER` as XLM macro-sheet routes to DLL
functions or commands. The scope remains static: FormulaFence inventories only
the selected stored-definition spellings `CALL`, `EXEC`, `EXECUTE`, `RUN`,
`SEND.KEYS`, `ON.DATA`, `ON.DOUBLECLICK`, `ON.ENTRY`, `ON.KEY`, `ON.RECALC`,
`ON.SHEET`, `ON.TIME`, and `ON.WINDOW`; it does not evaluate a formula,
resolve a target or handler, load a DLL, send DDE, execute a macro or program,
or interpret arbitrary XLM commands.

Four fresh controlled `.xlsx` files were generated in a standalone validation
directory without opening Excel. The baseline (SHA-256
`4d43bb04d24e690bbe1199f8d2d5253880c955d78dfde8c27573c3d127a8c174`)
uses a named `LAMBDA` with `EXEC` and `RUN`, a nested named `LAMBDA`, and a
formula-defined `ON.TIME` value. Its dedicated public ledger reports three
invoking formula cells, five selected calls, and three relevant
formula-defined names. Changing only a private `RUN` target (SHA-256
`2954c95263ec4234cfabe1082e3c5bcab2478bb64f7e9dc543523483709cc93b`)
kept those public counts fixed while emitting `FF073`; the narrow policy
exited 1 with `FFP073`. Changing only a statically visible payload input
(SHA-256
`6ea7e7e6a3a465e2a9581b8a03c9231a54abc7c9db241e2107d09b6bcb23631b`)
kept the action ledger equal and emitted `FF073` with a static-input count of
one.

The state-only candidate (SHA-256
`9559f91fb58e5961ee6ac6991be1576a5d05b10d363c689c3f7ef9c86d897db3`)
adds an unrelated worksheet while retaining every stored action definition. It
emitted no `FF073` or `FFP073` and passed the narrow policy with exit 0,
demonstrating that FormulaFence does not simulate event dispatch or claim that
unrelated workbook state changes an action. The suite separately covers every
selected spelling, nested names, recursive names, sheet-local resolution,
workbook-defined callable shadowing, direct worksheet-call exclusion,
uninvoked definitions, static inputs, private signatures, policy enforcement,
and report rendering. The full suite passed with 545 tests.

The dedicated `formula_defined_xlm_actions` profile object and `FF073` /
`FFP073` SARIF results excluded the controlled formula-name identities and
action/input sentinels. Ordinary defined-name and semantic-diff output remains
normal reviewer context, so that redaction claim is deliberately limited to the
dedicated ledger and policy-facing results. No formula was evaluated and no
macro, program, DLL, DDE command, or event handler was invoked during
validation.

The staged `formulafence-0.79.0-py3-none-any.whl` (SHA-256
`2d5a67246ad36558607b242924b4ea46aa6b9af5fd9d9e74a3c104c4e0e1d26e`)
and its source distribution were built from the release tree and passed `twine
check`. The wheel installed with declared dependencies into a fresh virtual
environment and returned `FormulaFence 0.79.0`; it reproduced the 3 / 5 / 3
ledger counts, emitted `FF073` / `FFP073` for the definition candidate, and
retained the unrelated-sheet policy control. Dedicated packaged SARIF results
remained redacted.

## Workbook tab information formula boundary — 2026-07-26

FormulaFence 0.78.0 was checked against Microsoft's [SHEET function
documentation](https://support.microsoft.com/en-us/excel/functions/sheet-function),
which documents sheet-number behavior and its optional value, and [SHEETS
function documentation](https://support.microsoft.com/en-us/excel/functions/sheets-function),
which documents that an omitted reference counts sheets in the containing
workbook. Both document that hidden, very-hidden, macro, chart, and dialog
sheets are included. The scope remains static: FormulaFence inventories native
`CELL`, `INFO`, `SHEET`, and `SHEETS` calls, and privately compares only the
raw all-tab catalog relevant to stored `SHEET` and omitted-reference `SHEETS()`
calls. It does not calculate a formula, resolve an explicit reference, infer a
3-D reference, or simulate workbook state.

Five fresh controlled `.xlsx` files were generated in a standalone validation
directory without opening Excel. The baseline (SHA-256
`8090ef0f9b3d90ff47617916c1493bfc58c4f99812c61ed46835f6fbefa4b971`)
contains a direct `SHEET()`, a direct `SHEETS()`, a named formula with omitted
`SHEETS()`, and an explicitly referenced `SHEETS(reference)` control. Its
dedicated public ledger reports four invoking formula cells, four native calls,
one relevant formula-defined name, one `SHEET` call, three `SHEETS` calls, and
two omitted-reference `SHEETS` calls. Inserting a valid chart sheet at tab index
one (SHA-256
`4621ab3d3de93f393ecadbb017bbfdadb3f548dc1e4c7959829ccee92a75d29b`)
kept that ledger unchanged while emitting `FF072` with the private
workbook-tab-catalog condition; the narrow policy exited 1 with `FFP072`.

The visibility-only candidate (SHA-256
`0b56ab5eeed3d6366f20e4d2b27e1621e7afbc4ee8ea02d5728b84a4787236db`)
hid an existing worksheet. It emitted ordinary `FF007` but no `FF072` or
`FFP072`, and the narrow policy exited 0. An explicit-reference-only baseline
(SHA-256
`93af5f41b78341ea279a8fdb094e7ca4f3da40b2b6dc2494351f6a9092ba0e79`)
and inserted-tab candidate (SHA-256
`8a3202f0391a00cb9912f4bb265511138e5754064eda2a58d456eb7667f67fc0`)
produced no `FF072` or `FFP072`; the narrow policy again exited 0. That control
demonstrates the intentional limit: FormulaFence inventories
`SHEETS(reference)` but does not guess whether a reference covers one sheet or
a 3-D span.

Dedicated `formula_environment_information_calls`, `FF072`, and `FFP072`
artifacts excluded the controlled defined-name identity and static input
sentinel. Ordinary defined-name and semantic-diff output remains normal
reviewer context, so that redaction claim is deliberately limited to the
dedicated ledger and policy-facing results. No formula was evaluated and no
workbook state was simulated during validation. The full suite passed with 535
tests.

The staged `formulafence-0.78.0-py3-none-any.whl` (SHA-256
`fc8b862c5453478e1ece554c5e7d927a6eca916719c55aa1594140f500c3e8b5`)
and source distribution were built from the release tree. The wheel installed with declared dependencies
into a fresh virtual environment and returned `FormulaFence 0.78.0`; it
reproduced the 4 / 4 / 1 / 1 / 3 / 2 ledger counts, emitted `FF072` / `FFP072`
for the chart-tab candidate, and retained the visibility-only and
explicit-reference policy controls. Dedicated packaged SARIF results remained
redacted.

## Formula data-provider boundary — 2026-07-26

FormulaFence 0.77.0 was checked against Microsoft's
[`STOCKHISTORY` documentation](https://support.microsoft.com/en-us/office/stockhistory-function-1ac8b5b3-5f62-4d94-8ab8-7504ec7239a8),
which describes retrieving historical financial-instrument data, and the
documented [Cube function category](https://support.microsoft.com/en-us/office/excel-functions-by-category-5f91f4e9-7b42-46d2-9bd1-63f26a86c0eb),
including `CUBESET` creating a server-side set and `CUBEVALUE` retrieving an
aggregate through a stored connection. The scope remains static: FormulaFence
records stored `STOCKHISTORY` and all seven Cube spellings, but does not
calculate them, resolve a connection, contact a market provider, query a cube,
or interpret a returned value.

Four fresh, controlled `.xlsx` artifacts were generated without opening Excel.
The baseline (SHA-256
`bd65d08ff6ab0eaa5f1f179bdc13d8038e6d298ffea10b1b63da2d9b649002b9`)
uses direct calls for every Cube function, direct `STOCKHISTORY`, a named
`LAMBDA`, and a nested named `LAMBDA`/formula-defined value. Its public
`formula_external_actions` ledger reports eleven invoking formula cells,
eleven calls, three relevant formula-defined names, three `STOCKHISTORY`
calls, and eight Cube calls. Changing only a private formula-defined Cube
connection (SHA-256
`ca7c236525ac677c968360b21cf2a77f8bab09e4cbdfd911b831c3174f54769e`)
kept every public count fixed while emitting `FF064` with the private
definition-material flag. Changing only a statically visible stock input
(SHA-256
`de6d3a630a25fdae0e6fa28e6f39b2638f73b6ff731b3aafa9bc3ce996b7a907`)
kept the ledger equal and emitted `FF064` with a static-input count of one.

The state-only candidate (SHA-256
`a2991bb14ed0c0a8b47409104b091671d65d3373cb419c463985b1affe9a76d0`)
adds an unrelated worksheet while retaining every stored provider call. It
emitted no `FF064` or `FFP064` and passed the narrow policy with exit 0,
demonstrating that FormulaFence does not simulate a provider refresh or infer
an external result from workbook state. The suite separately covers direct
calls, the entire Cube family, `_xlfn.` compatibility spelling, named chains,
definition-only changes, static inputs, and native-name shadowing.

Dedicated `FF064` / `FFP064` SARIF results and the profile excluded controlled
formula-name identities, stock symbols, Cube connections, MDX expressions, and
field/property text. Ordinary semantic and defined-name changes intentionally
remain normal reviewer context, so that redaction claim is limited to the
dedicated ledger and policy-facing results. No formula was evaluated and no
market provider, Cube, connection, or external service was contacted during
validation. The full suite passed with 527 tests.

The staged `formulafence-0.77.0-py3-none-any.whl` (SHA-256
`d6e4cfa72c8d4f87f3d76064e05781d2c134c23372433a315b2b7512554f7ce8`)
was built from the release tree and installed with its declared dependencies
into a fresh virtual environment. Its CLI returned `FormulaFence 0.77.0`; the
definition candidate exited 1 with `FF064` / `FFP064`, while the state-only
candidate exited 0 without either rule. The packaged profile retained the
11 / 11 / 3 / 3 / 8 formula-cell, call, defined-name, `STOCKHISTORY`, and
Cube-call counts; dedicated packaged SARIF results remained redacted.

## Native CELL and INFO environment-information boundary — 2026-07-26

FormulaFence 0.76.0 was checked against Microsoft's [CELL function
documentation](https://support.microsoft.com/en-us/office/cell-function-51bd39a5-f338-4dbe-a33f-955d67c2b2cf),
which documents the optional reference and selected-cell behavior, and its
[INFO function documentation](https://support.microsoft.com/en-au/office/info-function-725f259a-0e4b-49b3-8b52-58815c69acae),
which lists current operating-environment values. The scope is intentionally
static: FormulaFence inventories native CELL and INFO calls in worksheet
formulas, formula-defined names, and named LAMBDAs, but never evaluates them
or simulates their file, client, workspace, or selected-cell state.

Four fresh, controlled `.xlsx` artifacts were generated without opening Excel.
The baseline (SHA-256
`e6ce60697af78490f8bbb5c4d0347018203eb03468857f6ec9fcfedc2f79cbdc`)
uses direct CELL/INFO formulas, a named LAMBDA, and a formula-defined value.
Its public ledger reports five invoking formula cells, six native calls, two
relevant formula-defined names, and three CELL calls without an explicit
reference. Changing only a private INFO definition (SHA-256
`4a25058ea767207e44322fd49374045ce6fbd12997d58e9ca79de86b69cfb70b`)
kept every public count fixed while emitting FF072 with the private
definition-material flag. Changing only a statically visible shared input
(SHA-256
`23ef5299f3b5ed1a382b23391e3584b07037a77e5605fbe228488351752a7963`)
kept the ledger equal and emitted FF072 with a static-input count of one.

The state-only candidate (SHA-256
`8fd70009b265b7601fe6b320e7c64a63c239885f88d8b05331cd64782b781580`)
adds a worksheet while retaining the stored calls. It produced no FF072 or
FFP072 and passed the narrow policy, proving the implementation does not claim
to infer the selected cell or simulate a file, folder, client, workspace, or
workbook state. The suite separately covers direct calls, explicit versus
omitted CELL references, uninvoked stored names, recursive named LAMBDAs,
sheet-local precedence, native-name shadowing, static inputs, and state-only
changes. The full suite passed with 522 tests.

The dedicated formula_environment_information_calls profile object and FF072 /
FFP072 SARIF results excluded controlled names, values, formulas, arguments,
and the private omitted-reference propagation marker. Ordinary defined-name and
semantic-diff output deliberately retains normal reviewer context, so that
redaction claim is limited to the dedicated ledger and policy-facing results.
No formula or information call was evaluated, and no client/workbook state was
simulated during validation.

The staged formulafence-0.76.0-py3-none-any.whl (SHA-256
`732d9994bd028a79d2a5da5ec5b0e41816c502ea7ec0c2002cd2a48c33ec64ca`)
was built from the release tree and installed into a fresh virtual environment
with its declared dependencies. Its CLI returned FormulaFence 0.76.0; its
generated starter policy retained
`no_formula_environment_information_changes: true`; the definition candidate
exited 1 with FF072 and FFP072; and the state-only candidate exited 0 without
either rule. The packaged profile retained the 5 / 6 / 2 / 3 ledger counts.

## Formula-defined XLM environment-information boundary — 2026-07-26

FormulaFence 0.75.0 was checked against Microsoft's [Excel C API
reference](https://learn.microsoft.com/en-us/office/client-developer/excel/programming-with-the-c-api-in-excel),
which identifies workspace information functions such as GET.CELL and
GET.WORKBOOK. Microsoft's [xlfFree
example](https://learn.microsoft.com/en-us/office/client-developer/excel/xlfree)
uses GET.WORKSPACE to return platform information, and its [expression-evaluation
reference](https://learn.microsoft.com/en-us/office/client-developer/excel/excel-worksheet-and-expression-evaluation)
identifies GET.DOCUMENT as an XLM information function. The scope is
deliberately narrow: FormulaFence records those three calls only when stored in
a formula-defined name or named LAMBDA; direct worksheet formulas and raw XLM
macro-sheet program parts remain separate boundaries.

Four fresh, controlled .xlsx artifacts were generated without opening Excel.
The baseline (SHA-256
`fba8cbcf6d3ca7c5f565e717a55affecd4e3fcb574947a2fe906f2778da5adcf`)
uses a formula-defined value, a named LAMBDA, and one stored call for each of
GET.WORKBOOK, GET.WORKSPACE, and GET.DOCUMENT. Its public ledger reports three
invoking formula cells, three information calls, and three relevant
formula-defined names; it reports no namespaced custom-function candidate or
raw XLM macro-sheet surface. Changing only the private GET.WORKBOOK definition
(SHA-256
`acf35574cb682f3bf1694218119b018481b5f84e917151536a8bbf125e057b27`)
kept every public count fixed while emitting FF071 with the private
definition-material flag. Changing only a shared statically visible input
(SHA-256
`250d0ebcc576d4b8239fe09ae67b615ad0044daf66e1573e8e446ee0a4d1611d`)
kept the environment-information snapshot equal and emitted FF071 with a
static-input count of one.

The state-only candidate (SHA-256
`2f04901502793638955c867143e46d6877caf6547579df3618e8fdda8463df5e`)
adds one worksheet while retaining the same stored calls. It produced no
FF071 or FFP071 and passed the narrow policy, proving the implementation does
not claim to evaluate an information type or simulate workbook state. The
suite separately covers uninvoked stored names, recursive named LAMBDAs,
sheet-local precedence, native-name shadowing, and direct worksheet calls
remaining outside the boundary. The full suite passed with 511 tests. The new
policy caused formulafence check to exit 1 and emit both FF071 and FFP071 for
the definition and static-input candidates.

The dedicated formula_defined_xlm_environment_information_calls profile object
and dedicated FF071/FFP071 SARIF results excluded controlled name identities,
arguments, and input values. Ordinary defined-name and semantic-diff output
deliberately retains normal reviewer context, so that redaction claim is
limited to the dedicated ledger and policy-facing results. No formula or
information call was evaluated, no macro was run, and no workbook, workspace,
document, client, add-in, or printer state was simulated during validation.

The staged formulafence-0.75.0-py3-none-any.whl (SHA-256
`f2229c04b4934d6a063528768eb1f5a6b6c9a65fa0e2c3ca33b32825af1f08ea`)
was installed into a fresh virtual environment with its declared dependencies.
Its CLI returned FormulaFence 0.75.0; the generated starter policy retained
`no_formula_defined_xlm_environment_information_changes: true`, and the
packaged check emitted both `FF071` and `FFP071` for the definition and
static-input candidates. The state-only candidate again passed without either
rule. The packaged profile and dedicated SARIF results remained redacted.

## Formula-defined XLM GET.CELL boundary — 2026-07-26

FormulaFence 0.74.0 was checked against Microsoft's [Excel C API
reference](https://learn.microsoft.com/en-us/office/client-developer/excel/programming-with-the-c-api-in-excel),
which identifies GET.CELL as xlfGetCell and includes it among XLM information
functions. The scope is deliberately narrow: FormulaFence records GET.CELL
only when it is stored in a formula-defined name or named LAMBDA; direct
worksheet formulas and raw XLM macro-sheet program parts remain separate
boundaries.

Three fresh, controlled .xlsx artifacts were generated without opening Excel.
The baseline (SHA-256
`51d8076ef6bc26670fff62df805611dbc4f072e30d1c85166f95cf32e90520d2`)
uses a formula-defined value, a named LAMBDA, and a nested named LAMBDA to
reach GET.CELL. Its public ledger reports three invoking formula cells, three
GET.CELL calls, and three relevant formula-defined names; it reports no
namespaced custom-function candidate or raw XLM macro-sheet surface. Changing
only the private information-call definition (SHA-256
`83b9146f2b6cb1e738cc9c333acb9608b36ccb39835e70eadef419020fa35655`)
kept every public count fixed while emitting FF070 with the private
definition-material flag. Changing only a shared statically visible input
(SHA-256
`95cded492d62256e86d3b1a84c9ba76b8d5ba59af25a606dbcd7055a4e29f262`)
kept the GET.CELL snapshot equal and emitted FF070 with a static-input count
of one.

The suite separately covers uninvoked stored names, recursive named LAMBDAs,
sheet-local precedence, native-name shadowing, and direct worksheet GET.CELL
remaining outside the boundary. The full suite passed with 500 tests. The new
policy caused formulafence check to exit 1 and emit both FF070 and FFP070 for
both controlled candidates. The dedicated formula_defined_xlm_get_cell_calls
profile object and dedicated FF070/FFP070 SARIF results excluded controlled
name identities, arguments, and input values. Ordinary defined-name and
semantic-diff output deliberately retains normal reviewer context, so that
redaction claim is limited to the dedicated ledger and policy-facing results.
No formula or information call was evaluated, no macro was run, and no Excel
display, formatting, comment, or protection state was simulated during
validation.

The staged formulafence-0.74.0-py3-none-any.whl (SHA-256
`7250bcfcb7a5c4c8d58ea383e0706489e04c932a5aa07e85b6b9d9c0fa1091aa`)
was installed into a fresh virtual environment with its declared dependencies.
Its CLI returned FormulaFence 0.74.0; the generated starter policy retained
`no_formula_defined_xlm_get_cell_changes: true`, and the packaged check
emitted both `FF070` and `FFP070` for both controlled candidates. The
packaged profile and dedicated SARIF results remained redacted.

## Formula-defined XLM `EVALUATE` boundary — 2026-07-26

FormulaFence 0.73.0 was checked against Microsoft's [Excel expression-evaluation
reference](https://learn.microsoft.com/en-us/office/client-developer/excel/excel-worksheet-and-expression-evaluation),
which identifies `EVALUATE` as the XLM function that reduces a valid character
string to a worksheet value. The scope is deliberately narrow: FormulaFence
records `EVALUATE` only when it is stored in a formula-defined name or named
`LAMBDA`; direct worksheet formulas and raw XLM macro-sheet program parts
remain separate boundaries.

Three fresh, controlled .xlsx artifacts were generated without opening Excel.
The baseline (SHA-256
`ff66659d6e61343c28e0bc39f9d85594f1d2eb9b61dd200461aa7ba8fbc5323e`)
uses a formula-defined value, a named `LAMBDA`, and a nested named `LAMBDA`
to reach `EVALUATE`. Its public ledger reports three invoking formula cells,
three EVALUATE calls, and three relevant formula-defined names; it reports no
`REGISTER` or raw XLM macro-sheet surface. Changing only the private stored
expression (SHA-256
`ab3493faad49689abfaa6c6416c3744f537186fe71752fc926f0fcf19df6db8b`)
kept every public count fixed while emitting `FF069` with the private
definition-material flag. Changing only a shared statically visible text input
(SHA-256
`b4257c3cf5236ff7d70a3a008644186ff6f658f7c98cc43cbd74db66d7ce7e12`)
kept the evaluation snapshot equal and emitted `FF069` with a static-input
count of one.

The suite separately covers uninvoked stored names, recursive named
`LAMBDA`s, sheet-local precedence, native-name shadowing, direct worksheet
`EVALUATE` remaining outside the boundary, and the explicit limit that
FormulaFence does not re-tokenize formula text parsed at runtime. The full
suite passed with 490 tests. The new policy caused `formulafence check` to
exit `1` and emit both `FF069` and `FFP069` for the definition-only
candidate. The public profile and dedicated `FF069`/`FFP069` SARIF results
excluded controlled name identities and expression values. Ordinary semantic
diffs deliberately retain changed defined-name context for reviewers, so that
redaction claim is limited to the dedicated ledger and policy-facing results.
No formula or text expression was evaluated, no macro was run, and no
runtime-generated expression was parsed during validation.

The staged `formulafence-0.73.0-py3-none-any.whl` (SHA-256
`61d47f405ac7b10a47b3e2cd4ea11892e3d64b9e0bf6dafe38df3706dda5d8ae`)
was installed into a fresh virtual environment with its declared dependencies.
Its CLI returned FormulaFence 0.73.0; the generated starter policy retained
`no_formula_defined_xlm_evaluation_changes: true`, and the packaged check
emitted both `FF069` and `FFP069` for the controlled definition-only
candidate. The packaged profile and dedicated SARIF results remained redacted.

## Formula-defined XLM `REGISTER` boundary — 2026-07-26

FormulaFence 0.72.0 was checked against Microsoft's current
[`xlfRegister` Form 1 reference](https://learn.microsoft.com/en-au/office/client-developer/excel/xlfregister-form-1),
which identifies `REGISTER` as the Excel XLM equivalent for registering DLL
functions or commands and documents macro types callable from a defined-name
definition, and its [`Form 2` reference](https://learn.microsoft.com/en-au/office/client-developer/excel/xlfregister-form-2),
which documents XLL loading and activation. The scope is intentionally narrow:
FormulaFence records `REGISTER` only when it is stored in a formula-defined
name or named `LAMBDA`; direct worksheet formulas and raw XLM macro-sheet
program parts remain separate boundaries.

Three fresh, controlled `.xlsx` artifacts were generated without opening
Excel. The baseline (SHA-256
`b0d0805a8ff3b65337bc9dc4d080378d1ab941e42424012398c8fd8aaca41952`)
uses a formula-defined value, a named `LAMBDA`, and a nested named `LAMBDA` to
reach `REGISTER`. Its public ledger reports three invoking formula cells,
three REGISTER calls, and three relevant formula-defined names; it reports no
worksheet `REGISTER.ID` or raw XLM macro-sheet surface. Changing only the
private type string in the inner definition (SHA-256
`ca0130dbbcdb89db3aa9f4e98acf231d59ff2261f780f6ab09c21df2a1c1d78a`)
kept every public count fixed while emitting `FF068` with the private
definition-material flag. Changing only a shared static input (SHA-256
`66d876650d76ae00aadca604f7a9739cf2605ff649393a8115498b51e0f58a01`)
kept the registration snapshot equal and emitted `FF068` with a static-input
count of one. The suite separately covers uninvoked stored names, recursive
named `LAMBDA`s, sheet-local precedence, native-name shadowing, and direct
worksheet `REGISTER` remaining outside this boundary.

The new policy caused `formulafence check` to exit `1` and emit both `FF068`
and `FFP068` for the definition-only candidate. The public profile and the
dedicated `FF068`/`FFP068` SARIF results excluded controlled name identities,
module values, and procedure values. Ordinary semantic diffs deliberately
retain changed defined-name context for reviewers, so that redaction claim is
limited to the dedicated ledger and policy-facing results. No formula was
evaluated, no macro was run, and no DLL/XLL was resolved or loaded during
validation.

The staged `formulafence-0.72.0-py3-none-any.whl` (SHA-256
`394a0b52b1914b3767fa8693b5998e56371e16cfcf6bf3f97c50bce46ceab8a4`)
was installed into a fresh virtual environment with its declared dependencies.
Its CLI returned `FormulaFence 0.72.0`; the generated starter policy retained
`no_formula_defined_xlm_registration_changes: true`, and the packaged check
emitted both `FF068` and `FFP068` for the controlled definition-only candidate.
The packaged profile and dedicated SARIF results remained redacted.

## Formula-defined external-action propagation — 2026-07-26

FormulaFence 0.71.0 extends the existing `FF064` boundary through
formula-defined names and named `LAMBDA` bodies. The scope remains grounded in
Microsoft's documented formula semantics for
[`HYPERLINK`](https://support.microsoft.com/en-US/Excel/work-with-links-in-excel),
[`WEBSERVICE`](https://support.microsoft.com/en-US/Excel/functions/webservice-function),
[`IMAGE`](https://support.microsoft.com/en-us/excel/functions/image-function),
and [`RTD`](https://support.microsoft.com/en-us/excel/functions/rtd-function):
the workbook can retain an action call even when a worksheet formula reaches it
through a stored name rather than spelling the call directly.

Three fresh controlled `.xlsx` artifacts were generated without opening Excel.
The baseline (SHA-256
`49cc5bd488e79e8e9372ab680c39e982e9f529d462db089c7eacb03821824de4`)
uses a direct formula-defined value, a named `LAMBDA`, and a nested named
`LAMBDA` to reach `HYPERLINK` and `WEBSERVICE`. Its public ledger reports three
formula cells, three relevant formula-defined names, two HYPERLINK calls, and
one WEBSERVICE call. Changing only private content inside the inner named
definition (SHA-256
`87c9a3f46af4499ef8f59bc923055fe4150474c119cc9e1e07ca0f9094274538`)
kept all public counts fixed while emitting `FF064` with the private
name-definition-material flag. Changing only the shared input (SHA-256
`6ff7dfd59a7fe9d4cc4616a431f83b8a5b6034008094dea6ca6a96a11752ee76`)
kept the action snapshot equal and emitted `FF064` with a static-input count of
one. The suite separately covers uninvoked stored names and cycle-safe
recursive named `LAMBDA`s.

The existing `no_formula_external_action_changes` policy emitted `FFP064` for
the named-definition candidate. Profile output and the dedicated `FF064` SARIF
result excluded all test name identities, labels, and endpoint strings. The
ordinary semantic diff intentionally remains separate reviewer context. No
formula was evaluated and no endpoint, image, link, or RTD provider was opened,
fetched, followed, or executed during validation.

The staged `formulafence-0.71.0-py3-none-any.whl` (SHA-256
`f3c66783da0b6517faa4083cb4b7b1e60f8d56fdf7aba9c1d3e02020ec8a324c`)
was installed into a fresh virtual environment with its declared dependencies.
Its CLI returned `FormulaFence 0.71.0`; the generated starter policy retained
`no_formula_external_action_changes: true`, and the packaged check emitted both
`FF064` and `FFP064` for the controlled named-definition candidate. The
packaged profile and dedicated SARIF result remained redacted.

## Worksheet code-resource registration boundary — 2026-07-26

FormulaFence 0.70.0 was checked against Microsoft's current
[`REGISTER.ID` reference](https://support.microsoft.com/en-us/office/register-id-function-f8f0af0f-fd66-4704-a0f2-87b27b175b50).
It documents `REGISTER.ID(module_text, procedure, [type_text])`, states that
the function registers a DLL or code resource when it has not already been
registered, and explicitly distinguishes worksheet-capable `REGISTER.ID` from
`REGISTER`. Microsoft's [`CALL` reference](https://support.microsoft.com/en-us/office/call-function-32d58445-e646-4ffd-8d5e-b45077a5e995)
places `CALL` on macro sheets, supporting the release's intentionally narrow
scope: inspect stored worksheet and formula-defined `REGISTER.ID` expressions while leaving raw
macro-sheet program material to the existing XLM scanner.

Six fresh, controlled `.xlsx` artifacts were generated without opening Excel.
The direct baseline (SHA-256
`3ece94fc58f6f3b345179229a463762eafcda57d7814e5ea7d26eeb42b5717b4`)
reported three registration formula cells, three `REGISTER.ID` calls, and zero
relevant formula-defined names. Its same-count call-material variant (SHA-256
`246d09e7b2e68df0a5858ad88620fd5063dee96c11d05dd548a39d7d5624275d`)
emitted `worksheet_code_resource_registrations_changed` with `FF067` and the
private formula-material flag. Changing only a static input (SHA-256
`0f2517337c837391ea79ccd1ac0f119d399b0a88af94ccec46179608029f692d`)
emitted `FF067` with a static-input count of one.

The named baseline (SHA-256
`e5a95f696e9fb5de6996f05b027bba7e389e625ab6508e62f55ecbd7fd82f72e`)
used a formula-defined value, a named `LAMBDA`, and a nested named `LAMBDA` to
reach the stored registration. It reported three formula cells, three calls,
and three relevant formula-defined names. Changing only the inner definition
(SHA-256 `3ebae9afec68e956cf1a9bc7acb20e4a312a758e29e3425f61f924c465a0322d`)
emitted the private named-definition flag without changing its callers;
changing only the shared input (SHA-256
`40e6f743b0f86744e6a85023433b783c6e8e67c77049d3deea796afb9bc6bd89`)
emitted a static-input count of one. The suite additionally verifies local
name precedence and cycle-safe recursive named `LAMBDA` propagation.

The new policy caused `formulafence check` to exit `1` and emit `FFP067` for
the direct same-count change. The controlled module/procedure values were
absent from the profile and dedicated `FF067` SARIF result. Ordinary semantic
diffs deliberately preserve normal reviewer context, so that privacy assertion
is limited to the dedicated ledger and policy-facing result. No formula was
evaluated, no DLL/XLL was loaded, and no trust configuration was inspected.

The staged wheel `formulafence-0.70.0-py3-none-any.whl` (SHA-256
`fb8e12eff0bb2727c479c89f925cbd005e15f2236ad64ec2d335adacdd45b9db`) was
built locally. The exact wheel was installed into a fresh virtual environment
with declared dependencies. Its CLI returned `FormulaFence 0.70.0`, its starter
policy included `no_worksheet_code_resource_registration_changes: true`, and
the controlled check exited `1` with both `FF067` and `FFP067`; the profile and
dedicated SARIF result remained redacted.

## Namespaced Office custom-function boundary — 2026-07-26

FormulaFence 0.69.0 was checked against Microsoft's [custom-functions
overview](https://learn.microsoft.com/en-us/office/dev/add-ins/excel/custom-functions-overview),
which documents JavaScript/TypeScript functions and manifest-configured
namespaces, its [custom-functions tutorial](https://learn.microsoft.com/en-us/office/dev/add-ins/tutorials/excel-tutorial-create-custom-functions),
which uses `=CONTOSO.ADD(10,200)`, and its [web-data guidance](https://learn.microsoft.com/en-us/office/dev/add-ins/excel/custom-functions-web-reqs),
which documents fetch, HTTP, WebSockets, and streaming functions. Those
references establish both the formula shape and the boundary: the workbook can
store a call while the manifest, code, and runtime live elsewhere.

An independently maintained [OfficeDev Excel Custom Functions sample at commit
`6143f4e326f45c186a891cc923e6e2cdb2136298`](https://github.com/OfficeDev/Excel-Custom-Functions/tree/6143f4e326f45c186a891cc923e6e2cdb2136298)
was examined through its immutable manifest and function source. Its manifest
uses `CustomFunctionsRuntime`, declares a `Namespace`, and points to separate
script/metadata/page URLs; the namespace resource is `CONTOSO`. The immutable
manifest blob is `128cad2cf3d262384a5dedbd3d3e4caf19d232fa` and its function
source blob is `5d327dd94481a298c70d9bbf47d0d0cd05b1460a`. No sample code,
manifest, URL, or runtime was placed in FormulaFence's test artifact or loaded
by FormulaFence.

The dotted-native exclusion list was also compared with the dotted entries in
Microsoft's current [alphabetical Excel function catalog](https://support.microsoft.com/en-US/Excel/excel-functions-alphabetical).
The implementation additionally excludes `ECMA.CEILING`, which Microsoft
exposes through the [Excel JavaScript function API](https://learn.microsoft.com/en-us/javascript/api/excel/excel.functions?view=excel-js-preview)
even though it is not listed in that alphabetical worksheet catalog.

A controlled `.xlsx` pair used the documented namespaced call shape alongside
native dotted formulas and a workbook-defined dotted `LAMBDA`. The baseline
(SHA-256 `b36f463ff29b031395728bc71b51a8547bd8c639bce2472a3cd27bc679b42e4d`)
reported four candidate formula cells, five candidate calls, and two namespaces.
Native `ECMA.CEILING` and `WORKDAY.INTL` calls plus the defined-name
`LOCAL.RATE` call were excluded by the direct-call classifier. Changing only
one same-count candidate callable produced
SHA-256 `65062d406ee8e1dd4bb5c41b7a76919025631ed4bee8e9ab2284dc05fabac8e6`
and emitted `office_custom_functions_changed` with `FF066` and the private
material flag. Changing only a static input used by unchanged candidates
produced SHA-256
`9006dc4a21961b5b1bb9012697baff7089f55a5286a8c6c4b10087abac3c96b8` and
emitted `FF066` with a static-input count of one.

A second controlled workbook used a formula-defined value plus a named
`LAMBDA` wrapper and a nested named `LAMBDA`, each reaching a stored namespaced
call. Its baseline (SHA-256
`1fa3d678b772ea1e9ed6d6661d3310a84fc8b809399ca61cbcafda59615dc308`)
reported three candidate formula cells, five candidate calls, and one namespace:
repeated callable tokens remained counted, the nested wrapper propagated to its
caller, and the formula-defined value remained an ordinary static dependency.
Rewriting only the inner named `LAMBDA` (SHA-256
`4d9c0040c4ed357046fb1d544dc16631979bccb8ad0a4acf503c98ac2bd617b2`)
emitted the private material flag without changing its callers; changing the
shared input (SHA-256
`4d052ebaa472b8d3ca4ca5d776f6838ba84801ec050338bf7da7e8b2cd136516`)
emitted the static-input flag. The suite also verifies worksheet-local name
precedence and treats a recursive named `LAMBDA` as a visible dependency
coverage gap while recording its one directly stored candidate without an
unbounded expansion.

The staged wheel
`formulafence-0.69.0-py3-none-any.whl` (SHA-256
`e23977a1e7b989de0dbe32d2ff4a106db3fbe0d13afa9af99bd3e38fb99eb353`) was
installed in a fresh virtual environment with its declared dependencies. The
installed CLI returned `FormulaFence 0.69.0`; its starter policy exited `1`
with `FFP066` for both candidates. The profile and the dedicated `FF066` SARIF
result were checked for the controlled namespace, function names, and private
input/query values and contained none. The ordinary semantic diff intentionally
retains changed formulas and formula-defined names for reviewer context, so the
privacy assertion is limited to the candidate ledger and policy-facing result.
No formula was evaluated, no
manifest or add-in was loaded, and no network request or custom-function
runtime was contacted during validation.

## Python in Excel code boundary — 2026-07-26

FormulaFence 0.68.0 was checked against Microsoft's [Python in Excel
introduction](https://support.microsoft.com/en-US/Excel/python/introduction-to-python-in-excel),
which documents its Microsoft Cloud runtime, and the OOXML [Python
part](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/151e4bcd-90a0-4d82-8b98-f16bf273e4ff)
and [Python function](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/28f8f9b3-e370-440c-9afc-2a1bba74cde6)
definitions. Those references establish the split boundary: a `PY` formula
refers to a separately stored Python script, while the code itself is not a
normal worksheet-cell formula.

An independently maintained [`xl/python.xml` template at OfficeLib commit
`7354a320505326be8ef66e47293a8696ecd57eb1`](https://github.com/Gryneon/OfficeLib/blob/7354a320505326be8ef66e47293a8696ecd57eb1/Template/xl/python.xml)
(SHA-256 `e0e9ad398d696fec0924be61207744a9b672c064951d456ba6c0bb0685b16a75`)
was placed in a controlled `.xlsx` package with its documented content-type
and workbook relationship. The public profile reported one Python part, one
environment definition, one initialization, zero PY formula cells, zero
scripts, and zero Python coverage gaps. The profile did not expose any source
text from the public template.

A separate controlled pair with two stored PY formulas, two stored scripts,
and one environment initialization established code/binding behavior. The
baseline SHA-256 was
`feeebaa7a08858d2e07a5f24fd73a859c9cdd0f64098c6d39a52a3c2f1c29722`; changing
only stored script text produced candidate SHA-256
`f7bf2780244794a38795905be7383c924cf88899120402eac19221d8f44de66f`.
Public counts stayed fixed while FormulaFence emitted exactly
`python_in_excel_changed` with `FF065` and its private-definition flag.
Changing only a static source cell used by an otherwise unchanged
`=_xlfn._xlws.PY(0,0,A9)` formula produced SHA-256
`76a473f6e40640ec36e884c01b2c5b11c121708242f87986d624150751578b28` and
emitted `FF065` with a static-input count of one.

A fresh virtual environment installed the staged
`formulafence-0.68.0-py3-none-any.whl`; the installed CLI returned
`FormulaFence 0.68.0`. Its starter policy exited `1` with `FFP065` for both
the code-only and static-input candidates. Profiles, JSON reports, and policy
findings were checked for the private source sentinels and contained none. The
ordinary semantic diff still intentionally retains changed PY formulas and
ordinary values; the privacy assertion is limited to the Python ledger and
finding details. No Python source was loaded as Python, no formula was
evaluated, and no Microsoft Cloud runtime was contacted during validation.

## Formula external-action ledger — 2026-07-26

FormulaFence 0.67.0 was checked against Microsoft's documented formula
semantics for [`HYPERLINK`](https://support.microsoft.com/en-US/Excel/work-with-links-in-excel),
[`WEBSERVICE`](https://support.microsoft.com/en-US/Excel/functions/webservice-function),
[`IMAGE`](https://support.microsoft.com/en-us/excel/functions/image-function),
and [`RTD`](https://support.microsoft.com/en-us/excel/functions/rtd-function).
Those references establish the review boundary: link destinations can be text
or cell references, `WEBSERVICE` calls a URL, `IMAGE` uses an HTTPS image
source, and `RTD` requests a COM-automation provider.

A controlled `.xlsx` pair was then built outside this repository from that
documented syntax using a clean openpyxl 3.1.5 environment. The baseline
(SHA-256 `8a407fd18ea30ddb80a128dd670bfe08a9ca7a9257e205b9f2ba749a6d085ddf`)
contains `HYPERLINK`, `WEBSERVICE`, `IMAGE`, namespaced `_xlfn.IMAGE`, `RTD`,
and `HYPERLINK(A9, ...)`. Its public ledger reports six formula cells, two
HYPERLINK calls, one WEBSERVICE call, two IMAGE calls, and one RTD call.

Changing only a literal HYPERLINK destination produced a candidate with SHA-256
`b9ac5b8c215ca61f85133b2d0bfe61cb6439ed179365435f19d8829e62fdf1d6`.
The public counts stayed fixed but FormulaFence emitted exactly
`formula_external_actions_changed` with `FF064` and its private-material flag.
Changing only the A9 value used by the unchanged `HYPERLINK(A9, ...)` formula
produced SHA-256
`ee88b14fda1d16e3a8182396cbdf1486c185a490f4a2752bab54d05c1f817ec2`.
The action snapshot remained equal, while the static dependency graph emitted
the same change and `FF064` with a static-input count of one. A policy enabling
`no_formula_external_action_changes` exited `1` with `FFP064`.

A fresh virtual environment installed the staged
`formulafence-0.67.0-py3-none-any.whl` after archive-integrity checks (SHA-256
`a445a2db2899665eac2a58559883120e37206d15cce3bb3b079ee06a53ba7510`); the
source archive passed the same integrity check.
The installed CLI returned `FormulaFence 0.67.0`. Profiles and FF064 details
were verified not to contain any test URL, image location, provider name,
server name, or referenced-source value. The normal full semantic diff
continues to include changed formulas by design, so the privacy assertion is
limited to the ledger and finding details. No formula was calculated and no
endpoint, image, link, or RTD provider was opened, fetched, followed, or
executed during validation.

## Package-wide external relationship ledger — 2026-07-26

FormulaFence 0.66.0 was checked against the Open Packaging Conventions
[relationship model](https://learn.microsoft.com/en-us/previous-versions/windows/desktop/opc/open-packaging-conventions-overview)
and [`TargetMode` semantics](https://learn.microsoft.com/en-us/dotnet/api/system.io.packaging.packagerelationship.targetmode),
then against two independently maintained Open XML SDK assets at immutable
commit
[`cd2b359ef824737edb93f1c6157c19551aae1e52`](https://github.com/dotnet/Open-XML-SDK/tree/cd2b359ef824737edb93f1c6157c19551aae1e52).

The SDK's strict Excel [Hyperlinks on OArt
Objects](https://github.com/dotnet/Open-XML-SDK/blob/cd2b359ef824737edb93f1c6157c19551aae1e52/test/DocumentFormat.OpenXml.Tests.Assets/assets/TestDataStorage/O14ISOStrict/Excel/Hyperlinks%20on%20OArt%20Objects-O12-XL-Hyperlinks.xlsx)
fixture (SHA-256
`4ddd4a4b3fbc58215be3e95723ce19d3bcd575a27ae26cb6f6f12ad166ffe4eb`)
contains two external relationship targets across two relationship parts. The
ledger reported two parts, two sources, two targets, two hyperlink targets,
zero image/other targets, and zero ledger coverage gaps. The source file also
has unrelated legacy-reader coverage warnings; the validation confirms that
those do not turn its valid external relationships into an unrecognized ledger
entry.

The SDK's [ExternalLink.xlsx](https://github.com/dotnet/Open-XML-SDK/blob/cd2b359ef824737edb93f1c6157c19551aae1e52/test/DocumentFormat.OpenXml.Tests.Assets/assets/TestDataStorage/v2FxTestFiles/spreadsheet/ExternalLink.xlsx)
fixture (SHA-256
`805c1f771ba788a99c51c4652675df8bbc28a816c2174735c98efff42476f7da`)
contains one separately modeled `externalLink` package and one package-level
external relationship. FormulaFence reported one relationship part, one source,
one other target, zero hyperlink/image targets, zero ledger coverage gaps, and
no parser warning. This confirms that the general ledger complements rather
than replaces the focused `FF025` external-link boundary.

For both public fixtures, every `TargetMode="External"` target extracted from
the raw relationship XML was checked absent from the generated JSON profile.
Controlled raw-OOXML fixtures then add an unbound opaque workbook relationship,
an unbound worksheet hyperlink, and an unbound worksheet image relationship.
Changing only the opaque target leaves all public counts fixed but emits exactly
`external_relationships_changed` with high-severity `FF063`; the
`no_external_relationship_changes` policy adds `FFP063`. Coordinated
relationship-ID rewrites stay quiet. Unknown attributes and deliberately
lowered byte limits remain fail-closed coverage evidence, and targets, source
paths, relationship types, identifiers, unknown attribute values, and raw XML
were checked absent from JSON, Markdown, SARIF, and policy output.

A clean virtual environment installed the staged
`formulafence-0.66.0-py3-none-any.whl` after archive-integrity checks (SHA-256
`8eedb3986784aa4034a027c0bb99069840a636a8a4cc44a163134a66d1eb727c`).
The installed CLI returned `FormulaFence 0.66.0`, reproduced both public
relationship inventories, and redacted every public raw external target. On
the controlled target-only pair it emitted only
`external_relationships_changed` / `FF063`; its generated starter policy made
the check exit `1` with `FFP063`.

## Office 2016+ ChartEx worksheet charts — 2026-07-26

FormulaFence 0.63.0 was checked against Microsoft's [ChartEx part
definition](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-odrawxml/5d0d453e-adac-43be-a797-59b9916593dd)
and [ChartEx relationship-ID
type](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-odrawxml/d8ede39e-a36c-48ad-8a17-0086a2d0889b),
then against the independently maintained LibreOffice core fixture at commit
[`a85bae573eeb9e1548176760c0bdb01509ec7c42`](https://github.com/LibreOffice/core/tree/a85bae573eeb9e1548176760c0bdb01509ec7c42):
[`sunburst.xlsx`](https://github.com/LibreOffice/core/blob/a85bae573eeb9e1548176760c0bdb01509ec7c42/chart2/qa/extras/data/xlsx/sunburst.xlsx)
(SHA-256
`8baf66751b6afa7b28f7e9fbc35ea5d9f7ce39d33c8840d9a67cf21a92dadfa4`).
Its worksheet drawing contains Excel's `mc:AlternateContent` form: a
`cx:chart` ChartEx graphic-frame choice and an older-client fallback shape. The
choice resolves through the Office 2014 `chartEx` relationship to
`xl/charts/chartEx1.xml`, which in turn has direct style and colour-style
parts. FormulaFence reported one host sheet, one drawing part, one ChartEx
reference/part, one series, one title, two data references, two fingerprinted
direct parts, and no parser warning; the fallback was not counted as a second
worksheet shape.

The same public commit's `waterfall.xlsx`, `treemap.xlsx`, `boxWhisker.xlsx`,
and `regionMap.xlsx` were also profiled. Each produced one clean ChartEx part;
their series/data-reference counts respectively demonstrated Waterfall (1/1),
Treemap (1/2), Box & Whisker (3/3), and Region Map (1/2) structures with the
same two direct style/colour payloads. This independently checks that the
scanner does not assume one visual ChartEx kind.

A controlled copy of `sunburst.xlsx` changed only private title text in
`xl/charts/chartEx1.xml`. Both archives passed `unzip -t`, retained the same
member set, and differed in uncompressed bytes only for that member. The source
and candidate SHA-256 values were respectively
`8baf66751b6afa7b28f7e9fbc35ea5d9f7ce39d33c8840d9a67cf21a92dadfa4` and
`2cf6fc0701388fb2e3ad460186e80ba7f1c7e4d22b05a68fb4af8e32e0c4dcca`.
The CLI emitted exactly one high-severity `chart_definitions_changed` change
with `FF030` and the safe `chart_definition_material_changed` detail. The
starter policy's `no_chart_definition_changes: true` rule exited `1` and added
`FFP030`. The
controlled suite separately changes a direct ChartEx style payload, rekeys the
drawing relationship without a finding, and fails closed for a malformed
ChartEx root, external chart binding, and unsupported direct relationship.
Private formulas, titles, fallback text, relationship targets, XML, and payload
bytes were checked absent from JSON, Markdown, SARIF, and policy output.

A clean Python virtual environment installed the staged 0.63.0 wheel
(SHA-256 `b27be099787164a1f7e28258a04453fc54687a9aaf18f695cfa3fbbbcd941d3b`)
after both distribution archives passed their archive-integrity checks. The
installed CLI returned `FormulaFence 0.63.0`, profiled one ChartEx part, one
series, two data references, two fingerprinted direct parts, and zero
unrecognized parts on the public Sunburst fixture. It emitted `FF030` for the
controlled ChartEx change, and its starter-policy check exited `1` with
`FFP030`.

This validates bounded stored-package comparison, relationship normalization,
and data minimisation. It does not calculate formulas, render a chart, assess
ChartEx visual semantics, follow external targets or second-hop relationships,
or parse direct media and embedded-package formats.

## SmartArt Diagram Data image payloads — 2026-07-26

FormulaFence 0.64.0 was checked against the OOXML
[Diagram Data Part](https://ooxml.info/docs/14/14.2/14.2.4/) rule and
Microsoft's [`DiagramDataPart.ImageParts` API](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.packaging.diagramdatapart.imageparts?view=openxml-2.8.1).
As an independent package-level check, the Open XML SDK project's immutable
[`SmartArt1.docx`](https://github.com/dotnet/Open-XML-SDK/blob/cd2b359ef824737edb93f1c6157c19551aae1e52/test/DocumentFormat.OpenXml.Tests.Assets/assets/TestDataStorage/v2FxTestFiles/wordprocessing/smart%20art/SmartArt1.docx)
fixture at commit
[`cd2b359ef824737edb93f1c6157c19551aae1e52`](https://github.com/dotnet/Open-XML-SDK/tree/cd2b359ef824737edb93f1c6157c19551aae1e52)
(SHA-256
`d67c50ee4f19528c2336f6d2cd4a6892792a15de78548f2d05ae85859ae193a9`)
contains two Diagram Data relationship parts, each with three direct internal
Image targets. Its Diagram Data XML uses `a:blip r:embed` bindings. It is a
WordprocessingML package, so it is not passed to FormulaFence's SpreadsheetML
reader; it independently verifies the package grammar used by the bounded
SpreadsheetML boundary.

The independently maintained
[`JanMarvin/openxlsx-data`](https://github.com/JanMarvin/openxlsx-data) SmartArt
fixture at immutable commit
[`b89fc5dec8cc9f03b8026a87cbdffe4f5b785207`](https://github.com/JanMarvin/openxlsx-data/tree/b89fc5dec8cc9f03b8026a87cbdffe4f5b785207),
[`diagram.xlsx`](https://github.com/JanMarvin/openxlsx-data/blob/b89fc5dec8cc9f03b8026a87cbdffe4f5b785207/diagram.xlsx)
(SHA-256
`cd708e028cb575f3b6821e721fde620bf15a307efda9893bbca0b96ec7a3b515`),
was profiled again. It retained its two worksheet DrawingML parts and three
SmartArt diagrams, with zero Diagram Data images and no parser warning. This
checks that the image path does not broaden or perturb ordinary SmartArt
workbooks.

A controlled SpreadsheetML package added one `a:blip r:embed` relationship
from a Diagram Data part to a direct internal PNG target. Both archives passed
`unzip -t`, retained the same member set, and differed in uncompressed bytes
only for that image payload. The baseline and candidate SHA-256 values were
respectively
`411c2fbe5d60a96786cf2123b7b4f65fbb093563ef96fb1f060f9b4ea2f4cffc` and
`1421949c84dc5c1ed811d996e6c4ffece71fa05c167138263df0d3b53c7d11cc`.
The baseline profile reported one SmartArt Diagram Data image, one
fingerprinted image, zero uninspected images, six related relationships, and
no parser warning. The ordinary cell and sheet inventories stayed equal while
the CLI emitted exactly one high-severity
`worksheet_drawing_shape_controls_changed` change with `FF044` and the safe
`worksheet_drawing_diagram_material_changed` detail. Profile, JSON, Markdown,
SARIF, and policy output were checked to ensure the image bytes, filename,
target, and relationship identifier remained absent.

A clean Python virtual environment installed the staged 0.64.0 wheel
(SHA-256 `f736125207ad789ca8b3c90e035282a0734fe47f0c40c30bf3182fd67a3d8eda`)
after wheel and source-distribution archive-integrity checks.
The installed CLI returned `FormulaFence 0.64.0`, reproduced the public
fixture's two drawing parts / three SmartArt diagrams / zero Diagram Data image
counts with no warning, and reproduced the controlled fixture's one
fingerprinted image with no uninspected image. It emitted `FF044` for the
byte-only image change, and the
`no_worksheet_drawing_shape_changes` policy exited `1` with `FFP044`.

The suite separately validates transitional and Strict relationship forms,
coordinated image relationship-ID rewrites, missing and external targets,
oversized parts, and byte/count budgets. FormulaFence hashes stored bytes only:
it does not decode or render an image, retrieve a target, or follow a hyperlink,
second-hop target, or relationship from another SmartArt component kind.

## Worksheet DrawingML SmartArt graphic frames — 2026-07-26

FormulaFence 0.62.0 was checked against Microsoft's documented
[Graphic Object Data](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oe376/f58e82a5-5590-4e36-b178-e12989960415)
and [Diagram relationship IDs](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.diagrams.relationshipids?view=openxml-3.0.1)
models, plus the independently maintained
[`JanMarvin/openxlsx-data`](https://github.com/JanMarvin/openxlsx-data) fixture
at commit
[`b89fc5dec8cc9f03b8026a87cbdffe4f5b785207`](https://github.com/JanMarvin/openxlsx-data/tree/b89fc5dec8cc9f03b8026a87cbdffe4f5b785207):
[`diagram.xlsx`](https://github.com/JanMarvin/openxlsx-data/blob/b89fc5dec8cc9f03b8026a87cbdffe4f5b785207/diagram.xlsx)
(SHA-256
`cd708e028cb575f3b6821e721fde620bf15a307efda9893bbca0b96ec7a3b515`).
The workbook contains two worksheet DrawingML parts and three non-chart
`xdr:graphicFrame` diagrams. FormulaFence profiled two participating
worksheets, two drawing parts, three anchors, three graphic frames / SmartArt
diagrams, and three each of the data, layout, quick-style, colour, and
`diagramDrawing` component parts, with no parser warning.

A controlled copy changed only one private text-bearing node in
`xl/diagrams/data1.xml`. Both archives passed `unzip -t`, had the same member
set, and differed in uncompressed bytes only for that diagram-data member. The
baseline and candidate SHA-256 values were respectively
`cd708e028cb575f3b6821e721fde620bf15a307efda9893bbca0b96ec7a3b515` and
`780c6515b626579dce966e76500df14a04739178a57b5e0c94b8ed43207eacf3`.
The ordinary cell and sheet inventories stayed equal while the CLI emitted
exactly one `worksheet_drawing_shape_controls_changed` change with high-
severity `FF044` and the safe detail
`worksheet_drawing_diagram_material_changed`. A policy enabling
`no_worksheet_drawing_shape_changes` exited `1` and added `FFP044`. Profile,
JSON, Markdown, SARIF, and policy-report output were checked to ensure the
private candidate text stayed absent.

A clean Python virtual environment installed the staged 0.62.0 wheel
(SHA-256 `ca0bfb7b1b4ab1e005035b6ac7ab9a5ad55031517929bd4eebe0f423c254019d`)
and returned `FormulaFence 0.62.0`. Its installed CLI profiled the public
fixture with the same three SmartArt frames and three parts of every supported
component kind, emitted exactly `worksheet_drawing_shape_controls_changed`
with `FF044`, and its policy check exited `1` with `FFP044`. The installed
profile, diff, and policy-report JSON were checked again for private candidate
text redaction.

The suite separately validates controlled data changes, transitional and Strict
DrawingML, coordinated non-visual and relationship-ID rewrites, malformed
diagram bindings, unknown non-chart graphic frames, redaction, and policy
enforcement. FormulaFence compares bounded stored declarations only: it does
not render SmartArt, calculate its final layout or visibility, resolve themes,
or follow component-side relationships to media or hyperlinks.

## Legacy shared-workbook revision history — 2026-07-26

FormulaFence 0.61.0 was checked against Microsoft's documented
[`headers`](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.headers?view=openxml-3.0.1),
[`header`](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.header?view=openxml-3.0.1),
and
[`revisions`](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.revisions?view=openxml-3.0.1)
SpreadsheetML models, plus the independently maintained
[`fil` revision-fixture generator](https://gea.i80.dk/hjess/fil/src/commit/bd5cbc68279b71839c566888020aac896ee01cef/tools/generate_test_fixtures/src/generate_test_fixtures/xlsx_revisions.py)
at commit `bd5cbc68279b71839c566888020aac896ee01cef`. Its generated
`xlsx_revisions_basic.xlsx` (SHA-256
`2714a53e96e653a3dfe6906824e6b8bce7291d8781b0ce42f4393ffd5321cfad`)
contains a workbook-bound `revisionHeaders` part with three historic headers
and no log relationship. FormulaFence reported one header part, three headers,
zero log parts/entries, and no shared-workbook revision coverage warning. This
also verifies that a header-only package is not incorrectly treated as a broken
log binding.

A controlled raw-OOXML pair used one revision header and one related log with
three record classes, shared/tracking/history-retention/protection controls,
and private historic cell values, author identity, date, GUID, and relationship
IDs. The candidate changed only a private historic old value in
`xl/revisions/revisionLog1.xml`; both ZIP archives passed `unzip -t`, and it
was the only member with different uncompressed bytes. The baseline and
candidate SHA-256 values were respectively
`5dcaca2c513819c970a88df0e457ee17a2636c14361a19de5e81ccd3f7ed1a8f` and
`591f422cf8885b57b73d99021ef193650a99a36446bf5775ea6bb2a94696bf60`.
The ordinary cell/sheet snapshots stayed equal while the report emitted exactly
one `shared_workbook_revisions_changed` change with high-severity `FF062`.

The suite separately validates tracking/retention/protection control changes,
Boolean and integer spelling plus coordinated relationship-ID normalization,
Strict SpreadsheetML and Strict relationship types, malformed/private unknown
revision metadata, redaction across JSON/Markdown/SARIF, and `FFP062` policy
enforcement. FormulaFence compares bounded stored declarations only: it does
not apply revisions, reconstruct a historic workbook state, resolve conflicts,
validate identity/timestamp claims, render Excel, or interpret arbitrary future
extensions.

A clean Python virtual environment installed the staged 0.61.0 wheel
(SHA-256 `01877e215bceaadd094b7f656a094a64f7ece0813dd9c37460e45232128db7a8`)
and returned `FormulaFence 0.61.0`. Its installed CLI profiled the controlled
baseline, emitted exactly `shared_workbook_revisions_changed` with `FF062`, and
its generated starter policy exited `1` with `FFP062`. JSON profile, diff, and
policy-report output were checked to confirm that the private historic values,
author identity, relationship ID, and log material remained absent.

## Excel Table Style controls — 2026-07-26

FormulaFence 0.60.0 was checked against the SpreadsheetML `TableStyles`,
`TableStyleInfo`, and `TableColumn` definitions and two independently maintained
LibreOffice regression workbooks at commit
[`a85bae573eeb9e1548176760c0bdb01509ec7c42`](https://github.com/LibreOffice/core/tree/a85bae573eeb9e1548176760c0bdb01509ec7c42):
[`Book1_custom.xlsx`](https://github.com/LibreOffice/core/blob/a85bae573eeb9e1548176760c0bdb01509ec7c42/sc/qa/unit/data/xlsx/Book1_custom.xlsx)
(SHA-256 `04fab5c43604fc59a084c240960bc16b35b23357ced56047944559ab44273747`)
and
[`tableStyleInnerBorders.xlsx`](https://github.com/LibreOffice/core/blob/a85bae573eeb9e1548176760c0bdb01509ec7c42/sc/qa/unit/data/xlsx/tableStyleInnerBorders.xlsx)
(SHA-256 `69cc43ed1761bf5e17ca93d5baea2192a75ba8afa7343e538682b6322dca1d46`).
Each contains one applied custom Table Style with seven applicable style
elements and a writer-generated `xr9:uid`; both profiles reported one style
binding, one styled table, one custom style, seven custom style elements, and
no Table Style coverage warning. `tableStyleInnerBorders.xlsx` also reported
one direct TableColumn Dxf assignment, confirming that table-local formatting
is not collapsed into the ordinary cell-style boundary.

A controlled raw-OOXML pair used an applied private custom Table Style backed
by three Dxf records, plus a direct TableColumn Dxf and named-cell-style
reference. The candidate changed only the TableColumn Dxf binding; both
archives had the same member set and only `xl/tables/table1.xml` differed in
uncompressed bytes. The baseline and candidate SHA-256 values were respectively
`81eabf8e5664eba1f5d70dbad042af676c3287047c292ee8474ae7a764dc1abf` and
`735ed1ca07c877defc43433339d590da9f23ddea1b915ae02baed0f1b0922086`.
The ordinary table inventory stayed equal, while the report emitted exactly one
`table_style_controls_changed` change with `FF061`; private table/style names
stayed absent from JSON, Markdown, and SARIF. The suite separately validates
custom Dxf-definition changes, Table Style toggles, Dxf reordering with
coordinated ID rewrites, `xr9:uid` changes, Strict SpreadsheetML raw parts,
malformed references with reader isolation, redaction, and `FFP061` policy
enforcement. FormulaFence compares declarations only: it does not render Excel's
final Table appearance, resolve themes, calculate values, apply conditional
formatting, or cover PivotTable-only style regions.

A clean Python virtual environment installed the staged 0.60.0 wheel with its
declared dependencies and returned `FormulaFence 0.60.0`. Its installed CLI
profiled the controlled baseline, emitted the high-severity
`table_style_controls_changed` change and `FF061`, and its generated starter
policy exited `1` with `FFP061`. The JSON profile and both reports were checked
to confirm that the private custom-style names, named-cell-style name, and Dxf
colour values remained absent.

## Legacy Excel Custom Views — 2026-07-26

FormulaFence 0.59.0 was checked against Microsoft's documented
[`customWorkbookView`](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.customworkbookview?view=openxml-3.0.1)
and
[`customSheetView`](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.customsheetview?view=openxml-3.0.1)
SpreadsheetML declarations, including the workbook `sheetId` /
`activeSheetId` linkage. A controlled raw-OOXML pair used two workbook Custom
Views and four GUID-linked worksheet Custom Views across two sheets. The
candidate changed only a private Custom View filter criterion: both archives
had the same member set and only `xl/worksheets/sheet1.xml` differed in
uncompressed bytes, leaving ordinary cells and formulas untouched. The
baseline and candidate SHA-256 values were respectively
`21c0ba8dc0a3eed867d26d40b206c56562cec523e27666e7d539c6bbae24dbbd` and
`51ed7dc5bcf34648ec9c874708a0206cf37dd3600591d9a0bee662fcaa790943`.

A clean Python virtual environment installed the staged 0.59.0 wheel
(SHA-256 `e1a1615c481b734ea08bebd940caab92973173252e83183e9950368cd05836c3`)
and returned `FormulaFence 0.59.0`. It profiled two workbook views, four
per-sheet views, two participating sheets, and, among per-sheet views, one
hidden-row/column view, one filtered view, two print-setting views, one
display-setting view, and no
coverage gap. The candidate emitted exactly one
`custom_workbook_views_changed` change with `FF060`; the installed starter
policy exited `1` and added `FFP060`. JSON profile and report output was
checked to ensure the Custom View names and private filter values stayed
absent.

The suite separately validates transitional and Strict worksheet views, real
chart-sheet alternate print state, a schema-valid empty chart-sheet container,
coordinated GUID/sheet-ID and writer-spelling normalization, malformed print
metadata, incomplete GUID binding, temporary-reader isolation, redaction, and
policy enforcement. FormulaFence compares stored declarations only: it does
not activate or render a view, calculate filtered results or final pagination,
interpret future extensions, or support Custom Views on unsupported sheet
types.

## Worksheet DrawingML connector controls — 2026-07-26

FormulaFence 0.58.0 was validated against an independently maintained,
Excel-style DrawingML part from
[`galirage/spreadsheet-intelligence`](https://github.com/galirage/spreadsheet-intelligence)
at commit
[`1762ec7b30714e43a85fa451cc97ed0b3e334dc3`](https://github.com/galirage/spreadsheet-intelligence/tree/1762ec7b30714e43a85fa451cc97ed0b3e334dc3):
[`tests/test_data/drawing1.xml`](https://github.com/galirage/spreadsheet-intelligence/blob/1762ec7b30714e43a85fa451cc97ed0b3e334dc3/tests/test_data/drawing1.xml).
A disposable workbook outside this repository used the source part unchanged
behind a normal worksheet drawing relationship. The profile found 24 regular
shapes, seven free `xdr:cxnSp` connectors, zero connector attachments, and no
DrawingML-shape coverage warning. This validates raw-part compatibility; it
does not claim that the upstream project distributes the enclosing workbook.

A separate controlled raw-OOXML pair exercised attached connectors without
changing cells: one endpoint was rebound while the connector's anchor,
geometry, and all cell material stayed fixed. A clean virtual environment
installed the staged 0.58.0 wheel (SHA-256
`2b9f932bed642db8ca88bed6f98971c0e730392cca70558216dd983c3ba0e666`),
returned `FormulaFence 0.58.0`, and emitted exactly one
`worksheet_drawing_shape_controls_changed` change with `FF044`. Its generated
starter policy exited `1` and added `FFP044`. JSON, Markdown, and SARIF output
was checked to ensure connector names, descriptions, and non-visual/endpoint
IDs stayed absent.

The suite separately validates connector line-presentation and attachment
changes, free and group-contained connectors, strict DrawingML, coordinated
non-visual/endpoint-ID rewrites, malformed endpoint coverage, and policy
enforcement. FormulaFence compares stored declarations only: it does not render
or route connectors, resolve themes or visibility, fetch external targets, or
cover `xdr:graphicFrame`, SmartArt, or other unsupported drawing objects.

## Native worksheet image controls — 2026-07-26

FormulaFence 0.57.0 was validated against three independently maintained
XlsxWriter examples at commit
[`cf3fe78d3eab5e4c7d825d4451af3a60e2a04011`](https://github.com/jmcnamara/XlsxWriter/tree/cf3fe78d3eab5e4c7d825d4451af3a60e2a04011):
[`images.py`](https://github.com/jmcnamara/XlsxWriter/blob/cf3fe78d3eab5e4c7d825d4451af3a60e2a04011/examples/images.py),
[`background.py`](https://github.com/jmcnamara/XlsxWriter/blob/cf3fe78d3eab5e4c7d825d4451af3a60e2a04011/examples/background.py),
and [`watermark.py`](https://github.com/jmcnamara/XlsxWriter/blob/cf3fe78d3eab5e4c7d825d4451af3a60e2a04011/examples/watermark.py).
Their baseline SHA-256 values were, respectively,
`86a71dac81cf456622f75aed0898f69c7c57bc51f158918081a8c06e60996e80`,
`c392aa76c41e7230c575e3d87b00efbe1f58ebaff44047f0e142981f31f4e8b7`,
and `7c919d1e5309d9dfc15469436b7ad4e356b59da21cef7d5b0d6338dcd673315b`.
The profiles reported three anchored pictures and one shared image part for the
first workbook, one worksheet background/image part for the second, and one
header/footer VML watermark/image part for the third.

A standalone raw-ZIP script outside this repository replaced only each
workbook's direct `xl/media/` member with a valid alternate PNG. It preserved
the member list and all non-media uncompressed member bytes. Candidate SHA-256
values were, respectively,
`d19dcce79da07308ce4d8842e14011a343b85976ecf9ddd1594524ed1a52d096`,
`362a10723336184ace11bfc5b7c206c870e57452dccc6e64a15be682ecc869e0`,
and `e92e1bdea58447212da730af49f4cf65aff74c2014b5056258d5c786ccc21093`.

A clean virtual environment installed the staged 0.57.0 wheel (SHA-256
`7cea4ab4b011319ded2df58a5b0ca63e57ed5cabcbb4e305d28db2109be7ce38`)
with its declared dependencies and returned `FormulaFence 0.57.0`. Each pair
emitted exactly one `worksheet_image_controls_changed` change and `FF059`; an
installed starter policy enabling `no_worksheet_image_changes` exited `1` and
added `FFP059`. JSON,
Markdown, and SARIF profiles/diffs were checked to ensure raw media member
names stayed absent. XlsxWriter's separate
[`embedded_images.py`](https://github.com/jmcnamara/XlsxWriter/blob/cf3fe78d3eab5e4c7d825d4451af3a60e2a04011/examples/embedded_images.py)
workbook remained outside this boundary (`worksheet_images.present == false`)
and inside the existing rich-data boundary (`rich_data.present == true`).

The suite separately validates anchored-picture presentation and payload
changes, direct backgrounds, VML header/footer watermarks, strict DrawingML,
external relationships without retrieval, ID normalization, privacy redaction,
policy enforcement, malformed roots, bounded XML reads, and separation from
chart and text-shape controls. FormulaFence compares stored package material;
it does not decode or render images, fetch targets, or calculate final layout
or pagination.

## Material worksheet-dimension controls — 2026-07-26

FormulaFence 0.56.0 was validated against two independently maintained
workbooks. The transitional baseline was generated locally from XlsxWriter's
public [`autofit.py`](https://github.com/jmcnamara/XlsxWriter/blob/cf3fe78d3eab5e4c7d825d4451af3a60e2a04011/examples/autofit.py)
example at commit `cf3fe78d3eab5e4c7d825d4451af3a60e2a04011`; its SHA-256 was
`f44bbb48d36c6f6cf0b08f8219370ea63d44948cb9b751be225538d3c57946f4` and
the candidate's was
`2e6e8c386dc05083c091d49cb52693d5a12e8d0987c1f8970d25d650431710a0`.
The strict baseline was the Open XML SDK's
[`2D Rotation-O12-XL-OartEffects.xlsx`](https://github.com/dotnet/Open-XML-SDK/blob/cd2b359ef824737edb93f1c6157c19551aae1e52/test/DocumentFormat.OpenXml.Tests.Assets/assets/TestDataStorage/O14ISOStrict/Excel/2D%20Rotation-O12-XL-OartEffects.xlsx)
fixture at commit `cd2b359ef824737edb93f1c6157c19551aae1e52`; its SHA-256 was
`0e0017c70a5362ef3c49be3fb82c3e80210cfda0e813413c2b28a5ee141c0ad3`
and the candidate's was
`151a0c2b48777865dbc060fe25827721753f6d656757899e21da712a7fd9defc`.

A standalone raw-ZIP script outside this repository changed a material
worksheet-default dimension in each package without using FormulaFence or a
workbook writer. ZIP-member comparison and archive integrity checks confirmed
that exactly `xl/worksheets/sheet1.xml` changed in each candidate; cells,
formulas, and every other member remained fixed. The XlsxWriter profile exposed
four effective positive-width/AutoFit columns; the strict profile exposed three
recognized Office 2010 baseline-adjustment sheets and no dimension coverage
gap.

A clean Python virtual environment installed the staged 0.56.0 wheel (SHA-256
`7fc133877c4cd095994d28dd1e593cc478ff4e22b1cb9db5f7a02144b334ee83`) and
returned `FormulaFence 0.56.0`. Each pair emitted exactly one
`worksheet_dimension_controls_changed` change and `FF058`; a policy enabling
`no_worksheet_dimension_changes` exited 1 and added `FFP058`. JSON and
Markdown profiles/diffs, SARIF diffs, and JSON policy reports were checked to
ensure changed dimension values, raw worksheet member names, `customHeight`,
`x14ac:dyDescent`, and source cell values were absent.

The suite separately validates default/direct sizing, AutoFit, layered ranges,
strict worksheet namespaces, Office 2010 baseline adjustment and custom-height
side effects, automatic thick-border adjustments, equivalent numeric/Boolean
spelling, inert writer declarations, malformed raw metadata, temporary-reader
isolation, redaction, policy enforcement, and separation from `FF036`
zero/hidden visibility controls. FormulaFence compares stored declarations; it
does not render final layout, calculate AutoFit, wrapped/merged-cell overflow,
or automatic page geometry.

## Effective cell-border controls — 2026-07-26

FormulaFence 0.55.0 was validated against two independently maintained Open
XML SDK fixtures at commit
[`cd2b359ef824737edb93f1c6157c19551aae1e52`](https://github.com/dotnet/Open-XML-SDK/tree/cd2b359ef824737edb93f1c6157c19551aae1e52).
The transitional
[`Styles.xlsx`](https://github.com/dotnet/Open-XML-SDK/blob/cd2b359ef824737edb93f1c6157c19551aae1e52/test/DocumentFormat.OpenXml.Tests.Assets/assets/TestDataStorage/v2FxTestFiles/spreadsheet/Styles.xlsx)
baseline SHA-256 was
`a1ca7e60befe2ca550cd4729d68028de2a96aa163574892ed6a0890595b26468`; its
candidate SHA-256 was
`450f8d9114ad526d0953b251d394e21db1c1ac37d47c04acf32d50572c3a458a`.
The strict-OOXML
[`2D Rotation-O12-XL-OartEffects.xlsx`](https://github.com/dotnet/Open-XML-SDK/blob/cd2b359ef824737edb93f1c6157c19551aae1e52/test/DocumentFormat.OpenXml.Tests.Assets/assets/TestDataStorage/O14ISOStrict/Excel/2D%20Rotation-O12-XL-OartEffects.xlsx)
baseline SHA-256 was
`0e0017c70a5362ef3c49be3fb82c3e80210cfda0e813413c2b28a5ee141c0ad3`; its
candidate SHA-256 was
`5894a56adf616acae79c851e32b93f632e3698c4d8f7c2df376d570e8004895e`.

A standalone raw-ZIP script outside this repository changed one stored border
definition in each fixture without using FormulaFence or a workbook writer.
ZIP-member comparison and archive integrity checks confirmed that exactly
`xl/styles.xml` changed in each candidate; ordinary cells and formulas remained
fixed.

A clean Python virtual environment installed the staged 0.55.0 wheel
(SHA-256 `d358993d00f19b7201e12a9d6dc1d284dbd46c66b05740453002dc16f91b9e33`).
Each pair emitted exactly one `cell_border_controls_changed` change and
`FF057`; a policy enabling `no_cell_border_changes` exited 1 and added
`FFP057`. JSON and Markdown profiles/diffs, SARIF diffs, and JSON policy
reports were checked to ensure changed border colours, raw `styles.xml` member
names, border indexes, and line-style spellings were absent.

The suite separately validates direct-cell, row, and column assignments;
default-XF controls; `xfId`/`applyBorder` inheritance; transitional and strict
worksheet namespaces; ordinary/logical/diagonal sides and outline; equivalent
omitted/`none`, Boolean/colour, diagonal, and empty-outline declarations;
malformed metadata; redaction; policy enforcement; and isolation from ordinary
workbook cells. FormulaFence compares stored declarations only: it does not
resolve theme/palette colours, choose adjacent-cell precedence, render final
styles, apply conditional-format/table/differential-style borders, calculate
print output, or infer Excel client behavior.

## Material worksheet print-layout controls — 2026-07-26

FormulaFence 0.54.0 was validated against two independently maintained Open
XML SDK fixtures at commit
[`cd2b359ef824737edb93f1c6157c19551aae1e52`](https://github.com/dotnet/Open-XML-SDK/tree/cd2b359ef824737edb93f1c6157c19551aae1e52).
The transitional
[`Styles.xlsx`](https://github.com/dotnet/Open-XML-SDK/blob/cd2b359ef824737edb93f1c6157c19551aae1e52/test/DocumentFormat.OpenXml.Tests.Assets/assets/TestDataStorage/v2FxTestFiles/spreadsheet/Styles.xlsx)
baseline SHA-256 was
`a1ca7e60befe2ca550cd4729d68028de2a96aa163574892ed6a0890595b26468`; its
candidate SHA-256 was
`b387281421f701dcc2e25630d3d0aa99d9519559511419477b6c3695b86a808e`.
The strict-OOXML
[`2D Rotation-O12-XL-OartEffects.xlsx`](https://github.com/dotnet/Open-XML-SDK/blob/cd2b359ef824737edb93f1c6157c19551aae1e52/test/DocumentFormat.OpenXml.Tests.Assets/assets/TestDataStorage/O14ISOStrict/Excel/2D%20Rotation-O12-XL-OartEffects.xlsx)
baseline SHA-256 was
`0e0017c70a5362ef3c49be3fb82c3e80210cfda0e813413c2b28a5ee141c0ad3`; its
candidate SHA-256 was
`1c24da193641bee1ef124331e346ac68dd9227e85af823ac29cfdd8e4ab7d37b`.

A standalone raw-ZIP script outside this repository changed only the stored
left page margin in `xl/worksheets/sheet1.xml` for the transitional fixture
and only `xl/worksheets/sheet2.xml` for the strict fixture. ZIP-member
comparison confirmed that exactly that one uncompressed member changed in each
candidate; ordinary cells and formulas remained fixed.

A clean Python virtual environment installed the staged 0.54.0 wheel
(SHA-256 `b23d9bd9d797f005e10f42b7598e4f5d622bbdecff7af405460c27d8adfe1d81`).
Each pair emitted exactly one `worksheet_print_layout_controls_changed` change
and `FF056`; a policy enabling `no_worksheet_print_layout_changes` exited 1
and added `FFP056`. The transitional source has an existing printer-settings
relationship coverage gap, but it was unchanged and did not create a false
`unrecognized_worksheet_print_layout_metadata_changed` detail. JSON and
Markdown profiles/diffs, SARIF diffs, and JSON policy reports were checked to
ensure the changed margin value, raw `pageMargins` identifier, worksheet-member
name, and printer relationship ID were absent.

The suite separately validates transitional and strict namespaces; print areas
and titles; gridlines, headings, centering, margins, page setup/fit-to-page,
headers/footers, and manual breaks; omitted/default, Boolean, integer, decimal,
and semantic no-op normalization; malformed metadata; redaction; policy
enforcement; and isolation from ordinary cells. FormulaFence compares stored
declarations only: it does not render or preview Excel, calculate page geometry
or automatic pagination, resolve printer/client defaults or `devMode`, or cover
custom/legacy sheet-view and extension print controls.

## Material worksheet-display controls — 2026-07-26

FormulaFence 0.53.0 was validated against the independently maintained
[`Styles.xlsx`](https://github.com/dotnet/Open-XML-SDK/blob/cd2b359ef824737edb93f1c6157c19551aae1e52/test/DocumentFormat.OpenXml.Tests.Assets/assets/TestDataStorage/v2FxTestFiles/spreadsheet/Styles.xlsx)
fixture from the Open XML SDK at commit
`cd2b359ef824737edb93f1c6157c19551aae1e52`. The downloaded transitional
baseline SHA-256 was
`a1ca7e60befe2ca550cd4729d68028de2a96aa163574892ed6a0890595b26468`.
A standalone raw-ZIP script outside this repository made a candidate with
identical ordinary cells and formulas, changing only
`xl/worksheets/sheet1.xml` to hide displayed zeroes. The candidate SHA-256 was
`c9f73b774b2f5d3a8438325dc2e8e42885901668176fa640726f056f58e4a548`.

The same proof was repeated on the independently maintained strict-OOXML
[`2D Rotation-O12-XL-OartEffects.xlsx`](https://github.com/dotnet/Open-XML-SDK/blob/cd2b359ef824737edb93f1c6157c19551aae1e52/test/DocumentFormat.OpenXml.Tests.Assets/assets/TestDataStorage/O14ISOStrict/Excel/2D%20Rotation-O12-XL-OartEffects.xlsx)
fixture at that commit. Its baseline and candidate SHA-256 values were
`0e0017c70a5362ef3c49be3fb82c3e80210cfda0e813413c2b28a5ee141c0ad3` and
`4245d8358761c90b89dd1a2eb781458a0a7d7f071776063493dc52bc8545fbca`.
Its raw-ZIP candidate changed only `xl/worksheets/sheet2.xml` to hide zeroes;
ordinary cells and formulas remained identical.

A clean Python virtual environment installed the staged 0.53.0 wheel
(SHA-256 `6e4e4c9b1725e7850da587b5a3297ef600801add22c156fb197066f782780d42`).
For each candidate it emitted exactly one
`worksheet_display_controls_changed` change and `FF055`; a policy enabling
`no_worksheet_display_control_changes` emitted `FFP055`. No worksheet-display
coverage warning was introduced. JSON/Markdown profile and diff plus SARIF
were checked to ensure raw `sheetView` control names, pane/selection targets,
gridline colour IDs, and worksheet-member names were absent.

The suite separately validates transitional and strict namespaces; hidden
zeroes, formula display, gridlines and custom gridline colours, headers,
outline symbols, rulers, page whitespace, direction, non-normal views, and
split/frozen panes; default/Boolean/unsigned-integer/decimal spelling;
navigation and zoom noise; malformed metadata; redaction; policy enforcement;
and isolation from ordinary cells. FormulaFence compares stored declarations
only: it does not render Excel, resolve the effective palette colour, calculate
viewport geometry or final visibility, inspect print settings, or interpret
extension-specific client behavior.

## Effective cell-alignment controls — 2026-07-26

FormulaFence 0.52.0 was validated against the independently maintained
[`Styles.xlsx`](https://github.com/dotnet/Open-XML-SDK/blob/cd2b359ef824737edb93f1c6157c19551aae1e52/test/DocumentFormat.OpenXml.Tests.Assets/assets/TestDataStorage/v2FxTestFiles/spreadsheet/Styles.xlsx)
fixture from the Open XML SDK at commit
`cd2b359ef824737edb93f1c6157c19551aae1e52`. The downloaded baseline
SHA-256 was
`a1ca7e60befe2ca550cd4729d68028de2a96aa163574892ed6a0890595b26468`.
It includes used cell-alignment XFs in a real formatted workbook alongside
unrelated presentation/control metadata.

A standalone raw-ZIP script outside this repository made a candidate with
identical ordinary cells, formulas, and logical package members except
`xl/styles.xml`: it changed one already-used alignment record while preserving
the source cell values and formulas. The candidate SHA-256 was
`aca7f4eff8aff38d8b066b52008442b6b514c796ced2cbd27ec70d1fc007ba36`.

A clean Python virtual environment installed the staged 0.52.0 wheel
(SHA-256 `ce4852b729f4d956bc0fcb3a376e7ef9964301a6660176ae3bf69faef1414d30`).
It emitted exactly one `cell_alignment_controls_changed` change and
`FF054`. The generated starter policy exited 1 with `FF054` and
`FFP054`. JSON and Markdown profile/diff, SARIF diff, and JSON policy
artifacts were checked to ensure the changed alignment values, attribute name,
target cell, and member name were absent.

The suite separately validates direct-cell, row, and column assignments;
default-XF controls; `xfId`/`applyAlignment` inheritance;
equivalent default/Boolean/integer spelling; inert `mergeCell` material;
malformed readable metadata; redaction; policy enforcement; and isolation from
ordinary workbook cells. FormulaFence compares stored effective declarations
only: it does not calculate layout/overflow/visibility, compose final visual
styles, or render Excel.

## Workbook DrawingML Theme controls — 2026-07-26

FormulaFence 0.51.0 was validated against the independently maintained
[`Blank.xlsx`](https://github.com/dotnet/Open-XML-SDK/blob/cd2b359ef824737edb93f1c6157c19551aae1e52/test/DocumentFormat.OpenXml.Tests.Assets/assets/TestDataStorage/v2FxTestFiles/spreadsheet/Blank.xlsx)
fixture from the Open XML SDK at commit
`cd2b359ef824737edb93f1c6157c19551aae1e52`. The baseline SHA-256 was
`7a9a2f9721f763d919eadbe30b7ecc1794bfdcc56e20dc12ba589a4ee8c70886`.
It contains one workbook-bound transitional DrawingML Theme with one colour,
font, and format scheme, and no direct Theme images or coverage warnings.

A standalone raw-ZIP script outside this repository created a candidate with
identical ordinary cells, formulas, and logical package members except
`xl/theme/theme1.xml`: it changed one stored Theme colour control. The
candidate SHA-256 was
`eb8387c473b5f15bc864ad887647a21802432392c25d4b6cd42b38322db485c0`.

A clean Python virtual environment installed the staged 0.51.0 wheel
(SHA-256 `6864555a5c113c7579bb6ef4302eb05920d84e13fe22887bdc7e73b90e051ee1`).
It emitted exactly one `workbook_theme_changed` change and `FF053`. The
starter policy exited 1 with `FF053` and `FFP053`. JSON profile/diff/policy,
Markdown profile/diff, and SARIF diff artifacts were checked to ensure the
before/after colour values, Theme member name, and relationship ID were absent.

The same staged wheel also profiled the independently maintained strict-OOXML
[`2D Rotation-O12-XL-OartEffects.xlsx`](https://github.com/dotnet/Open-XML-SDK/blob/cd2b359ef824737edb93f1c6157c19551aae1e52/test/DocumentFormat.OpenXml.Tests.Assets/assets/TestDataStorage/O14ISOStrict/Excel/2D%20Rotation-O12-XL-OartEffects.xlsx)
fixture at that commit with no Theme coverage warning. A raw-ZIP candidate
whose only logical member change was `xl/theme/theme1.xml` emitted exactly
`FF053`; its baseline and candidate SHA-256 values were
`0e0017c70a5362ef3c49be3fb82c3e80210cfda0e813413c2b28a5ee141c0ad3` and
`5f31ae0a08869f89329c46c6d00181feab455468dd8424e75a535a55208817ea`.

The suite separately validates transitional and strict Theme namespaces,
stored scheme changes, direct image-payload changes, relationship-ID
normalization, malformed metadata, bounded reads, redaction, policy
enforcement, and isolation from ordinary workbook cells. FormulaFence compares
stored package controls only: it does not resolve effective styles, render
cells/charts/drawings, calculate contrast, decode images, fetch targets, or
infer Excel client behavior.

## Custom workbook data stores — 2026-07-26

FormulaFence 0.50.0 was validated against the independently maintained
[`NoExtDataE6.xlsx`](https://github.com/dotnet/Open-XML-SDK/blob/cd2b359ef824737edb93f1c6157c19551aae1e52/test/DocumentFormat.OpenXml.Tests.Assets/assets/TestDataStorage/v2FxTestFiles/spreadsheet/NoExtDataE6.xlsx)
fixture from the Open XML SDK at commit
`cd2b359ef824737edb93f1c6157c19551aae1e52`. The downloaded baseline SHA-256
was `f2a375a46ec133bed66dccc28ca5914800d049aadb1326b19204ebf64ecb6287`.
FormulaFence inspected two real generic Custom XML parts, two Custom XML
property parts with six schema references, and one custom document-properties
part with two properties, with no custom-store coverage warnings.

A standalone raw-ZIP script outside this repository created a candidate with
identical ordinary cells, formulas, and package graph. A ZIP-member comparison
confirmed that exactly one uncompressed member changed:
`customXml/item2.xml`. The candidate SHA-256 was
`e6062a9e302db465b8d3c83b2252eea25876181f6cd419ea9fa747797e8bfde4`.

A clean Python virtual environment installed the staged 0.50.0 wheel
(SHA-256 `13313f994dd1de7b4713f829359b5b62b54eb1dc7484a24879c2ef5cbb1d5543`).
It emitted exactly one `custom_data_store_changed` change and `FF052`. A
policy enabling `no_custom_data_store_changes` emitted `FFP052`. JSON profile,
diff, and policy artifacts were checked to ensure the changed custom XML,
property names and values, identifiers, and relationship material were absent.

The suite separately validates generic Custom XML, Custom XML
property/schema/relationship material, workbook-bound Custom Data properties
and opaque binary payloads, custom document properties, identifier
normalization for relationship IDs and document-property `pid` values,
private storage-identity changes, malformed metadata, bounded reads, redaction,
policy enforcement, and Power Query `DataMashup` isolation. FormulaFence
compares stored package state only: it does not execute an add-in, resolve a
property, fetch a target, or interpret a binary payload.

## Excel rich-data controls — 2026-07-24

FormulaFence 0.49.0 was validated against the independently maintained
[`richData_datatypes.xlsx`](https://github.com/JanMarvin/openxlsx-data/raw/main/richData_datatypes.xlsx)
fixture from the `openxlsx-data` project. The downloaded baseline SHA-256 was
`c3064cba084e0d6d3aa1da246d0f4eb02ce270ea1fd0d39a16e501081659afaf`.
It contains real Rich Value Data, structures, types, arrays, supporting
property bags/structures, styles, rich-value metadata bindings, web-image
references, and external web-image relationships.

A standalone raw-ZIP script outside this repository created a candidate with
the same ordinary cells, formulas, and package graph. A ZIP-member comparison
confirmed that exactly one uncompressed member changed:
`xl/richData/rdrichvalue.xml`. The candidate SHA-256 was
`d9b89963f9ec55c4f0411cc3e95271694f7f28cfe9d6785915fb1a3b3316f3bd`.

A clean Python virtual environment installed the staged 0.49.0 wheel
(SHA-256 `7514c1b58332dd97021aa5cbac2a426d5ed6dd5afefa5cfcf9451867946557f6`).
It profiled one data/structure/type/array/property-bag/style/web-image part,
362 rich values, 10 structures including 6 linked-entity structures, 20
arrays, 4 supporting property bags, 12 metadata bindings and bound cells, 6
web images, and 12 external web-image relationship references—with zero
coverage warnings.

The clean wheel emitted exactly one `rich_data_controls_changed` change and
`FF051`. A policy enabling `no_rich_data_changes` exited 1 with `FF051` and
`FFP051`. JSON profile, diff, and policy artifacts were checked to ensure rich
value material, external endpoints, relationship identifiers, and bound-cell
locations were absent.

The suite separately validates value, metadata-binding, external web-image,
and rich-value-relationship changes; relationship ID/order normalization;
malformed metadata; bounded reads; redaction; policy enforcement; and
isolation from ordinary cells. FormulaFence compares stored rich-data
declarations only: it does not contact providers, refresh values, calculate
formulas, fetch endpoints, validate target content, or infer Excel client
behavior.

## Digital-signature controls — 2026-07-24

FormulaFence 0.48.0 is validated with a controlled `.xlsx` pair built outside
this repository by a standalone OpenPyXL 3.1.5 and raw-ZIP script, independent
of the test helpers. The baseline SHA-256 was
`1e4e186f57203d96afc1be1e34e496e809258b5ab2f846e477d57538cfaf0464`;
the candidate SHA-256 was
`428e2952167faf049dfa2960a3655ba914318b4a5054c2aa4b1a114ac87ace12`.
The package carries an empty OPC signature-origin part, one XMLDSIG envelope
with one signed reference and embedded certificate value, one certificate part
linked from the XML signature, and classic, Agile, and V3 VBA signature
payloads linked from a VBA project. The fixture is structurally shaped for
inspection, not a claim of a cryptographically valid signature; OpenPyXL could
still load the baseline normally.

All ordinary cells and formulas stayed fixed. A ZIP-member comparison confirmed
that the only uncompressed member changed in the candidate was
`package/services/digital-signature/certificate/cert1.cer`, whose private
payload was replaced.

A clean Python 3.13 virtual environment installed from the staged 0.48.0 wheel
(SHA-256 `8e92783f822c68271bea598769ab0096667ca62ce2ab32d56ee0c313a7e1c9cf`).
It emitted exactly one `digital_signature_controls_changed` change and
`FF050`. The profile reported one origin, one XML signature, one signed
reference, one embedded certificate value, one certificate part/relationship,
and three VBA signature payloads/relationships. A policy enabling
`no_digital_signature_changes` exited 1 with `FF050` and `FFP050`. JSON
diff, policy, and profile artifacts were checked to ensure private signature
values, certificate content, relationship identifiers, and signed-reference
URIs were absent.

The suite separately validates package-reference, certificate-payload, classic/
Agile/V3 VBA-payload, and relationship changes; equivalent relationship ID,
target spelling, and XMLDSIG base64 whitespace; malformed/unsafe metadata;
bounded reads; redaction; policy enforcement; and isolation from ordinary
workbook cells and macro payload hashing. FormulaFence inventories signature
envelopes only; it does not validate cryptography, certificate trust, expiry,
revocation, timestamps, signed contents, or VBA-code validity.

## SpreadsheetML XML Maps — 2026-07-24

FormulaFence 0.47.0 is validated with a controlled XML Maps pair built outside
this repository by a standalone OpenPyXL 3.1.5 and raw-ZIP script, independent
of the test helpers. The baseline SHA-256 was
`7bd3e109b69f67bf4e3defed8fc8ca6b8ccaf2b37506fe68c8ddfb161214dcde`;
the candidate SHA-256 was
`966b12e85cdb1b11de0b407d8373951761cb9aede8f6685762ecbd6efeb4f9a7`.
The package carries a real XML Maps part, an XML table-column property, a
single-cell XML table part, and the required workbook/worksheet relationships.
OpenPyXL could load the baseline normally. All ordinary cells, formulas, and
every uncompressed package member except `xl/tables/table1.xml` stayed fixed;
that member changed only the mapped field path.

A clean virtual environment installed from the staged 0.47.0 wheel (SHA-256
`26479d35f233b1c9f911e9e8c6033c3faf550997dfbaea41e4ec0a0d0a0806d1`).
It emitted exactly one `xml_mapping_controls_changed` change and FF049.
The XML-mapping profile reported one map part/schema/map/data binding, one
mapped table binding, and one mapped single-cell binding. A policy enabling
`no_xml_mapping_changes` exited 1 with FF049 and FFP049. JSON diff,
policy, and profile artifacts were checked to ensure schemas, map names,
XPath expressions, target cells, connection identities, and relationship
targets were absent.

The suite separately validates mapped-field and refresh behavior changes,
private relationship rebinding, unsafe relationships, equivalent Boolean and
unsigned-integer spelling, malformed single-cell references, bounded reads,
redaction, policy enforcement, and isolation from ordinary table-definition
changes. FormulaFence compares declarations only; it does not import/export or
validate XML data, open bindings, fetch data, calculate a refresh, or verify
Excel rendering.

## Worksheet sparklines — 2026-07-24

FormulaFence 0.46.0 is validated with a controlled `.xlsx` pair built outside
this repository with XlsxWriter 3.2.9. The workbook contains a real Office 2010
line sparkline with marker, axis, custom min/max, date-axis, and colour
controls. The baseline SHA-256 was
`b84f2bbd16d070ffc440c6880d77ab6259d383768c4b538e60ac8b2c773ae659`;
the candidate SHA-256 was
`798fb57dcfbed545522b0fd6516d7f4ced9a1aeaf701dc7fddf52280874f312d`.
All ordinary cells and every other uncompressed package member stayed fixed.
Comparing ZIP members showed exactly one changed member:
`xl/worksheets/sheet1.xml`, where only the sparkline source changed from one
stored row range to another.

A clean virtual environment using the published 0.45.1 wheel (SHA-256
`34cc951a5ecad227b46f832a36ef172d5932b674eff36a6c7984326536da837a`)
reported zero changes and zero findings. A clean environment using the staged
0.46.0 wheel (SHA-256
`da4a0582b39a15c6cf217652cbbe26b1450aeee35cfd91afd20fe1b0f48092b9`)
emitted exactly one `worksheet_sparkline_controls_changed` change and `FF048`.
A policy enabling `no_worksheet_sparkline_changes` exited `1` with `FF048` and
`FFP048`. JSON reports, policy output, and profiles were checked to ensure the
old/new source formulas, date-axis formula, output cell, and colour value were
absent.

The suite separately validates source-only and presentation-only changes,
equivalent source/destination/Boolean/numeric/colour spelling, declaration
reordering, malformed destinations, bounded XML reads, private redaction, and
reader isolation of the unsupported Sparkline Group extension. The scanner
compares stored declarations only; it does not calculate source values, render
the visual result, resolve names/external sources, or assess visual
accessibility.

## Worksheet cell hyperlinks — 2026-07-24

FormulaFence 0.45.1 is validated with a controlled `.xlsx` pair built outside
this repository from a clean openpyxl 3.1.5 workbook and its ordinary
worksheet-cell hyperlink surface. The baseline SHA-256 was
`48a6c07a70b72195b8efb3d881930dbbbf03b8f36439252c60eefda8130efbf5`; the
candidate SHA-256 was
`6ae5ade1d219a79839a334e7d9786d0839a6b0db4d5bb30f19974f62a98ff44d`.
The friendly cell value, style, ScreenTip, and every other uncompressed package
member stayed fixed. Comparing every ZIP member showed exactly one changed
member: `xl/worksheets/_rels/sheet1.xml.rels`, where only the external
hyperlink target changed.

A clean virtual environment using the published 0.44.0 wheel (SHA-256
`d8fc0abcad15991a09c290239ab62f20f499eb73e94f54ab4fc2788606dd7ff5`)
reported zero changes and zero findings for that pair. A clean environment
using the staged 0.45.1 wheel (SHA-256
`34cc951a5ecad227b46f832a36ef172d5932b674eff36a6c7984326536da837a`)
emitted exactly one `cell_hyperlink_controls_changed` change and `FF047`.
A policy enabling `no_cell_hyperlink_changes` exited `1` with `FF047` and
`FFP047`. JSON reports and profiles were checked to ensure the old and new
targets, ScreenTip, cell reference, and relationship ID were absent.

The suite separately validates standard and Office 2016 revision declarations,
target-, location-, display-override-, and ScreenTip-only changes, harmless
relationship-ID/revision-UID rewrites, unbound relationships, malformed
references, bounded XML reads, and composition with the legacy-Note reader
overlay. The scanner compares stored declarations only; it does not render,
fetch, follow, or test a target, inspect linked content, or infer client,
trust-zone, or redirect behavior.

## Legacy Excel Notes and threaded placeholders — 2026-07-24

FormulaFence 0.44.0 is validated with a controlled .xlsx pair built outside
this repository from a clean openpyxl 3.1.5 workbook and a standard
SpreadsheetML comments/VML Note package. The baseline SHA-256 was
`d173971d92bf1d2db0e31001c1e2141f6207051610f561061150006ccf749d86`; the
candidate SHA-256 was
`422eb34e927d4ce395df25dab1b7e66854fadc5b024a27bce253e14fea1f6d04`.
Ordinary cells, the author record, cell binding, comment property/layout, and
all VML stayed fixed. Comparing uncompressed ZIP members showed exactly one
changed member: `xl/comments/comment1.xml`, where only the Note body changed.

A clean virtual environment using the published 0.43.0 wheel (SHA-256
`0e686551d7a6df9edaa71bc1e44f4a177334bd736ffe8953b970d95be39e7ad5`)
reported zero changes and zero findings for that pair. A clean environment
using the staged 0.44.0 wheel (SHA-256
`d8fc0abcad15991a09c290239ab62f20f499eb73e94f54ab4fc2788606dd7ff5`)
emitted exactly one `legacy_comment_controls_changed` change and `FF046`.
A policy enabling `no_legacy_comment_changes` exited 1 with `FF046` and
`FFP046`. JSON reports were checked to ensure Note text, the author, and the
cell reference were absent.

The suite separately validates a conventional Note text edit, VML
visibility-only change, a threaded-comment Note placeholder, consistent VML/
comment/relationship/placeholder identifier rewrites, malformed comments XML,
bounded XML reads, external comments relationships, and external VML
relationships. External Note relationships are quarantined for the ordinary
reader only after raw inspection, so the candidate remains reviewable and
fail-closed rather than raising a reader error. The scanner compares stored
package declarations; it does not prove Excel rendering, author identity,
notification, reconciliation behavior, or cloud/client state.

## Modern threaded comments — 2026-07-24

FormulaFence 0.43.0 is validated with a controlled `.xlsx` pair built outside
this repository from a clean openpyxl 3.1.5 workbook and standard SpreadsheetML
threaded-comment/person package declarations. The baseline SHA-256 was
`a9b75a159b36a799991a6b80810771a71d2e270a97feb1c706c3e52353af81ee`; the
candidate SHA-256 was
`381106caa1fafdddf97933dd9b0070899d674257b20600ed3706896f71cf2bbf`.
Ordinary worksheet cells, person records, thread/reply structure, timestamps,
and mention bindings stayed fixed. Comparing every ZIP member showed exactly
one changed member: `xl/threadedComments/threadedComment1.xml`, where only a
reply body changed.

A clean virtual environment using the published 0.42.0 wheel (SHA-256
`8a6fef91078ebc9db6c927bdacd1935ba1d4bf153e0c8275b56adbf6bef50f40`)
reported zero changes and zero findings for that standard package pair. A clean
environment using the 0.43.0 release wheel (SHA-256
`0e686551d7a6df9edaa71bc1e44f4a177334bd736ffe8953b970d95be39e7ad5`)
emitted exactly one `threaded_comment_controls_changed` change and `FF045`.
Its generated starter policy exited `1` with `FF045` and `FFP045`. The release
JSON reports were checked to ensure comment text, cell reference, timestamp,
email-like user identity, and raw GUIDs were absent.

The suite also validates a zero-change self-diff, person-definition-only
change, standard `mentionpersonId` / `mentionId` handling, harmless comment/
parent/person/mention/relationship-ID rewrites, an unsafe external relationship,
malformed roots, and a deliberately lowered XML budget. This scanner compares
the stored package declarations only; it does not prove rendered visibility,
legacy comment-placeholder behavior, account resolution, mention notification,
or cloud collaboration behavior.

## Worksheet DrawingML shape controls — 2026-07-24

FormulaFence 0.42.0 is validated with a controlled `.xlsx` pair generated by
XlsxWriter 3.2.9 through its documented
[worksheet text-box surface](https://xlsxwriter.readthedocs.io/working_with_textboxes.html),
outside this repository. The baseline SHA-256 was
`349a68a289fe2db93c5cc17970d173989e118cf69c502a40ea029f4b0c68aaad`; the
candidate SHA-256 was
`e9d6790a3139f430d4c56a242da4e66ff791f416f8c333679f9901d7a8ae659a`.
The workbook cells and text-box text stayed fixed. Comparing every ZIP member
showed exactly one changed member: `xl/drawings/drawing1.xml`, where the
review-warning text's stored run colour changed from black to white.

The public 0.41.0 wheel (SHA-256
`c9f3f37b35e27db9f96a7ca6827bc006cd8fba649461a3a12a1199ef710f3144`)
reported zero changes and zero findings for that pair. A fresh 0.42.0 wheel
(SHA-256 `8a6fef91078ebc9db6c927bdacd1935ba1d4bf153e0c8275b56adbf6bef50f40`)
installed into an otherwise clean virtual environment emitted exactly one
`worksheet_drawing_shape_controls_changed` change and `FF044`. Its generated
starter policy exited `1` with `FF044` and `FFP044`. The release JSON reports
were checked to ensure the text-box text, colours, non-visual name, and
relationship identifiers were absent.

The fixture also validates a zero-change self-diff, a relationship-target-only
mutation, group-shape inventory, harmless relationship/non-visual-ID rewrites,
malformed XML, and a deliberately lowered XML budget. The scanner compares
stored supported `xdr:sp` / `xdr:grpSp` declarations; it does not assert screen
rendering, theme/contrast resolution, macro execution, text-link evaluation,
or media and other non-regular DrawingML-object behavior.

## Rich-text run controls — 2026-07-24

FormulaFence 0.41.0 is validated with controlled raw-OOXML `.xlsx` packages
whose normal cell text remains unchanged while character-level presentation
changes. The fixture has one relationship-backed shared-string item and one
inline string, each split into two rich `<r>` runs. Its safe profile exposes
only one referenced shared item/cell/two runs, one inline cell/two runs, zero
phonetic controls, and zero malformed controls. The suite verifies a zero-change
self-diff, exactly one `FF043` and `FFP043` for either shared or inline
colour-only mutation, and `FF043` when the styled character boundary moves
while the concatenated text remains unchanged.

The suite also verifies that an ordinary text edit in an otherwise unchanged
run-property sequence is reported as the normal cell edit rather than a second
rich-text control finding. Equivalent property ordering, color-case spelling,
and explicit `b val="false"` normalize without a finding. An unsupported
namespaced run-property attribute produces a parser-coverage warning, `FF010`,
and `FF043` rather than a silent omission. Text, fonts, colours, shared-string
indexes, and locations are verified absent from profiles, Markdown, ordinary
reports, and SARIF.

For a package-level compatibility reproduction, a controlled baseline with the
same visible warning text had SHA-256
`cf74ab0a768b98acd7297ff66faf390bc1c27f6d425c00c6e80a16e6152e484c`.
The candidate had SHA-256
`a73b1940eb36357809c951105f3955dbd137192b2f63ea7226d5a81bd48284a5`
and changed only the rich-run RGB value for the warning phrase, from opaque
black to opaque white. A fresh published 0.40.0 wheel reported zero changes
and zero findings for that pair. A fresh 0.41.0 wheel (SHA-256
`c9f3f37b35e27db9f96a7ca6827bc006cd8fba649461a3a12a1199ef710f3144`)
emitted exactly one `rich_text_run_controls_changed` change and `FF043`;
the starter policy exited `1` with `FF043` and `FFP043`. Both release
artifacts were checked to ensure neither contained the warning text, colour
values, or cell coordinate. This boundary checks stored XML only; it does not
assert screen rendering, theme resolution, foreground/background contrast, or
whether Excel will make the phrase visible.

## Stored formula results — 2026-07-24

FormulaFence 0.40.0 is validated with controlled raw-OOXML `.xlsx` packages
whose formula text and visible inputs stay fixed while a saved formula result
changes. The fixture contains numeric, string, Boolean, error, and intentionally
missing results under manual-calculation settings. Its safe profile exposes only
five formula cells, four cached results, one missing result, and result-type
counts. The suite verifies a zero-change self-diff, exactly one `FF042` and
`FFP042` for a cache-only numeric-result mutation, and no `FF042` when a visible
input change reaches the changed caches through the static graph.

Equivalent finite numeric and Boolean serializations are exercised without a
finding. An invalid numeric cache produces a parser-coverage warning, `FF010`,
and `FF042` rather than a silent omission. Result values, error text, result
digests, and formula-cell locations are verified absent from profiles, Markdown,
ordinary reports, and SARIF. The scanner compares saved XML only; these tests do
not assert formula correctness, stale-result provenance, volatile/dynamic/
external recalculation behavior, or Excel rendering.

As an independent package-compatibility reproduction, FormulaFence used
XlsxWriter 3.2.9's public
[`tutorial2.py` source](https://github.com/jmcnamara/XlsxWriter/blob/cf3fe78d3eab5e4c7d825d4451af3a60e2a04011/examples/tutorial2.py)
at commit `cf3fe78d3eab5e4c7d825d4451af3a60e2a04011`, generated locally and
not bundled with this repository. A controlled raw-package baseline with manual
calculation settings had SHA-256
`3b67f44e25555dd3172a441d6fc4c14a921e5c974401360e1e98f3935bf0e09a`.
The candidate changed only the stored result beside the unchanged `SUM(B2:B5)`
formula, from `1450` to `999999`, and had SHA-256
`49cbaf7f6474ebd26dba3d604270fbc0ec2c079b1319707c7c5e90b62f431958`.
Both packages use `calcMode="manual"` with full-recalculation and
calculate-on-save flags disabled. The published 0.39.0 wheel (SHA-256
`bff31fb99a49c0f257156dba35819ca408828ab50635b752d4c4ac16d706c4c3`)
reported zero changes and no findings. A fresh 0.40.0 wheel (SHA-256
`c7286a50775d8157795a2e4954701b1b69d0248f2b5424f5efa4bfb673d41c5e`)
emits exactly one `formula_cached_result_changed` change and `FF042`; the
starter policy exits `1` with `FF042` and `FFP042`. The release report
was also checked to ensure it contains none of the public result values, formula
text, or formula-cell coordinate from the mutation.

## Excel zero-dimension visibility controls — 2026-07-24

FormulaFence 0.39.0 was validated with controlled raw-OOXML `.xlsx` packages
containing a zero-height populated row, a zero-width populated column range with
a later positive-width override, zero worksheet-default row and column
dimensions, and ordinary positive resizes. The safe profile records only
zero-height/zero-width/default-zero counts and visible-row overrides; raw dimensions,
row/column targets, and raw declarations stay private. The suite verifies a
zero-change self-diff, `FF036` and `FFP036` for direct zero dimensions, an
effective all-column zero-width default with a later positive-width reveal,
equivalent zero spellings, and no `FF036` for ordinary positive resizes. As of
0.56.0, those ordinary positive resizes are guarded separately as `FF058`
worksheet-dimension controls rather than being silently out of scope.

Negative, non-finite, and out-of-range dimensions produce an explicit
parser-coverage warning, `FF010`, and `FF036` rather than a silent omission.
Raw dimension values, `customHeight`/`customWidth` flags, and row/column targets are
verified absent from JSON, Markdown, ordinary reports, and SARIF. The scanner
compares the documented zero-sized concealment states only; the tests do not
assert Excel rendering, near-zero display behavior, text overflow, arbitrary
positive layout changes, formula calculation, or print layout.

As an independent package-compatibility check, FormulaFence used XlsxWriter
3.2.9's public
[`tutorial2.py`](https://github.com/jmcnamara/XlsxWriter/blob/cf3fe78d3eab5e4c7d825d4451af3a60e2a04011/examples/tutorial2.py)
example at commit `cf3fe78d3eab5e4c7d825d4451af3a60e2a04011`, generated locally
and not bundled with this repository. The baseline `Expenses02.xlsx` SHA-256
was `272036dfdfc75257483b8a8509827cb677c2bd641c0e1f6059825391d0893225`.
A controlled raw-package mutation added only an effective `width="0"` column
declaration for the money column, without a `hidden` attribute or an ordinary
cell/formula edit; its SHA-256 was
`710cb371c9775064ef2a1a5f9c2c24e8d3f74829fb89f4862b25c92730d4a503`.
The published 0.38.0 wheel (SHA-256
`a07f638f9afd6861cd6b2127b62f572e2285174a6c7e2f7eace82ac8a18a83b0`)
reported `0` changes and no findings. A fresh 0.39.0 wheel emits exactly one
`filter_visibility_controls_changed` change and `FF036` (wheel SHA-256
`bff31fb99a49c0f257156dba35819ca408828ab50635b752d4c4ac16d706c4c3`);
the starter policy exits `1` with `FF036` and `FFP036`.

## Excel cell-fill controls — 2026-07-24

FormulaFence 0.38.0 was validated with controlled raw-OOXML `.xlsx` packages
containing two private solid direct fills, one private gradient direct fill, one
`customFormat=1` row fill, and a two-column raw style default. The safe profile
records three direct-cell, one row, and two effective-column assignments; no
fill colour, pattern type, gradient stop, style ID, or target. The suite
verifies a zero-change self-diff, `FF041` for a private-colour-only fill change,
`FFP041` under `no_cell_fill_changes`, `FF041` for a gradient-direction-only
change, and `FF041` when only the default fill definition changes—without a
cell value or formula change.

Equivalent fill-ID reallocation, valid pattern-child ordering, semantically
inert no-fill/solid-background declarations, explicit versus omitted
`applyFill`, base-XF inheritance, and equivalent split column-style ranges are
exercised without a finding. An out-of-bounds column maximum produces an
explicit parser-coverage warning, `FF010`, and `FF041` rather than a silent
omission. Fill colours, pattern/gradient material, style IDs, and cell/row/
column targets are verified absent from JSON, Markdown, ordinary reports, and
SARIF. The scanner compares fill declarations only; the tests do not assert
Excel's theme-colour resolution, pattern/gradient rendering, text/background
contrast, conditional-format differential styles, table styling, formula
calculation, width/overflow, or arbitrary visual formatting.

As an independent package-compatibility check, FormulaFence profiled
XlsxWriter 3.2.9's public
[`tutorial2.py`](https://github.com/jmcnamara/XlsxWriter/blob/cf3fe78d3eab5e4c7d825d4451af3a60e2a04011/examples/tutorial2.py)
example at commit `cf3fe78d3eab5e4c7d825d4451af3a60e2a04011`, generated locally
and not bundled with this repository. The resulting `Expenses02.xlsx` SHA-256
was `272036dfdfc75257483b8a8509827cb677c2bd641c0e1f6059825391d0893225`.
FormulaFence found no fill assignments and no coverage warning. Changing only
the first money cell's fill to a black solid fill produced `candidate.xlsx`
SHA-256 `648a600bb5f3824420288a0f179355d23dcb0a5e07a57839002aa3834bcb36a9`.
The published 0.37.0 wheel (SHA-256
`9db5e438bb501986b81d3d32d5e80b0996f09dc79e944cd321ebec36c154def0`)
reported `0` changes and no findings, while a fresh 0.38.0 wheel emits exactly
one `fill_controls_changed` change and `FF041` (wheel SHA-256
`a07f638f9afd6861cd6b2127b62f572e2285174a6c7e2f7eace82ac8a18a83b0`);
the starter policy exits `1` with `FF041` and `FFP041`.

## Excel cell-font controls — 2026-07-24

FormulaFence 0.37.0 was validated with controlled raw-OOXML `.xlsx` packages
containing a default font definition, two private direct font assignments
(including a white font), one `customFormat=1` row font, and a two-column raw
style default. The safe profile records one default definition, two direct-cell,
one row, and two effective-column assignments; no font names, colour values,
effects, style IDs, or targets. The suite verifies a zero-change self-diff,
`FF040` for a private-colour-only change, `FFP040` under
`no_cell_font_changes`, and `FF040` when only the default font definition
changes—without a cell value or formula change.

Equivalent font-ID reallocation, font-child ordering, explicit versus omitted
`applyFont`, base-XF inheritance, and equivalent split column-style ranges are
exercised without a finding. An out-of-bounds column maximum produces an
explicit parser-coverage warning, `FF010`, and `FF040` rather than a silent
omission. Font names, colour values, effects, style IDs, and cell/row/column
targets are verified absent from JSON, Markdown, ordinary reports, and SARIF.
The scanner compares font declarations only; the tests do not assert Excel's
theme-colour resolution, rendering, background/fill contrast, width/overflow,
rich-text behavior, table styles, or arbitrary visual formatting.

As an independent package-compatibility check, FormulaFence profiled
XlsxWriter 3.2.9's public
[`tutorial2.py`](https://github.com/jmcnamara/XlsxWriter/blob/cf3fe78d3eab5e4c7d825d4451af3a60e2a04011/examples/tutorial2.py)
example at commit `cf3fe78d3eab5e4c7d825d4451af3a60e2a04011`, generated locally
and not bundled with this repository. The resulting `Expenses02.xlsx` SHA-256
was `272036dfdfc75257483b8a8509827cb677c2bd641c0e1f6059825391d0893225`.
FormulaFence found one default font definition, three direct font assignments,
and no coverage warning. Changing only the first money cell's font colour to
white produced `candidate.xlsx` SHA-256
`b47e2970ef8f5745bd4f111c40fb478be557919981d36c57893c33c2cf942e36`.
The published 0.36.0 wheel (SHA-256
`407f6d4d19ddab87549cf46fdb18f6785bc8aecd464541d6ab4d8941a32f4f4f`)
reported `0` changes and no findings, while a fresh 0.37.0 wheel emits exactly
one `font_controls_changed` change and `FF040` (wheel SHA-256
`9db5e438bb501986b81d3d32d5e80b0996f09dc79e944cd321ebec36c154def0`);
the starter policy exits `1` with `FF040` and `FFP040`.

## Excel number-format controls — 2026-07-24

FormulaFence 0.36.0 was validated with controlled raw-OOXML `.xlsx` packages
containing one built-in direct format, two private direct custom formats
(including `;;;`), one `customFormat=1` row style, and a two-column raw style
default. The safe profile records three direct-cell, one row, and two effective
column assignments; one built-in and five custom assignments; and no raw codes,
style IDs, or targets. The suite verifies a zero-change self-diff, `FF039` for
a private-code-only change, `FFP039` under `no_number_format_changes`, and
`FF039` when only the base `cellXfs[0]` number format changes—without a cell
value or formula change.

Equivalent custom-format ID reallocation, explicit versus omitted
`applyNumberFormat`, base-XF inheritance, and equivalent split column-style
ranges are exercised without a finding. An out-of-bounds column maximum
produces an explicit parser-coverage warning, `FF010`, and `FF039` rather than
a silent omission. Format codes, style IDs, and cell/row/column targets are
verified absent from JSON, Markdown, ordinary reports, and SARIF. The scanner
compares number-format declarations only; the tests do not assert Excel's
locale-specific rendering, width/overflow behavior, format-code validity, or
non-number-format visual styles.

As an independent package-compatibility check, FormulaFence profiled
XlsxWriter 3.2.9's public
[`tutorial2.py`](https://github.com/jmcnamara/XlsxWriter/blob/cf3fe78d3eab5e4c7d825d4451af3a60e2a04011/examples/tutorial2.py)
example at commit `cf3fe78d3eab5e4c7d825d4451af3a60e2a04011`, generated locally
and not bundled with this repository. The resulting `Expenses02.xlsx` SHA-256
was `fa1cdb9fab4b703c04b5d79b55c6a9c348e1391ca8dabef4e7412c2f8370e553`.
FormulaFence found five direct custom-format assignments and no coverage warning.
Changing only the public example's first money cell to `;;;` produced
`candidate.xlsx` SHA-256
`6de8b067a32ee790a2992a5000fbde5498851072fb17976563982775266e014f`;
the published 0.35.0 wheel (SHA-256
`0ee611b8fd3c7fe4cc78d9a0c12a2c307fa06d554914b57c0b1dc2024f5401c7`)
reported `0` changes and no findings, while a fresh 0.36.0 wheel emits exactly
one `number_format_controls_changed` change and `FF039` (wheel SHA-256
`407f6d4d19ddab87549cf46fdb18f6785bc8aecd464541d6ab4d8941a32f4f4f`);
the starter policy exits `1` with `FF039` and `FFP039`.

## Excel column visibility — 2026-07-25

FormulaFence 0.35.0 was validated with controlled raw-OOXML `.xlsx` packages
containing a hidden/outlined column range, a width-only range layered over it,
a later explicit visible/outlined override, and one collapsed-outline marker.
The effective safe profile records three hidden columns, four outlined columns,
and one collapsed column. The suite verifies a zero-change self-diff, `FF036`
and `FFP036` when only the base hidden-column declaration changes, and no
ordinary cell or formula change.

Equivalent Boolean/default and unsigned-integer spellings, plus semantically
equivalent split column ranges, are exercised without a finding. An out-of-bounds
column maximum produces an explicit parser-coverage warning, `FF010`, and
`FF036` rather than a silent omission. Column ranges, raw XML, filter criteria,
selected values, custom sort lists, and row/range references are verified absent
from JSON, Markdown, ordinary reports, and SARIF. The controlled column-only
mutation is invisible to the published 0.34.0 wheel (`0` changes, no findings;
wheel SHA-256
`f5c19456e577f66ae45720b9ee3c43d1cd9a446ed298257437433bb602cf412b`),
whereas a freshly installed 0.35.0 wheel emits exactly one
`filter_visibility_controls_changed` change and `FF036`.

As an independent package-compatibility check, FormulaFence profiled
XlsxWriter 3.2.9's public
[`outline_collapsed.py`](https://github.com/jmcnamara/XlsxWriter/blob/cf3fe78d3eab5e4c7d825d4451af3a60e2a04011/examples/outline_collapsed.py)
example at commit `cf3fe78d3eab5e4c7d825d4451af3a60e2a04011`, generated locally
and not bundled with this repository. The resulting `outline_collapsed.xlsx`
SHA-256 was
`c60737867155dc18d46dc5e960ab8b6129acd511eadd81eb1d1d53a93e378fac`.
FormulaFence found six hidden columns, twelve outlined columns, one collapsed
column, and no visibility-control coverage warning. This validates static
declaration comparison, layered-column normalization, and data minimisation—not
whether Excel renders an outline, recalculates formulas, applies a filter, or
models column width/style or outline-display settings. The boundary follows the
Open XML [`cols`](https://c-rex.net/samples/ooxml/e1/Part4/OOXML_P4_DOCX_cols_topic_ID0E5XR4.html)
and [`col`](https://c-rex.net/samples/ooxml/e1/part4/OOXML_P4_DOCX_col_topic_ID0ELFQ4.html)
definitions.

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
embedded-package parsing, source trust, or nested-chart semantics. ChartEx
coverage is validated separately above; FormulaFence fingerprints its stored
package graph without interpreting its visualization semantics. The fixture
follows the OOXML [chart-part model](https://c-rex.net/samples/ooxml/e1/Part1/OOXML_P1_Fundamentals_Chart_topic_ID0ELZLM.html),
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

## Office Web Add-in worksheet and in-content controls — 2026-07-26

FormulaFence 0.65.0 extends that bounded boundary to the two Excel-hosted forms
that do not need a task pane. Microsoft's [Worksheet
definition](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/07d607af-5618-4ca2-b683-6a78dc0d9627)
documents the `{F7C9EE02-42E1-4005-9D12-6889AFFD525C}` worksheet extension
containing `x15:webExtensions`; the [CT_WebExtension
definition](https://learn.microsoft.com/en-nz/openspecs/office_standards/ms-xlsx/386851b6-b7b6-42b8-8cf1-d94bab7b0731)
links each local `appRef` to a definition binding. FormulaFence fingerprints
that relationship and its private range formula, but emits only worksheet and
binding counts. A controlled raw-OOXML fixture changed only the local formula:
ordinary cells were unchanged, `FF028` carried
`worksheet_binding_material_changed`, and the appRef and both formulas were
absent from JSON, Markdown, SARIF, and policy output. Missing appRefs and a
deliberately lowered 16 MiB XML-part cap remained explicit coverage warnings.
The temporary ordinary-reader copy removes the already-recorded supported
extension, preventing `openpyxl`'s lossy unsupported-extension warning without
weakening raw-package coverage.

For in-content frames, the immutable Open XML SDK fixtures
[`Bing.xlsx`](https://github.com/dotnet/Open-XML-SDK/blob/cd2b359ef824737edb93f1c6157c19551aae1e52/test/DocumentFormat.OpenXml.Tests.Assets/assets/TestDataStorage/O15Conformance/XL/WebExtension/Bing.xlsx)
(SHA-256 `a00b0b906a3b1300e1e0892ddbaafa3d8804e23b29d8e14adcb2b5a383eb64f9`)
and
[`Youtube.xlsx`](https://github.com/dotnet/Open-XML-SDK/blob/cd2b359ef824737edb93f1c6157c19551aae1e52/test/DocumentFormat.OpenXml.Tests.Assets/assets/TestDataStorage/O15Conformance/XL/WebExtension/Youtube.xlsx)
(SHA-256 `f794e20899fed3af919e52a5cf1352019f27c6828b54c53c1ae0d37f89210cb8`)
at commit
[`cd2b359ef824737edb93f1c6157c19551aae1e52`](https://github.com/dotnet/Open-XML-SDK/tree/cd2b359ef824737edb93f1c6157c19551aae1e52)
each contain a worksheet DrawingML `mc:AlternateContent` choice with an active
`a:graphicData` Web Extension frame, a direct `webextension1.xml` relationship,
and a static native-picture fallback. FormulaFence profiled each with one
definition part, one in-content drawing part, one in-content reference, one
in-content definition, zero unrecognized add-in parts, and no parser warning.
It skipped the inactive fallback as a duplicate shape but retained it as one
bounded native worksheet image, so static preview-byte changes remain covered
by `FF059` rather than disappearing from review.

The controlled suite also changes only an active-frame anchor (emitting `FF028`
with `in_content_drawing_binding_changed`), proves stable frame/relationship-ID
renumbering produces no finding, and fails closed for a missing frame
relationship ID. Task-pane and definition XML reads retain their 16 MiB per
part, 32 MiB per-workbook, 64-part limits. Worksheet-binding and in-content
DrawingML reads are separately bounded to 16 MiB per part, 64 MiB per workbook,
and 512 parts. This validates static stored-package comparison, branch
selection, relationship normalization, and data minimisation—not manifest
retrieval, Office.js execution, add-in activation/rendering, source trust, or
unrecognized worksheet extension and graphic-frame forms.

The staged `formulafence-0.65.0-py3-none-any.whl` archive passed `unzip -t`
(SHA-256 `98ba721d995a0e4fc5eae7a366de363955ff975c2a2a49a99e15f6784e517cfe`).
A clean virtual environment installed that wheel, returned `FormulaFence
0.65.0`, and reproduced the one-definition/one-drawing/one-reference/
one-in-content-definition, zero-unrecognized result for both public fixtures.

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
