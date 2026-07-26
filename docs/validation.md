# External validation notes

FormulaFence's test suite builds small fixtures that isolate individual risks.
Those tests are necessary but insufficient for confidence in an Office-file
reader, so each release should also be exercised on independently maintained
workbooks without copying their contents into this repository.

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
