# External validation notes

FormulaFence's test suite builds small fixtures that isolate individual risks.
Those tests are necessary but insufficient for confidence in an Office-file
reader, so each release should also be exercised on independently maintained
workbooks without copying their contents into this repository.

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
equivalent zero spellings, and no `FF036` for ordinary positive resizes.

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
