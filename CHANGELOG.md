# Changelog

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
