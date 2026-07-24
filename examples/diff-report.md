# FormulaFence change report

- **Baseline:** `examples/baseline.xlsx`
- **Candidate:** `examples/candidate-risky.xlsx`
- **Changes:** 1
- **Findings:** 1
- **Highest severity:** `high`

## Findings

| Severity | Rule | Location | Finding |
| --- | --- | --- | --- |
| high | `FF001` | `Model!B2` | Formula was replaced with a value. |

## Semantic changes

| Risk | Change | Location | Downstream formulas |
| --- | --- | --- | ---: |
| high | `formula_to_value` | `Model!B2` | 1 |

## Impact samples

- `Model!B2` affects: `Dashboard!B12`

## Dependency paths

- `Model!B2` → `Dashboard!B12`
