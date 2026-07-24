# Example

```bash
python generate_workbooks.py
formulafence diff baseline.xlsx candidate-risky.xlsx --format markdown
formulafence check baseline.xlsx candidate-risky.xlsx --policy formulafence.yml
```

The candidate intentionally replaces `Model!B2`'s formula with a hard-coded
value. The diff traces its effect through `Dashboard!B12`; the policy rejects
the change.
