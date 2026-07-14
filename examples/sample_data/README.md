# Sample data

`reconciliations.csv` is committed as a sample. The other two inputs are generated
deterministically (seed 74015) so they are not stored in git — create them with:

```bash
python ../../scripts/generate_demo_data.py .
```

That writes `trial_balance.csv`, `journal_entries.csv`, and `reconciliations.csv`
(257 / 220 / 48 rows) with the demo's planted issues included.
