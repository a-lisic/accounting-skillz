# Close Accelerator

A Claude [Agent Skill](https://github.com/anthropics/skills) that runs a month-end / quarter-end
**corporate close review** over a trial balance, journal-entry register, and balance-sheet
reconciliations — and produces a review-ready Excel workpaper plus a management flux memo in seconds,
across every entity at once.

It does what a senior accountant does on the first read of a close — tie out the balance sheet,
explain the big swings, and sniff out the entries that don't smell right — but automated and at scale,
so the human spends time on judgment instead of ticking and tying.

> ⚠️ **This is a preparer's assistant, not an approver.** Everything it produces is a *draft for
> review*. It never posts entries and never files anything. Every exception it raises must be signed
> off by a qualified accountant before close. Anomaly flags mean "worth a look," not "fraud found."

---

## What it checks

| Check | What it catches |
|-------|-----------------|
| **Reconciliation review** | Unexplained GL-vs-subledger variance, large reconciling items, and items aged past 60/90 days (a rec that "ties" only via a big stale item is still a problem). |
| **Flux / variance analysis** | Month-over-month and vs-budget swings that clear **both** a % and a $ threshold, each with drafted commentary the preparer refines. |
| **Journal-entry testing** | Duplicate payments, round-dollar top-side entries, and manual entries posted off-hours or on weekends (timing rules apply only to manual entries to keep the signal clean). |
| **Benford's Law screen** | Leading-digit distribution across JE amounts, plus a per-vendor concentration test that catches split / fabricated invoicing. |
| **Intercompany netting** | Due From vs Due To across all entities must net to ~zero before consolidation; any imbalance is surfaced with per-entity detail. |

## Output

- **`close_workpaper.xlsx`** — opens on a Summary dashboard (exception counts + $ exposure by
  category), then a tab per check. High-severity rows shaded red, medium amber.
- **`flux_memo.md`** — the narrative version for management: executive summary, material fluxes with
  commentary, and the exceptions requiring sign-off.

See [`examples/sample_outputs/`](examples/sample_outputs/) for a full generated set.

---

## Quick start

```bash
pip install -r requirements.txt

# Option A — run on your own data (three CSVs, schema below)
python scripts/analyze_close.py <input_dir> <output_dir> --period 2026-06

# Option B — see it work on generated demo data (issues planted on purpose)
python scripts/generate_demo_data.py demo
python scripts/analyze_close.py demo demo --period 2026-06
```

Thresholds are tunable to the entity's materiality:

```bash
python scripts/analyze_close.py in out \
  --flux_pct 0.30 --flux_amt 50000 \
  --recon_unexplained 1000 --recon_item 25000 \
  --round_dollar_min 50000 --intercompany_tol 1000
```

## Input schema

Real ERP exports (QuickBooks, NetSuite, Xero) map to these with light column renaming. Don't edit
source numbers to fit — rename headers on a copy.

**`trial_balance.csv`** — `entity, account, account_name, account_type, current_balance, prior_balance, budget_balance`
(`account_type` ∈ Asset/Liability/Equity/Revenue/Expense; intercompany accounts use `1900` = Due From
Affiliates, `2400` = Due To Affiliates)

**`journal_entries.csv`** — `je_id, entity, date, time, user, account, account_name, description, debit, credit, vendor`
(`date` ISO `YYYY-MM-DD`; `time` 24h `HH:MM`; blank `vendor` = manual entry)

**`reconciliations.csv`** — `entity, account, account_name, gl_balance, subledger_balance, reconciling_items, aging, prepared_by`

If a file or column is missing, the corresponding check simply reports fewer flags — note it in the
summary so the reviewer knows a check was limited.

---

## Install as a Claude skill

This repo *is* a skill. To use it inside Claude (Cowork / Claude Code):

- **Cowork:** package it with the skill-creator's `package_skill.py` and click **Save skill** on the
  resulting `.skill` file, **or**
- **Claude Code:** drop this folder into your skills directory.

The [`SKILL.md`](SKILL.md) frontmatter drives when Claude auto-invokes it (month-end close, flux
analysis, reconciliation review, journal-entry testing, Benford, intercompany, consolidation).

## How it works

The analysis is deterministic and lives in a script rather than being done conversationally — this
keeps the numbers reproducible and auditable, which matters for close work. Claude orchestrates:
maps the user's columns to the schema, runs the analyzer, reads the exceptions back, and drafts the
narrative. See [`references/methodology.md`](references/methodology.md) for the forensic rationale
behind each check, threshold guidance, and how to extend the rules (new anomaly tests, different
intercompany account codes, revenue-specific tests like healthcare contractual allowances).

## Repo layout

```
close-accelerator/
├── SKILL.md                     # skill definition + triggering
├── scripts/
│   ├── analyze_close.py         # the analysis engine
│   └── generate_demo_data.py    # synthetic multi-entity demo data
├── references/
│   └── methodology.md           # forensic rationale + extension guide
├── examples/
│   ├── sample_data/             # demo CSV inputs
│   └── sample_outputs/          # generated workpaper, memo, dashboard
└── requirements.txt
```

## Roadmap ideas

- Live QuickBooks / NetSuite / Xero pull via MCP (skip the CSV export step)
- Scheduled run on the 1st that emails the exception list
- Prior-year comparison in flux analysis
- ASC 842 lease roll-forwards and noncontrolling-interest split checks for JV structures

## License

[MIT](LICENSE). The sample data is synthetic — no real company or patient information.
