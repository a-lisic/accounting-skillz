---
name: close-accelerator
description: >
  Runs a month-end / quarter-end corporate close review over a trial balance, journal-entry register,
  and balance-sheet reconciliations, and produces a review-ready Excel workpaper plus a management
  flux memo. Performs reconciliation review, flux/variance analysis (month-over-month and vs budget)
  with drafted commentary, forensic journal-entry testing (duplicate payments, round-dollar top-side
  entries, off-hours/weekend manual postings), a Benford's Law screen, and an intercompany netting
  check across entities. Use whenever the user is closing the books, reviewing a close package, doing
  flux or variance analysis, reviewing reconciliations, hunting journal-entry anomalies or fraud risk,
  consolidating entities, or prepping audit support. Especially strong on multi-entity or JV structures.
  Trigger on: month-end close, quarter-end, flux/variance analysis, reconciliation review, journal
  entry testing, Benford, intercompany, consolidation, or trial balance review.
---

# Close Accelerator

Turn a pile of close inputs into a reviewed, exception-flagged workpaper in one pass. The goal
is to do what a senior accountant does on the first read of a close — tie out the balance sheet,
explain the big swings, and sniff out the entries that don't smell right — but in seconds and
across every entity at once, so the human spends their time on judgment instead of ticking and tying.

**This is a preparer's assistant, not an approver.** Everything it produces is a draft for review.
It never posts entries, never files anything, and every exception it raises must be signed off by a
qualified accountant before close. Say this plainly in outputs — it protects the user and sets the
right expectation.

## When to use it

Reach for this whenever someone is closing or reviewing the books: a month-end or quarter-end close,
a flux/variance review, a reconciliation tie-out, journal-entry testing, an intercompany or
consolidation check, or audit/PBC prep. It shines on multi-entity data where doing this by hand
across every entity is the painful part.

## Inputs

The skill expects three CSVs (real exports work — QuickBooks, NetSuite, Xero, or an ERP dump all
produce these; map columns if headers differ). If the user only has one or two, run what you can and
say which checks were skipped.

`trial_balance.csv` — one row per entity/account:
`entity, account, account_name, account_type, current_balance, prior_balance, budget_balance`
(`account_type` in Asset/Liability/Equity/Revenue/Expense. Intercompany accounts should use
account `1900` = Due From Affiliates and `2400` = Due To Affiliates, or tell the script the codes.)

`journal_entries.csv` — the period's JE register:
`je_id, entity, date, time, user, account, account_name, description, debit, credit, vendor`
(`date` ISO `YYYY-MM-DD`; `time` 24h `HH:MM`; blank `vendor` = manual entry.)

`reconciliations.csv` — balance-sheet account recs:
`entity, account, account_name, gl_balance, subledger_balance, reconciling_items, aging, prepared_by`

If the user's file uses different column names, don't hand-edit their data — either rename columns
to this schema in a copy, or adjust the reading logic. Ask for the mapping if it's ambiguous.

## How to run it

The analysis is deterministic, so it lives in a script rather than being done by hand — this keeps
the numbers reproducible and auditable, which matters for close work.

```bash
python scripts/analyze_close.py <input_dir> <output_dir> --period 2026-06
```

Thresholds are tunable via flags (match them to the entity's materiality):
`--flux_pct 0.30` `--flux_amt 50000` `--recon_unexplained 1000` `--recon_item 25000`
`--round_dollar_min 50000` `--intercompany_tol 1000`.

It writes `close_workpaper.xlsx` and `flux_memo.md` to the output directory. Deliver both, and
lead with the exception count so the user knows how much needs their attention.

If the user wants a demo or has no data yet, generate a realistic multi-entity dataset with
`python scripts/generate_demo_data.py <dir>` (it plants a few issues on purpose) and run the
analyzer on it to show what the output looks like.

## What each check does (and why)

- **Reconciliation review** — flags recs with unexplained GL-vs-subledger variance, large
  reconciling items, or items aged past 60/90 days. A rec that "ties" only because of a big stale
  open item is still a problem; that's why aging matters as much as the variance.
- **Flux analysis** — month-over-month and vs-budget swings that clear both a % and a $ threshold
  (both, so you don't drown in immaterial percentage noise on tiny accounts). Each flag gets drafted
  commentary the preparer can refine — the memo is a starting draft, not the final word.
- **Journal-entry anomalies** — duplicate payments (same vendor/amount across dates), round-dollar
  top-side entries, and manual entries posted off-hours or on weekends. Timing flags are applied
  **only** to manual entries, because routine vendor invoices are often batch-posted off-hours and
  would otherwise bury the real signals.
- **Benford's Law screen** — leading-digit distribution across JE amounts, plus a per-vendor
  concentration test that catches a vendor whose invoices cluster on one leading digit (a classic
  fabricated- or split-invoice tell).
- **Intercompany netting** — Due From vs Due To across all entities must net to ~zero before
  consolidation; any imbalance is surfaced with the per-entity detail so it can be chased down.

See `references/methodology.md` for the forensic rationale, threshold guidance, and how to extend
the checks (new anomaly rules, different intercompany account codes, revenue-specific tests).

## Output structure

The workbook opens on a **Summary** dashboard (exception counts and $ exposure by category), then a
tab per check: Recon Review, Flux Analysis, JE Anomalies, Benford Screen, Intercompany. High-severity
rows are shaded red, medium amber. The **flux_memo.md** is the narrative version for management —
executive summary, material fluxes with commentary, and the exceptions requiring sign-off.

## Guardrails

Do not present results as conclusions. Frame anomalies as "worth a look", not "fraud found" — the
skill surfaces statistical and structural oddities that a human then investigates. Keep the
sign-off language in every deliverable. If the data looks incomplete (missing entities, a TB that
doesn't balance), say so rather than analyzing around it silently.
