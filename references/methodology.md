# Close Accelerator — methodology & extension guide

This file explains the *why* behind each check and how to extend them. Read it when a user asks to
tune thresholds, add a new anomaly rule, or adapt the skill to a different chart of accounts.

## Reconciliation review

A balance-sheet reconciliation is "clean" only when GL ties to the supporting subledger and any
difference is explained by legitimate, current reconciling items. Two failure modes matter:

1. **Unexplained variance** — GL minus subledger minus documented reconciling items ≠ 0. Any
   material residual means the account isn't actually reconciled.
2. **Stale reconciling items** — the rec "ties," but only because a large item has been sitting open
   for months. Aged items (>60/90 days) are where write-offs, missed cutoffs, and errors hide, so
   aging is weighted as heavily as the dollar variance.

Tune `--recon_unexplained` and `--recon_item` to the entity's materiality. For a small clinic,
$1,000 unexplained is worth a look; for a consolidated parent, raise it.

## Flux / variance analysis

The dual-threshold design (`--flux_pct` AND `--flux_amt`) is deliberate. A percentage-only filter
screams about a $400 account that tripled; a dollar-only filter misses a 40% jump in a mid-size
expense. Requiring both keeps the list to things a controller would actually ask about. Commentary
is drafted, not asserted — the preparer knows the business reason and edits it in.

Extend by adding a third comparison (prior-year same period) or by rolling up to account level across
entities to spot company-wide trends, not just entity-level ones.

## Journal-entry anomaly detection

These mirror standard audit journal-entry testing (JET):

- **Duplicate payments** — same entity/account/vendor/amount appearing on more than one date. The
  single most common recoverable-dollars finding in AP.
- **Round-dollar top-side entries** — large exact-round amounts (multiples of $10k above a floor)
  are the signature of manual estimates and management overrides, which is exactly where earnings
  management and error concentrate.
- **Off-hours / weekend manual postings** — applied *only* to manual entries (blank vendor). A
  vendor invoice posted Saturday is probably a batch job; a manual $250k top-side accrual posted at
  2:47am on a Sunday by one person is a control red flag. Restricting timing rules to manual entries
  is what keeps the signal-to-noise usable — without it, routine AP buries everything.

Severity escalates to High when a single entry combines round-dollar + manual + off-hours/weekend —
the classic override profile. Add rules here: entries to seldom-used accounts, entries by users
outside an approved-preparer list, or debits/credits to income-statement accounts near period end.
## Benford's Law screen

Naturally occurring financial amounts follow a logarithmic leading-digit distribution (≈30.1% start
with 1, ≈4.6% with 9). Fabricated or manipulated numbers usually don't. Two layers are run:

1. **Population level** — the whole JE debit set vs Benford expectation, with a chi-square-style
   deviation. On small populations this is noisy, so treat it as directional.
2. **Vendor concentration** — a vendor with many invoices dominated by one leading digit (≥70%) is
   the sharper signal: it catches split invoicing to stay under approval limits and fabricated
   round-ish amounts that a population test would wash out.

Benford is a *screen*, never proof. It points a human at where to look.

## Intercompany netting

In a multi-entity / JV structure, every intercompany receivable in one entity must have a matching
payable in another, so company-wide Due From must equal Due To before consolidation. A residual means
a one-sided entry, an FX/timing difference, or a missed elimination — all of which distort the
consolidated statements and the noncontrolling-interest split. The check reports the net plus the
per-entity balances so the break can be traced. If the chart of accounts uses different codes for the
intercompany accounts, change the `1900`/`2400` references in `analyze_close.py`.

## Adapting to real ERP exports

QuickBooks, NetSuite, and Xero all export trial balances and GL detail that map to the three input
schemas with light column renaming. Never edit the client's source numbers to fit — rename headers on
a copy, or adjust the reader. When a column is missing (e.g., no `budget_balance`), the corresponding
check simply reports fewer flags; note that in the summary so the reviewer knows a check was limited.
