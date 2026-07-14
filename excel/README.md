# Close Control Center — Excel edition (zero install)

The whole Close Accelerator review, running **inside Microsoft 365 Excel** as an
[Office Script](https://learn.microsoft.com/en-us/office/dev/scripts/). No installation,
no admin rights, no macros — Office Scripts live in the *Automate* tab that's already in
your work Excel, and your data never leaves your company's Microsoft environment.

## Setup (one time, ~2 minutes)

1. Open Excel (desktop with a Microsoft 365 work account, or Excel on the web).
2. Go to the **Automate** tab → **New Script**.
3. Delete the starter code, paste the entire contents of
   [`close_control_center.ts`](close_control_center.ts), rename the script
   **Close Control Center**, and hit **Save script**.

If there is no Automate tab, your admin has Office Scripts turned off — that's a
one-line request to IT ("enable Office Scripts"), not a software install.

## Each close (~1 minute)

1. In your workbook, make three sheets named exactly `TrialBalance`,
   `JournalEntries`, `Reconciliations`, with these row-1 headers:

   | Sheet | Headers |
   |---|---|
   | TrialBalance | `entity, account, account_name, account_type, current_balance, prior_balance, budget_balance` |
   | JournalEntries | `je_id, entity, date, time, user, account, account_name, description, debit, credit, vendor` |
   | Reconciliations | `entity, account, account_name, gl_balance, subledger_balance, reconciling_items, aging, prepared_by` |

2. Paste your ERP exports under the headers (Workday report exports paste straight in;
   dates work as either real Excel dates or `YYYY-MM-DD` text, times as Excel times or `HH:MM`).
3. **Automate tab → Close Control Center → Run.**

It writes six tabs: **CC Summary** (exception counts), **CC Recon**, **CC Flux**,
**CC JE Flags**, **CC Benford**, **CC Intercompany** — high-severity rows shaded red,
medium amber. Re-running replaces the CC tabs; your input sheets are never modified.

Thresholds (materiality, round-dollar floor, intercompany account codes) are the
`CFG` block at the top of the script — edit and save.

Tip: pair with **Power Query** (Data → Get Data) to pull the Workday export files into
the three input sheets automatically, so each close becomes *refresh → Run → review*.

> Same rules as the rest of this repo: every flag is a draft for preparer review.
> Nothing posts, nothing files, and a qualified accountant signs off before close.
