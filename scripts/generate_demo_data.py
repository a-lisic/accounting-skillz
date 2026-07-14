#!/usr/bin/env python3
"""
Generate a realistic month-end close dataset for a joint-venture dialysis operator.
Multiple clinic entities + Corporate + Eliminations. Several issues are deliberately
planted so the Close Accelerator skill has something to catch.

Outputs (CSV) into the target directory:
  - trial_balance.csv       entity-level current / prior / budget balances
  - journal_entries.csv     the period's JE register
  - reconciliations.csv     balance-sheet account recs (GL vs subledger)
"""
import csv, os, random, sys

random.seed(74015)  # deterministic

OUT = sys.argv[1] if len(sys.argv) > 1 else "."
os.makedirs(OUT, exist_ok=True)

PERIOD = "2026-06"
PRIOR = "2026-05"

# --- Entities: 8 clinic JVs + Corporate + Eliminations ------------------------
CLINICS = [
    "Riverside Dialysis LLC", "Lakeshore Renal Care LLC", "Summit Kidney Center LLC",
    "Prairie Dialysis Partners LLC", "Gulf Coast Renal LLC", "Cascade Dialysis LLC",
    "Heartland Kidney Care LLC", "Bayou Renal Associates LLC",
]
CORP = "ARC Corporate"
ELIM = "Eliminations"
ENTITIES = CLINICS + [CORP, ELIM]

# --- Chart of accounts --------------------------------------------------------
# (account, name, type) ; type in Asset/Liability/Equity/Revenue/Expense
COA = [
    ("1000", "Cash - Operating", "Asset"),
    ("1010", "Cash - Patient Trust", "Asset"),
    ("1200", "Patient Accounts Receivable", "Asset"),
    ("1210", "Allowance for Contractual Adj", "Asset"),
    ("1400", "Supplies Inventory", "Asset"),
    ("1500", "Prepaid Expenses", "Asset"),
    ("1600", "Property & Equipment", "Asset"),
    ("1610", "Accumulated Depreciation", "Asset"),
    ("1700", "Right-of-Use Asset (ASC 842)", "Asset"),
    ("1900", "Due From Affiliates", "Asset"),
    ("2000", "Accounts Payable", "Liability"),
    ("2100", "Accrued Payroll", "Liability"),
    ("2150", "Accrued Expenses", "Liability"),
    ("2200", "Lease Liability (ASC 842)", "Liability"),
    ("2400", "Due To Affiliates", "Liability"),
    ("3000", "Members' Equity", "Equity"),
    ("3100", "Noncontrolling Interest", "Equity"),
    ("4000", "Patient Service Revenue", "Revenue"),
    ("4100", "Capitation Revenue", "Revenue"),
    ("4900", "Contractual Allowances", "Revenue"),
    ("5000", "Medical Supplies Expense", "Expense"),
    ("5100", "Clinical Payroll", "Expense"),
    ("5200", "Physician Fees", "Expense"),
    ("5300", "Facility Rent", "Expense"),
    ("5400", "Utilities", "Expense"),
    ("5500", "Depreciation Expense", "Expense"),
    ("5600", "Repairs & Maintenance", "Expense"),
    ("5700", "Insurance", "Expense"),
    ("5800", "Administrative Expense", "Expense"),
]

def base_amount(acct_type, account):
    """Reasonable per-clinic monthly magnitudes."""
    scale = {
        "1000": 480_000, "1010": 22_000, "1200": 1_250_000, "1210": -520_000,
        "1400": 165_000, "1500": 78_000, "1600": 3_400_000, "1610": -1_150_000,
        "1700": 2_100_000, "1900": 0, "2000": 240_000, "2100": 195_000,
        "2150": 88_000, "2200": 2_050_000, "2400": 0, "3000": 1_900_000,
        "3100": 640_000, "4000": -2_450_000, "4100": -310_000, "4900": 940_000,
        "5000": 610_000, "5100": 720_000, "5200": 205_000, "5300": 138_000,
        "5400": 41_000, "5500": 96_000, "5600": 28_000, "5700": 33_000,
        "5800": 71_000,
    }
    return scale.get(account, 50_000)

# ------------------------------------------------------------------------------
# 1) TRIAL BALANCE  (current, prior, budget) per entity/account
# ------------------------------------------------------------------------------
tb_rows = []
flux_target = None
for ci, clinic in enumerate(CLINICS):
    for account, name, atype in COA:
        if account in ("1900", "2400"):  # intercompany handled separately below
            continue
        base = base_amount(atype, account)
        jitter = 1 + random.uniform(-0.06, 0.06)
        cur = base * jitter * (1 + ci * 0.015)
        prior = cur * (1 + random.uniform(-0.04, 0.04))
        budget = cur * (1 + random.uniform(-0.05, 0.05))
        tb_rows.append([clinic, account, name, atype, round(cur, 2), round(prior, 2), round(budget, 2)])

# Corporate entity (lighter, mostly overhead + equity)
for account, name, atype in COA:
    if account in ("1900", "2400"):
        continue
    if atype in ("Revenue",):
        continue
    base = base_amount(atype, account) * 0.35
    cur = base * (1 + random.uniform(-0.05, 0.05))
    prior = cur * (1 + random.uniform(-0.03, 0.03))
    budget = cur * (1 + random.uniform(-0.04, 0.04))
    tb_rows.append([CORP, account, name, atype, round(cur, 2), round(prior, 2), round(budget, 2)])

# ---- PLANT #1: material unexplained flux -> Medical Supplies at Gulf Coast +58% MoM
for r in tb_rows:
    if r[0] == "Gulf Coast Renal LLC" and r[1] == "5000":
        r[5] = round(r[4] / 1.58, 2)  # prior far below current => +58% jump (~$220k)
        flux_target = (r[0], r[1])

# ---- Intercompany: Due From / Due To across affiliates.
# These should net to ~zero company-wide. We PLANT #2: an out-of-balance of $312,400.
due_from_total = 0.0
due_to_total = 0.0
inter_rows = []
for ci, clinic in enumerate(CLINICS):
    df = round(random.uniform(120_000, 480_000), 2)   # clinic owed by others
    dt = round(random.uniform(120_000, 480_000), 2)   # clinic owes others
    due_from_total += df
    due_to_total += dt
    inter_rows.append([clinic, "1900", "Due From Affiliates", "Asset", df, round(df*0.98,2), df])
    inter_rows.append([clinic, "2400", "Due To Affiliates", "Liability", dt, round(dt*0.98,2), dt])
# Corporate plug so that WITHOUT the planted error it would net to zero:
corp_df = due_to_total - due_from_total  # makes From==To before planting
# PLANT the imbalance: shrink corporate Due From by 312,400 so it no longer nets
corp_df_planted = corp_df - 312_400.00
inter_rows.append([CORP, "1900", "Due From Affiliates", "Asset", round(corp_df_planted,2), round(corp_df_planted*0.99,2), round(corp_df_planted,2)])
tb_rows.extend(inter_rows)

with open(os.path.join(OUT, "trial_balance.csv"), "w", newline="") as f:
    w = csv.writer(f)
    w.writerow(["entity","account","account_name","account_type","current_balance","prior_balance","budget_balance"])
    w.writerows(tb_rows)

# ------------------------------------------------------------------------------
# 2) JOURNAL ENTRIES  (the period register)
# ------------------------------------------------------------------------------
VENDORS = ["McKesson Medical","Fresenius Supply","Baxter Intl","Gambro Renal","DaVita Labs",
           "City Power & Light","Cintas Facility","Grainger","Iron Mountain","Aramark",
           "Quest Diagnostics","Stericycle","ADP Payroll","CBRE Facilities","Henry Schein"]
USERS = ["j.torres","m.nguyen","s.patel","d.oconnell","controller.review"]
je_rows = []
jeid = 5000
def add_je(entity, date, time, user, account, name, desc, debit, credit, vendor=""):
    global jeid
    jeid += 1
    je_rows.append([f"JE{jeid}", entity, date, time, user, account, name, desc,
                    round(debit,2) if debit else "", round(credit,2) if credit else "", vendor])

# routine entries
for clinic in CLINICS:
    for _ in range(random.randint(9, 14)):
        day = random.randint(1, 28)
        acct, name, atype = random.choice([c for c in COA if c[2] == "Expense"])
        amt = round(random.uniform(3_000, 85_000), 2)
        v = random.choice(VENDORS)
        add_je(clinic, f"{PERIOD}-{day:02d}", f"{random.randint(8,17):02d}:{random.randint(0,59):02d}",
               random.choice(USERS[:4]), acct, name, f"Invoice {v}", amt, 0, v)
        add_je(clinic, f"{PERIOD}-{day:02d}", f"{random.randint(8,17):02d}:{random.randint(0,59):02d}",
               random.choice(USERS[:4]), "2000", "Accounts Payable", f"Invoice {v}", 0, amt, v)

# ---- PLANT #3: duplicate payment (same vendor, same amount, 2 days apart)
add_je("Riverside Dialysis LLC","2026-06-11","10:14","j.torres","5000","Medical Supplies Expense","Invoice McKesson Medical #INV-88421",48750.00,0,"McKesson Medical")
add_je("Riverside Dialysis LLC","2026-06-11","10:14","j.torres","2000","Accounts Payable","Invoice McKesson Medical #INV-88421",0,48750.00,"McKesson Medical")
add_je("Riverside Dialysis LLC","2026-06-13","09:02","j.torres","5000","Medical Supplies Expense","Invoice McKesson Medical #INV-88421",48750.00,0,"McKesson Medical")
add_je("Riverside Dialysis LLC","2026-06-13","09:02","j.torres","2000","Accounts Payable","Invoice McKesson Medical #INV-88421",0,48750.00,"McKesson Medical")

# ---- PLANT #4: suspicious round-dollar top-side accrual, off-hours, unusual user
add_je("ARC Corporate","2026-06-28","02:47","d.oconnell","5800","Administrative Expense","Management accrual - true up",250000.00,0,"")
add_je("ARC Corporate","2026-06-28","02:47","d.oconnell","2150","Accrued Expenses","Management accrual - true up",0,250000.00,"")

# ---- PLANT #5: weekend entry
add_je("Summit Kidney Center LLC","2026-06-14","23:31","m.nguyen","5200","Physician Fees","Physician bonus accrual",90000.00,0,"")  # 6/14/2026 is a Sunday
add_je("Summit Kidney Center LLC","2026-06-14","23:31","m.nguyen","2150","Accrued Expenses","Physician bonus accrual",0,90000.00,"")

# ---- PLANT #6: Benford-violating vendor (Grainger invoices all start with 9)
for amt in [9120.00, 9340.55, 9875.20, 9410.00, 9990.75, 9230.10, 9705.40, 9560.00, 9880.25, 9145.60]:
    add_je("Cascade Dialysis LLC","2026-06-"+f"{random.randint(1,28):02d}",f"{random.randint(8,17):02d}:15",
           "s.patel","5600","Repairs & Maintenance","Invoice Grainger",amt,0,"Grainger")
    add_je("Cascade Dialysis LLC","2026-06-"+f"{random.randint(1,28):02d}",f"{random.randint(8,17):02d}:15",
           "s.patel","2000","Accounts Payable","Invoice Grainger",0,amt,"Grainger")

with open(os.path.join(OUT, "journal_entries.csv"), "w", newline="") as f:
    w = csv.writer(f)
    w.writerow(["je_id","entity","date","time","user","account","account_name","description","debit","credit","vendor"])
    w.writerows(je_rows)

# ------------------------------------------------------------------------------
# 3) RECONCILIATIONS  (balance-sheet account recs: GL vs subledger)
# ------------------------------------------------------------------------------
REC_ACCTS = [("1200","Patient Accounts Receivable"),("2000","Accounts Payable"),
             ("1400","Supplies Inventory"),("1500","Prepaid Expenses"),
             ("2100","Accrued Payroll"),("2200","Lease Liability (ASC 842)")]
rec_rows = []
for clinic in CLINICS:
    for acct, name in REC_ACCTS:
        gl = next((r[4] for r in tb_rows if r[0]==clinic and r[1]==acct), 0) or round(random.uniform(50_000,900_000),2)
        gl = abs(gl)
        # most recs tie out within a small aged reconciling item
        recon_item = round(random.uniform(0, 4_000), 2)
        sub = round(gl - recon_item, 2)
        rec_rows.append([clinic, acct, name, round(gl,2), sub, recon_item, "0-30 days", random.choice(USERS[:4])])

# ---- PLANT #7: large unreconciled AR variance at Prairie ($188,900, aged >90 days)
for r in rec_rows:
    if r[0]=="Prairie Dialysis Partners LLC" and r[1]=="1200":
        r[4] = round(r[3] - 188_900.00, 2)   # subledger far below GL
        r[5] = 188_900.00
        r[6] = ">90 days"

with open(os.path.join(OUT, "reconciliations.csv"), "w", newline="") as f:
    w = csv.writer(f)
    w.writerow(["entity","account","account_name","gl_balance","subledger_balance","reconciling_items","aging","prepared_by"])
    w.writerows(rec_rows)

print(f"Wrote demo data to {OUT}/")
print(f"  trial_balance.csv     {len(tb_rows)} rows")
print(f"  journal_entries.csv   {len(je_rows)} rows")
print(f"  reconciliations.csv   {len(rec_rows)} rows")
print(f"  intercompany Due From total (incl corp): {due_from_total+corp_df_planted:,.2f}")
print(f"  intercompany Due To total:               {due_to_total:,.2f}")
print(f"  planted flux target: {flux_target}")
