#!/usr/bin/env python3
"""
Close Accelerator - analyzer engine.

Reads three close inputs (trial_balance.csv, journal_entries.csv, reconciliations.csv),
runs balance-sheet reconciliation review, flux/variance analysis, forensic JE anomaly
detection (duplicate / round-dollar / off-hours / weekend / manual top-side), a Benford's
Law screen, and an intercompany netting check. Writes a formatted multi-tab Excel
workpaper and a management flux memo (Markdown).

Usage:
  python analyze_close.py <input_dir> <output_dir> [--period 2026-06]

Thresholds are CLI-overridable; defaults are sensible for a mid-size operator.
"""
import csv, os, sys, argparse, datetime, math
from collections import defaultdict, Counter

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ----------------------------- config ----------------------------------------
ap = argparse.ArgumentParser()
ap.add_argument("input_dir")
ap.add_argument("output_dir")
ap.add_argument("--period", default="2026-06")
ap.add_argument("--flux_pct", type=float, default=0.30)      # % change to flag
ap.add_argument("--flux_amt", type=float, default=50_000)    # $ change to flag
ap.add_argument("--recon_unexplained", type=float, default=1_000)
ap.add_argument("--recon_item", type=float, default=25_000)
ap.add_argument("--round_dollar_min", type=float, default=50_000)
ap.add_argument("--intercompany_tol", type=float, default=1_000)
args = ap.parse_args()

IN, OUT = args.input_dir, args.output_dir
os.makedirs(OUT, exist_ok=True)

def read_csv(name):
    with open(os.path.join(IN, name)) as f:
        return list(csv.DictReader(f))

tb = read_csv("trial_balance.csv")
je = read_csv("journal_entries.csv")
rec = read_csv("reconciliations.csv")

def fnum(x):
    try: return float(x)
    except (TypeError, ValueError): return 0.0

def _sev_rank(s): return {"OK":0,"Low":1,"Medium":2,"High":3}.get(s,0)

# ----------------------------- 1) RECON REVIEW --------------------------------
recon_findings = []
for r in rec:
    gl, sub, item = fnum(r["gl_balance"]), fnum(r["subledger_balance"]), fnum(r["reconciling_items"])
    variance = gl - sub
    unexplained = variance - item
    aging = r.get("aging","")
    reasons, sev = [], "OK"
    if abs(unexplained) > args.recon_unexplained:
        reasons.append(f"Unexplained variance ${unexplained:,.0f}"); sev = "High"
    if abs(item) > args.recon_item:
        reasons.append(f"Large reconciling item ${item:,.0f}")
        sev = sev if _sev_rank(sev) >= _sev_rank("Medium") else "Medium"
    if aging in (">90 days",">60 days") and abs(item) > 5_000:
        reasons.append(f"Aged reconciling item ({aging})"); sev = "High"
    if reasons:
        recon_findings.append({
            "entity": r["entity"], "account": r["account"], "name": r["account_name"],
            "gl": gl, "sub": sub, "item": item, "unexplained": unexplained,
            "aging": aging, "prepared_by": r.get("prepared_by",""),
            "severity": sev, "issue": "; ".join(reasons)})

# ----------------------------- 2) FLUX ANALYSIS -------------------------------
flux_findings = []
for r in tb:
    cur, prior, budget = fnum(r["current_balance"]), fnum(r["prior_balance"]), fnum(r["budget_balance"])
    mom = cur - prior
    mom_pct = (mom / abs(prior)) if prior else 0.0
    bud_var = cur - budget
    bud_pct = (bud_var / abs(budget)) if budget else 0.0
    flag = (abs(mom_pct) >= args.flux_pct and abs(mom) >= args.flux_amt) or \
           (abs(bud_pct) >= args.flux_pct and abs(bud_var) >= args.flux_amt)
    if flag:
        drivers = []
        if abs(mom_pct) >= args.flux_pct and abs(mom) >= args.flux_amt:
            drivers.append(f"{'up' if mom>0 else 'down'} {mom_pct*100:,.0f}% MoM (${mom:,.0f})")
        if abs(bud_pct) >= args.flux_pct and abs(bud_var) >= args.flux_amt:
            drivers.append(f"{'over' if bud_var>0 else 'under'} budget {bud_pct*100:,.0f}% (${bud_var:,.0f})")
        commentary = f"{r['account_name']} at {r['entity']} is " + " and ".join(drivers) + \
                     ". Obtain supporting detail and management explanation before sign-off."
        flux_findings.append({
            "entity": r["entity"], "account": r["account"], "name": r["account_name"],
            "type": r["account_type"], "current": cur, "prior": prior, "budget": budget,
            "mom": mom, "mom_pct": mom_pct, "bud_var": bud_var, "bud_pct": bud_pct,
            "severity": "High" if (abs(mom_pct)>=0.5 or abs(bud_pct)>=0.5) else "Medium",
            "commentary": commentary})
flux_findings.sort(key=lambda x: -abs(x["mom"]))

# ----------------------------- 3) JE ANOMALIES --------------------------------
def parse_date(d):
    try: return datetime.date.fromisoformat(d)
    except Exception: return None
def parse_hour(t):
    try: return int(t.split(":")[0])
    except Exception: return 12

je_findings = []
# duplicates: same entity/account/vendor/amount, appearing on >1 date (debits only)
dup_groups = defaultdict(list)
for e in je:
    deb = fnum(e["debit"])
    if deb > 0:
        key = (e["entity"], e["account"], e.get("vendor",""), round(deb,2), e.get("description","")[:40])
        dup_groups[key].append(e)
for key, entries in dup_groups.items():
    dates = {x["date"] for x in entries}
    if len(entries) >= 2 and len(dates) >= 2 and key[2]:  # vendor present, >1 date
        for x in entries:
            je_findings.append({"je": x["je_id"], "entity": x["entity"], "date": x["date"],
                "user": x["user"], "amount": fnum(x["debit"]), "account": x["account_name"],
                "vendor": x.get("vendor",""), "severity": "High",
                "issue": f"Possible duplicate payment: {len(entries)} identical charges (${key[3]:,.0f}) to {key[2]} across {len(dates)} dates"})

flagged_je = {f["je"] for f in je_findings}
for e in je:
    amt = max(fnum(e["debit"]), fnum(e["credit"]))
    if amt == 0: continue
    d = parse_date(e["date"]); hr = parse_hour(e.get("time","12:00"))
    reasons, sev = [], "Low"
    is_round = amt >= args.round_dollar_min and abs(amt % 10_000) < 0.005
    is_offhours = hr < 6 or hr >= 20
    is_weekend = d.weekday() >= 5 if d else False
    is_manual = not e.get("vendor","")
    # Round-dollar large entries are worth a look regardless of source.
    if is_round: reasons.append(f"Round-dollar ${amt:,.0f}"); sev = "Medium"
    # Timing anomalies are only meaningful for MANUAL top-side entries -- routine
    # vendor invoices are often batch-posted off-hours/weekends and would be noise.
    if is_manual and is_offhours: reasons.append(f"Posted off-hours ({e.get('time','')})"); sev = max(sev,"Medium",key=_sev_rank)
    if is_manual and is_weekend: reasons.append(f"Posted on a weekend ({e['date']})"); sev = max(sev,"Medium",key=_sev_rank)
    if is_round and is_manual and (is_offhours or is_weekend):
        sev = "High"; reasons.append("manual top-side entry, no vendor support")
    if reasons and e["je_id"] not in flagged_je:
        je_findings.append({"je": e["je_id"], "entity": e["entity"], "date": e["date"],
            "user": e["user"], "amount": amt, "account": e["account_name"],
            "vendor": e.get("vendor",""), "severity": sev, "issue": "; ".join(reasons)})
je_findings.sort(key=lambda x: (-_sev_rank(x["severity"]), -x["amount"]))

# ----------------------------- 4) BENFORD SCREEN ------------------------------
BENFORD = {d: math.log10(1 + 1/d) for d in range(1,10)}
lead = Counter()
amounts = [abs(fnum(e["debit"])) for e in je if fnum(e["debit"]) >= 100]
for a in amounts:
    s = f"{a:.2f}".lstrip("0").lstrip(".")
    for ch in s:
        if ch.isdigit() and ch != "0":
            lead[int(ch)] += 1; break
n = sum(lead.values()) or 1
benford_rows = []
chi = 0.0
for d in range(1,10):
    actual = lead[d]/n
    expected = BENFORD[d]
    exp_ct = expected*n
    chi += ((lead[d]-exp_ct)**2)/exp_ct if exp_ct else 0
    benford_rows.append({"digit": d, "actual_ct": lead[d], "actual_pct": actual,
                         "expected_pct": expected, "deviation": actual-expected})
# vendor-level concentration (>=8 invoices dominated by one leading digit)
vendor_lead = defaultdict(Counter)
for e in je:
    deb = fnum(e["debit"]); v = e.get("vendor","")
    if deb >= 100 and v:
        s = f"{deb:.2f}".lstrip("0").lstrip(".")
        for ch in s:
            if ch.isdigit() and ch != "0":
                vendor_lead[v][int(ch)] += 1; break
benford_vendor_flags = []
for v, c in vendor_lead.items():
    total = sum(c.values())
    if total >= 8:
        top_digit, top_ct = c.most_common(1)[0]
        if top_ct/total >= 0.7:
            benford_vendor_flags.append({"vendor": v, "count": total, "digit": top_digit,
                "share": top_ct/total, "severity": "High",
                "issue": f"{top_ct}/{total} invoices lead with digit {top_digit} ({top_ct/total*100:.0f}%) - possible fabricated/split invoicing"})

# ----------------------------- 5) INTERCOMPANY --------------------------------
due_from = sum(fnum(r["current_balance"]) for r in tb if r["account"]=="1900")
due_to   = sum(fnum(r["current_balance"]) for r in tb if r["account"]=="2400")
ic_net = due_from - due_to
ic_flag = abs(ic_net) > args.intercompany_tol
ic_by_entity = []
for r in tb:
    if r["account"] in ("1900","2400"):
        ic_by_entity.append({"entity": r["entity"], "account": r["account_name"],
                             "balance": fnum(r["current_balance"])})

# ----------------------------- write memo -------------------------------------
def money(x): return f"${x:,.0f}"
memo = []
memo.append(f"# Month-End Close Review Memo — {args.period}\n")
memo.append(f"**Prepared by:** Close Accelerator (automated draft — requires preparer sign-off)  ")
memo.append(f"**Scope:** {len({r['entity'] for r in tb})} entities · {len(je)} journal entries · {len(rec)} account reconciliations\n")
total_exc = len(recon_findings)+len(flux_findings)+len(je_findings)+len(benford_vendor_flags)+(1 if ic_flag else 0)
memo.append(f"## Executive summary\n")
memo.append(f"The automated review identified **{total_exc} exceptions** requiring attention before the books are finalized: "
            f"{len(recon_findings)} reconciliation issue(s), {len(flux_findings)} material flux(es), "
            f"{len([f for f in je_findings if f['severity']=='High'])} high-risk journal entr(ies), "
            f"{len(benford_vendor_flags)} vendor pattern flag(s), and an intercompany imbalance of {money(abs(ic_net))}.\n")

memo.append("## Material fluctuations\n")
if flux_findings:
    for f in flux_findings:
        memo.append(f"- **{f['name']} — {f['entity']}:** {f['commentary']}")
else:
    memo.append("- None above threshold.")
memo.append("")
memo.append("## Reconciliation exceptions\n")
if recon_findings:
    for f in recon_findings:
        memo.append(f"- **{f['name']} — {f['entity']}** ({f['severity']}): {f['issue']}. GL {money(f['gl'])} vs subledger {money(f['sub'])}.")
else:
    memo.append("- All reconciliations tie within tolerance.")
memo.append("")
memo.append("## Journal-entry / fraud-risk flags\n")
for f in je_findings[:12]:
    memo.append(f"- **{f['je']} — {f['entity']}** ({f['severity']}): {f['issue']}. {money(f['amount'])}, user {f['user']}.")
for f in benford_vendor_flags:
    memo.append(f"- **Vendor {f['vendor']}** ({f['severity']}): {f['issue']}.")
memo.append("")
memo.append("## Intercompany\n")
memo.append(f"Company-wide Due From Affiliates ({money(due_from)}) less Due To Affiliates ({money(due_to)}) "
            f"= **{money(ic_net)}**. {'**Out of balance — investigate before consolidating.**' if ic_flag else 'In balance.'}\n")
memo.append("---\n*Every item above is a preparer-level draft. Nothing has been posted. A qualified accountant must review and approve before close.*")
with open(os.path.join(OUT,"flux_memo.md"),"w") as f:
    f.write("\n".join(memo))

# ----------------------------- write workbook ---------------------------------
wb = Workbook()
# styles
NAVY="1F2A44"; STEEL="2E4374"; RED="C0392B"; AMBER="E67E22"; GREEN="1E8449"; LIGHT="EAEFF7"
hdr_fill = PatternFill("solid", fgColor=STEEL)
hdr_font = Font(color="FFFFFF", bold=True, size=11)
title_font = Font(color=NAVY, bold=True, size=16)
sub_font = Font(color="555555", size=10)
sev_fill = {"High":PatternFill("solid",fgColor="F5B7B1"),
            "Medium":PatternFill("solid",fgColor="FAD7A0"),
            "Low":PatternFill("solid",fgColor="D5F5E3"),"OK":PatternFill("solid",fgColor="D5F5E3")}
thin = Side(style="thin", color="D5DCE8")
border = Border(left=thin,right=thin,top=thin,bottom=thin)
MONEY="#,##0;(#,##0)"; PCT="0.0%"

def style_header(ws, row, ncols):
    for c in range(1,ncols+1):
        cell = ws.cell(row=row, column=c)
        cell.fill=hdr_fill; cell.font=hdr_font; cell.border=border
        cell.alignment=Alignment(horizontal="center", vertical="center", wrap_text=True)

def autosize(ws, widths):
    for i,w in enumerate(widths,1):
        ws.column_dimensions[get_column_letter(i)].width = w

# ---- Summary tab
ws = wb.active; ws.title="Summary"
ws.sheet_view.showGridLines=False
ws["A1"]="Month-End Close — Exception Dashboard"; ws["A1"].font=title_font
ws["A2"]=f"Period {args.period}  ·  {len({r['entity'] for r in tb})} entities  ·  {len(je)} journal entries  ·  {len(rec)} reconciliations"; ws["A2"].font=sub_font
ws["A3"]="Automated draft — every exception requires preparer review and sign-off before close."; ws["A3"].font=Font(color=RED,italic=True,size=10)

cats = [
    ("Reconciliation exceptions", len(recon_findings), sum(abs(f["item"]) for f in recon_findings)),
    ("Material fluxes", len(flux_findings), sum(abs(f["mom"]) for f in flux_findings)),
    ("Journal-entry flags", len(je_findings), sum(f["amount"] for f in je_findings)),
    ("Benford vendor flags", len(benford_vendor_flags), 0),
    ("Intercompany imbalance", 1 if ic_flag else 0, abs(ic_net) if ic_flag else 0),
]
r0=5
ws.cell(row=r0,column=1,value="Category"); ws.cell(row=r0,column=2,value="Exceptions"); ws.cell(row=r0,column=3,value="$ Exposure")
style_header(ws,r0,3)
for i,(name,cnt,exp) in enumerate(cats):
    rr=r0+1+i
    ws.cell(row=rr,column=1,value=name).border=border
    c=ws.cell(row=rr,column=2,value=cnt); c.border=border; c.alignment=Alignment(horizontal="center")
    c.fill = sev_fill["High"] if cnt>0 and name!="Benford vendor flags" or (name=="Benford vendor flags" and cnt>0) else sev_fill["OK"]
    m=ws.cell(row=rr,column=3,value=exp); m.number_format=MONEY; m.border=border
tot=ws.cell(row=r0+1+len(cats),column=1,value="TOTAL EXCEPTIONS"); tot.font=Font(bold=True)
tc=ws.cell(row=r0+1+len(cats),column=2,value=total_exc); tc.font=Font(bold=True); tc.alignment=Alignment(horizontal="center")
ws.cell(row=r0+1+len(cats)+2,column=1,value="High-severity items are highlighted red on each tab. See the flux_memo.md for the narrative write-up.").font=sub_font
autosize(ws,[34,14,18])

# ---- helper to dump a findings tab
def dump_tab(title, headers, rows, widths, sev_col=None, num_cols=(), pct_cols=()):
    ws = wb.create_sheet(title)
    ws.sheet_view.showGridLines=False
    ws.cell(row=1,column=1,value=title).font=title_font
    hr=3
    for c,h in enumerate(headers,1): ws.cell(row=hr,column=c,value=h)
    style_header(ws,hr,len(headers))
    for i,row in enumerate(rows):
        rr=hr+1+i
        for c,val in enumerate(row,1):
            cell=ws.cell(row=rr,column=c,value=val); cell.border=border
            if c in num_cols: cell.number_format=MONEY
            if c in pct_cols: cell.number_format=PCT
        if sev_col:
            sv=row[sev_col-1]
            fill=sev_fill.get(sv)
            if fill: ws.cell(row=rr,column=sev_col).fill=fill
    ws.freeze_panes=ws.cell(row=hr+1,column=1)
    autosize(ws,widths)
    return ws

# Reconciliation tab
dump_tab("Recon Review",
    ["Entity","Acct","Account","GL Balance","Subledger","Recon Items","Unexplained","Aging","Prepared By","Severity","Issue"],
    [[f["entity"],f["account"],f["name"],f["gl"],f["sub"],f["item"],f["unexplained"],f["aging"],f["prepared_by"],f["severity"],f["issue"]] for f in recon_findings],
    [22,7,26,15,15,14,14,11,13,10,46], sev_col=10, num_cols=(4,5,6,7))

# Flux tab
dump_tab("Flux Analysis",
    ["Entity","Acct","Account","Type","Current","Prior","Budget","MoM $","MoM %","Bud Var $","Bud %","Severity","Commentary"],
    [[f["entity"],f["account"],f["name"],f["type"],f["current"],f["prior"],f["budget"],f["mom"],f["mom_pct"],f["bud_var"],f["bud_pct"],f["severity"],f["commentary"]] for f in flux_findings],
    [22,7,26,10,14,14,14,13,9,13,9,10,60], sev_col=12, num_cols=(5,6,7,8,10), pct_cols=(9,11))

# JE anomalies tab
dump_tab("JE Anomalies",
    ["JE ID","Entity","Date","User","Account","Vendor","Amount","Severity","Issue"],
    [[f["je"],f["entity"],f["date"],f["user"],f["account"],f["vendor"],f["amount"],f["severity"],f["issue"]] for f in je_findings],
    [10,22,12,16,26,18,14,10,60], sev_col=8, num_cols=(7,))

# Benford tab
wsb = wb.create_sheet("Benford Screen")
wsb.sheet_view.showGridLines=False
wsb.cell(row=1,column=1,value="Benford's Law Screen — leading-digit distribution").font=title_font
wsb.cell(row=2,column=1,value=f"Population: {n} journal-entry debit amounts ≥ $100.  Chi-square vs Benford: {chi:.1f} (higher = more deviation).").font=sub_font
hb=4
for c,h in enumerate(["Leading Digit","Actual Count","Actual %","Benford Expected %","Deviation"],1):
    wsb.cell(row=hb,column=c,value=h)
style_header(wsb,hb,5)
for i,row in enumerate(benford_rows):
    rr=hb+1+i
    wsb.cell(row=rr,column=1,value=row["digit"]).border=border
    wsb.cell(row=rr,column=2,value=row["actual_ct"]).border=border
    a=wsb.cell(row=rr,column=3,value=row["actual_pct"]); a.number_format=PCT; a.border=border
    e=wsb.cell(row=rr,column=4,value=row["expected_pct"]); e.number_format=PCT; e.border=border
    dv=wsb.cell(row=rr,column=5,value=row["deviation"]); dv.number_format=PCT; dv.border=border
    if abs(row["deviation"])>0.06: dv.fill=sev_fill["High"]
autosize(wsb,[14,14,12,18,12])
# vendor concentration flags below
vr=hb+len(benford_rows)+3
wsb.cell(row=vr,column=1,value="Vendor concentration flags").font=Font(bold=True,color=NAVY,size=12)
for c,h in enumerate(["Vendor","# Invoices","Lead Digit","Share","Severity","Issue"],1):
    wsb.cell(row=vr+1,column=c,value=h)
style_header(wsb,vr+1,6)
for i,f in enumerate(benford_vendor_flags):
    rr=vr+2+i
    wsb.cell(row=rr,column=1,value=f["vendor"]).border=border
    wsb.cell(row=rr,column=2,value=f["count"]).border=border
    wsb.cell(row=rr,column=3,value=f["digit"]).border=border
    sh=wsb.cell(row=rr,column=4,value=f["share"]); sh.number_format=PCT; sh.border=border
    sv=wsb.cell(row=rr,column=5,value=f["severity"]); sv.border=border; sv.fill=sev_fill["High"]
    wsb.cell(row=rr,column=6,value=f["issue"]).border=border

# Intercompany tab
wsi = wb.create_sheet("Intercompany")
wsi.sheet_view.showGridLines=False
wsi.cell(row=1,column=1,value="Intercompany Netting Check").font=title_font
wsi.cell(row=2,column=1,value="Due From Affiliates and Due To Affiliates must net to zero company-wide before consolidation.").font=sub_font
wsi.cell(row=4,column=1,value="Total Due From Affiliates"); wsi.cell(row=4,column=2,value=due_from).number_format=MONEY
wsi.cell(row=5,column=1,value="Total Due To Affiliates"); wsi.cell(row=5,column=2,value=due_to).number_format=MONEY
wsi.cell(row=6,column=1,value="Net imbalance").font=Font(bold=True)
nc=wsi.cell(row=6,column=2,value=ic_net); nc.number_format=MONEY; nc.font=Font(bold=True)
nc.fill = sev_fill["High"] if ic_flag else sev_fill["OK"]
wsi.cell(row=7,column=1,value=("OUT OF BALANCE — investigate before consolidating." if ic_flag else "In balance.")).font=Font(color=RED if ic_flag else GREEN, bold=True)
hh=9
for c,h in enumerate(["Entity","Account","Balance"],1): wsi.cell(row=hh,column=c,value=h)
style_header(wsi,hh,3)
for i,r in enumerate(ic_by_entity):
    rr=hh+1+i
    wsi.cell(row=rr,column=1,value=r["entity"]).border=border
    wsi.cell(row=rr,column=2,value=r["account"]).border=border
    b=wsi.cell(row=rr,column=3,value=r["balance"]); b.number_format=MONEY; b.border=border
autosize(wsi,[24,26,18])

out_xlsx=os.path.join(OUT,"close_workpaper.xlsx")
wb.save(out_xlsx)

# ----------------------------- console recap ----------------------------------
print("=== Close Accelerator — exception recap ===")
print(f"Reconciliation exceptions : {len(recon_findings)}")
for f in recon_findings: print(f"   [{f['severity']}] {f['entity']} {f['name']}: {f['issue']}")
print(f"Material fluxes           : {len(flux_findings)}")
for f in flux_findings: print(f"   [{f['severity']}] {f['entity']} {f['name']}: MoM {f['mom_pct']*100:,.0f}%")
print(f"JE anomalies              : {len(je_findings)}")
for f in je_findings: print(f"   [{f['severity']}] {f['je']} {f['entity']} {f['account']} ${f['amount']:,.0f}: {f['issue']}")
print(f"Benford vendor flags      : {len(benford_vendor_flags)}")
for f in benford_vendor_flags: print(f"   [{f['severity']}] {f['vendor']}: {f['issue']}")
print(f"Intercompany net imbalance: ${ic_net:,.2f}  flagged={ic_flag}")
print(f"\nWrote: {out_xlsx}")
print(f"Wrote: {os.path.join(OUT,'flux_memo.md')}")
