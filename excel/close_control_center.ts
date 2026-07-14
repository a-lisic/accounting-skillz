/**
 * CLOSE CONTROL CENTER — Office Script (zero install, no admin rights needed)
 * ---------------------------------------------------------------------------
 * Runs a month-end close review INSIDE Excel: reconciliation review, flux
 * analysis, journal-entry anomaly testing (duplicates, round-dollar top-side,
 * off-hours/weekend manual postings), a Benford vendor screen, and an
 * intercompany netting check. Writes color-coded exception tabs + a summary.
 *
 * HOW TO INSTALL (Microsoft 365 Excel, work or web):
 *   1. Open Excel → "Automate" tab → "New Script".
 *   2. Delete the starter code, paste this entire file, rename it
 *      "Close Control Center", Save.
 *   3. In your workbook, create three sheets with these exact names and
 *      headers in row 1 (paste your ERP exports under them):
 *      TrialBalance:    entity | account | account_name | account_type |
 *                       current_balance | prior_balance | budget_balance
 *      JournalEntries:  je_id | entity | date | time | user | account |
 *                       account_name | description | debit | credit | vendor
 *      Reconciliations: entity | account | account_name | gl_balance |
 *                       subledger_balance | reconciling_items | aging | prepared_by
 *   4. Click Run. Review the "CC •" tabs it creates.
 *
 * EVERYTHING IS A DRAFT FOR REVIEW. This script posts nothing and files
 * nothing; a qualified accountant signs off before close. Flags mean
 * "worth a look," not "fraud found."
 * ---------------------------------------------------------------------------
 */

// ======================= CONFIG — tune to your materiality ==================
const CFG = {
  FLUX_PCT: 0.30,           // % change to flag (must ALSO clear FLUX_AMT)
  FLUX_AMT: 50000,          // $ change to flag
  RECON_UNEXPLAINED: 1000,  // unexplained GL-vs-subledger residual
  RECON_ITEM: 25000,        // large reconciling item
  ROUND_DOLLAR_MIN: 50000,  // round-$ floor for manual-JE scrutiny
  INTERCOMPANY_TOL: 1000,   // Due From vs Due To net tolerance
  IC_FROM_ACCT: "1900",     // Due From Affiliates account code
  IC_TO_ACCT: "2400",       // Due To Affiliates account code
  SHEET_TB: "TrialBalance",
  SHEET_JE: "JournalEntries",
  SHEET_REC: "Reconciliations"
};
// ============================================================================

function main(workbook: ExcelScript.Workbook) {
  const tb = readTable(workbook, CFG.SHEET_TB);
  const je = readTable(workbook, CFG.SHEET_JE);
  const rec = readTable(workbook, CFG.SHEET_REC);
  const missing: string[] = [];
  if (tb.length === 0) missing.push(CFG.SHEET_TB);
  if (je.length === 0) missing.push(CFG.SHEET_JE);
  if (rec.length === 0) missing.push(CFG.SHEET_REC);

  // ---------- 1) RECONCILIATION REVIEW ----------
  const reconRows: (string | number)[][] = [];
  for (const r of rec) {
    const gl = num(r["gl_balance"]), sub = num(r["subledger_balance"]), item = num(r["reconciling_items"]);
    const unexplained = gl - sub - item;
    const aging = str(r["aging"]);
    const reasons: string[] = [];
    let sev = "OK";
    if (Math.abs(unexplained) > CFG.RECON_UNEXPLAINED) { reasons.push(`Unexplained variance ${money(unexplained)}`); sev = "High"; }
    if (Math.abs(item) > CFG.RECON_ITEM) { reasons.push(`Large reconciling item ${money(item)}`); if (sevRank(sev) < 2) sev = "Medium"; }
    if ((aging.indexOf(">90") >= 0 || aging.indexOf(">60") >= 0) && Math.abs(item) > 5000) { reasons.push(`Aged reconciling item (${aging})`); sev = "High"; }
    if (reasons.length > 0) reconRows.push([str(r["entity"]), str(r["account"]), str(r["account_name"]), gl, sub, item, unexplained, aging, str(r["prepared_by"]), sev, reasons.join("; ")]);
  }

  // ---------- 2) FLUX ANALYSIS ----------
  const fluxRows: (string | number)[][] = [];
  for (const r of tb) {
    const cur = num(r["current_balance"]), prior = num(r["prior_balance"]), budget = num(r["budget_balance"]);
    const mom = cur - prior, momPct = prior !== 0 ? mom / Math.abs(prior) : 0;
    const budVar = cur - budget, budPct = budget !== 0 ? budVar / Math.abs(budget) : 0;
    const hitMoM = Math.abs(momPct) >= CFG.FLUX_PCT && Math.abs(mom) >= CFG.FLUX_AMT;
    const hitBud = Math.abs(budPct) >= CFG.FLUX_PCT && Math.abs(budVar) >= CFG.FLUX_AMT;
    if (hitMoM || hitBud) {
      const drivers: string[] = [];
      if (hitMoM) drivers.push(`${mom > 0 ? "up" : "down"} ${pct(momPct)} MoM (${money(mom)})`);
      if (hitBud) drivers.push(`${budVar > 0 ? "over" : "under"} budget ${pct(budPct)} (${money(budVar)})`);
      const sev = (Math.abs(momPct) >= 0.5 || Math.abs(budPct) >= 0.5) ? "High" : "Medium";
      fluxRows.push([str(r["entity"]), str(r["account"]), str(r["account_name"]), str(r["account_type"]), cur, prior, budget, mom, momPct, budVar, budPct, sev,
        `${str(r["account_name"])} at ${str(r["entity"])} is ${drivers.join(" and ")}. Obtain support and management explanation before sign-off.`]);
    }
  }
  fluxRows.sort((a, b) => Math.abs(num(b[7])) - Math.abs(num(a[7])));

  // ---------- 3) JOURNAL-ENTRY ANOMALIES ----------
  const jeRows: (string | number)[][] = [];
  const flagged: { [id: string]: boolean } = {};
  // duplicates: same entity/account/vendor/amount (debits), >1 distinct date
  const dupGroups: { [key: string]: { ids: string[]; dates: { [d: string]: boolean }; rows: { [k: string]: string | number | boolean }[] } } = {};
  for (const e of je) {
    const deb = num(e["debit"]); const vendor = str(e["vendor"]);
    if (deb > 0 && vendor !== "") {
      const key = [str(e["entity"]), str(e["account"]), vendor, deb.toFixed(2), str(e["description"]).slice(0, 40)].join("|");
      if (!dupGroups[key]) dupGroups[key] = { ids: [], dates: {}, rows: [] };
      dupGroups[key].ids.push(str(e["je_id"]));
      dupGroups[key].dates[dateKey(e["date"])] = true;
      dupGroups[key].rows.push(e);
    }
  }
  for (const key of Object.keys(dupGroups)) {
    const g = dupGroups[key];
    const nDates = Object.keys(g.dates).length;
    if (g.ids.length >= 2 && nDates >= 2) {
      for (const e of g.rows) {
        flagged[str(e["je_id"])] = true;
        jeRows.push([str(e["je_id"]), str(e["entity"]), dateKey(e["date"]), str(e["user"]), str(e["account_name"]), str(e["vendor"]), num(e["debit"]), "High",
          `Possible duplicate payment: ${g.ids.length} identical charges (${money(num(e["debit"]))}) to ${str(e["vendor"])} across ${nDates} dates`]);
      }
    }
  }
  // round-dollar / off-hours / weekend (timing rules apply ONLY to manual entries)
  for (const e of je) {
    const amt = Math.max(num(e["debit"]), num(e["credit"]));
    if (amt === 0 || flagged[str(e["je_id"])]) continue;
    const isManual = str(e["vendor"]) === "";
    const isRound = amt >= CFG.ROUND_DOLLAR_MIN && Math.abs(amt % 10000) < 0.005;
    const hr = hourOf(e["time"]);
    const isOffHours = hr < 6 || hr >= 20;
    const wd = weekdayOf(e["date"]);
    const isWeekend = wd === 0 || wd === 6;
    const reasons: string[] = [];
    let sev = "Low";
    if (isRound) { reasons.push(`Round-dollar ${money(amt)}`); sev = "Medium"; }
    if (isManual && isOffHours) { reasons.push(`Posted off-hours (${timeStr(e["time"])})`); if (sevRank(sev) < 2) sev = "Medium"; }
    if (isManual && isWeekend) { reasons.push(`Posted on a weekend (${dateKey(e["date"])})`); if (sevRank(sev) < 2) sev = "Medium"; }
    if (isRound && isManual && (isOffHours || isWeekend)) { sev = "High"; reasons.push("manual top-side entry, no vendor support"); }
    if (reasons.length > 0) jeRows.push([str(e["je_id"]), str(e["entity"]), dateKey(e["date"]), str(e["user"]), str(e["account_name"]), str(e["vendor"]), amt, sev, reasons.join("; ")]);
  }
  jeRows.sort((a, b) => (sevRank(str(b[7])) - sevRank(str(a[7]))) || (num(b[6]) - num(a[6])));

  // ---------- 4) BENFORD VENDOR SCREEN ----------
  const vendorLead: { [v: string]: number[] } = {};
  for (const e of je) {
    const deb = num(e["debit"]); const v = str(e["vendor"]);
    if (deb >= 100 && v !== "") {
      const d = leadDigit(deb);
      if (d > 0) { if (!vendorLead[v]) vendorLead[v] = [0,0,0,0,0,0,0,0,0,0]; vendorLead[v][d]++; }
    }
  }
  const benfordRows: (string | number)[][] = [];
  for (const v of Object.keys(vendorLead)) {
    const counts = vendorLead[v];
    let total = 0, topD = 1;
    for (let d = 1; d <= 9; d++) { total += counts[d]; if (counts[d] > counts[topD]) topD = d; }
    if (total >= 8 && counts[topD] / total >= 0.7) {
      benfordRows.push([v, total, topD, counts[topD] / total, "High",
        `${counts[topD]}/${total} invoices lead with digit ${topD} (${pct(counts[topD] / total)}) — possible fabricated/split invoicing`]);
    }
  }

  // ---------- 5) INTERCOMPANY NETTING ----------
  let dueFrom = 0, dueTo = 0;
  const icRows: (string | number)[][] = [];
  for (const r of tb) {
    const acct = str(r["account"]);
    if (acct === CFG.IC_FROM_ACCT) { dueFrom += num(r["current_balance"]); icRows.push([str(r["entity"]), str(r["account_name"]), num(r["current_balance"])]); }
    if (acct === CFG.IC_TO_ACCT)   { dueTo   += num(r["current_balance"]); icRows.push([str(r["entity"]), str(r["account_name"]), num(r["current_balance"])]); }
  }
  const icNet = dueFrom - dueTo;
  const icFlag = Math.abs(icNet) > CFG.INTERCOMPANY_TOL;

  // ---------- WRITE OUTPUT TABS ----------
  const highJe = jeRows.filter(r => r[7] === "High").length;
  const totalExc = reconRows.length + fluxRows.length + jeRows.length + benfordRows.length + (icFlag ? 1 : 0);

  writeSheet(workbook, "CC Recon",
    ["Entity", "Acct", "Account", "GL Balance", "Subledger", "Recon Items", "Unexplained", "Aging", "Prepared By", "Severity", "Issue"],
    reconRows, 10, [4, 5, 6, 7], []);
  writeSheet(workbook, "CC Flux",
    ["Entity", "Acct", "Account", "Type", "Current", "Prior", "Budget", "MoM $", "MoM %", "Bud Var $", "Bud %", "Severity", "Commentary"],
    fluxRows, 12, [5, 6, 7, 8, 10], [9, 11]);
  writeSheet(workbook, "CC JE Flags",
    ["JE ID", "Entity", "Date", "User", "Account", "Vendor", "Amount", "Severity", "Issue"],
    jeRows, 8, [7], []);
  writeSheet(workbook, "CC Benford",
    ["Vendor", "# Invoices", "Lead Digit", "Share", "Severity", "Issue"],
    benfordRows, 5, [], [4]);
  writeSheet(workbook, "CC Intercompany",
    ["Entity", "Account", "Balance"],
    icRows, 0, [3], []);

  // Summary tab
  const sumName = "CC Summary";
  const old = workbook.getWorksheet(sumName); if (old) old.delete();
  const ws = workbook.addWorksheet(sumName);
  const sumRows: (string | number)[][] = [
    ["CLOSE CONTROL CENTER — Exception Summary", "", ""],
    ["Automated draft — every exception requires preparer review and sign-off before close.", "", ""],
    ["", "", ""],
    ["Category", "Exceptions", "Notes"],
    ["Reconciliation exceptions", reconRows.length, "see CC Recon"],
    ["Material fluxes", fluxRows.length, "see CC Flux"],
    ["Journal-entry flags", jeRows.length, `${highJe} high severity — see CC JE Flags`],
    ["Benford vendor flags", benfordRows.length, "see CC Benford"],
    ["Intercompany imbalance", icFlag ? 1 : 0, `Due From ${money(dueFrom)} vs Due To ${money(dueTo)} → net ${money(icNet)}`],
    ["TOTAL EXCEPTIONS", totalExc, missing.length > 0 ? `WARNING: missing/empty input sheet(s): ${missing.join(", ")}` : "all three inputs loaded"]
  ];
  ws.getRangeByIndexes(0, 0, sumRows.length, 3).setValues(sumRows);
  ws.getRange("A1").getFormat().getFont().setBold(true);
  ws.getRange("A1").getFormat().getFont().setSize(14);
  ws.getRange("A2").getFormat().getFont().setItalic(true);
  ws.getRange("A2").getFormat().getFont().setColor("#C0392B");
  const hdr = ws.getRange("A4:C4");
  hdr.getFormat().getFill().setColor("#2E4374");
  hdr.getFormat().getFont().setColor("#FFFFFF");
  hdr.getFormat().getFont().setBold(true);
  for (let i = 4; i <= 8; i++) {
    const cnt = num(sumRows[i][1]);
    ws.getRangeByIndexes(i, 1, 1, 1).getFormat().getFill().setColor(cnt > 0 ? "#F5B7B1" : "#D5F5E3");
  }
  ws.getRange("A10:C10").getFormat().getFont().setBold(true);
  ws.getUsedRange().getFormat().autofitColumns();
  ws.activate();

  console.log(`Close Control Center: ${totalExc} exceptions (${highJe} high-severity JE flags). Review the CC tabs.`);
}

// =============================== helpers ====================================
function readTable(workbook: ExcelScript.Workbook, name: string): { [k: string]: string | number | boolean }[] {
  const ws = workbook.getWorksheet(name);
  if (!ws) return [];
  const used = ws.getUsedRange();
  if (!used) return [];
  const vals = used.getValues();
  if (vals.length < 2) return [];
  const headers = vals[0].map(h => String(h).trim().toLowerCase());
  const out: { [k: string]: string | number | boolean }[] = [];
  for (let i = 1; i < vals.length; i++) {
    let empty = true;
    const row: { [k: string]: string | number | boolean } = {};
    for (let c = 0; c < headers.length; c++) {
      row[headers[c]] = vals[i][c];
      if (String(vals[i][c]) !== "") empty = false;
    }
    if (!empty) out.push(row);
  }
  return out;
}

function writeSheet(workbook: ExcelScript.Workbook, name: string, headers: string[], rows: (string | number)[][], sevCol: number, moneyCols: number[], pctCols: number[]) {
  const old = workbook.getWorksheet(name); if (old) old.delete();
  const ws = workbook.addWorksheet(name);
  ws.getRangeByIndexes(0, 0, 1, headers.length).setValues([headers]);
  const hdr = ws.getRangeByIndexes(0, 0, 1, headers.length);
  hdr.getFormat().getFill().setColor("#2E4374");
  hdr.getFormat().getFont().setColor("#FFFFFF");
  hdr.getFormat().getFont().setBold(true);
  if (rows.length > 0) {
    ws.getRangeByIndexes(1, 0, rows.length, headers.length).setValues(rows);
    for (const c of moneyCols) ws.getRangeByIndexes(1, c - 1, rows.length, 1).setNumberFormatLocal("#,##0;(#,##0)");
    for (const c of pctCols) ws.getRangeByIndexes(1, c - 1, rows.length, 1).setNumberFormatLocal("0.0%");
    if (sevCol > 0) {
      for (let i = 0; i < rows.length; i++) {
        const sev = str(rows[i][sevCol - 1]);
        const color = sev === "High" ? "#F5B7B1" : sev === "Medium" ? "#FAD7A0" : "";
        if (color) ws.getRangeByIndexes(i + 1, sevCol - 1, 1, 1).getFormat().getFill().setColor(color);
      }
    }
  } else {
    ws.getRangeByIndexes(1, 0, 1, 1).setValues([["No exceptions — clean."]]);
  }
  ws.getUsedRange().getFormat().autofitColumns();
}

function num(x: string | number | boolean): number { const n = typeof x === "number" ? x : parseFloat(String(x).replace(/[$,]/g, "")); return isNaN(n) ? 0 : n; }
function str(x: string | number | boolean): string { return x === undefined || x === null ? "" : String(x).trim(); }
function sevRank(s: string): number { return s === "High" ? 3 : s === "Medium" ? 2 : s === "Low" ? 1 : 0; }
function money(x: number): string { const s = Math.round(Math.abs(x)).toString().replace(/\B(?=(\d{3})+(?!\d))/g, ","); return (x < 0 ? "-$" : "$") + s; }
function pct(x: number): string { return Math.round(x * 100) + "%"; }
function leadDigit(x: number): number { let s = Math.abs(x).toString(); for (const ch of s) { if (ch >= "1" && ch <= "9") return parseInt(ch, 10); } return 0; }

/** Excel may hand us dates as serial numbers or as strings — handle both. */
function toDate(x: string | number | boolean): Date | null {
  if (typeof x === "number") return new Date(Math.round((x - 25569) * 86400 * 1000)); // Excel serial → UTC ms
  const s = str(x);
  const m = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (m) return new Date(Date.UTC(parseInt(m[1], 10), parseInt(m[2], 10) - 1, parseInt(m[3], 10)));
  const d = new Date(s);
  return isNaN(d.getTime()) ? null : d;
}
function dateKey(x: string | number | boolean): string {
  const d = toDate(x); if (!d) return str(x);
  return `${d.getUTCFullYear()}-${String(d.getUTCMonth() + 1).padStart(2, "0")}-${String(d.getUTCDate()).padStart(2, "0")}`;
}
function weekdayOf(x: string | number | boolean): number { const d = toDate(x); return d ? d.getUTCDay() : -1; }
function hourOf(x: string | number | boolean): number {
  if (typeof x === "number") return Math.floor((x % 1) * 24); // Excel time fraction
  const m = str(x).match(/^(\d{1,2}):/); return m ? parseInt(m[1], 10) : 12;
}
function timeStr(x: string | number | boolean): string {
  if (typeof x === "number") { const mins = Math.round((x % 1) * 1440); return `${String(Math.floor(mins / 60)).padStart(2, "0")}:${String(mins % 60).padStart(2, "0")}`; }
  return str(x);
}
