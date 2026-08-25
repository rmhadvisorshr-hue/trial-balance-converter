import PDFDocument from "pdfkit";
import type { ConvertPayload } from "../../types";
import { FIGURES_UNIT_DIVISOR, FIGURES_UNIT_HEADING } from "../../types";
import { computeStatements } from "../../statements";
import type { Statements } from "../../statements";
import { round2 } from "../../classify";
import { LIABILITY_NOTES, ASSET_NOTES, PL_NOTES } from "./notes";
import type { NoteDef } from "./notes";
import { capitalShareholders } from "./shareCapitalReserves";

const MARGIN = 40;
const PORTRAIT = { w: 595.28, h: 841.89 };
const LANDSCAPE = { w: 841.89, h: 595.28 };
const FACE_VALUE = 10;
const DEFAULT_TAX_RATE = 0.26;

function money(value: number, divisor: number): string {
  const scaled = value / divisor;
  const text = Math.abs(scaled).toLocaleString("en-IN", {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  });
  return scaled < 0 ? `(${text})` : text;
}

// A slimmed-down version of the layout cursor used by the "statutory" PDF
// (../pdf.ts), extended with a generic multi-column table() renderer for
// the Notes/11.FA/DTL/DTA/Ageing schedules, and orientation switching for
// the wide fixed-asset tables. pdfkit doesn't auto-paginate absolutely-
// positioned content, so every row tracks its own y and page-breaks itself.
class Cursor {
  y: number;
  page: { w: number; h: number } = PORTRAIT;
  private bottom = PORTRAIT.h - MARGIN - 24;
  private contentW = PORTRAIT.w - MARGIN * 2;

  constructor(private doc: PDFKit.PDFDocument) {
    this.y = doc.y;
  }

  newPage(orientation: "portrait" | "landscape" = "portrait") {
    this.page = orientation === "landscape" ? LANDSCAPE : PORTRAIT;
    this.bottom = this.page.h - MARGIN - 24;
    this.contentW = this.page.w - MARGIN * 2;
    this.doc.addPage({ size: "A4", layout: orientation, margin: MARGIN });
    this.y = MARGIN;
  }

  space(needed: number) {
    if (this.y + needed > this.bottom) this.newPage(this.page === LANDSCAPE ? "landscape" : "portrait");
  }

  get contentWidth() {
    return this.contentW;
  }

  // Two-column row: a label on the left, a value on the right - used for
  // the signature block, where left/right lines don't line up 1:1.
  twoCol(left: string, right: string, opts: { bold?: boolean; size?: number } = {}) {
    const size = opts.size ?? 9;
    this.space(14);
    if (left) this.text(left, MARGIN, 280, { size });
    if (right) this.text(right, MARGIN + 300, this.contentW - 300, { size, bold: opts.bold });
    this.y += 14;
  }

  text(str: string, x: number, width: number, opts: { bold?: boolean; italic?: boolean; size?: number; align?: "left" | "right" | "center"; color?: string } = {}) {
    const font = opts.bold ? "Helvetica-Bold" : opts.italic ? "Helvetica-Oblique" : "Helvetica";
    this.doc.font(font).fontSize(opts.size ?? 9.5).fillColor(opts.color ?? "#111827").text(str, x, this.y, { width, align: opts.align ?? "left" });
  }

  heading(text: string, opts: { size?: number; color?: string } = {}) {
    this.space((opts.size ?? 12) + 6);
    this.text(text, MARGIN, this.contentW, { bold: true, size: opts.size ?? 12, color: opts.color });
    this.y += (opts.size ?? 12) + 6;
  }

  note(text: string, opts: { italic?: boolean } = { italic: true }) {
    const h = Math.max(11, this.doc.font("Helvetica-Oblique").fontSize(8).heightOfString(text, { width: this.contentW }) + 3);
    this.space(h);
    this.text(text, MARGIN, this.contentW, { italic: opts.italic ?? true, size: 8, color: "#6b7280" });
    this.y += h;
  }

  gap(h = 8) {
    this.y += h;
  }

  rule() {
    this.space(6);
    this.doc.moveTo(MARGIN, this.y).lineTo(MARGIN + this.contentW, this.y).lineWidth(0.75).strokeColor("#9ca3af").stroke();
    this.y += 6;
  }

  // A simple label/note-number/amount row, matching the Excel BS/P&L layout.
  line(label: string, amount: number | null, divisor: number, opts: { bold?: boolean; indent?: number; noteNo?: number | string; size?: number } = {}) {
    const size = opts.size ?? 9.5;
    const indent = opts.indent ?? 0;
    const labelW = this.contentW - 260;
    this.doc.font(opts.bold ? "Helvetica-Bold" : "Helvetica").fontSize(size);
    const h = Math.max(14, this.doc.heightOfString(label, { width: labelW - indent }) + 4);
    this.space(h);
    this.text(label, MARGIN + indent, labelW - indent, { bold: opts.bold, size });
    if (opts.noteNo !== undefined) this.text(String(opts.noteNo), MARGIN + labelW + 20, 40, { size, align: "center" });
    if (amount !== null) this.text(money(amount, divisor), MARGIN + this.contentW - 160, 160, { bold: opts.bold, size, align: "right" });
    this.y += h;
  }

  // Generic table with fixed column widths (proportional to contentW).
  table(headers: string[], widths: number[], rows: (string | number | null)[][], opts: { moneyCols?: number[]; divisor?: number } = {}) {
    const totalW = widths.reduce((s, w) => s + w, 0);
    const colX = (i: number) => MARGIN + widths.slice(0, i).reduce((s, w) => s + (w / totalW) * this.contentW, 0);
    const colW = (i: number) => (widths[i] / totalW) * this.contentW;

    this.space(16);
    headers.forEach((h, i) => this.text(h, colX(i), colW(i), { bold: true, size: 8.5, align: i === 0 ? "left" : "right" }));
    this.y += 14;
    this.doc.moveTo(MARGIN, this.y - 2).lineTo(MARGIN + this.contentW, this.y - 2).lineWidth(0.5).strokeColor("#d1d5db").stroke();

    for (const row of rows) {
      this.space(13);
      row.forEach((cell, i) => {
        if (cell === null) return;
        const isMoney = opts.moneyCols?.includes(i);
        const text = isMoney ? money(Number(cell), opts.divisor ?? 1) : String(cell);
        this.text(text, colX(i), colW(i), { size: 8.5, align: i === 0 ? "left" : "right" });
      });
      this.y += 13;
    }
  }
}

interface RenderCtx {
  st: Statements;
  divisor: number;
  unitHeading: string | null;
  payload: ConvertPayload;
}

function noteTotal(st: Statements, def: NoteDef): number {
  if (def.net) return round2(st.total("OPENING_STOCK") - st.total("CLOSING_STOCK_PL"));
  return round2(def.codes.reduce((s, c) => s + st.total(c), 0));
}

function renderCover(doc: PDFKit.PDFDocument, rc: RenderCtx) {
  const { payload, unitHeading } = rc;
  doc.y = 220;
  doc.font("Helvetica-Bold").fontSize(22).fillColor("#111827").text(payload.meta.firmName || "COMPANY NAME", MARGIN, doc.y, { width: PORTRAIT.w - MARGIN * 2, align: "center" });
  doc.moveDown(0.5);
  if (payload.meta.cin) {
    doc.font("Helvetica").fontSize(11).fillColor("#374151").text(`CIN : ${payload.meta.cin}`, MARGIN, doc.y, { width: PORTRAIT.w - MARGIN * 2, align: "center" });
    doc.moveDown(0.8);
  }
  doc.font("Helvetica").fontSize(14).fillColor("#374151").text("SCHEDULE III FINANCIAL STATEMENTS", MARGIN, doc.y, { width: PORTRAIT.w - MARGIN * 2, align: "center" });
  doc.moveDown(0.6);
  doc.fontSize(11).text(payload.meta.periodLabel || "", MARGIN, doc.y, { width: PORTRAIT.w - MARGIN * 2, align: "center" });
  if (unitHeading) {
    doc.moveDown(1.2);
    doc.font("Helvetica-Oblique").fontSize(10).fillColor("#6b7280").text(unitHeading, MARGIN, doc.y, { width: PORTRAIT.w - MARGIN * 2, align: "center" });
  }
}

function renderSignature(c: Cursor, rc: RenderCtx) {
  const { payload } = rc;
  const meta = payload.meta;
  c.gap(16);
  c.rule();
  c.twoCol(`Place : ${meta.place || "Vasai"}`, `For ${meta.firmName || "the Company"}`, { bold: true });
  c.twoCol(`Date : ${meta.date || ""}`, "");
  c.twoCol(`UDIN : ${meta.udin || ""}`, "For and on behalf of the Board", { bold: true });
  c.gap(20);
  c.twoCol("", "Director");
  c.twoCol("", "DIN :");
  c.gap(10);
  c.twoCol("As per our report on even date", "");
  c.twoCol(`For ${meta.caName || "the Chartered Accountants"}`, "");
  c.twoCol("(Chartered Accountants)", "");
  c.gap(6);
  c.twoCol(`Proprietor : CA ${meta.caName || ""}`, "");
  c.twoCol(`M No : ${meta.caMembershipNo || ""}`, "");
  c.twoCol(`FRN No : ${meta.caFirmRegNo || ""}`, "");
}

function renderBalanceSheet(doc: PDFKit.PDFDocument, c: Cursor, rc: RenderCtx) {
  const { st, divisor, unitHeading, payload } = rc;
  c.newPage("portrait");
  c.heading(payload.meta.firmName || "COMPANY NAME", { size: 13 });
  if (payload.meta.cin) c.heading(`CIN : ${payload.meta.cin}`, { size: 9 });
  c.heading(`BALANCE SHEET AS AT ${(payload.meta.asAtLabel || "").replace(/^as at\s*/i, "")}`.trim(), { size: 11 });
  if (unitHeading) c.heading(unitHeading, { size: 9, color: "#6b7280" });
  c.gap(6);

  c.line("I  EQUITY AND LIABILITIES", null, divisor, { bold: true });
  c.line("1  Shareholders' funds", null, divisor, { bold: true });
  c.line("(a) Share capital", st.total("CAPITAL"), divisor, { indent: 14, noteNo: 2 });
  const reserves = round2(st.total("RESERVES") + st.netProfit);
  c.line("(b) Reserves and surplus", reserves, divisor, { indent: 14, noteNo: 3 });

  c.line("2  Non-current liabilities", null, divisor, { bold: true });
  c.line("(a) Long term borrowings", st.total("LONG_TERM_BORROW"), divisor, { indent: 14, noteNo: 4 });
  c.line("(b) Other long term liabilities", 0, divisor, { indent: 14, noteNo: 5 });
  c.line("(c) Deferred tax liabilities (net)", st.total("DEFERRED_TAX_LIAB"), divisor, { indent: 14, noteNo: 6 });

  c.line("3  Current liabilities", null, divisor, { bold: true });
  c.line("(a) Short term borrowings", st.total("SHORT_TERM_BORROW"), divisor, { indent: 14, noteNo: 7 });
  c.line("(b) Trade payables", st.total("TRADE_PAYABLES"), divisor, { indent: 14, noteNo: 8 });
  c.line("(c) Other current liabilities", st.total("OTHER_CURR_LIAB"), divisor, { indent: 14, noteNo: 9 });
  c.line("(d) Short term provisions", st.total("SHORT_TERM_PROV"), divisor, { indent: 14, noteNo: 10 });
  c.line("TOTAL", round2(st.totalLiabilities), divisor, { bold: true });
  c.rule();
  c.gap(12);

  c.line("II  ASSETS", null, divisor, { bold: true });
  c.line("1  Non-current assets", null, divisor, { bold: true });
  c.line("(a) Property, Plant and Equipment and Intangible Assets", st.total("PPE"), divisor, { indent: 14, noteNo: 11 });
  c.line("(b) Non-current investments", st.total("NONCURR_INVEST"), divisor, { indent: 14, noteNo: 12 });
  c.line("(c) Long-term loans and advances", round2(st.total("LONG_TERM_LOANS_ADV")), divisor, { indent: 14 });
  c.line("(d) Deferred tax assets (net)", st.total("DEFERRED_TAX_ASSET"), divisor, { indent: 14, noteNo: 13 });

  c.line("2  Current assets", null, divisor, { bold: true });
  c.line("(a) Inventories", st.total("INVENTORY"), divisor, { indent: 14, noteNo: 14 });
  c.line("(b) Trade receivables", st.total("TRADE_RECV"), divisor, { indent: 14, noteNo: 15 });
  c.line("(c) Cash and Bank Balances", st.total("CASH_BANK"), divisor, { indent: 14, noteNo: 16 });
  c.line("(d) Short term loans and advances", st.total("SHORT_TERM_LOANS_ADV"), divisor, { indent: 14, noteNo: 17 });
  c.line("(e) Other current assets", st.total("OTHER_CURR_ASSETS"), divisor, { indent: 14, noteNo: 18 });
  c.line("TOTAL", round2(st.totalAssets), divisor, { bold: true });
  c.rule();
  c.gap(6);

  c.note(`Balancing check (Total Assets - Total Liabilities, should be 0): ${money(round2(st.totalAssets - st.totalLiabilities), divisor)}`, { italic: true });
  c.gap(6);
  c.note("The accompanying notes 2-31 are an integral part of these financial statements.");
}

function renderProfitAndLoss(doc: PDFKit.PDFDocument, c: Cursor, rc: RenderCtx) {
  const { st, divisor, unitHeading, payload } = rc;
  c.newPage("portrait");
  c.heading(payload.meta.firmName || "COMPANY NAME", { size: 13 });
  if (payload.meta.cin) c.heading(`CIN : ${payload.meta.cin}`, { size: 9 });
  c.heading("PROFIT & LOSS STATEMENT FOR THE YEAR ENDED", { size: 11 });
  if (payload.meta.periodLabel) c.heading(payload.meta.periodLabel, { size: 9.5, color: "#374151" });
  if (unitHeading) c.heading(unitHeading, { size: 9, color: "#6b7280" });
  c.gap(6);

  c.line("I  Revenue from operations", st.revenue, divisor, { noteNo: 19 });
  c.line("II  Other income", st.otherIncome, divisor, { noteNo: 20 });
  const totalIncome = round2(st.revenue + st.otherIncome);
  c.line("III  Total Income (I + II)", totalIncome, divisor, { bold: true });
  c.gap(6);

  c.line("IV  Expenses:", null, divisor, { bold: true });
  const costMaterials = round2(st.total("PURCHASES") + st.total("DIRECT_EXP"));
  const changeInv = round2(st.total("OPENING_STOCK") - st.total("CLOSING_STOCK_PL"));
  c.line("Cost of materials consumed / Purchases", costMaterials, divisor, { indent: 14, noteNo: 21 });
  c.line("Changes in inventory", changeInv, divisor, { indent: 14, noteNo: 22 });
  c.line("Employee benefit expenses", st.employeeBenefit, divisor, { indent: 14, noteNo: 23 });
  c.line("Finance costs", st.financeCost, divisor, { indent: 14, noteNo: 24 });
  c.line("Depreciation and amortisation", st.depreciation, divisor, { indent: 14, noteNo: 11 });
  c.line("Other expenses", st.otherExpense, divisor, { indent: 14, noteNo: 25 });
  const totalExpense = round2(costMaterials + changeInv + st.employeeBenefit + st.financeCost + st.depreciation + st.otherExpense);
  c.line("V  Total expenses", totalExpense, divisor, { bold: true });
  c.gap(6);

  const pbt = round2(totalIncome - totalExpense);
  c.line("VI  Profit before tax (III - V)", pbt, divisor, { bold: true });
  c.gap(6);
  c.line("VII  Tax expense:", null, divisor, { bold: true });
  c.line("(1) Current tax (enter manually)", 0, divisor, { indent: 14 });
  c.line("(2) Deferred tax", 0, divisor, { indent: 14 });
  c.line("(3) (Excess)/Short provision of income tax", 0, divisor, { indent: 14 });
  const netProfit = pbt;
  c.line("VIII  Profit for the year", netProfit, divisor, { bold: true });
  c.rule();
  c.gap(6);

  const capitalTotal = round2(st.total("CAPITAL"));
  const issuedShares = Math.max(1, Math.round(capitalTotal / FACE_VALUE));
  c.line(`IX  Earnings per equity share (face value Rs.${FACE_VALUE}/share, assumed)`, round2(netProfit / issuedShares), divisor);
  c.note("Share count for EPS is estimated from Capital / face value - confirm against the share register.");
  c.note("Current tax, and any deferred tax movement beyond the default (see DTA), must be entered by the CA - not derivable from a trial balance.");
}

function renderNotesSection(doc: PDFKit.PDFDocument, c: Cursor, rc: RenderCtx, title: string, defs: NoteDef[]) {
  const { st, divisor, payload } = rc;
  c.newPage("portrait");
  c.heading(payload.meta.firmName || "COMPANY NAME", { size: 12 });
  c.heading(title, { size: 11 });
  c.gap(6);

  for (const def of defs) {
    c.space(30);
    c.line(`Note "${def.no}" : ${def.title}`, null, divisor, { bold: true });
    if (def.codes.length === 0) {
      c.line("-", 0, divisor, { indent: 14 });
      c.gap(8);
      continue;
    }
    if (def.net) {
      c.line("Opening stock", round2(st.total("OPENING_STOCK")), divisor, { indent: 14 });
      c.line("Less: Closing stock", round2(-st.total("CLOSING_STOCK_PL")), divisor, { indent: 14 });
    } else {
      let any = false;
      for (const code of def.codes) {
        for (const item of st.byCode.get(code) ?? []) {
          c.line(item.name, item.amount, divisor, { indent: 14 });
          any = true;
        }
      }
      if (!any) c.line("-", 0, divisor, { indent: 14 });
    }
    if (def.no === 8) c.note("(reclassify above into MSME / Others per the MSME Act - defaults to Others)");
    c.line("Total", noteTotal(st, def), divisor, { bold: true, indent: 14 });
    c.gap(8);
  }
}

function renderShareCapitalReserves(doc: PDFKit.PDFDocument, c: Cursor, rc: RenderCtx) {
  const { st, divisor, payload } = rc;
  c.newPage("portrait");
  c.heading(payload.meta.firmName || "COMPANY NAME", { size: 12 });
  c.heading('Note "2" : SHARE CAPITAL', { size: 11 });
  c.gap(4);

  const capitalTotal = round2(st.total("CAPITAL"));
  const issuedShares = Math.round(capitalTotal / FACE_VALUE);
  c.line("Authorized Shares (enter manually - not present in a trial balance)", 0, divisor, { indent: 14 });
  c.line(`Issued, Subscribed & Fully Paid up - Equity Shares of Rs.${FACE_VALUE} each (${issuedShares} shares, estimated)`, capitalTotal, divisor, { indent: 14 });
  c.line("Total (Note 2)", capitalTotal, divisor, { bold: true });
  c.gap(8);

  c.note("Terms/Rights: one class of equity shares, one vote per share, residual entitlement on liquidation in proportion to holding.");
  c.note("Shares held by holding/ultimate holding company - NIL.");
  c.gap(6);

  c.line("Shareholders holding more than 5% shares:", null, divisor, { bold: true });
  const shareholders = capitalShareholders(rc.st, issuedShares);
  if (shareholders.length) {
    for (const sh of shareholders) c.line(`${sh.name} - ${sh.shares} shares (${sh.pct})`, null, divisor, { indent: 14 });
  } else {
    c.line("(refer share register - capital ledgers were not individually named)", null, divisor, { indent: 14 });
  }
  c.gap(10);

  c.heading('Note "3" : RESERVES & SURPLUS', { size: 11 });
  c.line("Opening balance (prior-year closing reserves not available - enter manually)", 0, divisor, { indent: 14 });
  c.line("(+) Net Profit/(Loss) for the current year", st.netProfit, divisor, { indent: 14 });
  c.line("(+) Opening Balance Difference", 0, divisor, { indent: 14 });
  c.line("Total (Note 3 - Closing Balance)", round2(st.netProfit), divisor, { bold: true });
}

function renderFixedAssetChain(doc: PDFKit.PDFDocument, c: Cursor, rc: RenderCtx) {
  const { st, divisor, payload } = rc;
  const ppeItems = st.byCode.get("PPE") ?? [];

  c.newPage("landscape");
  c.heading(payload.meta.firmName || "COMPANY NAME", { size: 12 });
  c.heading('Note "11" : PROPERTY, PLANT & EQUIPMENT AND INTANGIBLE ASSETS', { size: 10.5 });
  c.note("A trial balance gives closing net book value only - Gross Block / Accumulated Depreciation split, opening balances, and additions/deletions are not available and default to 0 (enter manually).");
  c.gap(6);

  const headers = ["Particular", "Gross Opening", "Additions", "Deletions", "Gross Closing", "Dep Opening", "Dep For Year", "Dep Disposals", "Dep Closing", "Net Closing"];
  const widths = [26, 8, 8, 8, 9, 8, 9, 8, 8, 10];
  const rows: (string | number)[][] = ppeItems.map((a) => [a.name, 0, 0, 0, 0, 0, 0, 0, 0, round2(a.amount)]);
  rows.push(["Depreciation per Trial Balance (unallocated)", 0, 0, 0, 0, 0, round2(st.depreciation), 0, 0, 0]);
  const totalNet = round2(ppeItems.reduce((s, a) => s + a.amount, 0));
  rows.push(["Total", 0, 0, 0, 0, 0, round2(st.depreciation), 0, 0, totalNet]);
  c.table(headers, widths, rows, { moneyCols: [1, 2, 3, 4, 5, 6, 7, 8, 9], divisor });

  c.gap(16);
  c.heading('Note "9" (DTL) : AS PER BOOKS OF ACCOUNTS', { size: 10 });
  c.note("Mirrors the '11. FA' schedule above.");
  c.gap(16);
  c.heading("AS PER INCOME TAX ACT", { size: 10 });
  c.note("Income-tax depreciation (WDV, block-wise per the Income Tax Rules) cannot be derived from a trial balance - defaults to the books figure (zero timing difference) below.");
  c.line("Depreciation for the year (enter manually; defaults to books figure)", round2(st.depreciation), divisor, { indent: 14 });

  c.newPage("portrait");
  c.heading(payload.meta.firmName || "COMPANY NAME", { size: 12 });
  c.heading("DTA : Deferred Tax Calculation", { size: 11 });
  c.gap(4);
  c.line("Tax rate (edit if different)", null, 1);
  c.text(`${(DEFAULT_TAX_RATE * 100).toFixed(0)}%`, MARGIN + 335, 160, { align: "right" });
  c.y += 14;
  c.gap(6);

  c.line("(A) Fixed Assets", null, divisor, { bold: true });
  c.line("Depreciation as per Books", round2(st.depreciation), divisor, { indent: 14 });
  c.line("Depreciation as per Income Tax Act (defaults to Books)", round2(st.depreciation), divisor, { indent: 14 });
  c.line("Difference", 0, divisor, { indent: 14 });
  c.line("Deferred Tax Liability / (Asset)", 0, divisor, { indent: 14, bold: true });
  c.gap(6);

  c.line("(B) Disallowances under Sec. 43B (enter manually)", null, divisor, { bold: true });
  c.line("(enter each disallowance as a separate line)", 0, divisor, { indent: 14 });
  c.line("Deferred Tax Asset (43B)", 0, divisor, { indent: 14, bold: true });
  c.gap(6);
  c.line("Net Deferred Tax Expense / (Income)", 0, divisor, { bold: true });
}

function renderRatios(doc: PDFKit.PDFDocument, c: Cursor, rc: RenderCtx) {
  const { st, divisor, payload } = rc;
  c.newPage("landscape");
  c.heading(payload.meta.firmName || "COMPANY NAME", { size: 12 });
  c.heading('Note "24" : Disclosure of Ratios', { size: 11 });
  c.note("Ratios use closing-balance figures in place of period averages (a trial balance has no prior-year data); confirm before filing.");
  c.gap(6);

  const currentAssets = round2(["INVENTORY", "TRADE_RECV", "CASH_BANK", "SHORT_TERM_LOANS_ADV", "OTHER_CURR_ASSETS"].reduce((s, code) => s + st.total(code as Parameters<Statements["total"]>[0]), 0));
  const currentLiab = round2(["SHORT_TERM_BORROW", "TRADE_PAYABLES", "OTHER_CURR_LIAB", "SHORT_TERM_PROV"].reduce((s, code) => s + st.total(code as Parameters<Statements["total"]>[0]), 0));
  const capitalAndReserves = round2(st.total("CAPITAL") + st.total("RESERVES") + st.netProfit);
  const borrowings = round2(st.total("LONG_TERM_BORROW") + st.total("SHORT_TERM_BORROW"));
  const ebitda = round2(st.grossProfit + st.otherIncome - st.employeeBenefit - st.otherExpense);
  const cogs = round2(st.total("PURCHASES") + st.total("DIRECT_EXP") + st.total("OPENING_STOCK") - st.total("CLOSING_STOCK_PL"));
  const purchases = round2(st.total("PURCHASES") + st.total("DIRECT_EXP"));
  const workingCapital = round2(currentAssets - currentLiab);
  const capitalEmployed = round2(capitalAndReserves + st.total("LONG_TERM_BORROW"));
  const safeDiv = (n: number, d: number) => (d === 0 ? 0 : round2(n / d));

  const rows: (string | number)[][] = [
    ["a.", "Current Ratio (times)", safeDiv(currentAssets, currentLiab)],
    ["b.", "Debt-Equity Ratio (times)", safeDiv(borrowings, capitalAndReserves)],
    ["c.", "Debt Service Coverage Ratio (times, interest only)", safeDiv(ebitda, st.financeCost)],
    ["d.", "Return on Equity Ratio (%)", safeDiv(st.netProfit, capitalAndReserves)],
    ["e.", "Inventory Turnover Ratio (times)", safeDiv(cogs, st.total("INVENTORY"))],
    ["f.", "Trade Receivables Turnover Ratio (times)", safeDiv(st.revenue, st.total("TRADE_RECV"))],
    ["g.", "Trade Payables Turnover Ratio (times)", safeDiv(purchases, st.total("TRADE_PAYABLES"))],
    ["h.", "Net Capital Turnover Ratio (times)", safeDiv(st.revenue, workingCapital)],
    ["i.", "Net Profit Ratio (%)", safeDiv(st.netProfit, st.revenue)],
    ["j.", "Return on Capital Employed (%)", safeDiv(round2(ebitda - st.depreciation), capitalEmployed)],
    ["k.", "Return on Investment (%, Other Income proxy)", safeDiv(st.otherIncome, st.total("NONCURR_INVEST"))],
  ];
  c.table(["Sr", "Particulars", "Value"], [4, 40, 12], rows, { moneyCols: [2], divisor: 1 });
  void divisor;
}

function renderStatutoryNotes(doc: PDFKit.PDFDocument, c: Cursor, rc: RenderCtx) {
  const { payload } = rc;
  c.newPage("portrait");
  c.heading(payload.meta.firmName || "COMPANY NAME", { size: 12 });
  c.heading("Notes 27-31", { size: 11 });
  c.gap(6);

  c.line("27 - Related Party Disclosures", null, 1, { bold: true });
  const shareholders = capitalShareholders(rc.st, 0);
  if (shareholders.length) {
    for (const sh of shareholders) c.line(`${sh.name} - Director / Promoter (confirm)`, null, 1, { indent: 14 });
  } else {
    c.note("(no individually-named capital ledgers found - enter related parties manually)");
  }
  c.note("Enter each related-party transaction (rent, remuneration, sales, purchases, loans, etc.) - not derivable from a trial balance.");
  c.gap(8);

  c.line("28 - MSME Disclosure", null, 1, { bold: true });
  c.note("This app cannot determine MSME-registration status of creditors from a trial balance - amounts default to 0; enter from the MSME register / vendor confirmations.");
  const msmeItems = [
    "(a) Principal + interest due and unpaid at year end",
    "(b) Interest paid under Section 16",
    "(c) Interest due for delayed payment",
    "(d) Interest accrued and unpaid at year end",
    "(e) Interest due and payable in succeeding years",
  ];
  for (const item of msmeItems) c.line(item, 0, 1, { indent: 14 });
  c.gap(8);

  c.line("29 - Realizability of Current Assets", null, 1, { bold: true });
  c.note("In the opinion of the Board, current assets, loans and advances have a realizable value in the ordinary course of business at least equal to the amount stated, and all known liabilities are provided for.");
  c.gap(8);

  c.line("30 - Other Statutory Information", null, 1, { bold: true });
  const checklist = [
    "(i) No benami property proceedings.",
    "(ii) No transactions with companies struck off under Sec. 248 (2013 Act) / Sec. 560 (1956 Act).",
    "(iii) No trading/investment in Crypto or Virtual Currency during the year.",
    "(iv) No funds advanced/loaned/invested with an understanding to lend/invest on the Company's behalf in Ultimate Beneficiaries.",
    "(v) No funds received with an understanding to lend/invest on the Funding Party's behalf in Ultimate Beneficiaries.",
    "(vi) No income surrendered/disclosed in tax assessments that was not recorded in the books.",
    "(vii) Not declared a wilful defaulter by any bank/financial institution.",
    "(viii) No charge registration/satisfaction pending beyond the statutory period.",
  ];
  for (const item of checklist) c.line(item, null, 1, { indent: 14 });
  c.gap(8);

  c.line("31 - Regrouping / Reclassification", null, 1, { bold: true });
  c.note("Previous year's figures have been regrouped/reclassified wherever necessary to conform to the current year's presentation.");
}

function renderAgeing(doc: PDFKit.PDFDocument, c: Cursor, rc: RenderCtx) {
  const { st, divisor, payload } = rc;
  c.newPage("landscape");
  c.heading(payload.meta.firmName || "COMPANY NAME", { size: 12 });
  c.heading("Ageing Schedule (Trade Payables & Trade Receivables)", { size: 11 });
  c.note("Ageing buckets cannot be derived from a trial balance (no invoice/posting-date detail). Each total is placed under 'Not bucketed' so the grand total ties to Notes 8/15; redistribute using the debtor/creditor sub-ledger.");
  c.gap(6);

  const headers = ["Particulars", "Not bucketed", "< 1 year", "1-2 years", "2-3 years", "> 3 years", "Total"];
  const widths = [20, 12, 12, 12, 12, 12, 12];

  const payables = round2(st.total("TRADE_PAYABLES"));
  c.line("Trade Payables Ageing", null, divisor, { bold: true });
  c.table(headers, widths, [
    ["Undisputed", payables, 0, 0, 0, 0, payables],
    ["Disputed", 0, 0, 0, 0, 0, 0],
  ], { moneyCols: [1, 2, 3, 4, 5, 6], divisor });
  c.gap(16);

  const receivables = round2(st.total("TRADE_RECV"));
  c.line("Trade Receivables Ageing", null, divisor, { bold: true });
  c.table(headers, widths, [
    ["Undisputed", receivables, 0, 0, 0, 0, receivables],
    ["Disputed", 0, 0, 0, 0, 0, 0],
  ], { moneyCols: [1, 2, 3, 4, 5, 6], divisor });
}

function renderLoanMaturity(doc: PDFKit.PDFDocument, c: Cursor, rc: RenderCtx) {
  const { payload } = rc;
  c.newPage("portrait");
  c.heading(payload.meta.firmName || "COMPANY NAME", { size: 12 });
  c.heading("Current maturities of long term borrowings", { size: 11 });
  c.note("Loan EMI/interest/principal split requires the actual loan amortisation schedule (not available from a trial balance) - enter from the lender's statement.");
  c.gap(6);
  const rows = Array.from({ length: 12 }, (_, i) => [`Month ${i + 1}`, 0, 0, 0]);
  rows.push(["Total", 0, 0, 0]);
  c.table(["Period", "EMI", "Interest", "Principal"], [20, 14, 14, 14], rows, { moneyCols: [1, 2, 3], divisor: 1 });
}

export async function buildScheduleIIIPdf(payload: ConvertPayload): Promise<Buffer> {
  const st = computeStatements(payload.ledgers);
  const figuresUnit = payload.meta.figuresUnit || "actual";
  const divisor = FIGURES_UNIT_DIVISOR[figuresUnit];
  const unitHeading = FIGURES_UNIT_HEADING[figuresUnit];
  const rc: RenderCtx = { st, divisor, unitHeading, payload };

  const doc = new PDFDocument({ size: "A4", margin: MARGIN, bufferPages: true, autoFirstPage: false });
  const chunks: Buffer[] = [];
  doc.on("data", (chunk: Buffer) => chunks.push(chunk));
  const done = new Promise<Buffer>((resolve) => doc.on("end", () => resolve(Buffer.concat(chunks))));

  doc.addPage();
  renderCover(doc, rc);

  const c = new Cursor(doc);
  renderBalanceSheet(doc, c, rc);
  renderProfitAndLoss(doc, c, rc);
  renderSignature(c, rc);
  renderNotesSection(doc, c, rc, "Notes 4-10 (Liabilities)", LIABILITY_NOTES);
  renderNotesSection(doc, c, rc, "Notes 12-18 (Assets)", ASSET_NOTES);
  renderNotesSection(doc, c, rc, "Notes 19-25 (P&L)", PL_NOTES);
  renderShareCapitalReserves(doc, c, rc);
  renderFixedAssetChain(doc, c, rc);
  renderRatios(doc, c, rc);
  renderStatutoryNotes(doc, c, rc);
  renderAgeing(doc, c, rc);
  renderLoanMaturity(doc, c, rc);

  const range = doc.bufferedPageRange();
  for (let i = range.start; i < range.start + range.count; i++) {
    doc.switchToPage(i);
    const pageW = doc.page.width;
    const pageH = doc.page.height;
    const bottomMargin = doc.page.margins.bottom;
    doc.page.margins.bottom = 0;
    doc.font("Helvetica").fontSize(8).fillColor("#6b7280").text(`Page ${i - range.start + 1} of ${range.count}`, MARGIN, pageH - 30, {
      width: pageW - MARGIN * 2,
      align: "right",
      lineBreak: false,
    });
    doc.page.margins.bottom = bottomMargin;
  }

  doc.end();
  return done;
}
