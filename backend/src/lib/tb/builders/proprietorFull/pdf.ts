import PDFDocument from "pdfkit";
import type { ConvertPayload } from "../../types";
import { FIGURES_UNIT_DIVISOR, FIGURES_UNIT_HEADING } from "../../types";
import { computeStatements } from "../../statements";
import type { Statements } from "../../statements";
import { round2 } from "../../classify";

const MARGIN = 40;
const PORTRAIT = { w: 595.28, h: 841.89 };
const LANDSCAPE = { w: 841.89, h: 595.28 };
const TAX_RATE_CESS = 0.04;

function money(value: number, divisor: number): string {
  const scaled = value / divisor;
  const text = Math.abs(scaled).toLocaleString("en-IN", {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  });
  return scaled < 0 ? `(${text})` : text;
}

function slab115bac(nti: number): number {
  if (nti <= 300000) return 0;
  if (nti <= 700000) return (nti - 300000) * 0.05;
  if (nti <= 1000000) return (nti - 700000) * 0.1 + 20000;
  if (nti <= 1200000) return (nti - 1000000) * 0.15 + 50000;
  if (nti <= 1500000) return (nti - 1200000) * 0.2 + 110000;
  return (nti - 1500000) * 0.3 + 140000;
}

// Same self-contained layout engine as scheduleiii/pdf.ts and
// partnershipFull/pdf.ts (kept per-module rather than shared, matching
// their own precedent).
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

  text(
    str: string,
    x: number,
    width: number,
    opts: { bold?: boolean; italic?: boolean; size?: number; align?: "left" | "right" | "center"; color?: string } = {},
  ) {
    const font = opts.bold ? "Helvetica-Bold" : opts.italic ? "Helvetica-Oblique" : "Helvetica";
    this.doc.font(font).fontSize(opts.size ?? 9.5).fillColor(opts.color ?? "#111827").text(str, x, this.y, {
      width,
      align: opts.align ?? "left",
    });
  }

  heading(text: string, opts: { size?: number; color?: string } = {}) {
    this.space((opts.size ?? 12) + 6);
    this.text(text, MARGIN, this.contentW, { bold: true, size: opts.size ?? 12, color: opts.color });
    this.y += (opts.size ?? 12) + 6;
  }

  note(text: string) {
    const h = Math.max(11, this.doc.font("Helvetica-Oblique").fontSize(8).heightOfString(text, { width: this.contentW }) + 3);
    this.space(h);
    this.text(text, MARGIN, this.contentW, { italic: true, size: 8, color: "#6b7280" });
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

  line(label: string, amount: number | null, divisor: number, opts: { bold?: boolean; indent?: number; size?: number } = {}) {
    const size = opts.size ?? 9.5;
    const indent = opts.indent ?? 0;
    const labelW = this.contentW - 160;
    this.doc.font(opts.bold ? "Helvetica-Bold" : "Helvetica").fontSize(size);
    const h = Math.max(14, this.doc.heightOfString(label, { width: labelW - indent }) + 4);
    this.space(h);
    this.text(label, MARGIN + indent, labelW - indent, { bold: opts.bold, size });
    if (amount !== null) this.text(money(amount, divisor), MARGIN + this.contentW - 140, 140, { bold: opts.bold, size, align: "right" });
    this.y += h;
  }

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
  capitalTotal: number; // collapsed single Capital + Reserves figure
  netProfit: number;
}

function renderCover(doc: PDFKit.PDFDocument, rc: RenderCtx) {
  const { payload, unitHeading } = rc;
  doc.y = 220;
  doc.font("Helvetica-Bold").fontSize(22).fillColor("#111827").text(payload.meta.firmName || "FIRM NAME", MARGIN, doc.y, {
    width: PORTRAIT.w - MARGIN * 2,
    align: "center",
  });
  doc.moveDown(0.8);
  doc.font("Helvetica").fontSize(14).fillColor("#374151").text("FINANCIAL STATEMENTS", MARGIN, doc.y, {
    width: PORTRAIT.w - MARGIN * 2,
    align: "center",
  });
  doc.moveDown(0.6);
  doc.fontSize(11).text(payload.meta.periodLabel || "", MARGIN, doc.y, { width: PORTRAIT.w - MARGIN * 2, align: "center" });
  if (unitHeading) {
    doc.moveDown(1.2);
    doc.font("Helvetica-Oblique").fontSize(10).fillColor("#6b7280").text(unitHeading, MARGIN, doc.y, {
      width: PORTRAIT.w - MARGIN * 2,
      align: "center",
    });
  }
}

function renderFirmCapBs(doc: PDFKit.PDFDocument, c: Cursor, rc: RenderCtx) {
  const { st, divisor, unitHeading, payload, capitalTotal, netProfit } = rc;
  c.newPage("portrait");
  c.heading(payload.meta.firmName || "FIRM NAME", { size: 13 });
  c.heading("BUSINESS CAPITAL ACCOUNT", { size: 11 });
  if (unitHeading) c.heading(unitHeading, { size: 9, color: "#6b7280" });
  c.gap(6);

  c.line("Balance b/d (Capital + Reserves per trial balance)", round2(capitalTotal - netProfit), divisor, { indent: 14 });
  c.line("(+) Net Profit", netProfit, divisor, { indent: 14 });
  for (const label of ["Gifts Received", "Drawings", "Self-Assessment Tax", "Personal Insurance Premium"]) {
    c.line(`(enter manually) ${label}`, 0, divisor, { indent: 14 });
  }
  c.line("Balance c/d (Closing Capital)", capitalTotal, divisor, { bold: true });
  c.rule();
  c.gap(10);

  c.heading(`BALANCE SHEET AS AT ${(payload.meta.asAtLabel || "").replace(/^as at\s*/i, "")}`.trim(), { size: 11 });
  c.gap(6);
  c.line("LIABILITIES", null, divisor, { bold: true });
  c.line("Capital", capitalTotal, divisor, { indent: 14 });
  const secured = round2(st.total("SHORT_TERM_BORROW"));
  if (secured) c.line("Secured Loans", secured, divisor, { indent: 14 });
  const unsecured = round2(st.total("LONG_TERM_BORROW"));
  c.line("Unsecured Loans (Sch 1)", unsecured, divisor, { indent: 14 });
  const creditors = round2(st.total("TRADE_PAYABLES"));
  c.line("Sundry Creditors (Sch 2)", creditors, divisor, { indent: 14 });
  const otherLiab = round2(st.total("OTHER_CURR_LIAB") + st.total("SHORT_TERM_PROV") + st.total("DEFERRED_TAX_LIAB"));
  c.line("Other Current Liabilities (Sch 3)", otherLiab, divisor, { indent: 14 });
  const totalLiabilities = round2(capitalTotal + secured + unsecured + creditors + otherLiab);
  c.line("TOTAL LIABILITIES", totalLiabilities, divisor, { bold: true });
  c.rule();
  c.gap(10);

  c.line("ASSETS", null, divisor, { bold: true });
  c.line("Fixed Assets", round2(st.total("PPE")), divisor, { indent: 14 });
  const investments = round2(st.total("NONCURR_INVEST"));
  if (investments) c.line("Investments and Advances", investments, divisor, { indent: 14 });
  const loansAdv = round2(st.total("SHORT_TERM_LOANS_ADV") + st.total("LONG_TERM_LOANS_ADV"));
  c.line("Loans & Advances (Sch 4)", loansAdv, divisor, { indent: 14 });
  const debtors = round2(st.total("TRADE_RECV"));
  c.line("Sundry Debtors (Sch 6)", debtors, divisor, { indent: 14 });
  const otherAssets = round2(st.total("INVENTORY") + st.total("OTHER_CURR_ASSETS") + st.total("DEFERRED_TAX_ASSET"));
  c.line("Other Current Assets (Sch 5)", otherAssets, divisor, { indent: 14 });
  c.line("Advance Tax (enter manually)", 0, divisor, { indent: 14 });
  const cash = round2(st.total("CASH_BANK"));
  c.line("Cash and Bank Balances", cash, divisor, { indent: 14 });
  const totalAssets = round2(st.total("PPE") + investments + loansAdv + debtors + otherAssets + cash);
  c.line("TOTAL ASSETS", totalAssets, divisor, { bold: true });
  c.rule();
  c.gap(6);
  c.note(`Balancing check (Total Assets - Total Liabilities, should be 0): ${money(round2(totalAssets - totalLiabilities), divisor)}`);
}

function renderTradingPL(doc: PDFKit.PDFDocument, c: Cursor, rc: RenderCtx) {
  const { st, divisor, unitHeading, payload } = rc;
  c.newPage("portrait");
  c.heading(payload.meta.firmName || "FIRM NAME", { size: 13 });
  c.heading("TRADING, PROFIT AND LOSS A/C FOR THE YEAR ENDED", { size: 11 });
  if (payload.meta.periodLabel) c.heading(payload.meta.periodLabel, { size: 9.5, color: "#374151" });
  if (unitHeading) c.heading(unitHeading, { size: 9, color: "#6b7280" });
  c.gap(6);

  c.line("Opening Stock", round2(st.total("OPENING_STOCK")), divisor, { indent: 14 });
  c.line("Purchases / Direct Expenses", round2(st.total("PURCHASES") + st.total("DIRECT_EXP")), divisor, { indent: 14 });
  c.line("Sales / Revenue from Operations", st.revenue, divisor, { indent: 14 });
  c.line("Closing Stock", round2(st.total("CLOSING_STOCK_PL")), divisor, { indent: 14 });
  c.line("Gross Profit c/d", st.grossProfit, divisor, { bold: true });
  c.rule();
  c.gap(8);

  c.line("Indirect Expenses", null, divisor, { bold: true });
  for (const code of ["EMP_BENEFIT", "FINANCE_COST", "DEPRECIATION", "OTHER_EXP"] as const) {
    for (const item of st.byCode.get(code) ?? []) c.line(item.name, item.amount, divisor, { indent: 14 });
  }
  c.line("Indirect Income (incl. Gross Profit b/d)", round2(st.grossProfit + st.otherIncome), divisor, { bold: true });
  c.line("Net Profit c/d", st.netProfit, divisor, { bold: true });
  c.rule();
}

function renderSchedulesAndFA(doc: PDFKit.PDFDocument, c: Cursor, rc: RenderCtx) {
  const { st, divisor, payload } = rc;
  c.newPage("portrait");
  c.heading(payload.meta.firmName || "FIRM NAME", { size: 12 });
  c.heading("FIRM SCHEDULES", { size: 11 });
  c.gap(6);

  const schedule = (title: string, codes: Parameters<Statements["total"]>[0][]) => {
    c.line(title, null, divisor, { bold: true });
    let any = false;
    for (const code of codes) {
      for (const item of st.byCode.get(code) ?? []) {
        c.line(item.name, item.amount, divisor, { indent: 14 });
        any = true;
      }
    }
    if (!any) c.line("-", 0, divisor, { indent: 14 });
    c.gap(8);
  };

  schedule("Sch 1  UNSECURED LOANS", ["LONG_TERM_BORROW"]);
  schedule("Sch 2  SUNDRY CREDITORS", ["TRADE_PAYABLES"]);
  schedule("Sch 3  OTHER CURRENT LIABILITIES", ["OTHER_CURR_LIAB", "SHORT_TERM_PROV", "DEFERRED_TAX_LIAB"]);
  schedule("Sch 4  LOANS AND ADVANCES", ["SHORT_TERM_LOANS_ADV", "LONG_TERM_LOANS_ADV"]);
  schedule("Sch 5  OTHER CURRENT ASSETS", ["INVENTORY", "OTHER_CURR_ASSETS", "DEFERRED_TAX_ASSET"]);
  schedule("Sch 6  SUNDRY DEBTORS", ["TRADE_RECV"]);
  if ((st.byCode.get("NONCURR_INVEST") ?? []).length) schedule("INVESTMENTS AND ADVANCES", ["NONCURR_INVEST"]);

  c.newPage("landscape");
  c.heading(payload.meta.firmName || "FIRM NAME", { size: 12 });
  c.heading("FIRM FIXED ASSETS", { size: 11 });
  c.note("A trial balance gives closing net book value only - Rate/Opening/Additions/pre-post-3-Oct split default to 0 (enter manually). Land/shop-premises entries (0% rate) can't be distinguished from depreciable assets by name alone.");
  c.gap(6);
  const ppeItems = st.byCode.get("PPE") ?? [];
  const headers = ["Particulars", "Rate", "Opening", "Addn upto 3 Oct", "Addn after 3 Oct", "Dep/Round Off", "Closing"];
  const widths = [26, 8, 12, 14, 14, 12, 14];
  const rows: (string | number)[][] = ppeItems.map((a) => [a.name, 0, 0, 0, 0, 0, round2(a.amount)]);
  rows.push(["Depreciation per Trial Balance (unallocated)", 0, 0, 0, 0, round2(st.depreciation), 0]);
  rows.push(["Total", 0, 0, 0, 0, round2(st.depreciation), round2(ppeItems.reduce((s, a) => s + a.amount, 0))]);
  c.table(headers, widths, rows, { moneyCols: [1, 2, 3, 4, 5, 6], divisor: rc.divisor });
}

function renderSelf(doc: PDFKit.PDFDocument, c: Cursor, rc: RenderCtx) {
  const { divisor, payload, netProfit } = rc;
  c.newPage("portrait");
  c.heading(payload.meta.firmName || "FIRM NAME", { size: 12 });
  c.heading("PERSONAL (SELF) BALANCE SHEET", { size: 11 });
  c.note("Personal assets (gold, residential property, personal FDs, personal bank accounts) are categorically outside a business trial balance - only Net Profit genuinely flows through below; everything else is a labeled template for the CA to fill in.");
  c.gap(8);

  c.line("Capital Account", null, divisor, { bold: true });
  c.line("Opening Capital (enter manually)", 0, divisor, { indent: 14 });
  c.line("Net Profit (from business)", netProfit, divisor, { indent: 14 });
  for (const label of ["Bank Interest", "Interest on FD", "Dividend Income", "LIC Survival Benefit"]) {
    c.line(`${label} (enter manually)`, 0, divisor, { indent: 14 });
  }
  for (const label of ["Drawings", "Personal Insurance Premium"]) {
    c.line(`Less: ${label} (enter manually)`, 0, divisor, { indent: 14 });
  }
  c.line("Total Capital", netProfit, divisor, { bold: true });
  c.gap(10);

  c.line("Assets", null, divisor, { bold: true });
  for (const label of ["Gold Ornaments", "Residential Property", "Furniture", "Vehicle (personal)"]) {
    c.line(`${label} (enter manually)`, 0, divisor, { indent: 14 });
  }
  c.line("Investments and Advances (enter manually)", 0, divisor, { indent: 14 });
  c.line("Bank Balances - personal (enter manually)", 0, divisor, { indent: 14 });
  c.line("Cash Balance (enter manually)", 0, divisor, { indent: 14 });
}

function renderComp(doc: PDFKit.PDFDocument, c: Cursor, rc: RenderCtx) {
  const { divisor, payload, netProfit } = rc;
  c.newPage("portrait");
  c.heading(payload.meta.firmName || "FIRM NAME", { size: 12 });
  c.heading("STATEMENT SHOWING COMPUTATION OF TOTAL INCOME", { size: 11 });
  c.gap(6);

  c.line("Income from Business (Net Profit as per P&L A/c)", netProfit, divisor);
  c.line("Income from Other Sources - Bank Interest (enter manually)", 0, divisor);
  c.line("Income from Other Sources - Dividend Income (enter manually)", 0, divisor);
  const gti = round2(netProfit);
  c.line("Gross Total Income", gti, divisor, { bold: true });
  c.gap(6);

  c.line("Less: Chapter VIA Deductions", null, divisor, { bold: true });
  c.line("80C - LIC (enter manually, capped at Rs.1,50,000)", 0, divisor, { indent: 14 });
  c.line("80D - Mediclaim (enter manually, capped at Rs.25,000)", 0, divisor, { indent: 14 });
  c.line("80TTA - Savings Bank Interest (capped at Rs.10,000)", 0, divisor, { indent: 14 });
  const nti = Math.round(gti / 10) * 10;
  c.line("Net Total Income", nti, divisor, { bold: true });
  c.gap(6);

  const tax = round2(slab115bac(nti));
  c.line("Income Tax as per Section 115BAC", tax, divisor);
  const cess = round2(tax * TAX_RATE_CESS);
  c.line(`Add: Education Cess @ ${TAX_RATE_CESS * 100}%`, cess, divisor);
  const totalTax = round2(tax + cess);
  c.line("Total Tax Payable", totalTax, divisor, { bold: true });
  c.line("Less: Advance Tax (enter manually)", 0, divisor);
  c.line("Add: Interest u/s 234A/234B/234C (enter manually)", 0, divisor);
  c.line("Less: Self-Assessment Tax Paid (enter manually)", 0, divisor);
  c.line("Tax Payable / (Refund)", round2(totalTax), divisor, { bold: true });
}

export async function buildProprietorFullPdf(payload: ConvertPayload): Promise<Buffer> {
  const st = computeStatements(payload.ledgers);
  const figuresUnit = payload.meta.figuresUnit || "actual";
  const divisor = FIGURES_UNIT_DIVISOR[figuresUnit];
  const unitHeading = FIGURES_UNIT_HEADING[figuresUnit];
  const capitalTotal = round2(st.total("CAPITAL") + st.total("RESERVES") + st.netProfit);
  const rc: RenderCtx = { st, divisor, unitHeading, payload, capitalTotal, netProfit: st.netProfit };

  const doc = new PDFDocument({ size: "A4", margin: MARGIN, bufferPages: true, autoFirstPage: false });
  const chunks: Buffer[] = [];
  doc.on("data", (chunk: Buffer) => chunks.push(chunk));
  const done = new Promise<Buffer>((resolve) => doc.on("end", () => resolve(Buffer.concat(chunks))));

  doc.addPage();
  renderCover(doc, rc);

  const c = new Cursor(doc);
  renderFirmCapBs(doc, c, rc);
  renderTradingPL(doc, c, rc);
  renderSchedulesAndFA(doc, c, rc);
  renderSelf(doc, c, rc);
  renderComp(doc, c, rc);

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
