import PDFDocument from "pdfkit";
import type { ConvertPayload } from "../../types";
import { FIGURES_UNIT_DIVISOR, FIGURES_UNIT_HEADING } from "../../types";
import { computeStatements } from "../../statements";
import type { Statements } from "../../statements";
import { round2 } from "../../classify";
import { detectPartners } from "./partners";
import type { Partner } from "./partners";
import { computeAppropriation } from "./appropriationMath";
import type { AppropriationMath } from "./appropriationMath";

const MARGIN = 40;
const PORTRAIT = { w: 595.28, h: 841.89 };
const LANDSCAPE = { w: 841.89, h: 595.28 };
const TAX_RATE = 0.3;
const CESS_RATE = 0.04;

function money(value: number, divisor: number): string {
  const scaled = value / divisor;
  const text = Math.abs(scaled).toLocaleString("en-IN", {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  });
  return scaled < 0 ? `(${text})` : text;
}

// Same layout engine as scheduleiii/pdf.ts (kept self-contained rather than
// shared, matching that file's own precedent of not sharing pdfkit layout
// code with the "statutory" ../pdf.ts renderer).
class Cursor {
  y: number;
  page: { w: number; h: number } = PORTRAIT;
  private bottom = PORTRAIT.h - MARGIN - 24;
  private contentW = PORTRAIT.w - MARGIN * 2;

  constructor(private doc: PDFKit.PDFDocument) {
    this.y = doc.y;
  }

  get contentWidth() {
    return this.contentW;
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
  partners: Partner[];
  math: AppropriationMath;
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
  c.line("Indirect Income (incl. Gross Profit b/d)", round2(st.grossProfit + st.otherIncome), divisor, {
    bold: true,
  });
  c.line("Net Profit c/d", st.netProfit, divisor, { bold: true });
  c.rule();
}

function renderPlAppropriation(doc: PDFKit.PDFDocument, c: Cursor, rc: RenderCtx) {
  const { st, divisor, payload, partners, math } = rc;
  c.newPage("portrait");
  c.heading(payload.meta.firmName || "FIRM NAME", { size: 12 });
  c.heading("PROFIT & LOSS APPROPRIATION A/C FOR THE YEAR ENDED", { size: 11 });
  c.gap(6);

  c.line("Net Profit b/d", st.netProfit, divisor, { bold: true });
  c.gap(6);

  c.line("Less: Interest on Capital", null, divisor, { bold: true });
  for (const p of partners) c.line(p.name, 0, divisor, { indent: 14 });
  c.note("Interest on Capital needs month-by-month capital movement data not available from a trial balance - defaults to 0 (enter manually; see Int on Capital).");
  c.gap(6);

  c.line("Less: Remuneration as Salary (Section 40(b))", null, divisor, { bold: true });
  for (const pm of math.perPartner) c.line(pm.name, pm.remuneration, divisor, { indent: 14 });
  c.gap(6);

  c.line("Less: Provision for Tax", math.taxProvision, divisor, { bold: true });
  c.gap(6);

  c.line("Share in Net Profit", null, divisor, { bold: true });
  for (const pm of math.perPartner) c.line(pm.name, pm.shareOfProfit, divisor, { indent: 14 });
  c.rule();
}

function renderCapitalAccounts(doc: PDFKit.PDFDocument, c: Cursor, rc: RenderCtx) {
  const { divisor, payload, partners, math } = rc;
  c.newPage("portrait");
  c.heading(payload.meta.firmName || "FIRM NAME", { size: 12 });
  c.heading("PARTNER CAPITAL ACCOUNTS", { size: 11 });
  c.note("Opening balance is the trial balance's Capital ledger figure (reflects the year's Drawings/Capital Introduced already posted, per Tally). Closing balance adds this year's Interest/Remuneration/Share of Profit, computed only at year end.");
  c.gap(6);

  for (const p of partners) {
    const pm = math.perPartner.find((x) => x.name === p.name)!;
    const closing = round2(p.capital + pm.remuneration + pm.shareOfProfit);
    c.line(`${p.name} (${(p.sharePct * 100).toFixed(2)}%)`, null, divisor, { bold: true });
    c.line("Balance b/d (per trial balance)", p.capital, divisor, { indent: 14 });
    c.line("(+) Interest on Capital", pm.interest, divisor, { indent: 14 });
    c.line("(+) Remuneration as Salary", pm.remuneration, divisor, { indent: 14 });
    c.line("(+) Share in Net Profit", pm.shareOfProfit, divisor, { indent: 14 });
    c.line("Balance c/d (closing)", closing, divisor, { bold: true, indent: 14 });
    c.gap(8);
  }
}

function renderSchedules(doc: PDFKit.PDFDocument, c: Cursor, rc: RenderCtx) {
  const { st, divisor, payload } = rc;
  c.newPage("portrait");
  c.heading(payload.meta.firmName || "FIRM NAME", { size: 12 });
  c.heading("SCHEDULES", { size: 11 });
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

  schedule("SCH-1  SUNDRY CREDITORS", ["TRADE_PAYABLES"]);
  schedule("SCH-2  OTHER CURRENT LIABILITIES", ["OTHER_CURR_LIAB", "SHORT_TERM_PROV", "DEFERRED_TAX_LIAB"]);
  schedule("SCH-3  OTHER CURRENT ASSETS", ["INVENTORY", "OTHER_CURR_ASSETS", "DEFERRED_TAX_ASSET"]);
  schedule("SCH-4  LOANS AND ADVANCES", ["SHORT_TERM_LOANS_ADV", "LONG_TERM_LOANS_ADV"]);
  schedule("SCH-5  SUNDRY DEBTORS", ["TRADE_RECV"]);
  if ((st.byCode.get("NONCURR_INVEST") ?? []).length) schedule("INVESTMENTS", ["NONCURR_INVEST"]);
}

function renderFixedAssets(doc: PDFKit.PDFDocument, c: Cursor, rc: RenderCtx) {
  const { st, divisor, payload } = rc;
  c.newPage("landscape");
  c.heading(payload.meta.firmName || "FIRM NAME", { size: 12 });
  c.heading("SCH-5  FIXED ASSETS", { size: 11 });
  c.note("A trial balance gives closing net book value only - Rate/Opening/Additions/pre-post-3-Oct split default to 0 (enter manually).");
  c.gap(6);

  const ppeItems = st.byCode.get("PPE") ?? [];
  const headers = ["Particulars", "Rate", "Opening", "Addn upto 3 Oct", "Addn after 3 Oct", "Dep/Round Off", "Closing"];
  const widths = [26, 8, 12, 14, 14, 12, 14];
  const rows: (string | number)[][] = ppeItems.map((a) => [a.name, 0, 0, 0, 0, 0, round2(a.amount)]);
  rows.push(["Depreciation per Trial Balance (unallocated)", 0, 0, 0, 0, round2(st.depreciation), 0]);
  const closingTotal = round2(ppeItems.reduce((s, a) => s + a.amount, 0));
  rows.push(["Total", 0, 0, 0, 0, round2(st.depreciation), closingTotal]);
  c.table(headers, widths, rows, { moneyCols: [1, 2, 3, 4, 5, 6], divisor });
}

function renderBalanceSheet(doc: PDFKit.PDFDocument, c: Cursor, rc: RenderCtx) {
  const { st, divisor, unitHeading, payload, partners, math } = rc;
  c.newPage("portrait");
  c.heading(payload.meta.firmName || "FIRM NAME", { size: 13 });
  c.heading(`BALANCE SHEET AS AT ${(payload.meta.asAtLabel || "").replace(/^as at\s*/i, "")}`.trim(), { size: 11 });
  if (unitHeading) c.heading(unitHeading, { size: 9, color: "#6b7280" });
  c.gap(6);

  c.line("LIABILITIES", null, divisor, { bold: true });
  c.line("Capital Account", null, divisor, { bold: true, indent: 14 });
  let totalCapital = 0;
  for (const p of partners) {
    const pm = math.perPartner.find((x) => x.name === p.name)!;
    const closing = round2(p.capital + pm.remuneration + pm.shareOfProfit);
    totalCapital = round2(totalCapital + closing);
    c.line(p.name, closing, divisor, { indent: 28 });
  }
  const secured = round2(st.total("LONG_TERM_BORROW") + st.total("SHORT_TERM_BORROW"));
  c.line("Secured / Unsecured Loans", secured, divisor, { indent: 14 });
  const otherCurrLiab = round2(st.total("OTHER_CURR_LIAB") + st.total("SHORT_TERM_PROV") + st.total("DEFERRED_TAX_LIAB"));
  c.line("Sundry Creditors (Sch 1)", round2(st.total("TRADE_PAYABLES")), divisor, { indent: 14 });
  c.line("Other Current Liabilities (Sch 2)", otherCurrLiab, divisor, { indent: 14 });
  c.line("Provision for Taxation", math.taxProvision, divisor, { indent: 14 });
  const totalLiabilities = round2(totalCapital + secured + st.total("TRADE_PAYABLES") + otherCurrLiab + math.taxProvision);
  c.line("TOTAL LIABILITIES", totalLiabilities, divisor, { bold: true });
  c.rule();
  c.gap(10);

  c.line("ASSETS", null, divisor, { bold: true });
  c.line("Fixed Assets (Sch 5)", round2(st.total("PPE")), divisor, { indent: 14 });
  if ((st.byCode.get("NONCURR_INVEST") ?? []).length) {
    c.line("Investments", round2(st.total("NONCURR_INVEST")), divisor, { indent: 14 });
  }
  c.line("Loans and Advances (Sch 4)", round2(st.total("SHORT_TERM_LOANS_ADV") + st.total("LONG_TERM_LOANS_ADV")), divisor, { indent: 14 });
  c.line("Sundry Debtors (Sch 5)", round2(st.total("TRADE_RECV")), divisor, { indent: 14 });
  const otherCurrAssets = round2(st.total("INVENTORY") + st.total("OTHER_CURR_ASSETS") + st.total("DEFERRED_TAX_ASSET"));
  c.line("Other Current Assets (Sch 3)", otherCurrAssets, divisor, { indent: 14 });
  c.line("Cash and Bank Balance", round2(st.total("CASH_BANK")), divisor, { indent: 14 });
  const totalAssets = round2(st.totalAssets);
  c.line("TOTAL ASSETS", totalAssets, divisor, { bold: true });
  c.rule();
  c.gap(6);
  c.note(`Balancing check (Total Assets - Total Liabilities, should be 0): ${money(round2(totalAssets - totalLiabilities), divisor)}`);
}

function renderStat(doc: PDFKit.PDFDocument, c: Cursor, rc: RenderCtx) {
  const { st, divisor, payload, math } = rc;
  c.newPage("portrait");
  c.heading(payload.meta.firmName || "FIRM NAME", { size: 12 });
  c.heading("STATEMENT SHOWING COMPUTATION OF TOTAL INCOME", { size: 11 });
  c.gap(6);

  const totalNetProfit = round2(math.shareOfProfitTotal + 0);
  c.line("Net Profit as per Profit and Loss Appropriation A/c", math.shareOfProfitTotal, divisor);
  c.line("Add: Disallowed Expenses (enter manually)", 0, divisor);
  c.line("Total Net Profit", totalNetProfit, divisor, { bold: true });
  c.line("Less: Unabsorbed Losses b/d (enter manually)", 0, divisor);
  const gti = totalNetProfit;
  c.line("Gross Total Income", gti, divisor, { bold: true });
  c.line("Less: Deductions Under Chapter VIA (enter manually)", 0, divisor);
  const nti = Math.round(gti / 10) * 10;
  c.line("Net Total Income", nti, divisor, { bold: true });
  const incomeTax = round2(nti * TAX_RATE);
  c.line(`Income Tax @ ${TAX_RATE * 100}%`, incomeTax, divisor);
  const cess = Math.round(round2(incomeTax * CESS_RATE));
  c.line(`Add: Cess @ ${CESS_RATE * 100}%`, cess, divisor);
  const totalTax = Math.round(round2(incomeTax + cess));
  c.line("Total Tax Payable", totalTax, divisor, { bold: true });
  c.line("Less: Advance Tax (enter manually)", 0, divisor);
  const tdsItem = (st.byCode.get("OTHER_CURR_ASSETS") ?? []).find((i) => /tds/i.test(i.name));
  const tdsAmount = round2(tdsItem?.amount ?? 0);
  c.line(`Less: TDS${tdsItem ? ` (${tdsItem.name})` : " (enter manually)"}`, tdsAmount, divisor);
  c.line("Less: Self-Assessment Tax Paid (enter manually)", 0, divisor);
  c.line("Net Tax Payable / (Refund Due)", round2(totalTax - tdsAmount), divisor, { bold: true });
}

function renderDetailsAndIntOnCapital(doc: PDFKit.PDFDocument, c: Cursor, rc: RenderCtx) {
  const { payload, partners } = rc;
  c.newPage("portrait");
  c.heading(payload.meta.firmName || "FIRM NAME", { size: 12 });
  c.heading("DETAILS OF CLIENT", { size: 11 });
  c.note("PAN, address, bank details, TAN/PTEC/PTRC and IT/TDS file numbers are not in a trial balance - enter manually.");
  c.gap(10);

  c.heading("CALCULATION OF INTEREST ON CAPITAL", { size: 11 });
  c.note("Needs month-by-month capital movement data not available from a trial balance - enter manually per partner (see the Excel version's editable template).");
  for (const p of partners) c.line(p.name, 0, 1, { indent: 14 });
}

export async function buildPartnershipFullPdf(payload: ConvertPayload): Promise<Buffer> {
  const st = computeStatements(payload.ledgers);
  const figuresUnit = payload.meta.figuresUnit || "actual";
  const divisor = FIGURES_UNIT_DIVISOR[figuresUnit];
  const unitHeading = FIGURES_UNIT_HEADING[figuresUnit];
  const partners = detectPartners(st);
  const math = computeAppropriation(st.netProfit, partners);
  const rc: RenderCtx = { st, divisor, unitHeading, payload, partners, math };

  const doc = new PDFDocument({ size: "A4", margin: MARGIN, bufferPages: true, autoFirstPage: false });
  const chunks: Buffer[] = [];
  doc.on("data", (chunk: Buffer) => chunks.push(chunk));
  const done = new Promise<Buffer>((resolve) => doc.on("end", () => resolve(Buffer.concat(chunks))));

  doc.addPage();
  renderCover(doc, rc);

  const c = new Cursor(doc);
  renderBalanceSheet(doc, c, rc);
  renderTradingPL(doc, c, rc);
  renderPlAppropriation(doc, c, rc);
  renderCapitalAccounts(doc, c, rc);
  renderSchedules(doc, c, rc);
  renderFixedAssets(doc, c, rc);
  renderStat(doc, c, rc);
  renderDetailsAndIntOnCapital(doc, c, rc);

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
