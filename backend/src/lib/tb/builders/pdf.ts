import PDFDocument from "pdfkit";
import type { ConvertPayload, StatementCode, StatementMeta } from "../types";
import { ENTITY_LABELS, FIGURES_UNIT_DIVISOR, FIGURES_UNIT_HEADING, entityFamily } from "../types";
import { computeStatements } from "../statements";
import type { Statements, LineItem } from "../statements";

const MARGIN = 40;
const PAGE_W = 595.28; // A4, points
const PAGE_H = 841.89;
const CONTENT_W = PAGE_W - MARGIN * 2;
const LABEL_W = 350;
const AMOUNT_W = CONTENT_W - LABEL_W;
const BOTTOM = PAGE_H - MARGIN - 24; // leave room for the page-number footer

function round2(v: number): number {
  return Math.round(v * 100) / 100;
}

function money(value: number, divisor: number): string {
  const scaled = value / divisor;
  const text = Math.abs(scaled).toLocaleString("en-IN", {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  });
  return scaled < 0 ? `(${text})` : text;
}

function items(st: Statements, ...codes: StatementCode[]): LineItem[] {
  const out: LineItem[] = [];
  for (const c of codes) out.push(...(st.byCode.get(c) ?? []));
  return out;
}

// Manual layout cursor: pdfkit does not auto-paginate absolutely-positioned
// text/drawing calls, so every row tracks its own y and page-breaks itself.
class Cursor {
  y: number;
  constructor(private doc: PDFKit.PDFDocument) {
    this.y = doc.y;
  }

  space(needed: number) {
    if (this.y + needed > BOTTOM) {
      this.doc.addPage();
      this.y = MARGIN;
    }
  }

  text(
    str: string,
    x: number,
    width: number,
    opts: {
      bold?: boolean;
      italic?: boolean;
      size?: number;
      align?: "left" | "right" | "center";
      color?: string;
    } = {},
  ) {
    const font = opts.bold ? "Helvetica-Bold" : opts.italic ? "Helvetica-Oblique" : "Helvetica";
    this.doc
      .font(font)
      .fontSize(opts.size ?? 9.5)
      .fillColor(opts.color ?? "#111827")
      .text(str, x, this.y, { width, align: opts.align ?? "left" });
  }

  row(
    label: string,
    amount: number | null,
    divisor: number,
    opts: { bold?: boolean; indent?: number; size?: number } = {},
  ) {
    const size = opts.size ?? 9.5;
    const indent = opts.indent ?? 0;
    const w = LABEL_W - indent;
    this.doc.font(opts.bold ? "Helvetica-Bold" : "Helvetica").fontSize(size);
    const h = Math.max(14, this.doc.heightOfString(label, { width: w }) + 4);
    this.space(h);
    this.text(label, MARGIN + indent, w, { bold: opts.bold, size });
    if (amount !== null) {
      this.text(money(amount, divisor), MARGIN + LABEL_W, AMOUNT_W, {
        bold: opts.bold,
        size,
        align: "right",
      });
    }
    this.y += h;
  }

  note(text: string, opts: { size?: number; italic?: boolean } = {}) {
    const size = opts.size ?? 8.5;
    this.doc.font(opts.italic ? "Helvetica-Oblique" : "Helvetica").fontSize(size);
    const h = Math.max(12, this.doc.heightOfString(text, { width: CONTENT_W }) + 4);
    this.space(h);
    this.doc.text(text, MARGIN, this.y, { width: CONTENT_W });
    this.y += h;
  }

  rule() {
    this.space(6);
    this.doc
      .moveTo(MARGIN, this.y)
      .lineTo(MARGIN + CONTENT_W, this.y)
      .lineWidth(0.75)
      .strokeColor("#9ca3af")
      .stroke();
    this.y += 6;
  }

  gap(h = 8) {
    this.y += h;
  }

  heading(text: string, opts: { size?: number; color?: string } = {}) {
    this.space((opts.size ?? 12) + 6);
    this.text(text, MARGIN, CONTENT_W, { bold: true, size: opts.size ?? 12, color: opts.color });
    this.y += (opts.size ?? 12) + 6;
  }
}

function section(c: Cursor, title: string, total: number, list: LineItem[], divisor: number) {
  c.row(title, total, divisor, { bold: true });
  for (const it of list) {
    c.row(it.name, it.amount, divisor, { indent: 14 });
  }
  c.gap(4);
}

interface RenderCtx {
  st: Statements;
  meta: StatementMeta;
  divisor: number;
  unitHeading: string | null;
}

function renderCover(
  doc: PDFKit.PDFDocument,
  meta: StatementMeta,
  entityLabel: string,
  unitHeading: string | null,
) {
  doc.y = 220;
  doc
    .font("Helvetica-Bold")
    .fontSize(22)
    .fillColor("#111827")
    .text(meta.firmName || "FIRM NAME", MARGIN, doc.y, { width: CONTENT_W, align: "center" });
  doc.moveDown(1);
  doc
    .font("Helvetica")
    .fontSize(14)
    .fillColor("#374151")
    .text("FINANCIAL STATEMENTS", MARGIN, doc.y, { width: CONTENT_W, align: "center" });
  doc.moveDown(0.6);
  doc.fontSize(11).text(meta.periodLabel || "", MARGIN, doc.y, { width: CONTENT_W, align: "center" });
  if (meta.asAtLabel) {
    doc.moveDown(0.3);
    doc.fontSize(10).text(meta.asAtLabel, MARGIN, doc.y, { width: CONTENT_W, align: "center" });
  }
  doc.moveDown(0.3);
  doc.fontSize(10).text(entityLabel, MARGIN, doc.y, { width: CONTENT_W, align: "center" });
  if (unitHeading) {
    doc.moveDown(1.2);
    doc
      .font("Helvetica-Oblique")
      .fontSize(10)
      .fillColor("#6b7280")
      .text(unitHeading, MARGIN, doc.y, { width: CONTENT_W, align: "center" });
  }
}

function renderBalanceSheet(doc: PDFKit.PDFDocument, c: Cursor, ctx: RenderCtx) {
  const { st, meta, divisor, unitHeading } = ctx;
  doc.addPage();
  c.y = MARGIN;

  c.heading(meta.firmName || "FIRM NAME", { size: 13 });
  c.heading(`BALANCE SHEET ${(meta.asAtLabel || "").toUpperCase()}`.trim(), { size: 11 });
  if (unitHeading) c.heading(unitHeading, { size: 9, color: "#6b7280" });
  c.gap(6);

  c.row("Liabilities", null, divisor, { bold: true, size: 10.5 });
  c.gap(2);
  const capList = [
    ...items(st, "CAPITAL", "RESERVES"),
    { name: "Net Profit", amount: st.netProfit },
  ];
  section(
    c,
    "Capital Account",
    round2(st.totalMany(["CAPITAL", "RESERVES"]) + st.netProfit),
    capList,
    divisor,
  );
  section(
    c,
    "Loans (Liability)",
    st.totalMany(["LONG_TERM_BORROW", "SHORT_TERM_BORROW"]),
    items(st, "LONG_TERM_BORROW", "SHORT_TERM_BORROW"),
    divisor,
  );
  section(
    c,
    "Current Liabilities",
    st.totalMany(["TRADE_PAYABLES", "OTHER_CURR_LIAB", "SHORT_TERM_PROV", "DEFERRED_TAX_LIAB"]),
    items(st, "TRADE_PAYABLES", "OTHER_CURR_LIAB", "SHORT_TERM_PROV", "DEFERRED_TAX_LIAB"),
    divisor,
  );
  c.row("Total Liabilities", st.totalLiabilities, divisor, { bold: true });
  c.rule();
  c.gap(12);

  c.row("Assets", null, divisor, { bold: true, size: 10.5 });
  c.gap(2);
  section(c, "Fixed Assets", st.total("PPE"), items(st, "PPE"), divisor);
  section(
    c,
    "Investments",
    st.totalMany(["NONCURR_INVEST", "LONG_TERM_LOANS_ADV", "DEFERRED_TAX_ASSET"]),
    items(st, "NONCURR_INVEST", "LONG_TERM_LOANS_ADV", "DEFERRED_TAX_ASSET"),
    divisor,
  );
  section(
    c,
    "Current Assets",
    st.totalMany([
      "INVENTORY",
      "TRADE_RECV",
      "CASH_BANK",
      "SHORT_TERM_LOANS_ADV",
      "OTHER_CURR_ASSETS",
    ]),
    items(st, "INVENTORY", "TRADE_RECV", "CASH_BANK", "SHORT_TERM_LOANS_ADV", "OTHER_CURR_ASSETS"),
    divisor,
  );
  c.row("Total Assets", st.totalAssets, divisor, { bold: true });
  c.rule();

  // Deliberately shown in actual rupees regardless of the selected display
  // unit, so a real trial-balance mismatch is never hidden by scaling.
  if (Math.abs(st.balanceDifference) > 0.5) {
    c.gap(8);
    c.note(
      `Note: Assets - Liabilities difference = ${st.balanceDifference.toFixed(2)} (review mapping / enter missing items).`,
      { italic: true },
    );
  }
}

function renderProfitAndLoss(doc: PDFKit.PDFDocument, c: Cursor, ctx: RenderCtx) {
  const { st, meta, divisor, unitHeading } = ctx;
  doc.addPage();
  c.y = MARGIN;

  c.heading(meta.firmName || "FIRM NAME", { size: 13 });
  c.heading("PROFIT & LOSS ACCOUNT", { size: 11 });
  if (meta.periodLabel) c.heading(meta.periodLabel, { size: 9.5, color: "#374151" });
  if (unitHeading) c.heading(unitHeading, { size: 9, color: "#6b7280" });
  c.gap(6);

  c.row("Trading Account", null, divisor, { bold: true, size: 10.5 });
  c.gap(2);
  section(c, "Sales Accounts", st.revenue, items(st, "REV_OPS"), divisor);
  const costTotal = round2(
    st.total("PURCHASES") + st.total("OPENING_STOCK") - st.total("CLOSING_STOCK_PL"),
  );
  section(
    c,
    "Cost of Sales",
    costTotal,
    items(st, "OPENING_STOCK", "PURCHASES", "CLOSING_STOCK_PL"),
    divisor,
  );
  const directList = items(st, "DIRECT_EXP");
  if (directList.length) section(c, "Direct Expenses", st.total("DIRECT_EXP"), directList, divisor);
  c.row("Gross Profit", st.grossProfit, divisor, { bold: true });
  c.rule();
  c.gap(12);

  c.row("Income Statement", null, divisor, { bold: true, size: 10.5 });
  c.gap(2);
  section(c, "Indirect Incomes", st.otherIncome, items(st, "OTHER_INCOME"), divisor);
  section(
    c,
    "Indirect Expenses",
    st.totalIndirectExpense,
    items(st, "EMP_BENEFIT", "FINANCE_COST", "DEPRECIATION", "OTHER_EXP"),
    divisor,
  );
  c.row("Net Profit", st.netProfit, divisor, { bold: true });
  c.rule();
}

function renderSignatureBlock(
  c: Cursor,
  meta: StatementMeta,
  forFirmText: string,
  designation: string,
) {
  c.gap(20);
  const leftX = MARGIN;
  const leftW = 280;
  const rightX = MARGIN + 300;
  const rightW = CONTENT_W - 300;

  const line = (left: string, right?: string, opts: { bold?: boolean } = {}) => {
    c.space(14);
    if (left) c.text(left, leftX, leftW, { size: 9 });
    if (right) c.text(right, rightX, rightW, { size: 9, bold: opts.bold });
    c.y += 14;
  };

  line(`Place : ${meta.place || "Vasai"}`, forFirmText, { bold: true });
  line(`Date : ${meta.date || ""}`);
  line(`UDIN : ${meta.udin || ""}`);
  c.gap(6);
  line("", designation);
  line("", "As per our report on even date");
  line("", `For ${meta.caName || "Namrata Prakash Sharma"}`);
  line("", "(Chartered Accountants)");
  c.gap(6);
  line("", `Proprietor : CA ${meta.caName || "Namrata Prakash Sharma"}`);
  line("", `M No : ${meta.caMembershipNo || "177309"}`);
  line("", `FRN No : ${meta.caFirmRegNo || "144860W"}`);
}

export async function buildFinancialStatementsPdf(payload: ConvertPayload): Promise<Buffer> {
  // Pvt Ltd companies always get the Schedule III PDF, Partnership always
  // gets the partner-wise PDF, LLP always gets its own PDF (Capital/Current
  // split, Sources/Application BS, FORM 8), Proprietor always gets the
  // Business + Personal split PDF - matching the invariants in
  // builders/index.ts.
  const family = entityFamily(payload.entity);
  if (family === "corporate") {
    const { buildScheduleIIIPdf } = await import("./scheduleiii/pdf");
    return buildScheduleIIIPdf(payload);
  }
  if (family === "noncorporate") {
    const { buildPartnershipFullPdf } = await import("./partnershipFull/pdf");
    return buildPartnershipFullPdf(payload);
  }
  if (family === "proprietor") {
    const { buildProprietorFullPdf } = await import("./proprietorFull/pdf");
    return buildProprietorFullPdf(payload);
  }
  if (family === "llp") {
    const { buildLlpFullPdf } = await import("./llpFull/pdf");
    return buildLlpFullPdf(payload);
  }

  // Statements are always computed on actual (unscaled) figures; the unit
  // only affects how amounts are written into the document below.
  const st = computeStatements(payload.ledgers);
  const designation =
    payload.entity === "proprietor" ? "Proprietor" : payload.entity === "llp" ? "Designated Partner" : "Partner";
  const figuresUnit = payload.meta.figuresUnit || "actual";
  const divisor = FIGURES_UNIT_DIVISOR[figuresUnit];
  const unitHeading = FIGURES_UNIT_HEADING[figuresUnit];
  const meta = payload.meta;

  const doc = new PDFDocument({ size: "A4", margin: MARGIN, bufferPages: true, autoFirstPage: false });
  const chunks: Buffer[] = [];
  doc.on("data", (chunk: Buffer) => chunks.push(chunk));
  const done = new Promise<Buffer>((resolve) => {
    doc.on("end", () => resolve(Buffer.concat(chunks)));
  });

  doc.addPage();
  renderCover(doc, meta, ENTITY_LABELS[payload.entity], unitHeading);

  const c = new Cursor(doc);
  const ctx: RenderCtx = { st, meta, divisor, unitHeading };
  renderBalanceSheet(doc, c, ctx);
  renderProfitAndLoss(doc, c, ctx);
  renderSignatureBlock(c, meta, `For ${meta.firmName || "the entity"}`, designation);

  // "Page X of Y" needs the total page count, which is only known once every
  // page has been laid out; bufferPages lets us go back and stamp footers.
  // The footer sits inside the page's bottom margin, which would otherwise
  // make pdfkit's auto-pagination insert a *new* trailing page for it on
  // every iteration; zeroing the margin for this one write avoids that.
  const range = doc.bufferedPageRange();
  for (let i = range.start; i < range.start + range.count; i++) {
    doc.switchToPage(i);
    const bottomMargin = doc.page.margins.bottom;
    doc.page.margins.bottom = 0;
    doc
      .font("Helvetica")
      .fontSize(8)
      .fillColor("#6b7280")
      .text(`Page ${i - range.start + 1} of ${range.count}`, MARGIN, PAGE_H - 30, {
        width: CONTENT_W,
        align: "right",
        lineBreak: false,
      });
    doc.page.margins.bottom = bottomMargin;
  }

  doc.end();
  return done;
}
