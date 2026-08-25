import type ExcelJS from "exceljs";
import type { BuildCtx } from "../common";
import { setCell, setFormula, applyColumnWidths, topBorder } from "../../excel/helpers";
import type { Partner } from "./partners";
import { PROVISION_FOR_TAX_RATE, computeAppropriation } from "./appropriationMath";

export interface PartnerRowRefs {
  interestRow: number;
  remunerationRow: number;
  remunerationAmount: number; // known cached value at remunerationRow, in actual rupees (not display-scaled)
  shareOfProfitRow: number;
  shareOfProfitAmount: number; // known cached value at shareOfProfitRow, in actual rupees
}

export interface PlAppropriationResult {
  byPartner: Map<string, PartnerRowRefs>; // keyed by partner name (matches `partners` array order)
  shareOfProfitStart: number; // first "Share in Net Profit" row - for STAT's SUM range
  shareOfProfitEnd: number; // last "Share in Net Profit" row - for STAT's SUM range
  shareOfProfitTotal: number; // SUM of all partners' Share in Net Profit - for STAT, computed once here
  taxProvisionRow: number; // "To Provision for Tax" row - this is a real liability, must appear on the BS too
  taxProvisionAmount: number; // known cached value at taxProvisionRow, in actual rupees
}

// 'p&l app': Interest on Capital -> Section 40(b) Remuneration -> Provision
// for Tax -> Share in Net Profit, all partner-wise. The Section 40(b)
// formula (first Rs.3L of book profit: higher of Rs.1,50,000 or 90%;
// balance: 60%) and the flat-rate tax provision are exact statutory
// formulas needing only Net Profit and each partner's share % - both are
// live Excel formulas here, not JS-side approximations. Interest on Capital
// itself defaults to 0 (see interestOnCapital.ts - a trial balance has no
// month-by-month capital movement data to compute it from).
export function writePlAppropriation(
  ws: ExcelJS.Worksheet,
  ctx: BuildCtx,
  partners: Partner[],
  netProfitFormulaRef: string, // e.g. "'P&L'!B36" - the Trading & P&L's Net Profit c/d cell
): PlAppropriationResult {
  applyColumnWidths(ws, [34, 16, 4, 24, 16]);
  const { meta, st, divisor } = ctx;
  const money = (v: number) => v / divisor;
  const netProfit = st.netProfit;
  const math = computeAppropriation(netProfit, partners);
  const mathByPartner = new Map(math.perPartner.map((p) => [p.name, p]));

  let r = 2;
  setCell(ws, r++, 1, meta.firmName || "FIRM NAME", { bold: true });
  setCell(ws, r++, 1, meta.periodLabel || "", { italic: true });
  if (ctx.unitHeading) setCell(ws, r++, 1, ctx.unitHeading, { italic: true });
  setCell(ws, r++, 1, "PROFIT & LOSS APPROPRIATION A/C FOR THE YEAR ENDED", { bold: true });
  r++;
  setCell(ws, r, 1, "PARTICULARS", { bold: true });
  setCell(ws, r, 2, "AMOUNT", { bold: true, align: "right" });
  setCell(ws, r, 4, "PARTICULARS", { bold: true });
  setCell(ws, r++, 5, "AMOUNT", { bold: true, align: "right" });
  r++;

  setCell(ws, r, 4, "By Net Profit b/d");
  const netProfitRow = r;
  setFormula(ws, r++, 5, netProfitFormulaRef, money(netProfit), { money: true });
  r++;

  setCell(ws, r++, 1, "To Interest on Capital", { bold: true, underline: true });
  const interestStart = r;
  for (const p of partners) {
    setCell(ws, r, 1, `      ${p.name}`);
    setCell(ws, r++, 2, 0, { money: true }); // see interestOnCapital.ts - not derivable from a TB
  }
  const interestEnd = r - 1;
  const interestRows = new Map<string, number>();
  partners.forEach((p, i) => interestRows.set(p.name, interestStart + i));
  r++;

  setCell(ws, r++, 1, "To Remuneration as Salary", { bold: true, underline: true });
  const bookProfitFormula = `E${netProfitRow}-SUM(B${interestStart}:B${interestEnd})`;
  const remunerationTotalFormula =
    `IF(${bookProfitFormula}<=166667,MIN(${bookProfitFormula},150000),` +
    `IF(${bookProfitFormula}<=300000,(${bookProfitFormula})*90%,270000+(${bookProfitFormula}-300000)*60%))`;
  const remunerationRows = new Map<string, number>();
  const remunerationAmounts = new Map<string, number>();
  for (const p of partners) {
    const amount = mathByPartner.get(p.name)!.remuneration;
    setCell(ws, r, 1, `      ${p.name}`);
    setFormula(ws, r, 2, `(${remunerationTotalFormula})*${p.sharePct}`, money(amount), {
      money: true,
    });
    remunerationRows.set(p.name, r);
    remunerationAmounts.set(p.name, amount);
    r++;
  }
  const remunerationEnd = r - 1;
  r++;

  setCell(ws, r, 1, "To Provision for Tax", { bold: true });
  const taxRow = r;
  const afterInterestRemunFormula = `E${netProfitRow}-SUM(B${interestStart}:B${remunerationEnd})`;
  setFormula(
    ws,
    r++,
    2,
    `ROUND((${afterInterestRemunFormula})*${PROVISION_FOR_TAX_RATE},-1)`,
    math.taxProvision,
    { money: true },
  );
  r++;

  setCell(ws, r++, 1, "To Share In Net Profit", { bold: true, underline: true });
  const shareStart = r;
  const shareFormula = `E${netProfitRow}-SUM(B${interestStart}:B${taxRow})`;
  const shareRows = new Map<string, number>();
  const shareAmounts = new Map<string, number>();
  for (const p of partners) {
    const amount = mathByPartner.get(p.name)!.shareOfProfit;
    setCell(ws, r, 1, `      ${p.name}`);
    setFormula(ws, r, 2, `(${shareFormula})*${p.sharePct}`, money(amount), {
      money: true,
    });
    shareRows.set(p.name, r);
    shareAmounts.set(p.name, amount);
    r++;
  }
  const shareEnd = r - 1;
  r++;

  setCell(ws, r, 1, "TOTAL", { bold: true });
  const totalDebitRow = r;
  setFormula(ws, r, 2, `SUM(B${interestStart}:B${shareEnd})`, money(netProfit), { bold: true, money: true });
  setCell(ws, r, 4, "TOTAL", { bold: true });
  setFormula(ws, r, 5, `E${netProfitRow}`, money(netProfit), { bold: true, money: true });
  topBorder(ws, r, 1, 5);
  void totalDebitRow;

  const byPartner = new Map<string, PartnerRowRefs>();
  for (const p of partners) {
    byPartner.set(p.name, {
      interestRow: interestRows.get(p.name)!,
      remunerationRow: remunerationRows.get(p.name)!,
      remunerationAmount: remunerationAmounts.get(p.name)!,
      shareOfProfitRow: shareRows.get(p.name)!,
      shareOfProfitAmount: shareAmounts.get(p.name)!,
    });
  }

  return {
    byPartner,
    shareOfProfitStart: shareStart,
    shareOfProfitEnd: shareEnd,
    shareOfProfitTotal: math.shareOfProfitTotal,
    taxProvisionRow: taxRow,
    taxProvisionAmount: math.taxProvision,
  };
}
