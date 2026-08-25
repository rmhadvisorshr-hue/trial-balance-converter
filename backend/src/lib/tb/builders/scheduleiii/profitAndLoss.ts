import type ExcelJS from "exceljs";
import type { BuildCtx } from "../common";
import { setCell, setFormula, applyColumnWidths, topBorder, signatureBlock } from "../../excel/helpers";
import { round2 } from "../../classify";
import type { NotesResult } from "./notes";
import type { FixedAssetChainResult } from "./fixedAssetChain";

export interface ProfitAndLossResult {
  netProfitRow: number; // for Note 2,3-SC's Reserves link
}

export function writeScheduleIIIProfitAndLoss(
  ws: ExcelJS.Worksheet,
  ctx: BuildCtx,
  notes: NotesResult,
  fa: FixedAssetChainResult,
  dtaNetExpenseRow: number,
): ProfitAndLossResult {
  applyColumnWidths(ws, [6, 42, 8, 20, 20]);
  const { meta, st, divisor } = ctx;
  const money = (v: number) => v / divisor;
  const noteRow = (n: number) => notes.byNote.get(n)?.row;

  let r = 2;
  setCell(ws, r++, 1, meta.firmName || "COMPANY NAME", { bold: true });
  if (meta.cin) setCell(ws, r++, 1, `CIN : ${meta.cin}`, {});
  setCell(ws, r++, 1, "PROFIT & LOSS STATEMENT FOR THE YEAR ENDED", { bold: true });
  setCell(ws, r++, 1, meta.periodLabel || "", { italic: true });
  if (ctx.unitHeading) setCell(ws, r++, 1, ctx.unitHeading, { italic: true });
  r++;
  setCell(ws, r, 2, "Particulars", { bold: true });
  setCell(ws, r, 3, "Notes", { bold: true, align: "center" });
  setCell(ws, r++, 4, meta.periodLabel || "", { bold: true, align: "right" });
  r++;

  const line = (
    roman: string,
    label: string,
    noteNo: number | null,
    formula: string | null,
    known: number,
    bold = false,
  ) => {
    if (roman) setCell(ws, r, 1, roman, { bold: true });
    setCell(ws, r, 2, label, { bold });
    if (noteNo) setCell(ws, r, 3, noteNo, { align: "center" });
    if (formula) setFormula(ws, r, 4, formula, money(known), { bold, money: true });
    r++;
  };

  const revNote = noteRow(19);
  const otherIncNote = noteRow(20);
  line("I", "Revenue from operations", 19, revNote ? `NOTES!B${revNote}` : null, st.revenue);
  line("II", "Other income", 20, otherIncNote ? `NOTES!B${otherIncNote}` : null, st.otherIncome);
  const totalIncomeRow = r;
  const totalIncome = round2(st.revenue + st.otherIncome);
  setCell(ws, r, 1, "III", { bold: true });
  setCell(ws, r, 2, "Total Income (I + II)", { bold: true });
  setFormula(ws, r, 4, `SUM(D${totalIncomeRow - 2}:D${totalIncomeRow - 1})`, money(totalIncome), {
    bold: true,
    money: true,
  });
  r += 2;

  setCell(ws, r, 1, "IV", { bold: true });
  setCell(ws, r++, 2, "Expenses:", { bold: true });
  const expenseStart = r;
  const costMaterials = round2(st.total("PURCHASES") + st.total("DIRECT_EXP"));
  const changeInv = round2(st.total("OPENING_STOCK") - st.total("CLOSING_STOCK_PL"));
  const n21 = noteRow(21);
  const n22 = noteRow(22);
  const n23 = noteRow(23);
  const n24 = noteRow(24);
  const n25 = noteRow(25);
  line("", "Cost of materials consumed / Purchases", 21, n21 ? `NOTES!B${n21}` : null, costMaterials);
  line("", "Changes in inventory", 22, n22 ? `NOTES!B${n22}` : null, changeInv);
  line("", "Employee benefit expenses", 23, n23 ? `NOTES!B${n23}` : null, st.employeeBenefit);
  line("", "Finance costs", 24, n24 ? `NOTES!B${n24}` : null, st.financeCost);
  line("", "Depreciation and amortisation", 11, `'11. FA'!G${fa.faDepForYearRow}`, st.depreciation);
  line("", "Other expenses", 25, n25 ? `NOTES!B${n25}` : null, st.otherExpense);
  const expenseEnd = r - 1;

  setCell(ws, r, 1, "V", { bold: true });
  setCell(ws, r, 2, "Total expenses", { bold: true });
  const totalExpense = round2(
    costMaterials + changeInv + st.employeeBenefit + st.financeCost + st.depreciation + st.otherExpense,
  );
  setFormula(ws, r, 4, `SUM(D${expenseStart}:D${expenseEnd})`, money(totalExpense), {
    bold: true,
    money: true,
  });
  const totalExpenseRow = r;
  r += 2;

  setCell(ws, r, 1, "VI", { bold: true });
  setCell(ws, r, 2, "Profit before tax (III - V)", { bold: true });
  const pbt = round2(totalIncome - totalExpense);
  setFormula(ws, r, 4, `D${totalIncomeRow}-D${totalExpenseRow}`, money(pbt), { bold: true, money: true });
  const pbtRow = r;
  r += 2;

  setCell(ws, r, 1, "VII", { bold: true });
  setCell(ws, r++, 2, "Tax expense:", { bold: true });
  const taxStart = r;
  setCell(ws, r, 2, "(1) Current tax");
  setCell(ws, r++, 4, 0, { money: true });
  setCell(ws, r, 2, "(2) Deferred tax");
  setFormula(ws, r++, 4, `DTA!C${dtaNetExpenseRow}`, 0, { money: true });
  setCell(ws, r, 2, "(3) (Excess)/Short provision of income tax");
  setCell(ws, r++, 4, 0, { money: true });
  const taxEnd = r - 1;
  r++;

  setCell(ws, r, 1, "VIII", { bold: true });
  setCell(ws, r, 2, "Profit for the year", { bold: true });
  const netProfit = round2(pbt - 0);
  setFormula(ws, r, 4, `D${pbtRow}-SUM(D${taxStart}:D${taxEnd})`, money(netProfit), {
    bold: true,
    money: true,
  });
  const netProfitRow = r;
  topBorder(ws, r, 1, 4);
  r += 2;

  const capitalTotal = round2(st.total("CAPITAL"));
  const issuedShares = Math.max(1, Math.round(capitalTotal / 10));
  setCell(ws, r, 1, "IX", { bold: true });
  setCell(ws, r, 2, "Earnings per equity share (face value Rs.10/share, assumed)");
  setFormula(ws, r, 4, `D${netProfitRow}/${issuedShares}`, round2(netProfit / issuedShares), {
    money: true,
  });
  r += 2;

  setCell(
    ws,
    r++,
    1,
    "Note: Share count for EPS is estimated from Capital / Rs.10 face value - confirm against the share register.",
    { italic: true },
  );
  r++;

  signatureBlock(ws, r, meta, ctx.forFirmText, ctx.designation);

  return { netProfitRow };
}
