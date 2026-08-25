import type ExcelJS from "exceljs";
import type { BuildCtx } from "../common";
import { writeVerticalPL } from "../common";
import { setCell, setFormula, applyColumnWidths, topBorder } from "../../excel/helpers";
import { round2 } from "../../classify";
import type { LineItem } from "../../statements";

export interface TradingPLResult {
  netProfitRow: number; // 'P&L' sheet row holding "To/By Net Profit c/d" - for p&l app
}

// Writes a two-sided (To/By) account section: debit items in col A/B,
// credit items in col D/E, independently listed (T-accounts don't require
// the two sides to line up row-for-row). The balancing figure lands on
// whichever side is short - "To ... c/d" on debit for a profit, "By ... c/d"
// on credit for a loss - and both column totals SUM over the same row range
// so blank cells on the shorter side just contribute 0.
function writeTSection(
  ws: ExcelJS.Worksheet,
  r: number,
  money: (v: number) => number,
  debitItems: LineItem[],
  creditItems: LineItem[],
  balancingLabel: string,
  balancingAmount: number,
): { nextRow: number; totalRow: number; balancingRow: number } {
  const start = r;
  debitItems.forEach((it, i) => {
    setCell(ws, start + i, 1, it.name);
    setCell(ws, start + i, 2, money(round2(it.amount)), { money: true });
  });
  creditItems.forEach((it, i) => {
    setCell(ws, start + i, 4, it.name);
    setCell(ws, start + i, 5, money(round2(it.amount)), { money: true });
  });
  const lastItemRow = Math.max(start + debitItems.length - 1, start + creditItems.length - 1, start - 1);
  const balRow = lastItemRow + 1;
  const debitItemsTotal = round2(debitItems.reduce((s, i) => s + i.amount, 0));
  const creditItemsTotal = round2(creditItems.reduce((s, i) => s + i.amount, 0));
  let debitTotal: number;
  let creditTotal: number;
  if (balancingAmount >= 0) {
    setCell(ws, balRow, 1, balancingLabel);
    setCell(ws, balRow, 2, money(round2(balancingAmount)), { money: true });
    debitTotal = round2(debitItemsTotal + balancingAmount);
    creditTotal = creditItemsTotal;
  } else {
    setCell(ws, balRow, 4, balancingLabel.replace("To ", "By ").replace("Profit", "Loss"));
    setCell(ws, balRow, 5, money(round2(-balancingAmount)), { money: true });
    debitTotal = debitItemsTotal;
    creditTotal = round2(creditItemsTotal - balancingAmount);
  }

  const totalRow = balRow + 1;
  setCell(ws, totalRow, 1, "TOTAL", { bold: true });
  setFormula(ws, totalRow, 2, `SUM(B${start}:B${balRow})`, money(debitTotal), {
    bold: true,
    money: true,
  });
  setCell(ws, totalRow, 4, "TOTAL", { bold: true });
  setFormula(ws, totalRow, 5, `SUM(E${start}:E${balRow})`, money(creditTotal), {
    bold: true,
    money: true,
  });
  return { nextRow: totalRow + 2, totalRow, balancingRow: balRow };
}

// 'P&L': the classic T-format (To/By) Trading, Profit & Loss Account.
// 'PL (V)' is the vertical restatement of the same figures - reuses the
// existing writeVerticalPL (already used by every other entity type), since
// it's the same underlying data laid out differently.
export function writeTradingProfitLoss(
  plWs: ExcelJS.Worksheet,
  plVWs: ExcelJS.Worksheet,
  ctx: BuildCtx,
): TradingPLResult {
  applyColumnWidths(plWs, [30, 14, 4, 30, 14]);
  const { meta, st, divisor } = ctx;
  const money = (v: number) => v / divisor;

  let r = 2;
  setCell(plWs, r++, 1, meta.firmName || "FIRM NAME", { bold: true });
  setCell(plWs, r++, 1, meta.periodLabel || "", { italic: true });
  if (ctx.unitHeading) setCell(plWs, r++, 1, ctx.unitHeading, { italic: true });
  setCell(plWs, r++, 1, "TRADING, PROFIT AND LOSS A/C FOR THE YEAR ENDED", { bold: true });
  r++;
  setCell(plWs, r, 1, "PARTICULARS", { bold: true });
  setCell(plWs, r, 2, "AMOUNT", { bold: true, align: "right" });
  setCell(plWs, r, 4, "PARTICULARS", { bold: true });
  setCell(plWs, r++, 5, "AMOUNT", { bold: true, align: "right" });
  r++;

  const openingStock = st.byCode.get("OPENING_STOCK") ?? [];
  const purchases = [...(st.byCode.get("PURCHASES") ?? []), ...(st.byCode.get("DIRECT_EXP") ?? [])];
  const sales = st.byCode.get("REV_OPS") ?? [];
  const closingStock = st.byCode.get("CLOSING_STOCK_PL") ?? [];
  const tradingDebit = [...openingStock, ...purchases];
  const tradingCredit = [...sales, ...closingStock];
  const { nextRow: r2 } = writeTSection(plWs, r, money, tradingDebit, tradingCredit, "To Gross Profit c/d", st.grossProfit);
  r = r2;

  const indirectExpItems = [
    ...(st.byCode.get("EMP_BENEFIT") ?? []),
    ...(st.byCode.get("FINANCE_COST") ?? []),
    ...(st.byCode.get("DEPRECIATION") ?? []),
    ...(st.byCode.get("OTHER_EXP") ?? []),
  ];
  const indirectIncomeItems: LineItem[] = [
    { name: "By Gross Profit b/d", amount: round2(st.grossProfit) },
    ...(st.byCode.get("OTHER_INCOME") ?? []),
  ];
  const { totalRow: plTotalRow, balancingRow } = writeTSection(
    plWs,
    r,
    money,
    indirectExpItems,
    indirectIncomeItems,
    "To Net Profit c/d",
    st.netProfit,
  );
  void plTotalRow;

  writeVerticalPL(plVWs, ctx, "Profit & Loss A/c");

  return { netProfitRow: balancingRow };
}
