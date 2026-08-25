import type ExcelJS from "exceljs";
import type { BuildCtx } from "../common";
import { setCell, setFormula, applyColumnWidths, topBorder, signatureBlock } from "../../excel/helpers";
import type { SelfSchResult } from "./selfSchedule";

// 'SELF CAP&BS(V)' (vertical) + 'SELF CAP&BS' (T-format): the proprietor's
// *personal* net worth statement, genuinely separate from the business one
// this time (unlike the app's previous output, which just relabeled the
// business BS under this sheet name). Net Profit is the one real linkage -
// formula-linked from 'FIRM CAP&BS' - because it's the only personal-net-
// worth figure a business trial balance can actually supply; every other
// line (personal bank interest, dividend, LIC survival benefit, drawings,
// personal insurance, gifts, and all personal assets via SELF SCH) is
// categorically outside a business TB and is emitted as a labeled 0/manual
// row, same honesty as the rest of this app's non-derivable figures.
export function writeSelfCapitalBalanceSheet(
  vWs: ExcelJS.Worksheet,
  tWs: ExcelJS.Worksheet,
  ctx: BuildCtx,
  netProfitFormulaRef: string, // e.g. "'FIRM CAP&BS'!F13"
  netProfitAmount: number,
  selfSch: SelfSchResult,
) {
  const { meta, divisor } = ctx;
  const money = (v: number) => v / divisor;

  // ---------------- SELF CAP&BS(V) ----------------
  applyColumnWidths(vWs, [34, 18]);
  let r = 1;
  setCell(vWs, r++, 1, meta.firmName || "FIRM NAME", { bold: true });
  setCell(vWs, r++, 1, "Personal Balance Sheet", { bold: true });
  setCell(vWs, r++, 1, meta.periodLabel || "", { italic: true });
  r++;
  setCell(vWs, r++, 1, "Liabilities:", { bold: true });
  setCell(vWs, r++, 1, "Capital Account", { bold: true });
  setCell(vWs, r, 1, "Opening Capital (enter manually)");
  setCell(vWs, r++, 2, 0, { money: true });
  setCell(vWs, r, 1, "Net Profit (from business)");
  setFormula(vWs, r++, 2, netProfitFormulaRef, money(netProfitAmount), { money: true });
  for (const label of ["Bank Interest", "Interest on FD", "Dividend Income", "LIC Survival Benefit"]) {
    setCell(vWs, r, 1, `${label} (enter manually)`);
    setCell(vWs, r++, 2, 0, { money: true });
  }
  for (const label of ["Drawings", "Personal Insurance Premium"]) {
    setCell(vWs, r, 1, `Less: ${label} (enter manually)`);
    setCell(vWs, r++, 2, 0, { money: true });
  }
  const capitalStart = r - 7;
  const capitalEnd = r - 1;
  setCell(vWs, r, 1, "Total Capital", { bold: true });
  setFormula(vWs, r, 2, `SUM(B${capitalStart}:B${capitalEnd})`, money(netProfitAmount), { bold: true, money: true });
  const vLiabTotalRow = r;
  topBorder(vWs, r, 1, 2);
  r += 2;

  setCell(vWs, r++, 1, "Assets:", { bold: true });
  setCell(vWs, r, 1, "Fixed Assets (Sch)");
  setFormula(vWs, r, 2, `'SELF SCH'!B${selfSch.fixedAssetsRow}`, 0, { money: true });
  const faRow = r;
  r++;
  setCell(vWs, r, 1, "Investments and Advances (enter manually)");
  setCell(vWs, r++, 2, 0, { money: true });
  const invRow = r - 1;
  setCell(vWs, r, 1, "Bank Balances (Sch)");
  setFormula(vWs, r, 2, `'SELF SCH'!B${selfSch.bankBalancesRow}`, 0, { money: true });
  const bankRow = r;
  r++;
  setCell(vWs, r, 1, "Cash Balance (enter manually)");
  setCell(vWs, r++, 2, 0, { money: true });
  const cashRow = r - 1;
  setCell(vWs, r, 1, "Total Assets", { bold: true });
  setFormula(vWs, r, 2, `SUM(B${faRow}:B${cashRow})`, 0, { bold: true, money: true });
  const vAssetTotalRow = r;
  topBorder(vWs, r, 1, 2);
  r += 2;
  setCell(vWs, r, 1, "Balancing check (Total Assets - Total Capital, should be 0 once personal figures above are filled in)", {
    italic: true,
  });
  setFormula(vWs, r, 2, `B${vAssetTotalRow}-B${vLiabTotalRow}`, money(netProfitAmount), { italic: true, money: true });
  void invRow;
  void bankRow;

  // ---------------- SELF CAP&BS (T-format) ----------------
  applyColumnWidths(tWs, [26, 18, 4, 26, 18]);
  r = 2;
  setCell(tWs, r++, 1, meta.firmName || "FIRM NAME", { bold: true });
  setCell(tWs, r++, 1, meta.periodLabel || "", { italic: true });
  r++;
  setCell(tWs, r++, 1, "Sch : Personal Capital Account", { bold: true });
  setCell(tWs, r, 1, "PARTICULARS", { bold: true });
  setCell(tWs, r, 2, "AMOUNT", { bold: true, align: "right" });
  setCell(tWs, r, 4, "PARTICULARS", { bold: true });
  setCell(tWs, r++, 5, "AMOUNT", { bold: true, align: "right" });

  const drawRow = r;
  setCell(tWs, r, 1, "To Drawings");
  setCell(tWs, r, 2, 0, { money: true });
  setCell(tWs, r, 4, "By Balance b/d (enter manually)");
  setCell(tWs, r, 5, 0, { money: true });
  r++;
  const insRow = r;
  setCell(tWs, r, 1, "To Personal Insurance Premium");
  setCell(tWs, r, 2, 0, { money: true });
  setCell(tWs, r, 4, "By Net Profit (from business)");
  const tNetProfitRow = r;
  setFormula(tWs, r, 5, netProfitFormulaRef, money(netProfitAmount), { money: true });
  r++;
  setCell(tWs, r, 4, "By Bank Interest / FD Interest / Dividend / LIC (enter manually)");
  setCell(tWs, r, 5, 0, { money: true });
  const otherIncomeRow = r;
  r++;
  setCell(tWs, r, 1, "To Balance c/d", { bold: true });
  const tClosingRow = r;
  setFormula(
    tWs,
    r,
    2,
    `E${drawRow}+E${insRow}+E${otherIncomeRow}-B${drawRow}-B${insRow}`,
    money(netProfitAmount),
    { bold: true, money: true },
  );
  void tClosingRow;
  r++;

  r++;
  setCell(tWs, r, 1, "Total", { bold: true });
  setFormula(tWs, r, 2, `SUM(B${drawRow}:B${tClosingRow})`, money(netProfitAmount), { bold: true, money: true });
  setCell(tWs, r, 4, "Total", { bold: true });
  setFormula(tWs, r, 5, `SUM(E${drawRow}:E${otherIncomeRow})`, money(netProfitAmount), { bold: true, money: true });
  topBorder(tWs, r, 1, 5);
  r += 2;
  void tNetProfitRow;

  signatureBlock(tWs, r, meta, ctx.forFirmText, "Proprietor");
}
