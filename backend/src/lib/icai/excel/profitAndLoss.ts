import type ExcelJS from "exceljs";
import type { WorkbookAnalysis } from "../types";
import { setCell, topBorder, applyColumnWidths } from "../../tb/excel/helpers";
import { totalFor, round2, costNoteTitle, depreciationTotal } from "./common";

export function writeProfitAndLoss(ws: ExcelJS.Worksheet, analysis: WorkbookAnalysis) {
  applyColumnWidths(ws, [50, 8, 20, 20]);
  let r = 1;
  setCell(ws, r++, 1, analysis.entityName || "ENTITY NAME", { bold: true });
  setCell(ws, r++, 1, `Statement of Profit and Loss for the year ended ${analysis.currentYearLabel}${analysis.previousYearLabel ? ` (with comparative figures for ${analysis.previousYearLabel})` : ""}`, { bold: true });
  setCell(ws, r++, 1, "(All amounts in Indian Rupees, unless otherwise stated)", { italic: true });
  if (analysis.unitHeading) setCell(ws, r++, 1, analysis.unitHeading, { italic: true });
  r++;
  setCell(ws, r, 1, "Particulars", { bold: true });
  setCell(ws, r, 2, "Note", { bold: true, align: "center" });
  setCell(ws, r, 3, analysis.currentYearLabel, { bold: true, align: "right" });
  setCell(ws, r++, 4, analysis.previousYearLabel || "-", { bold: true, align: "right" });

  const line = (label: string, note: string, cur: number, prev: number, opts: { bold?: boolean } = {}) => {
    setCell(ws, r, 1, label, opts);
    if (note) setCell(ws, r, 2, note, { align: "center" });
    setCell(ws, r, 3, cur, { money: true, bold: opts.bold });
    setCell(ws, r++, 4, prev, { money: true, bold: opts.bold });
  };

  const revenueCur = totalFor(analysis, "N19_REVENUE", "amountCurrent");
  const revenuePrev = totalFor(analysis, "N19_REVENUE", "amountPrevious");
  const otherIncCur = totalFor(analysis, "N20_OTHER_INCOME", "amountCurrent");
  const otherIncPrev = totalFor(analysis, "N20_OTHER_INCOME", "amountPrevious");
  line("Revenue from operations", "19", revenueCur, revenuePrev);
  line("Other income", "20", otherIncCur, otherIncPrev);
  const totalIncomeCur = round2(revenueCur + otherIncCur);
  const totalIncomePrev = round2(revenuePrev + otherIncPrev);
  line("Total Income", "", totalIncomeCur, totalIncomePrev, { bold: true });
  r++;

  setCell(ws, r++, 1, "Expenses:", { bold: true });
  const costCur = totalFor(analysis, "N21_COST_OF_CONSTRUCTION", "amountCurrent");
  const costPrev = totalFor(analysis, "N21_COST_OF_CONSTRUCTION", "amountPrevious");
  const invChangeCur = totalFor(analysis, "N22_INVENTORY_CHANGE", "amountCurrent");
  const invChangePrev = totalFor(analysis, "N22_INVENTORY_CHANGE", "amountPrevious");
  const empCur = totalFor(analysis, "N23_EMPLOYEE_BENEFIT", "amountCurrent");
  const empPrev = totalFor(analysis, "N23_EMPLOYEE_BENEFIT", "amountPrevious");
  const financeCur = totalFor(analysis, "N24_FINANCE_COST", "amountCurrent");
  const financePrev = totalFor(analysis, "N24_FINANCE_COST", "amountPrevious");
  const depCur = depreciationTotal(analysis, "amountCurrent");
  const depPrev = depreciationTotal(analysis, "amountPrevious");
  const otherExpCur = totalFor(analysis, "N26_OTHER_EXPENSE", "amountCurrent");
  const otherExpPrev = totalFor(analysis, "N26_OTHER_EXPENSE", "amountPrevious");

  line(costNoteTitle(analysis), "21", costCur, costPrev);
  line("Changes in inventories of work-in-progress", "22", invChangeCur, invChangePrev);
  line("Employee benefits expense", "23", empCur, empPrev);
  line("Finance costs", "24", financeCur, financePrev);
  line("Depreciation and amortization expense", "25", depCur, depPrev);
  line("Other expenses", "26", otherExpCur, otherExpPrev);
  const totalExpCur = round2(costCur + invChangeCur + empCur + financeCur + depCur + otherExpCur);
  const totalExpPrev = round2(costPrev + invChangePrev + empPrev + financePrev + depPrev + otherExpPrev);
  line("Total expenses", "", totalExpCur, totalExpPrev, { bold: true });
  topBorder(ws, r - 1, 1, 4);
  r++;

  const profitBeforeRemCur = round2(totalIncomeCur - totalExpCur);
  const profitBeforeRemPrev = round2(totalIncomePrev - totalExpPrev);
  line("Profit before partners' remuneration and tax", "", profitBeforeRemCur, profitBeforeRemPrev, { bold: true });
  r++;

  const remunerationCur = totalFor(analysis, "PARTNERS_REMUNERATION", "amountCurrent");
  const remunerationPrev = totalFor(analysis, "PARTNERS_REMUNERATION", "amountPrevious");
  line("Partners' remuneration", "3", remunerationCur, remunerationPrev);
  const profitBeforeTaxCur = round2(profitBeforeRemCur - remunerationCur);
  const profitBeforeTaxPrev = round2(profitBeforeRemPrev - remunerationPrev);
  line("Profit before tax", "", profitBeforeTaxCur, profitBeforeTaxPrev, { bold: true });
  r++;

  setCell(ws, r++, 1, "Tax expense:", { bold: true });
  const taxCur = totalFor(analysis, "N8_PROVISION", "amountCurrent");
  const taxPrev = totalFor(analysis, "N8_PROVISION", "amountPrevious");
  line("Current tax", "8", taxCur, taxPrev);
  line("Total tax expense", "", taxCur, taxPrev, { bold: true });
  r++;

  const profitForYearCur = round2(profitBeforeTaxCur - taxCur);
  const profitForYearPrev = round2(profitBeforeTaxPrev - taxPrev);
  line("Profit for the year", "", profitForYearCur, profitForYearPrev, { bold: true });
  topBorder(ws, r - 1, 1, 4);
  r += 2;

  setCell(ws, r++, 1, "Allocated as under (Note 3):", {});
  line("Share of profit credited to Partners' Capital Accounts", "", round2(analysis.owners.reduce((s, o) => s + o.shareOfProfit, 0)), round2(analysis.owners.reduce((s, o) => s + o.shareOfProfitPrevious, 0)));
  r++;
  setCell(ws, r++, 1, "The accompanying notes are an integral part of these financial statements.", { italic: true });
}
