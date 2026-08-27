import type ExcelJS from "exceljs";
import type { WorkbookAnalysis } from "../types";
import { setCell, topBorder, applyColumnWidths, signatureBlock } from "../../tb/excel/helpers";
import { totalMany, ownersFundsTotal, fixedAssetsTotal, round2, entitySignatoryDesignation } from "./common";

// Columns: A particulars, B note, C current year, D previous year.
export function writeBalanceSheet(ws: ExcelJS.Worksheet, analysis: WorkbookAnalysis) {
  applyColumnWidths(ws, [50, 8, 20, 20]);
  let r = 1;
  setCell(ws, r++, 1, analysis.entityName || "ENTITY NAME", { bold: true });
  setCell(ws, r++, 1, `Balance Sheet as at ${analysis.currentYearLabel}${analysis.previousYearLabel ? ` (with comparative figures as at ${analysis.previousYearLabel})` : ""}`, { bold: true });
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
  const heading = (label: string) => {
    setCell(ws, r++, 1, label, { bold: true });
  };

  heading("I  OWNERS' FUNDS AND LIABILITIES");
  heading("1  Owners' Funds");
  const ownersCur = ownersFundsTotal(analysis, "closingBalance");
  const ownersPrev = ownersFundsTotal(analysis, "closingBalancePrevious");
  const reservesCur = totalMany(analysis, ["N4_RESERVES"], "amountCurrent");
  const reservesPrev = totalMany(analysis, ["N4_RESERVES"], "amountPrevious");
  line("Owners'/Partners' Capital Account", "3", ownersCur, ownersPrev);
  line("Reserves and surplus", "4", reservesCur, reservesPrev);
  line("Total Owners' Funds", "", round2(ownersCur + reservesCur), round2(ownersPrev + reservesPrev), { bold: true });
  r++;

  heading("2  Non-current liabilities");
  const ltBorrowCur = totalMany(analysis, ["N5_LT_BORROWING"], "amountCurrent");
  const ltBorrowPrev = totalMany(analysis, ["N5_LT_BORROWING"], "amountPrevious");
  const dtlCur = totalMany(analysis, ["N6_DEFERRED_TAX"], "amountCurrent");
  const dtlPrev = totalMany(analysis, ["N6_DEFERRED_TAX"], "amountPrevious");
  const otherLtCur = totalMany(analysis, ["N7_OTHER_LT_LIAB"], "amountCurrent");
  const otherLtPrev = totalMany(analysis, ["N7_OTHER_LT_LIAB"], "amountPrevious");
  const ltProvCur = 0; // this pipeline only distinguishes short-term provisions - see N8_PROVISION below
  const ltProvPrev = 0;
  line("Long-term borrowings", "5", ltBorrowCur, ltBorrowPrev);
  line("Deferred tax liabilities (Net)", "6", dtlCur, dtlPrev);
  line("Other long-term liabilities", "7", otherLtCur, otherLtPrev);
  line("Long-term provisions", "8", ltProvCur, ltProvPrev);
  const nonCurrLiabCur = round2(ltBorrowCur + dtlCur + otherLtCur + ltProvCur);
  const nonCurrLiabPrev = round2(ltBorrowPrev + dtlPrev + otherLtPrev + ltProvPrev);
  line("Total Non-current liabilities", "", nonCurrLiabCur, nonCurrLiabPrev, { bold: true });
  r++;

  heading("3  Current liabilities");
  const stBorrowCur = totalMany(analysis, ["N5_ST_BORROWING"], "amountCurrent");
  const stBorrowPrev = totalMany(analysis, ["N5_ST_BORROWING"], "amountPrevious");
  const payableCur = totalMany(analysis, ["N9_TRADE_PAYABLE"], "amountCurrent");
  const payablePrev = totalMany(analysis, ["N9_TRADE_PAYABLE"], "amountPrevious");
  const otherCurLiabCur = totalMany(analysis, ["N10_OTHER_CURR_LIAB"], "amountCurrent");
  const otherCurLiabPrev = totalMany(analysis, ["N10_OTHER_CURR_LIAB"], "amountPrevious");
  const stProvCur = totalMany(analysis, ["N8_PROVISION"], "amountCurrent");
  const stProvPrev = totalMany(analysis, ["N8_PROVISION"], "amountPrevious");
  line("Short-term borrowings", "5", stBorrowCur, stBorrowPrev);
  line("Trade payables", "9", payableCur, payablePrev);
  line("Other current liabilities", "10", otherCurLiabCur, otherCurLiabPrev);
  line("Short-term provisions", "8", stProvCur, stProvPrev);
  const currLiabCur = round2(stBorrowCur + payableCur + otherCurLiabCur + stProvCur);
  const currLiabPrev = round2(stBorrowPrev + payablePrev + otherCurLiabPrev + stProvPrev);
  line("Total Current liabilities", "", currLiabCur, currLiabPrev, { bold: true });
  r++;

  const totalLiabSideCur = round2(ownersCur + reservesCur + nonCurrLiabCur + currLiabCur);
  const totalLiabSidePrev = round2(ownersPrev + reservesPrev + nonCurrLiabPrev + currLiabPrev);
  line("TOTAL (I)", "", totalLiabSideCur, totalLiabSidePrev, { bold: true });
  topBorder(ws, r - 1, 1, 4);
  r++;

  heading("II  ASSETS");
  heading("1  Non-current assets");
  const faCur = fixedAssetsTotal(analysis, "current");
  const faPrev = fixedAssetsTotal(analysis, "previous");
  const investCur = totalMany(analysis, ["N12_INVESTMENT"], "amountCurrent");
  const investPrev = totalMany(analysis, ["N12_INVESTMENT"], "amountPrevious");
  const dtaCur = 0;
  const dtaPrev = 0;
  const ltLoansCur = 0; // this pipeline classifies all loans & advances as short-term unless a source line says otherwise
  const ltLoansPrev = 0;
  const otherNcaCur = totalMany(analysis, ["N14_OTHER_NONCURR_ASSET"], "amountCurrent");
  const otherNcaPrev = totalMany(analysis, ["N14_OTHER_NONCURR_ASSET"], "amountPrevious");
  line("Property, Plant and Equipment", "11", faCur, faPrev);
  line("Non-current investments", "12", investCur, investPrev);
  line("Deferred tax assets (Net)", "6", dtaCur, dtaPrev);
  line("Long-term loans and advances", "13", ltLoansCur, ltLoansPrev);
  line("Other non-current assets", "14", otherNcaCur, otherNcaPrev);
  const nonCurrAssetCur = round2(faCur + investCur + dtaCur + ltLoansCur + otherNcaCur);
  const nonCurrAssetPrev = round2(faPrev + investPrev + dtaPrev + ltLoansPrev + otherNcaPrev);
  line("Total Non-current assets", "", nonCurrAssetCur, nonCurrAssetPrev, { bold: true });
  r++;

  heading("2  Current assets");
  const curInvestCur = 0;
  const curInvestPrev = 0;
  const inventoryCur = totalMany(analysis, ["N15_INVENTORY"], "amountCurrent");
  const inventoryPrev = totalMany(analysis, ["N15_INVENTORY"], "amountPrevious");
  const receivableCur = totalMany(analysis, ["N16_TRADE_RECEIVABLE"], "amountCurrent");
  const receivablePrev = totalMany(analysis, ["N16_TRADE_RECEIVABLE"], "amountPrevious");
  const cashCur = totalMany(analysis, ["N17_CASH_BANK"], "amountCurrent");
  const cashPrev = totalMany(analysis, ["N17_CASH_BANK"], "amountPrevious");
  const stLoansCur = totalMany(analysis, ["N13_LOANS_ADVANCES"], "amountCurrent");
  const stLoansPrev = totalMany(analysis, ["N13_LOANS_ADVANCES"], "amountPrevious");
  const otherCaCur = totalMany(analysis, ["N18_OTHER_CURR_ASSET"], "amountCurrent");
  const otherCaPrev = totalMany(analysis, ["N18_OTHER_CURR_ASSET"], "amountPrevious");
  line("Current investments", "12", curInvestCur, curInvestPrev);
  line("Inventories", "15", inventoryCur, inventoryPrev);
  line("Trade receivables", "16", receivableCur, receivablePrev);
  line("Cash and bank balances", "17", cashCur, cashPrev);
  line("Short-term loans and advances", "13", stLoansCur, stLoansPrev);
  line("Other current assets", "18", otherCaCur, otherCaPrev);
  const currAssetCur = round2(curInvestCur + inventoryCur + receivableCur + cashCur + stLoansCur + otherCaCur);
  const currAssetPrev = round2(curInvestPrev + inventoryPrev + receivablePrev + cashPrev + stLoansPrev + otherCaPrev);
  line("Total Current assets", "", currAssetCur, currAssetPrev, { bold: true });
  r++;

  const totalAssetSideCur = round2(nonCurrAssetCur + currAssetCur);
  const totalAssetSidePrev = round2(nonCurrAssetPrev + currAssetPrev);
  line("TOTAL (II)", "", totalAssetSideCur, totalAssetSidePrev, { bold: true });
  topBorder(ws, r - 1, 1, 4);
  r += 2;

  setCell(ws, r, 1, "Check: Total Assets less Total Liabilities (should be Nil)", { italic: true });
  setCell(ws, r, 3, round2(totalAssetSideCur - totalLiabSideCur), { money: true, italic: true });
  setCell(ws, r++, 4, round2(totalAssetSidePrev - totalLiabSidePrev), { money: true, italic: true });
  r += 2;

  setCell(ws, r++, 1, "Brief about the entity", {});
  setCell(ws, r++, 1, "Significant accounting policies", {});
  r++;
  setCell(ws, r++, 1, "The accompanying notes are an integral part of these financial statements.", { italic: true });

  signatureBlock(
    ws,
    r,
    analysis,
    `For ${analysis.entityName || "the entity"}`,
    entitySignatoryDesignation(analysis.entityKind),
  );
}
