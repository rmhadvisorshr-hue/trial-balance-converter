import type ExcelJS from "exceljs";
import type { BuildCtx } from "../common";
import { writeNotesSheet } from "./notes";
import { writeFixedAssetChain } from "./fixedAssetChain";
import { writeScheduleIIIProfitAndLoss } from "./profitAndLoss";
import { writeShareCapitalReservesSheet } from "./shareCapitalReserves";
import { writeScheduleIIIBalanceSheet } from "./balanceSheet";
import { writeRatiosSheet } from "./ratios";
import { writeStatutoryNotesSheet } from "./statutoryNotes";
import { writeAgingSheet } from "./agingSchedule";
import { writeLoanMaturitySheet } from "./loanMaturity";

// Build order matters here (unlike the other statement styles): each sheet
// below is written as a live Excel formula chain, so a sheet that
// references another (e.g. P&L's Depreciation line references '11. FA')
// needs that sheet's row numbers already computed - hence NOTES/11.FA/DTL/
// DTA first, then P&L (needs NOTES + 11.FA + DTA), then Note 2,3-SC (needs
// P&L's Net Profit row), then BS/Ratio last (need all of the above).
// Ageing/Notes 27-31/Loan Maturity are additive disclosure sheets that
// don't feed back into any other sheet's formulas, so they can be written
// any time after the sheets they reference.
export function buildScheduleIII(wb: ExcelJS.Workbook, ctx: BuildCtx) {
  wb.calcProperties.fullCalcOnLoad = true;

  // Sheets are created up front in the tab order a reader expects (BS/P&L
  // first), but written into below in dependency order.
  const bsWs = wb.addWorksheet("BS");
  const plWs = wb.addWorksheet("P&L");
  const notesWs = wb.addWorksheet("NOTES");
  const scWs = wb.addWorksheet("Note 2,3- SC");
  const faWs = wb.addWorksheet("11. FA");
  const dtlWs = wb.addWorksheet("DTL");
  const dtaWs = wb.addWorksheet("DTA");
  const ratioWs = wb.addWorksheet("Ratio");
  const notes27Ws = wb.addWorksheet("Notes 27-31");
  const agingWs = wb.addWorksheet("AGING SC");
  const loanMaturityWs = wb.addWorksheet("Current Liability Maturity");

  const notes = writeNotesSheet(notesWs, ctx);
  const fa = writeFixedAssetChain(faWs, dtlWs, dtaWs, ctx);
  const pl = writeScheduleIIIProfitAndLoss(plWs, ctx, notes, fa, fa.dtaNetExpenseRow);
  const shareCap = writeShareCapitalReservesSheet(scWs, ctx, `'P&L'!D${pl.netProfitRow}`);
  writeScheduleIIIBalanceSheet(bsWs, ctx, notes, shareCap, fa.faNetClosingRow);
  writeRatiosSheet(ratioWs, ctx, notes, shareCap, fa, `'P&L'!D${pl.netProfitRow}`);
  writeStatutoryNotesSheet(notes27Ws, ctx);
  writeAgingSheet(agingWs, ctx, notes);
  writeLoanMaturitySheet(loanMaturityWs, ctx);
}
