import type ExcelJS from "exceljs";
import type { WorkbookAnalysis } from "../types";
import { setCell, setFormula, topBorder } from "../../tb/excel/helpers";
import { accountsFor, totalFor, round2 } from "./common";
import { writeNotesSheetHeader, writeNote } from "./noteWriter";

// Note 5 shows Long-term and Short-term borrowings side by side, each with
// its own current/previous pair (4 amount columns total), unlike every
// other note here which is a plain current/previous pair - so it gets its
// own writer instead of going through writeNote().
function writeBorrowingsNote(ws: ExcelJS.Worksheet, analysis: WorkbookAnalysis, startRow: number): number {
  let r = startRow;
  setCell(ws, r++, 1, "Note 5: Borrowings", { bold: true });
  setCell(ws, r, 2, "Long-term", { bold: true, align: "center" });
  setCell(ws, r++, 4, "Short-term", { bold: true, align: "center" });
  setCell(ws, r, 1, "Particulars");
  setCell(ws, r, 2, analysis.currentYearLabel, { bold: true, align: "right" });
  setCell(ws, r, 3, analysis.previousYearLabel || "-", { bold: true, align: "right" });
  setCell(ws, r, 4, analysis.currentYearLabel, { bold: true, align: "right" });
  setCell(ws, r++, 5, analysis.previousYearLabel || "-", { bold: true, align: "right" });

  const lt = accountsFor(analysis, "N5_LT_BORROWING");
  const st = accountsFor(analysis, "N5_ST_BORROWING");
  if (lt.length === 0 && st.length === 0) {
    setCell(ws, r++, 1, "NIL");
    return r + 1;
  }

  const startData = r;
  for (const item of lt) {
    setCell(ws, r, 1, item.name, item.confident ? {} : { italic: true });
    setCell(ws, r, 2, item.amountCurrent, { money: true });
    setCell(ws, r, 3, item.amountPrevious, { money: true });
    setCell(ws, r, 4, 0, { money: true });
    setCell(ws, r++, 5, 0, { money: true });
  }
  for (const item of st) {
    setCell(ws, r, 1, item.name, item.confident ? {} : { italic: true });
    setCell(ws, r, 2, 0, { money: true });
    setCell(ws, r, 3, 0, { money: true });
    setCell(ws, r, 4, item.amountCurrent, { money: true });
    setCell(ws, r++, 5, item.amountPrevious, { money: true });
  }
  const endData = r - 1;
  setCell(ws, r, 1, "Total", { bold: true });
  setFormula(ws, r, 2, `SUM(B${startData}:B${endData})`, round2(totalFor(analysis, "N5_LT_BORROWING", "amountCurrent")), { bold: true, money: true });
  setFormula(ws, r, 3, `SUM(C${startData}:C${endData})`, round2(totalFor(analysis, "N5_LT_BORROWING", "amountPrevious")), { bold: true, money: true });
  setFormula(ws, r, 4, `SUM(D${startData}:D${endData})`, round2(totalFor(analysis, "N5_ST_BORROWING", "amountCurrent")), { bold: true, money: true });
  setFormula(ws, r, 5, `SUM(E${startData}:E${endData})`, round2(totalFor(analysis, "N5_ST_BORROWING", "amountPrevious")), { bold: true, money: true });
  topBorder(ws, r, 1, 5);
  r += 2;
  return r;
}

// Note 9 (Trade Payables): this pipeline has no source signal for MSME
// registration status, so - like the reference case - the whole balance
// defaults to "Others" with the MSME split left for the CA to confirm,
// rather than guessing.
function writeTradePayablesNote(ws: ExcelJS.Worksheet, analysis: WorkbookAnalysis, startRow: number): number {
  let r = startRow;
  setCell(ws, r++, 1, "Note 9: Trade Payables", { bold: true });
  setCell(ws, r, 1, "Particulars", { bold: true });
  setCell(ws, r, 2, analysis.currentYearLabel, { bold: true, align: "right" });
  setCell(ws, r++, 3, analysis.previousYearLabel || "-", { bold: true, align: "right" });

  const cur = round2(totalFor(analysis, "N9_TRADE_PAYABLE", "amountCurrent"));
  const prev = round2(totalFor(analysis, "N9_TRADE_PAYABLE", "amountPrevious"));
  setCell(ws, r, 1, "Total outstanding dues of micro, small and medium enterprises (MSME)");
  setCell(ws, r, 2, 0, { money: true });
  setCell(ws, r++, 3, 0, { money: true });
  setCell(ws, r, 1, "Total outstanding dues of creditors other than MSME");
  setCell(ws, r, 2, cur, { money: true });
  setCell(ws, r++, 3, prev, { money: true });
  setCell(ws, r, 1, "Total Trade Payables", { bold: true });
  setCell(ws, r, 2, cur, { money: true, bold: true });
  setCell(ws, r, 3, prev, { money: true, bold: true });
  topBorder(ws, r, 1, 3);
  r += 1;
  setCell(
    ws,
    r++,
    1,
    "MSMED Act, 2006 disclosure: no supplier has been confirmed as a registered micro or small enterprise, so the entire balance is shown as 'Others' above - confirm supplier registration status before relying on this split.",
    { italic: true },
  );
  return r + 1;
}

export function writeNotesBs4to10(ws: ExcelJS.Worksheet, analysis: WorkbookAnalysis) {
  let r = writeNotesSheetHeader(ws, analysis);
  r = writeNote(ws, analysis, { noteNo: "4", title: "Reserves and Surplus", codes: ["N4_RESERVES"] }, r);
  r = writeBorrowingsNote(ws, analysis, r);
  r = writeNote(ws, analysis, { noteNo: "6", title: "Deferred Tax Liabilities/(Assets) (Net)", codes: ["N6_DEFERRED_TAX"] }, r);
  r = writeNote(ws, analysis, { noteNo: "7", title: "Other Long-term Liabilities", codes: ["N7_OTHER_LT_LIAB"] }, r);
  r = writeNote(ws, analysis, { noteNo: "8", title: "Provisions", codes: ["N8_PROVISION"] }, r);
  r = writeTradePayablesNote(ws, analysis, r);
  writeNote(ws, analysis, { noteNo: "10", title: "Other Current Liabilities", codes: ["N10_OTHER_CURR_LIAB"] }, r);
}
