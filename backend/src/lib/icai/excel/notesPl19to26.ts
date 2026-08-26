import type ExcelJS from "exceljs";
import type { WorkbookAnalysis } from "../types";
import { setCell, setFormula, topBorder } from "../../tb/excel/helpers";
import { costNoteTitle, depreciationTotal } from "./common";
import { writeNotesSheetHeader, writeNote, type NoteDef } from "./noteWriter";

// Depreciation for the year is read off Note 11's own fixed-asset roll-
// forward rather than matched by P&L line name across years - a firm's own
// books don't always spell the P&L line the same way from one year to the
// next (e.g. "Depreciation" vs "Depreciation account"), but there is always
// exactly one depreciation figure per year and it must tie to Note 11.
function writeDepreciationNote(ws: ExcelJS.Worksheet, analysis: WorkbookAnalysis, startRow: number): number {
  let r = startRow;
  setCell(ws, r++, 1, "Note 25: Depreciation and Amortization Expense", { bold: true });
  const depCur = depreciationTotal(analysis, "amountCurrent");
  const depPrev = depreciationTotal(analysis, "amountPrevious");
  if (!depCur && !depPrev) {
    setCell(ws, r++, 1, "NIL");
    return r + 1;
  }
  setCell(ws, r, 1, "Particulars", { bold: true });
  setCell(ws, r, 2, analysis.currentYearLabel, { bold: true, align: "right" });
  setCell(ws, r++, 3, analysis.previousYearLabel || "-", { bold: true, align: "right" });
  const dataRow = r;
  setCell(ws, r, 1, "Depreciation on tangible assets (Refer Note 11)");
  setCell(ws, r, 2, depCur, { money: true });
  setCell(ws, r++, 3, depPrev, { money: true });
  setCell(ws, r, 1, "Total Depreciation and Amortization Expense", { bold: true });
  setFormula(ws, r, 2, `SUM(B${dataRow}:B${dataRow})`, depCur, { bold: true, money: true });
  setFormula(ws, r, 3, `SUM(C${dataRow}:C${dataRow})`, depPrev, { bold: true, money: true });
  topBorder(ws, r, 1, 3);
  r += 2;
  return r;
}

export function writeNotesPl19to26(ws: ExcelJS.Worksheet, analysis: WorkbookAnalysis) {
  let r = writeNotesSheetHeader(ws, analysis);
  const defs: NoteDef[] = [
    { noteNo: "19", title: "Revenue from Operations", codes: ["N19_REVENUE"] },
    { noteNo: "20", title: "Other Income", codes: ["N20_OTHER_INCOME"] },
    { noteNo: "21", title: costNoteTitle(analysis), codes: ["N21_COST_OF_CONSTRUCTION"] },
  ];
  for (const def of defs) r = writeNote(ws, analysis, def, r);

  const hasWipFlag = analysis.warnings.some((w) => w.toLowerCase().includes("work-in-progress roll-forward"));
  setCell(ws, r++, 1, "Note 22: Changes in Inventories of Work-in-progress", { bold: true });
  if (hasWipFlag) {
    setCell(
      ws,
      r++,
      1,
      "NIL - movement in work-in-progress during the year is recognised directly on the Balance Sheet rather than routed through this statement - see the Basis of Preparation sheet.",
      { italic: true },
    );
  } else {
    setCell(ws, r++, 1, "NIL");
  }
  r++;

  r = writeNote(ws, analysis, { noteNo: "23", title: "Employee Benefits Expense", codes: ["N23_EMPLOYEE_BENEFIT"] }, r);
  r = writeNote(ws, analysis, { noteNo: "24", title: "Finance Costs", codes: ["N24_FINANCE_COST"] }, r);
  r = writeDepreciationNote(ws, analysis, r);
  writeNote(ws, analysis, { noteNo: "26", title: "Other Expenses", codes: ["N26_OTHER_EXPENSE"] }, r);
}
