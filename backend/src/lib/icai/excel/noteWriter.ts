import type ExcelJS from "exceljs";
import type { WorkbookAnalysis, IcaiNoteCode } from "../types";
import { setCell, setFormula, topBorder, applyColumnWidths } from "../../tb/excel/helpers";
import { accountsFor, round2 } from "./common";

export interface NoteDef {
  noteNo: string;
  title: string;
  codes: IcaiNoteCode[];
}

export function writeNotesSheetHeader(ws: ExcelJS.Worksheet, analysis: WorkbookAnalysis): number {
  applyColumnWidths(ws, [50, 20, 20]);
  let r = 1;
  setCell(ws, r++, 1, analysis.entityName || "ENTITY NAME", { bold: true });
  setCell(ws, r++, 1, `Notes forming part of the Financial Statements for the year ended ${analysis.currentYearLabel}`, { bold: true });
  setCell(ws, r++, 1, "(All amounts in Indian Rupees, unless otherwise stated)", { italic: true });
  if (analysis.unitHeading) setCell(ws, r++, 1, analysis.unitHeading, { italic: true });
  r++;
  return r;
}

// Writes one standard two-column (current/previous) note: a title row, a
// Particulars/amount header, one line per matching account, and a SUM-
// formula total. Used for every note whose figures are simply "every
// account mapped to this code" - Notes 5 (borrowings) and 3 (capital
// accounts) don't fit this shape and are written by their own dedicated
// functions instead.
export function writeNote(ws: ExcelJS.Worksheet, analysis: WorkbookAnalysis, def: NoteDef, startRow: number): number {
  let r = startRow;
  setCell(ws, r++, 1, `Note ${def.noteNo}: ${def.title}`, { bold: true });
  const items = def.codes.flatMap((c) => accountsFor(analysis, c));
  const totalCurrent = round2(items.reduce((s, a) => s + a.amountCurrent, 0));
  const totalPrevious = round2(items.reduce((s, a) => s + a.amountPrevious, 0));

  if (items.length === 0) {
    setCell(ws, r++, 1, "NIL");
    r++;
    return r;
  }

  setCell(ws, r, 1, "Particulars", { bold: true });
  setCell(ws, r, 2, analysis.currentYearLabel, { bold: true, align: "right" });
  setCell(ws, r++, 3, analysis.previousYearLabel || "-", { bold: true, align: "right" });

  const startData = r;
  for (const item of items) {
    setCell(ws, r, 1, item.name, item.confident ? {} : { italic: true });
    setCell(ws, r, 2, item.amountCurrent, { money: true });
    setCell(ws, r++, 3, item.amountPrevious, { money: true });
  }
  const endData = r - 1;
  setCell(ws, r, 1, "Total", { bold: true });
  setFormula(ws, r, 2, `SUM(B${startData}:B${endData})`, totalCurrent, { bold: true, money: true });
  setFormula(ws, r, 3, `SUM(C${startData}:C${endData})`, totalPrevious, { bold: true, money: true });
  topBorder(ws, r, 1, 3);
  r += 2;
  return r;
}

export function writeNotesSequence(ws: ExcelJS.Worksheet, analysis: WorkbookAnalysis, defs: NoteDef[]) {
  let r = writeNotesSheetHeader(ws, analysis);
  for (const def of defs) r = writeNote(ws, analysis, def, r);
}
