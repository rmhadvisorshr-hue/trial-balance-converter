import type ExcelJS from "exceljs";
import type { WorkbookAnalysis, SupportingSchedule } from "../types";
import { setCell, setFormula, topBorder, applyColumnWidths } from "../../tb/excel/helpers";
import { round2 } from "./common";

// One sheet per supporting schedule - however many the source data actually
// supports (not a fixed A/B/C), each with a sub-section per source block so
// a multi-source schedule (e.g. two separate "loans and advances" listings
// in the books) keeps its own headings instead of being silently flattened.
export function writeScheduleSheet(ws: ExcelJS.Worksheet, analysis: WorkbookAnalysis, schedule: SupportingSchedule) {
  applyColumnWidths(ws, [50, 20, 20]);
  let r = 1;
  setCell(ws, r++, 1, analysis.entityName || "ENTITY NAME", { bold: true });
  setCell(ws, r++, 1, `Schedule ${schedule.letter} - ${schedule.title} - Party-wise detail`, { bold: true });
  setCell(ws, r++, 1, `Supports Note ${schedule.supportsNote}`, { italic: true });
  r++;

  let grandCur = 0;
  let grandPrev = 0;
  for (const section of schedule.sections) {
    if (schedule.sections.length > 1) setCell(ws, r++, 1, section.heading, { bold: true });
    setCell(ws, r, 1, "Particulars", { bold: true });
    setCell(ws, r, 2, analysis.currentYearLabel, { bold: true, align: "right" });
    setCell(ws, r++, 3, analysis.previousYearLabel || "-", { bold: true, align: "right" });
    const start = r;
    for (const line of section.lines) {
      setCell(ws, r, 1, line.name);
      setCell(ws, r, 2, round2(line.amountCurrent), { money: true });
      setCell(ws, r++, 3, round2(line.amountPrevious), { money: true });
    }
    const end = r - 1;
    setCell(ws, r, 1, schedule.sections.length > 1 ? `Sub-total (${section.heading})` : "Total", { bold: true });
    setFormula(ws, r, 2, `SUM(B${start}:B${end})`, round2(section.subtotalCurrent), { bold: true, money: true });
    setFormula(ws, r, 3, `SUM(C${start}:C${end})`, round2(section.subtotalPrevious), { bold: true, money: true });
    topBorder(ws, r, 1, 3);
    r += 2;
    grandCur += section.subtotalCurrent;
    grandPrev += section.subtotalPrevious;
  }

  if (schedule.sections.length > 1) {
    setCell(ws, r, 1, "Total", { bold: true });
    setCell(ws, r, 2, round2(grandCur), { bold: true, money: true });
    setCell(ws, r, 3, round2(grandPrev), { bold: true, money: true });
    topBorder(ws, r, 1, 3);
  }
}

export function writeAllSchedules(wb: ExcelJS.Workbook, analysis: WorkbookAnalysis) {
  for (const schedule of analysis.schedules) {
    const ws = wb.addWorksheet(`Schedule ${schedule.letter}`);
    writeScheduleSheet(ws, analysis, schedule);
  }
}
