import ExcelJS from "exceljs";
import type { WorkbookAnalysis } from "../types";
import { applyWorkbookFormatting } from "../../tb/excel/formatting";
import { writeCoverSheet } from "./coverSheet";
import { writeBalanceSheet } from "./balanceSheet";
import { writeProfitAndLoss } from "./profitAndLoss";
import { writeNotes1to3 } from "./notes1to3";
import { writeNotesBs4to10 } from "./notesBs4to10";
import { writeNote11FixedAssets } from "./note11FixedAssets";
import { writeNotesBs12to18 } from "./notesBs12to18";
import { writeNotesPl19to26 } from "./notesPl19to26";
import { writeAllSchedules } from "./schedules";
import { writeBasisOfPreparation } from "./basisOfPreparation";

export async function buildIcaiWorkbook(analysis: WorkbookAnalysis): Promise<Buffer> {
  const wb = new ExcelJS.Workbook();
  wb.creator = "RMH Advisors - ICAI Financial Statement Generator";
  wb.created = new Date();

  writeCoverSheet(wb.addWorksheet("Cover"), analysis);
  writeBalanceSheet(wb.addWorksheet("Balance Sheet"), analysis);
  writeProfitAndLoss(wb.addWorksheet("Statement of P&L"), analysis);
  writeNotes1to3(wb.addWorksheet("Notes 1-3"), analysis);
  writeNotesBs4to10(wb.addWorksheet("Notes 4-10"), analysis);
  writeNote11FixedAssets(wb.addWorksheet("Note 11 - Fixed Assets"), analysis);
  writeNotesBs12to18(wb.addWorksheet("Notes 12-18"), analysis);
  writeNotesPl19to26(wb.addWorksheet("Notes 19-26"), analysis);
  writeAllSchedules(wb, analysis);
  writeBasisOfPreparation(wb.addWorksheet("Basis of Preparation"), analysis);

  // Same content-driven formatting pass the Trial Balance pipeline uses -
  // it works purely from what's already on the sheet, so it applies
  // unchanged to this pipeline's very different sheet layouts too.
  applyWorkbookFormatting(wb);

  const buf = await wb.xlsx.writeBuffer();
  return Buffer.from(buf);
}
