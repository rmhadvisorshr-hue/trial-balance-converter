import type ExcelJS from "exceljs";
import type { WorkbookAnalysis } from "../types";
import { setCell, applyColumnWidths } from "../../tb/excel/helpers";

// The one sheet in the workbook that is explicitly NOT a statutory note: it
// records judgement calls and anything the engine could not confidently
// classify, so the CA reviews it before relying on the statements - the
// software's contribution is surfacing these, not resolving them silently.
export function writeBasisOfPreparation(ws: ExcelJS.Worksheet, analysis: WorkbookAnalysis) {
  applyColumnWidths(ws, [90]);
  let r = 1;
  setCell(ws, r++, 1, analysis.entityName || "ENTITY NAME", { bold: true });
  setCell(ws, r++, 1, "Basis of Preparation, Judgements and Assumptions", { bold: true });
  setCell(
    ws,
    r++,
    1,
    "This sheet is not a statutory note - it records matters that need the partners'/statutory auditor's confirmation. It should be read before relying on the statements.",
    { italic: true },
  );
  r += 2;

  if (analysis.warnings.length > 0) {
    setCell(ws, r++, 1, "1. Accounting judgements and cross-checks", { bold: true });
    for (const w of analysis.warnings) setCell(ws, r++, 1, `•  ${w}`, { wrap: true });
    r++;
  }

  const lowConfidence = analysis.accounts.filter((a) => !a.confident);
  if (lowConfidence.length > 0) {
    setCell(ws, r++, 1, "2. Classifications made with lower confidence - please confirm", { bold: true });
    for (const a of lowConfidence) {
      setCell(ws, r++, 1, `•  "${a.name}" (${a.sourceSheet}): mapped to ${a.code} - ${a.reason}`, { wrap: true });
    }
    r++;
  }

  if (analysis.reviewItems.length > 0) {
    setCell(ws, r++, 1, "3. Unmapped items requiring manual classification", { bold: true });
    for (const item of analysis.reviewItems) {
      const amount = item.amountCurrent || item.amountPrevious;
      setCell(
        ws,
        r++,
        1,
        `•  "${item.label}" (${item.sourceSheet}): Rs. ${amount.toLocaleString("en-IN")} - ${item.reason}`,
        { wrap: true },
      );
    }
  }

  if (analysis.warnings.length === 0 && lowConfidence.length === 0 && analysis.reviewItems.length === 0) {
    setCell(ws, r++, 1, "No accounting judgements were flagged and every line item was classified with high confidence.", {});
  }
}
