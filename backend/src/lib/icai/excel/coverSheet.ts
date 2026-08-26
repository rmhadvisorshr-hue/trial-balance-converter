import type ExcelJS from "exceljs";
import type { WorkbookAnalysis } from "../types";
import { setCell, applyColumnWidths } from "../../tb/excel/helpers";

export function writeCoverSheet(ws: ExcelJS.Worksheet, analysis: WorkbookAnalysis) {
  applyColumnWidths(ws, [50, 30]);
  let r = 2;
  setCell(ws, r++, 1, analysis.entityName || "ENTITY NAME", { bold: true, size: 14 });
  r++;
  setCell(ws, r++, 1, "FINANCIAL STATEMENTS", { bold: true });
  setCell(ws, r++, 1, `For the year ended ${analysis.currentYearLabel}`, {});
  if (analysis.previousYearLabel) {
    setCell(ws, r++, 1, `(with comparative figures for the year ended ${analysis.previousYearLabel})`, {});
  }
  r++;
  setCell(ws, r++, 1, "Prepared in the format prescribed by the Institute of Chartered Accountants of India's", {});
  setCell(ws, r++, 1, "Guidance Note on Financial Statements of Non-Corporate Entities", {});
  r += 2;

  setCell(ws, r++, 1, "CONTENTS", { bold: true });
  const contents: [string, string][] = [
    ["Balance Sheet", ""],
    ["Statement of Profit and Loss", ""],
    ["Brief about the entity, Accounting Policies, Owners' Capital Account", "Notes 1 - 3"],
    ["Reserves, Borrowings, Deferred Tax, Other Liabilities, Provisions, Trade Payables", "Notes 4 - 10"],
    ["Property, Plant and Equipment", "Note 11"],
    ["Investments, Loans & Advances, Inventories, Receivables, Cash & Bank, Other Assets", "Notes 12 - 18"],
    ["Revenue, Other Income, Cost, Employee Cost, Finance Cost, Other Expenses", "Notes 19 - 26"],
    ...analysis.schedules.map((s) => [`${s.title} - Party-wise detail`, `Schedule ${s.letter}`] as [string, string]),
    ["Basis of Preparation, Judgements and Assumptions", ""],
  ];
  for (const [label, ref] of contents) {
    setCell(ws, r, 1, label);
    if (ref) setCell(ws, r, 2, ref);
    r++;
  }
}
