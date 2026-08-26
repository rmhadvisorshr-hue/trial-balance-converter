import type ExcelJS from "exceljs";
import type { WorkbookAnalysis } from "../types";
import { writeNotesSequence, type NoteDef } from "./noteWriter";

const DEFS: NoteDef[] = [
  { noteNo: "12", title: "Investments - Non-current and Current", codes: ["N12_INVESTMENT"] },
  { noteNo: "13", title: "Loans and Advances (Unsecured, considered good)", codes: ["N13_LOANS_ADVANCES"] },
  { noteNo: "14", title: "Other Non-current Assets", codes: ["N14_OTHER_NONCURR_ASSET"] },
  { noteNo: "15", title: "Inventories", codes: ["N15_INVENTORY"] },
  { noteNo: "16", title: "Trade Receivables", codes: ["N16_TRADE_RECEIVABLE"] },
  { noteNo: "17", title: "Cash and Bank Balances", codes: ["N17_CASH_BANK"] },
  { noteNo: "18", title: "Other Current Assets", codes: ["N18_OTHER_CURR_ASSET"] },
];

export function writeNotesBs12to18(ws: ExcelJS.Worksheet, analysis: WorkbookAnalysis) {
  writeNotesSequence(ws, analysis, DEFS);
}
