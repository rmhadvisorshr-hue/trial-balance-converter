import type ExcelJS from "exceljs";
import type { BuildCtx } from "../common";
import { setCell, setFormula, applyColumnWidths, topBorder, signatureBlock } from "../../excel/helpers";
import { round2 } from "../../classify";
import type { NotesResult } from "./notes";
import type { ShareCapitalReservesResult } from "./shareCapitalReserves";

export function writeScheduleIIIBalanceSheet(
  ws: ExcelJS.Worksheet,
  ctx: BuildCtx,
  notes: NotesResult,
  shareCap: ShareCapitalReservesResult,
  faNetClosingRow: number,
) {
  applyColumnWidths(ws, [6, 42, 8, 20, 20]);
  const { meta, st, divisor } = ctx;
  const money = (v: number) => v / divisor;
  const noteRow = (n: number) => notes.byNote.get(n)!.row;

  let r = 2;
  setCell(ws, r++, 1, meta.firmName || "COMPANY NAME", { bold: true });
  if (meta.cin) setCell(ws, r++, 1, `CIN : ${meta.cin}`, {});
  setCell(ws, r++, 1, `BALANCE SHEET AS AT ${(meta.asAtLabel || "").replace(/^as at\s*/i, "")}`, {
    bold: true,
  });
  if (ctx.unitHeading) setCell(ws, r++, 1, ctx.unitHeading, { italic: true });
  r++;
  setCell(ws, r, 2, "Particulars", { bold: true });
  setCell(ws, r, 3, "Notes", { bold: true, align: "center" });
  setCell(ws, r++, 4, meta.asAtLabel || meta.periodLabel || "", { bold: true, align: "right" });
  r++;

  const liabRows: number[] = [];
  const line = (label: string, noteNo: number | null, formula: string, known: number, prefix = "") => {
    setCell(ws, r, 2, prefix ? `${prefix} ${label}` : label);
    if (noteNo) setCell(ws, r, 3, noteNo, { align: "center" });
    setFormula(ws, r, 4, formula, money(known), { money: true });
    liabRows.push(r);
    r++;
  };

  setCell(ws, r, 1, "I", { bold: true });
  setCell(ws, r++, 2, "EQUITY AND LIABILITIES", { bold: true });
  setCell(ws, r, 1, 1, { bold: true });
  setCell(ws, r++, 2, "Shareholders' funds", { bold: true });
  line("Share capital", 2, `'Note 2,3- SC'!C${shareCap.shareCapitalRow}`, st.total("CAPITAL"), "(a)");
  const reserves = round2(st.total("RESERVES") + st.netProfit);
  line("Reserves and surplus", 3, `'Note 2,3- SC'!C${shareCap.reservesRow}`, reserves, "(b)");

  setCell(ws, r, 1, 2, { bold: true });
  setCell(ws, r++, 2, "Non-current liabilities", { bold: true });
  line("Long term borrowings", 4, `NOTES!B${noteRow(4)}`, st.total("LONG_TERM_BORROW"), "(a)");
  line("Other long term liabilities", 5, `NOTES!B${noteRow(5)}`, 0, "(b)");
  line("Deferred tax liabilities (net)", 6, `NOTES!B${noteRow(6)}`, st.total("DEFERRED_TAX_LIAB"), "(c)");

  setCell(ws, r, 1, 3, { bold: true });
  setCell(ws, r++, 2, "Current liabilities", { bold: true });
  line("Short term borrowings", 7, `NOTES!B${noteRow(7)}`, st.total("SHORT_TERM_BORROW"), "(a)");
  line("Trade payables", 8, `NOTES!B${noteRow(8)}`, st.total("TRADE_PAYABLES"), "(b)");
  line("Other current liabilities", 9, `NOTES!B${noteRow(9)}`, st.total("OTHER_CURR_LIAB"), "(c)");
  line("Short term provisions", 10, `NOTES!B${noteRow(10)}`, st.total("SHORT_TERM_PROV"), "(d)");
  r++;

  const liabFirst = liabRows[0];
  const liabLast = liabRows[liabRows.length - 1];
  setCell(ws, r, 2, "TOTAL", { bold: true });
  const totalLiabilities = round2(st.totalLiabilities);
  setFormula(ws, r, 4, `SUM(D${liabFirst}:D${liabLast})`, money(totalLiabilities), {
    bold: true,
    money: true,
  });
  const equityLiabTotalRow = r;
  topBorder(ws, r, 1, 4);
  r += 2;

  const assetRows: number[] = [];
  const assetLine = (
    label: string,
    noteNo: number | null,
    formula: string,
    known: number,
    prefix = "",
  ) => {
    setCell(ws, r, 2, prefix ? `${prefix} ${label}` : label);
    if (noteNo) setCell(ws, r, 3, noteNo, { align: "center" });
    setFormula(ws, r, 4, formula, money(known), { money: true });
    assetRows.push(r);
    r++;
  };

  setCell(ws, r, 1, "II", { bold: true });
  setCell(ws, r++, 2, "ASSETS", { bold: true });
  setCell(ws, r, 1, 1, { bold: true });
  setCell(ws, r++, 2, "Non-current assets", { bold: true });
  assetLine(
    "Property, Plant and Equipment and Intangible Assets",
    11,
    `'11. FA'!J${faNetClosingRow}`,
    st.total("PPE"),
    "(a)",
  );
  assetLine("Non-current investments", 12, `NOTES!B${noteRow(12)}`, st.total("NONCURR_INVEST"), "(b)");
  const ltla = round2(st.total("LONG_TERM_LOANS_ADV"));
  setCell(ws, r, 2, "(c) Long-term loans and advances");
  setCell(ws, r, 4, money(ltla), { money: true });
  assetRows.push(r);
  r++;
  assetLine(
    "Deferred tax assets (net)",
    13,
    `NOTES!B${noteRow(13)}`,
    st.total("DEFERRED_TAX_ASSET"),
    "(d)",
  );

  setCell(ws, r, 1, 2, { bold: true });
  setCell(ws, r++, 2, "Current assets", { bold: true });
  assetLine("Inventories", 14, `NOTES!B${noteRow(14)}`, st.total("INVENTORY"), "(a)");
  assetLine("Trade receivables", 15, `NOTES!B${noteRow(15)}`, st.total("TRADE_RECV"), "(b)");
  assetLine("Cash and Bank Balances", 16, `NOTES!B${noteRow(16)}`, st.total("CASH_BANK"), "(c)");
  assetLine(
    "Short term loans and advances",
    17,
    `NOTES!B${noteRow(17)}`,
    st.total("SHORT_TERM_LOANS_ADV"),
    "(d)",
  );
  assetLine("Other current assets", 18, `NOTES!B${noteRow(18)}`, st.total("OTHER_CURR_ASSETS"), "(e)");
  r++;

  const assetFirst = assetRows[0];
  const assetLast = assetRows[assetRows.length - 1];
  setCell(ws, r, 2, "TOTAL", { bold: true });
  // st.totalAssets already includes LONG_TERM_LOANS_ADV (ltla) - do not add it again.
  const totalAssets = round2(st.totalAssets);
  setFormula(ws, r, 4, `SUM(D${assetFirst}:D${assetLast})`, money(totalAssets), {
    bold: true,
    money: true,
  });
  const assetsTotalRow = r;
  topBorder(ws, r, 1, 4);
  r += 2;

  setCell(ws, r, 2, "Balancing check (Total Assets - Total Liabilities, should be 0)", { italic: true });
  setFormula(
    ws,
    r,
    4,
    `D${assetsTotalRow}-D${equityLiabTotalRow}`,
    money(round2(totalAssets - totalLiabilities)),
    { italic: true, money: true },
  );
  r += 2;

  setCell(ws, r++, 1, "Summary of Significant Accounting Policies & Other Disclosures", {});
  setCell(ws, r++, 1, "The accompanying notes are an integral part of the Financial Statements", {});
  r++;

  const boardRow = signatureBlock(ws, r, meta, ctx.forFirmText, ctx.designation);
  setCell(ws, r + 2, 6, "For and on behalf of the Board", { bold: true });
  setCell(ws, r + 6, 6, "Director");
  setCell(ws, r + 7, 6, "DIN :", {});
  void boardRow;
}
