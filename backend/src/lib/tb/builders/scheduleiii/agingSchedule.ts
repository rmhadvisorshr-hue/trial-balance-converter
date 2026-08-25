import type ExcelJS from "exceljs";
import type { BuildCtx } from "../common";
import { setCell, setFormula, applyColumnWidths, topBorder } from "../../excel/helpers";
import type { NotesResult } from "./notes";

// AGING SC - Trade Payables & Trade Receivables ageing (mandatory since the
// FY 2021-22 Schedule III amendment). A trial balance only carries closing
// balances, not invoice/posting dates, so the actual age-wise split cannot
// be derived - the full Note 8 / Note 15 balance is placed under a single
// "not bucketed" column (so the grand total still ties to the Balance
// Sheet) and the CA must redistribute it across the real ageing buckets
// from the underlying sub-ledger.
export function writeAgingSheet(ws: ExcelJS.Worksheet, ctx: BuildCtx, notes: NotesResult) {
  applyColumnWidths(ws, [26, 14, 14, 14, 14, 14, 14]);
  const { meta, st, divisor } = ctx;
  const money = (v: number) => v / divisor;
  const noteVal = (n: number) => `NOTES!B${notes.byNote.get(n)!.row}`;

  let r = 2;
  setCell(ws, r++, 1, meta.firmName || "COMPANY NAME", { bold: true });
  setCell(ws, r++, 1, `Notes to Financial Statements for the year ended ${meta.periodLabel || ""}`, {
    bold: true,
  });
  setCell(
    ws,
    r++,
    1,
    "Ageing buckets cannot be derived from a trial balance (no invoice/posting-date detail) - each total below is placed under 'Not bucketed (redistribute)' so the grand total ties to Notes 8/15; the CA must redistribute using the debtor/creditor sub-ledger.",
    { italic: true },
  );
  r++;

  const writeAgingTable = (title: string, noteNo: number, total: number) => {
    setCell(ws, r++, 1, title, { bold: true });
    const headers = [
      "Particulars",
      "Not bucketed (redistribute)",
      "Less than 1 year",
      "1-2 years",
      "2-3 years",
      "More than 3 years",
      "Total",
    ];
    headers.forEach((h, i) => setCell(ws, r, i + 1, h, { bold: true, align: i === 0 ? "left" : "right", wrap: true }));
    r++;
    const rowStart = r;
    setCell(ws, r, 1, "Undisputed");
    setFormula(ws, r, 2, noteVal(noteNo), money(total), { money: true });
    for (const c of [3, 4, 5, 6]) setCell(ws, r, c, 0, { money: true });
    setFormula(ws, r, 7, `SUM(B${r}:F${r})`, money(total), { money: true });
    r++;
    setCell(ws, r, 1, "Disputed");
    for (const c of [2, 3, 4, 5, 6]) setCell(ws, r, c, 0, { money: true });
    setFormula(ws, r, 7, `SUM(B${r}:F${r})`, 0, { money: true });
    r++;
    const rowEnd = r - 1;
    setCell(ws, r, 1, "Total", { bold: true });
    for (const c of [2, 3, 4, 5, 6, 7]) {
      const col = String.fromCharCode(64 + c);
      const known = c === 2 || c === 7 ? total : 0;
      setFormula(ws, r, c, `SUM(${col}${rowStart}:${col}${rowEnd})`, money(known), { bold: true, money: true });
    }
    topBorder(ws, r, 1, 7);
    r += 2;
  };

  writeAgingTable("Trade Payables Ageing", 8, st.total("TRADE_PAYABLES"));
  writeAgingTable("Trade Receivables Ageing", 15, st.total("TRADE_RECV"));
}
