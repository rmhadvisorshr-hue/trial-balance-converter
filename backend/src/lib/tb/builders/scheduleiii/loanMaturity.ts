import type ExcelJS from "exceljs";
import type { BuildCtx } from "../common";
import { setCell, setFormula, applyColumnWidths, topBorder } from "../../excel/helpers";

// Current maturities of long-term borrowings (EMI schedule). This needs the
// actual loan agreement terms (EMI amount, interest rate, tenure) which a
// trial balance never carries - only a closing loan balance is available.
// Emitted as an empty, correctly-formula'd template (Principal = EMI -
// Interest, Total = SUM) for the CA to fill in from the loan schedule/bank
// amortisation statement.
const MONTHS = 12;

export function writeLoanMaturitySheet(ws: ExcelJS.Worksheet, ctx: BuildCtx) {
  applyColumnWidths(ws, [22, 16, 16, 16]);
  const { meta } = ctx;

  let r = 1;
  setCell(ws, r++, 1, meta.firmName || "COMPANY NAME", { bold: true });
  setCell(ws, r++, 1, "Current maturities of long term borrowings", { bold: true });
  setCell(
    ws,
    r++,
    1,
    "Loan EMI/interest/principal split requires the actual loan amortisation schedule (not available from a trial balance) - enter from the lender's statement.",
    { italic: true },
  );
  r++;
  setCell(ws, r, 1, "Period", { bold: true });
  setCell(ws, r, 2, "EMI", { bold: true, align: "right" });
  setCell(ws, r, 3, "Interest", { bold: true, align: "right" });
  setCell(ws, r++, 4, "Principal", { bold: true, align: "right" });

  const start = r;
  for (let i = 0; i < MONTHS; i++) {
    setCell(ws, r, 1, `Month ${i + 1}`);
    setCell(ws, r, 2, 0, { money: true });
    setCell(ws, r, 3, 0, { money: true });
    setFormula(ws, r, 4, `B${r}-C${r}`, 0, { money: true });
    r++;
  }
  const end = r - 1;

  setCell(ws, r, 1, "Total", { bold: true });
  for (const c of [2, 3, 4]) {
    const col = String.fromCharCode(64 + c);
    setFormula(ws, r, c, `SUM(${col}${start}:${col}${end})`, 0, { bold: true, money: true });
  }
  topBorder(ws, r, 1, 4);
}
