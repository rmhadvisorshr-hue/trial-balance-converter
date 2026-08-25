import type ExcelJS from "exceljs";
import type { BuildCtx } from "../common";
import { setCell, setFormula, applyColumnWidths, topBorder } from "../../excel/helpers";
import type { Partner } from "./partners";

const MONTHS = ["APR", "MAY", "JUN", "JUL", "AUG", "SEP", "OCT", "NOV", "DEC", "JAN", "FEB", "MAR"];

// 'Int on Capital': month-by-month fluctuating capital balance and interest
// per partner. The reference file's own monthly balances are typed-in
// figures (not derived from any formula), because a trial balance only
// carries a single closing balance per ledger, never a month-by-month
// movement history - so this is emitted as an empty, correctly-formula'd
// template (Interest = ROUND(Balance*Rate/12,0)) for the CA to fill in from
// the partner capital sub-ledger.
export function writeInterestOnCapital(ws: ExcelJS.Worksheet, ctx: BuildCtx, partners: Partner[]) {
  const colsPerPartner = 2;
  applyColumnWidths(
    ws,
    [10, ...partners.flatMap(() => [16, 12])],
  );
  const { meta } = ctx;

  let r = 2;
  setCell(ws, r++, 1, meta.firmName || "FIRM NAME", { bold: true });
  setCell(ws, r++, 1, meta.periodLabel || "", { italic: true });
  r++;
  setCell(ws, r++, 1, "CALCULATION OF INTEREST ON CAPITAL", { bold: true });
  r++;

  partners.forEach((p, i) => {
    setCell(ws, r, 2 + i * colsPerPartner, p.name, { bold: true, align: "center" });
  });
  r++;
  setCell(ws, r, 1, "MONTH", { bold: true });
  partners.forEach((_p, i) => {
    setCell(ws, r, 2 + i * colsPerPartner, "AMOUNT", { bold: true, align: "right" });
    setCell(ws, r, 3 + i * colsPerPartner, "INTEREST", { bold: true, align: "right" });
  });
  r++;
  partners.forEach((_p, i) => {
    setCell(ws, r, 3 + i * colsPerPartner, 0.12, { align: "right" }); // rate/annum, editable per partner
  });
  r++;

  const start = r;
  for (const month of MONTHS) {
    setCell(ws, r, 1, month);
    partners.forEach((_p, i) => {
      const balCol = 2 + i * colsPerPartner;
      const rateCell = `${colLetter(3 + i * colsPerPartner)}${start - 1}`;
      setCell(ws, r, balCol, 0, { money: true }); // partner capital sub-ledger not available from a TB
      setFormula(ws, r, balCol + 1, `ROUND(${colLetter(balCol)}${r}*${rateCell}/12,0)`, 0, { money: true });
    });
    r++;
  }
  const end = r - 1;

  setCell(ws, r, 1, "TOTAL", { bold: true });
  partners.forEach((_p, i) => {
    const col = 3 + i * colsPerPartner;
    setFormula(ws, r, col, `SUM(${colLetter(col)}${start}:${colLetter(col)}${end})`, 0, {
      bold: true,
      money: true,
    });
  });
  topBorder(ws, r, 1, 2 + partners.length * colsPerPartner - 1);
}

function colLetter(col: number): string {
  let s = "";
  let n = col;
  while (n > 0) {
    const rem = (n - 1) % 26;
    s = String.fromCharCode(65 + rem) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}
