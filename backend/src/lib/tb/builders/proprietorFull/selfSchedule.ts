import type ExcelJS from "exceljs";
import type { BuildCtx } from "../common";
import { setCell, setFormula, applyColumnWidths, topBorder } from "../../excel/helpers";

export interface SelfSchResult {
  bankBalancesRow: number;
  fixedAssetsRow: number;
}

// 'SELF SCH': personal bank balances and personal fixed assets. Neither is
// derivable from a business trial balance at all - the business's books
// never see the proprietor's personal bank accounts, gold, or residential
// property - so both are emitted as labeled 0/manual templates (a few
// common categories seeded as placeholders, same "structurally correct,
// empty" honesty as Int on Capital / Loan Maturity in the Partnership and
// Pvt Ltd builds) rather than left completely blank.
export function writeSelfSchedule(ws: ExcelJS.Worksheet, ctx: BuildCtx): SelfSchResult {
  applyColumnWidths(ws, [40, 20]);
  const { meta } = ctx;

  let r = 2;
  setCell(ws, r++, 1, meta.firmName || "FIRM NAME", { bold: true });
  setCell(ws, r++, 1, meta.periodLabel || "", { italic: true });
  r++;

  setCell(ws, r++, 1, "Sch : Bank Balances (Personal)", { bold: true });
  setCell(ws, r, 1, "PARTICULARS", { bold: true });
  setCell(ws, r++, 2, "AMOUNT", { bold: true, align: "right" });
  const bankStart = r;
  setCell(ws, r, 1, "(enter each personal bank account separately - not present in the business trial balance)", {
    italic: true,
  });
  setCell(ws, r++, 2, 0, { money: true });
  const bankEnd = r - 1;
  r++;
  setCell(ws, r, 1, "TOTAL", { bold: true });
  setFormula(ws, r, 2, `SUM(B${bankStart}:B${bankEnd})`, 0, { bold: true, money: true });
  topBorder(ws, r, 1, 2);
  const bankBalancesRow = r;
  r += 2;

  setCell(ws, r++, 1, "Sch : Fixed Assets (Personal)", { bold: true });
  setCell(ws, r, 1, "PARTICULARS", { bold: true });
  setCell(ws, r++, 2, "AMOUNT", { bold: true, align: "right" });
  const faStart = r;
  for (const label of ["Gold Ornaments", "Residential Property", "Furniture", "Vehicle (personal)"]) {
    setCell(ws, r, 1, `${label} (enter manually)`);
    setCell(ws, r++, 2, 0, { money: true });
  }
  const faEnd = r - 1;
  r++;
  setCell(ws, r, 1, "TOTAL", { bold: true });
  setFormula(ws, r, 2, `SUM(B${faStart}:B${faEnd})`, 0, { bold: true, money: true });
  topBorder(ws, r, 1, 2);
  const fixedAssetsRow = r;

  return { bankBalancesRow, fixedAssetsRow };
}
