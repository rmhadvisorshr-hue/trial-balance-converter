import type ExcelJS from "exceljs";
import type { BuildCtx } from "../common";
import { setCell, setFormula, applyColumnWidths, topBorder } from "../../excel/helpers";
import { round2 } from "../../classify";

const CESS_RATE = 0.04;

function slab115bac(nti: number): number {
  if (nti <= 300000) return 0;
  if (nti <= 700000) return (nti - 300000) * 0.05;
  if (nti <= 1000000) return (nti - 700000) * 0.1 + 20000;
  if (nti <= 1200000) return (nti - 1000000) * 0.15 + 50000;
  if (nti <= 1500000) return (nti - 1200000) * 0.2 + 110000;
  return (nti - 1500000) * 0.3 + 140000;
}

// 'COMP': computation of total income for the individual (business + other
// sources), Chapter VIA deductions, and tax under Section 115BAC (the new
// regime slab rates - an exact, current statutory formula, fully
// replicable the same way Section 40(b) was for Partnership). Income from
// Other Sources (personal bank interest, dividend) and Chapter VIA (80C/
// 80D/80TTA - LIC, Mediclaim, personal savings interest) are personal-name
// figures a business trial balance doesn't carry - default 0/manual, same
// pattern as every other non-derivable figure in this app.
export function writeCompSheet(ws: ExcelJS.Worksheet, ctx: BuildCtx, netProfitFormulaRef: string, netProfitAmount: number) {
  applyColumnWidths(ws, [4, 36, 20, 20]);
  const { meta, divisor } = ctx;
  const money = (v: number) => v / divisor;

  let r = 2;
  setCell(ws, r++, 1, meta.firmName || "FIRM NAME", { bold: true });
  setCell(ws, r++, 1, meta.periodLabel || "", { italic: true });
  r++;
  setCell(ws, r++, 1, "STATEMENT SHOWING COMPUTATION OF TOTAL INCOME", { bold: true });
  r++;

  setCell(ws, r++, 2, "INCOME FROM BUSINESS", { bold: true });
  setCell(ws, r, 2, "Net Profit as per Profit & Loss A/c");
  const businessRow = r;
  setFormula(ws, r++, 4, netProfitFormulaRef, money(netProfitAmount), { money: true });
  r++;

  setCell(ws, r++, 2, "INCOME FROM OTHER SOURCES", { bold: true });
  setCell(ws, r, 2, "Bank Interest (enter manually)");
  const bankIntRow = r;
  setCell(ws, r++, 3, 0, { money: true });
  setCell(ws, r, 2, "Dividend Income (enter manually)");
  const divRow = r;
  setCell(ws, r++, 3, 0, { money: true });
  const otherSourcesRow = r;
  setFormula(ws, r++, 4, `SUM(C${bankIntRow}:C${divRow})`, 0, { money: true });
  r++;

  setCell(ws, r, 2, "Gross Total Income", { bold: true });
  const gtiRow = r;
  setFormula(ws, r++, 4, `ROUND(D${businessRow}+D${otherSourcesRow},0)`, money(round2(netProfitAmount)), {
    bold: true,
    money: true,
  });
  r++;

  setCell(ws, r++, 2, "Less: Deductions Under Chapter VIA", { bold: true });
  setCell(ws, r, 2, "U/S 80C - LIC (enter manually)");
  setCell(ws, r++, 3, 0, { money: true });
  setCell(ws, r, 2, "Restricted to");
  setCell(ws, r, 3, 150000, { money: true });
  const c80cRow = r - 1;
  setFormula(ws, r, 4, `MIN(C${c80cRow}:C${r})`, 0, { money: true });
  const d80c = r;
  r++;
  setCell(ws, r, 2, "U/S 80D - Mediclaim (enter manually)");
  setCell(ws, r++, 3, 0, { money: true });
  setCell(ws, r, 2, "Restricted to");
  setCell(ws, r, 3, 25000, { money: true });
  const c80dRow = r - 1;
  setFormula(ws, r, 4, `MIN(C${c80dRow}:C${r})`, 0, { money: true });
  const d80d = r;
  r++;
  setCell(ws, r, 2, "U/S 80TTA - Savings Bank Interest");
  setFormula(ws, r++, 3, `+C${bankIntRow}`, 0, { money: true });
  setCell(ws, r, 2, "Restricted to");
  setCell(ws, r, 3, 10000, { money: true });
  const c80ttaRow = r - 1;
  setFormula(ws, r, 4, `MIN(C${c80ttaRow}:C${r})`, 0, { money: true });
  const d80tta = r;
  r++;
  r++;

  setCell(ws, r, 2, "Net Total Income", { bold: true });
  const ntiRow = r;
  const nti = Math.round(round2(netProfitAmount) / 10) * 10;
  setFormula(ws, r++, 4, `ROUND(D${gtiRow}-D${d80c}-D${d80d}-D${d80tta},-1)`, money(nti), {
    bold: true,
    money: true,
  });
  r++;

  setCell(ws, r, 2, "Income Tax as per Section 115BAC");
  const taxRow = r;
  const tax115bac = round2(slab115bac(nti));
  setFormula(
    ws,
    r++,
    4,
    `ROUND(IF(D${ntiRow}<=300000,0,IF(D${ntiRow}<=700000,(D${ntiRow}-300000)*5%,` +
      `IF(D${ntiRow}<=1000000,(D${ntiRow}-700000)*10%+20000,` +
      `IF(D${ntiRow}<=1200000,(D${ntiRow}-1000000)*15%+50000,` +
      `IF(D${ntiRow}<=1500000,(D${ntiRow}-1200000)*20%+110000,(D${ntiRow}-1500000)*30%+140000))))),0)`,
    money(tax115bac),
    { money: true },
  );

  setCell(ws, r, 2, `Add: Education Cess @ ${CESS_RATE * 100}%`);
  const cessRow = r;
  const cess = round2(tax115bac * CESS_RATE);
  setFormula(ws, r++, 4, `ROUND(D${taxRow}*${CESS_RATE},0)`, money(cess), { money: true });

  setCell(ws, r, 2, "Total Tax Payable", { bold: true });
  const totalTaxRow = r;
  const totalTax = round2(tax115bac + cess);
  setFormula(ws, r++, 4, `SUM(D${taxRow}:D${cessRow})`, money(totalTax), { bold: true, money: true });
  r++;

  setCell(ws, r, 2, "Less: Advance Tax (enter manually)");
  const advTaxRow = r;
  setCell(ws, r++, 4, 0, { money: true });
  r++;

  setCell(ws, r, 2, "Add: Interest u/s 234A/234B/234C (enter manually)");
  const interestRow = r;
  setCell(ws, r++, 4, 0, { money: true });
  r++;

  setCell(ws, r, 2, "Less: Self-Assessment Tax Paid (enter manually)");
  const saTaxRow = r;
  setCell(ws, r++, 4, 0, { money: true });
  r++;

  setCell(ws, r, 2, "Tax Payable / (Refund)", { bold: true });
  setFormula(
    ws,
    r,
    4,
    `ROUND(D${totalTaxRow}-D${advTaxRow}+D${interestRow}-D${saTaxRow},-1)`,
    money(round2(totalTax)),
    { bold: true, money: true },
  );
  topBorder(ws, r, 1, 4);
}
