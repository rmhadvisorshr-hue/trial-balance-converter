import type ExcelJS from "exceljs";
import type { BuildCtx } from "../common";
import { setCell, setFormula, applyColumnWidths, topBorder } from "../../excel/helpers";
import { round2 } from "../../classify";
import type { PlAppropriationResult } from "./plAppropriation";

const TAX_RATE = 0.3; // flat rate for a partnership firm
const CESS_RATE = 0.04;

// 'STAT': computation of total income and tax payable. Net Profit as per
// the Appropriation A/c, the flat 30%+4%-cess rate, and TDS (from any
// TDS-named OTHER_CURR_ASSETS ledger) are all derivable from the trial
// balance; disallowed expenses, unabsorbed losses, Chapter VIA deductions,
// advance tax and self-assessment tax are external facts a TB can't supply,
// same "manual/0" pattern as the Pvt Ltd DTA disallowances. Every cached
// value below is computed step-by-step in the same order as the formula
// chain, so the displayed figures match what each formula actually
// evaluates to on recalculation.
export function writeStatSheet(ws: ExcelJS.Worksheet, ctx: BuildCtx, plApp: PlAppropriationResult) {
  applyColumnWidths(ws, [40, 20]);
  const { meta, divisor } = ctx;
  const money = (v: number) => v / divisor;

  let r = 2;
  setCell(ws, r++, 1, meta.firmName || "FIRM NAME", { bold: true });
  setCell(ws, r++, 1, meta.periodLabel || "", { italic: true });
  r++;
  setCell(ws, r++, 1, "STATEMENT SHOWING COMPUTATION OF TOTAL INCOME", { bold: true });
  r++;

  setCell(ws, r, 1, "Net Profit as per Profit and Loss Appropriation A/c");
  const netProfitRow = r;
  setFormula(
    ws,
    r++,
    2,
    `SUM('p&l app'!B${plApp.shareOfProfitStart}:B${plApp.shareOfProfitEnd})`,
    money(plApp.shareOfProfitTotal),
    { money: true },
  );

  setCell(ws, r, 1, "Add: Disallowed Expenses (enter manually - donations, etc.)");
  const disallowedRow = r;
  setCell(ws, r++, 2, 0, { money: true });
  r++;

  setCell(ws, r, 1, "Total Net Profit", { bold: true });
  const totalNetProfitRow = r;
  const totalNetProfit = round2(plApp.shareOfProfitTotal + 0);
  setFormula(ws, r++, 2, `B${netProfitRow}+B${disallowedRow}`, money(totalNetProfit), {
    bold: true,
    money: true,
  });
  r++;

  setCell(ws, r, 1, "Less: Unabsorbed Losses b/d (enter manually)");
  const lossesRow = r;
  setCell(ws, r++, 2, 0, { money: true });
  r++;

  setCell(ws, r, 1, "Gross Total Income", { bold: true });
  const gtiRow = r;
  const gti = round2(totalNetProfit - 0);
  setFormula(ws, r++, 2, `B${totalNetProfitRow}-B${lossesRow}`, money(gti), { bold: true, money: true });
  r++;

  setCell(ws, r, 1, "Less: Deductions Under Chapter VIA (enter manually)");
  const chapterViaRow = r;
  setCell(ws, r++, 2, 0, { money: true });
  r++;

  setCell(ws, r, 1, "Net Total Income", { bold: true });
  const ntiRow = r;
  const nti = Math.round(round2(gti - 0) / 10) * 10;
  setFormula(ws, r++, 2, `ROUND(B${gtiRow}-B${chapterViaRow},-1)`, money(nti), { bold: true, money: true });
  r++;

  setCell(ws, r, 1, `Income Tax @ ${TAX_RATE * 100}%`);
  const taxRow = r;
  const incomeTax = round2(nti * TAX_RATE);
  setFormula(ws, r++, 2, `B${ntiRow}*${TAX_RATE}`, money(incomeTax), { money: true });

  setCell(ws, r, 1, `Add: Cess @ ${CESS_RATE * 100}%`);
  const cessRow = r;
  const cess = Math.round(round2(incomeTax * CESS_RATE));
  setFormula(ws, r++, 2, `ROUND(B${taxRow}*${CESS_RATE},0)`, money(cess), { money: true });

  setCell(ws, r, 1, "Total Tax Payable", { bold: true });
  const totalTaxRow = r;
  const totalTax = Math.round(round2(incomeTax + cess));
  setFormula(ws, r++, 2, `ROUND(SUM(B${taxRow}:B${cessRow}),0)`, money(totalTax), { bold: true, money: true });
  r++;

  setCell(ws, r, 1, "Less: Advance Tax (enter manually)");
  const advanceTaxRow = r;
  setCell(ws, r++, 2, 0, { money: true });

  const tdsItem = (ctx.st.byCode.get("OTHER_CURR_ASSETS") ?? []).find((i) => /tds/i.test(i.name));
  const tdsAmount = round2(tdsItem?.amount ?? 0);
  setCell(ws, r, 1, "Less: TDS" + (tdsItem ? ` (${tdsItem.name})` : " (enter manually - no TDS ledger detected)"));
  const tdsRow = r;
  setCell(ws, r++, 2, money(tdsAmount), { money: true });

  setCell(ws, r, 1, "Less: Self-Assessment Tax Paid (enter manually)");
  const saTaxRow = r;
  setCell(ws, r++, 2, 0, { money: true });
  r++;

  setCell(ws, r, 1, "Net Tax Payable / (Refund Due)", { bold: true });
  setFormula(
    ws,
    r,
    2,
    `B${totalTaxRow}-B${advanceTaxRow}-B${tdsRow}-B${saTaxRow}`,
    money(round2(totalTax - 0 - tdsAmount - 0)),
    { bold: true, money: true },
  );
  topBorder(ws, r, 1, 2);
}
