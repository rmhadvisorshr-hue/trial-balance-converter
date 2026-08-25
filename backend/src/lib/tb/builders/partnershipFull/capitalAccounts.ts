import type ExcelJS from "exceljs";
import type { BuildCtx } from "../common";
import { setCell, setFormula, applyColumnWidths, topBorder } from "../../excel/helpers";
import { round2 } from "../../classify";
import type { Partner } from "./partners";
import type { PlAppropriationResult } from "./plAppropriation";

export interface CapitalAccountsResult {
  closingRowByPartner: Map<string, number>; // "Balance c/d" row per partner - for BS
  closingAmountByPartner: Map<string, number>; // known closing amount, in actual rupees - for BS's cached totals
}

// 'Cap': one Debit/Credit block per partner. A live (pre-closing) trial
// balance's Capital ledger already reflects everything posted *during* the
// year (drawings, capital introduced) but not the current year's Interest/
// Remuneration/Share of Profit, since those can only be computed now, at
// year end, via the P&L Appropriation account - that's the entire reason
// this account exists. So "By Balance b/d" is the known TB capital figure,
// and "To Balance c/d" (the true closing balance the Balance Sheet needs)
// is *computed* by adding this year's appropriations - the reverse of what
// might look natural, but matches both the reference file's own formula
// (`Balance c/d = Total credit - other debits`) and the fix already applied
// to Pvt Ltd's Reserves & Surplus (known opening, computed closing).
// Drawings/PTEC are assumed already reflected in the known TB balance and
// default to 0 here to avoid double-counting; if the CA has a separate
// drawings sub-ledger, entering it here will correctly flow through.
export function writeCapitalAccounts(
  ws: ExcelJS.Worksheet,
  ctx: BuildCtx,
  partners: Partner[],
  plApp: PlAppropriationResult,
): CapitalAccountsResult {
  applyColumnWidths(ws, [30, 18, 4, 26, 18]);
  const { meta, divisor } = ctx;
  const money = (v: number) => v / divisor;
  const closingRowByPartner = new Map<string, number>();
  const closingAmountByPartner = new Map<string, number>();

  let r = 2;
  setCell(ws, r++, 1, meta.firmName || "FIRM NAME", { bold: true });
  setCell(ws, r++, 1, meta.periodLabel || "", { italic: true });
  if (ctx.unitHeading) setCell(ws, r++, 1, ctx.unitHeading, { italic: true });
  r++;

  for (const p of partners) {
    const refs = plApp.byPartner.get(p.name)!;
    setCell(ws, r++, 1, `${p.name.toUpperCase()} CAPITAL ACCOUNT (${(p.sharePct * 100).toFixed(2)}%)`, {
      bold: true,
    });
    setCell(ws, r, 1, "PARTICULARS", { bold: true });
    setCell(ws, r, 2, "AMOUNT", { bold: true, align: "right" });
    setCell(ws, r, 4, "PARTICULARS", { bold: true });
    setCell(ws, r++, 5, "AMOUNT", { bold: true, align: "right" });

    const drawingsRow = r;
    setCell(ws, r, 1, "To Drawings (assume already reflected in the TB balance below - 0 to avoid double-counting)");
    setCell(ws, r, 2, 0, { money: true });
    setCell(ws, r, 4, "By Balance b/d (per trial balance)");
    const openingRow = r;
    setCell(ws, r, 5, money(p.capital), { money: true }); // known: TB capital ledger balance
    r++;

    const ptecRow = r;
    setCell(ws, r, 1, "To PTEC");
    setCell(ws, r, 2, 0, { money: true });
    setCell(ws, r, 4, "By Capital Introduced");
    setCell(ws, r, 5, 0, { money: true });
    r++;

    setCell(ws, r, 4, "By Interest on Capital");
    setFormula(ws, r, 5, `+'p&l app'!B${refs.interestRow}`, 0, { money: true });
    const interestCreditRow = r;
    r++;

    setCell(ws, r, 4, "By Remuneration as Salary");
    setFormula(ws, r, 5, `+'p&l app'!B${refs.remunerationRow}`, money(refs.remunerationAmount), { money: true });
    const remunerationCreditRow = r;
    r++;

    setCell(ws, r, 4, "By Share in Net Profit");
    const shareCreditRow = r;
    setFormula(ws, r, 5, `+'p&l app'!B${refs.shareOfProfitRow}`, money(refs.shareOfProfitAmount), { money: true });
    r++;

    // Closing = Opening + CapIntroduced + Interest + Remuneration + ShareOfProfit - Drawings - PTEC.
    const closingKnown = round2(p.capital + refs.remunerationAmount + refs.shareOfProfitAmount);
    setCell(ws, r, 1, "To Balance c/d", { bold: true });
    const closingRow = r;
    setFormula(
      ws,
      r,
      2,
      `E${openingRow}+E${ptecRow}+E${interestCreditRow}+E${remunerationCreditRow}+E${shareCreditRow}-B${drawingsRow}-B${ptecRow}`,
      money(closingKnown),
      { bold: true, money: true },
    );
    r++;

    r++;
    setCell(ws, r, 1, "TOTAL", { bold: true });
    setFormula(ws, r, 2, `SUM(B${drawingsRow}:B${closingRow})`, money(closingKnown), { bold: true, money: true });
    setCell(ws, r, 4, "TOTAL", { bold: true });
    setFormula(ws, r, 5, `SUM(E${openingRow}:E${shareCreditRow})`, money(closingKnown), {
      bold: true,
      money: true,
    });
    topBorder(ws, r, 1, 5);
    r += 2;

    closingRowByPartner.set(p.name, closingRow);
    closingAmountByPartner.set(p.name, closingKnown);
  }

  return { closingRowByPartner, closingAmountByPartner };
}
