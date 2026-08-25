import type ExcelJS from "exceljs";
import type { BuildCtx } from "../common";
import { setCell, setFormula, applyColumnWidths, topBorder } from "../../excel/helpers";
import { round2 } from "../../classify";
import type { Partner } from "../partnershipFull/partners";
import type { PlAppropriationResult } from "../partnershipFull/plAppropriation";

export interface CapitalCurrentResult {
  fixedClosingRowByPartner: Map<string, number>;
  fixedClosingAmountByPartner: Map<string, number>;
  currentClosingRowByPartner: Map<string, number>;
  currentClosingAmountByPartner: Map<string, number>;
  contributionDuringYearRowByPartner: Map<string, number>; // for FORM 8's "Contribution received during the year"
}

// 'Cap & Current': two accounts per partner, per the LLP Rules Second
// Schedule convention - NOT Partnership's single Cap account. The Fixed
// Capital/Contribution Account holds only the formal contribution (rarely
// moves, per the LLP Agreement); the Current Account absorbs every routine
// annual movement (accumulated reserves brought forward, Interest on
// Capital, Remuneration, Share of Profit, Drawings).
//
// detectPartners() folds the trial balance's accumulated Reserves & Surplus
// proportionally into each partner's `capital` figure (needed for
// Partnership's single-account model). For an LLP that fold is reversed
// here: `capital - reservesShare` recovers the raw TB Capital-ledger
// balance (the true Fixed Capital), and `reservesShare` becomes the Current
// Account's opening balance - both real, TB-derived figures, not guesses.
// Drawings and mid-year Contribution movements are the only genuinely
// non-derivable pieces (same "manual/0" pattern as Partnership's Cap
// account), since a trial balance carries no month-by-month movement detail.
export function writeCapitalCurrentAccounts(
  ws: ExcelJS.Worksheet,
  ctx: BuildCtx,
  partners: Partner[],
  plApp: PlAppropriationResult,
): CapitalCurrentResult {
  applyColumnWidths(ws, [36, 18, 4, 30, 18]);
  const { meta, st, divisor } = ctx;
  const money = (v: number) => v / divisor;
  const reservesTotal = round2(st.total("RESERVES"));

  const fixedClosingRowByPartner = new Map<string, number>();
  const fixedClosingAmountByPartner = new Map<string, number>();
  const currentClosingRowByPartner = new Map<string, number>();
  const currentClosingAmountByPartner = new Map<string, number>();
  const contributionDuringYearRowByPartner = new Map<string, number>();

  let r = 2;
  setCell(ws, r++, 1, meta.firmName || "LLP NAME", { bold: true });
  setCell(ws, r++, 1, meta.periodLabel || "", { italic: true });
  if (ctx.unitHeading) setCell(ws, r++, 1, ctx.unitHeading, { italic: true });
  r++;

  for (const p of partners) {
    const refs = plApp.byPartner.get(p.name)!;
    const reservesShare = round2(reservesTotal * p.sharePct);
    const rawCapital = round2(p.capital - reservesShare); // undo detectPartners' reserves fold

    // ---- Fixed Capital / Contribution Account ----
    setCell(
      ws,
      r++,
      1,
      `${p.name.toUpperCase()} - FIXED CAPITAL / CONTRIBUTION ACCOUNT (${(p.sharePct * 100).toFixed(2)}%)`,
      { bold: true },
    );
    setCell(ws, r, 1, "PARTICULARS", { bold: true });
    setCell(ws, r, 2, "AMOUNT", { bold: true, align: "right" });
    setCell(ws, r, 4, "PARTICULARS", { bold: true });
    setCell(ws, r++, 5, "AMOUNT", { bold: true, align: "right" });

    const withdrawnRow = r;
    setCell(ws, r, 1, "To Contribution Withdrawn During the Year (enter manually)");
    setCell(ws, r, 2, 0, { money: true });
    setCell(ws, r, 4, "By Balance b/d (per trial balance)");
    const fixedOpeningRow = r;
    setCell(ws, r, 5, money(rawCapital), { money: true });
    r++;

    setCell(ws, r, 4, "By Contribution Introduced During the Year (enter manually)");
    const contribRow = r;
    setCell(ws, r, 5, 0, { money: true });
    r++;

    setCell(ws, r, 1, "To Balance c/d", { bold: true });
    const fixedClosingRow = r;
    setFormula(ws, r, 2, `E${fixedOpeningRow}+E${contribRow}-B${withdrawnRow}`, money(rawCapital), {
      bold: true,
      money: true,
    });
    r++;

    r++;
    setCell(ws, r, 1, "TOTAL", { bold: true });
    setFormula(ws, r, 2, `SUM(B${withdrawnRow}:B${fixedClosingRow})`, money(rawCapital), {
      bold: true,
      money: true,
    });
    setCell(ws, r, 4, "TOTAL", { bold: true });
    setFormula(ws, r, 5, `SUM(E${fixedOpeningRow}:E${contribRow})`, money(rawCapital), {
      bold: true,
      money: true,
    });
    topBorder(ws, r, 1, 5);
    r += 2;

    fixedClosingRowByPartner.set(p.name, fixedClosingRow);
    fixedClosingAmountByPartner.set(p.name, rawCapital);
    contributionDuringYearRowByPartner.set(p.name, contribRow);

    // ---- Current Account ----
    setCell(ws, r++, 1, `${p.name.toUpperCase()} - CURRENT ACCOUNT`, { bold: true });
    setCell(ws, r, 1, "PARTICULARS", { bold: true });
    setCell(ws, r, 2, "AMOUNT", { bold: true, align: "right" });
    setCell(ws, r, 4, "PARTICULARS", { bold: true });
    setCell(ws, r++, 5, "AMOUNT", { bold: true, align: "right" });

    const drawingsRow = r;
    setCell(ws, r, 1, "To Drawings (enter manually)");
    setCell(ws, r, 2, 0, { money: true });
    setCell(ws, r, 4, "By Balance b/d (accumulated Reserves & Surplus per trial balance)");
    const currentOpeningRow = r;
    setCell(ws, r, 5, money(reservesShare), { money: true });
    r++;

    setCell(ws, r, 4, "By Interest on Capital");
    const interestCreditRow = r;
    setFormula(ws, r, 5, `+'p&l app'!B${refs.interestRow}`, 0, { money: true });
    r++;

    setCell(ws, r, 4, "By Remuneration as Salary");
    const remunerationCreditRow = r;
    setFormula(ws, r, 5, `+'p&l app'!B${refs.remunerationRow}`, money(refs.remunerationAmount), { money: true });
    r++;

    setCell(ws, r, 4, "By Share in Net Profit");
    const shareCreditRow = r;
    setFormula(ws, r, 5, `+'p&l app'!B${refs.shareOfProfitRow}`, money(refs.shareOfProfitAmount), { money: true });
    r++;

    const currentClosingKnown = round2(reservesShare + refs.remunerationAmount + refs.shareOfProfitAmount);
    setCell(ws, r, 1, "To Balance c/d", { bold: true });
    const currentClosingRow = r;
    setFormula(
      ws,
      r,
      2,
      `E${currentOpeningRow}+E${interestCreditRow}+E${remunerationCreditRow}+E${shareCreditRow}-B${drawingsRow}`,
      money(currentClosingKnown),
      { bold: true, money: true },
    );
    r++;

    r++;
    setCell(ws, r, 1, "TOTAL", { bold: true });
    setFormula(ws, r, 2, `SUM(B${drawingsRow}:B${currentClosingRow})`, money(currentClosingKnown), {
      bold: true,
      money: true,
    });
    setCell(ws, r, 4, "TOTAL", { bold: true });
    setFormula(ws, r, 5, `SUM(E${currentOpeningRow}:E${shareCreditRow})`, money(currentClosingKnown), {
      bold: true,
      money: true,
    });
    topBorder(ws, r, 1, 5);
    r += 2;

    currentClosingRowByPartner.set(p.name, currentClosingRow);
    currentClosingAmountByPartner.set(p.name, currentClosingKnown);
  }

  return {
    fixedClosingRowByPartner,
    fixedClosingAmountByPartner,
    currentClosingRowByPartner,
    currentClosingAmountByPartner,
    contributionDuringYearRowByPartner,
  };
}
