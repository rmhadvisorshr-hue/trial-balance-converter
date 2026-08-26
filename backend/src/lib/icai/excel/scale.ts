import type { WorkbookAnalysis } from "../types";
import { round2 } from "./common";

// Mirrors lib/tb/types.ts's FiguresUnit - kept as a plain divisor here rather
// than importing that module, since this pipeline has no other dependency on
// the trial-balance package.
export type FiguresUnit = "actual" | "thousands" | "lakhs";

export const FIGURES_UNIT_DIVISOR: Record<FiguresUnit, number> = {
  actual: 1,
  thousands: 1_000,
  lakhs: 100_000,
};

export const FIGURES_UNIT_HEADING: Record<FiguresUnit, string | null> = {
  actual: null,
  thousands: "Figures in Thousands ('000)",
  lakhs: "Figures in Lakhs",
};

function d(v: number, divisor: number): number {
  return round2(v / divisor);
}

// Analysis is always computed and reviewed in actual rupees; scaling only
// applies to the numbers written into the generated Excel, so this returns a
// new object rather than mutating the analysis the review screen is showing.
export function scaleAnalysis(analysis: WorkbookAnalysis, unit: FiguresUnit): WorkbookAnalysis {
  const divisor = FIGURES_UNIT_DIVISOR[unit];
  if (divisor === 1) return analysis;

  return {
    ...analysis,
    unitHeading: FIGURES_UNIT_HEADING[unit],
    accounts: analysis.accounts.map((a) => ({
      ...a,
      amountCurrent: d(a.amountCurrent, divisor),
      amountPrevious: d(a.amountPrevious, divisor),
    })),
    owners: analysis.owners.map((o) => ({
      ...o,
      openingBalance: d(o.openingBalance, divisor),
      openingBalancePrevious: d(o.openingBalancePrevious, divisor),
      capitalIntroduced: d(o.capitalIntroduced, divisor),
      capitalIntroducedPrevious: d(o.capitalIntroducedPrevious, divisor),
      remuneration: d(o.remuneration, divisor),
      remunerationPrevious: d(o.remunerationPrevious, divisor),
      interest: d(o.interest, divisor),
      interestPrevious: d(o.interestPrevious, divisor),
      shareOfProfit: d(o.shareOfProfit, divisor),
      shareOfProfitPrevious: d(o.shareOfProfitPrevious, divisor),
      otherCredits: d(o.otherCredits, divisor),
      otherCreditsPrevious: d(o.otherCreditsPrevious, divisor),
      withdrawals: d(o.withdrawals, divisor),
      withdrawalsPrevious: d(o.withdrawalsPrevious, divisor),
      withdrawalBreakup: o.withdrawalBreakup.map((w) => ({ ...w, amount: d(w.amount, divisor) })),
      closingBalance: d(o.closingBalance, divisor),
      closingBalancePrevious: d(o.closingBalancePrevious, divisor),
    })),
    fixedAssets: {
      current: analysis.fixedAssets.current.map((f) => ({
        ...f,
        opening: d(f.opening, divisor),
        additions: d(f.additions, divisor),
        depreciation: d(f.depreciation, divisor),
        closing: d(f.closing, divisor),
      })),
      previous: analysis.fixedAssets.previous.map((f) => ({
        ...f,
        opening: d(f.opening, divisor),
        additions: d(f.additions, divisor),
        depreciation: d(f.depreciation, divisor),
        closing: d(f.closing, divisor),
      })),
    },
    schedules: analysis.schedules.map((s) => ({
      ...s,
      sections: s.sections.map((sec) => ({
        ...sec,
        lines: sec.lines.map((l) => ({ ...l, amountCurrent: d(l.amountCurrent, divisor), amountPrevious: d(l.amountPrevious, divisor) })),
        subtotalCurrent: d(sec.subtotalCurrent, divisor),
        subtotalPrevious: d(sec.subtotalPrevious, divisor),
      })),
    })),
  };
}
