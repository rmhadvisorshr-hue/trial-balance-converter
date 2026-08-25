import { round2 } from "../../classify";
import type { Partner } from "./partners";

export const PROVISION_FOR_TAX_RATE = 0.312; // 30% + 4% cess, the standard partnership firm flat rate

export interface AppropriationMath {
  netProfit: number;
  bookProfit: number; // net profit less interest on capital (interest defaults to 0 - see interestOnCapital.ts)
  remunerationTotal: number; // Section 40(b) statutory ceiling
  taxProvision: number;
  shareOfProfitTotal: number;
  perPartner: { name: string; interest: number; remuneration: number; shareOfProfit: number }[];
}

// The Section 40(b) remuneration ceiling and the flat-rate tax provision -
// single source of truth shared by the Excel ('p&l app') and PDF renderers
// so this legally-defined calculation can't drift between the two outputs.
export function computeAppropriation(netProfit: number, partners: Partner[]): AppropriationMath {
  const interestTotal = 0; // not derivable from a TB - see interestOnCapital.ts
  const bookProfit = round2(netProfit - interestTotal);
  const remunerationTotal =
    bookProfit <= 166667
      ? Math.min(bookProfit, 150000)
      : bookProfit <= 300000
        ? bookProfit * 0.9
        : 270000 + (bookProfit - 300000) * 0.6;
  const taxProvision = Math.round(round2((netProfit - remunerationTotal) * PROVISION_FOR_TAX_RATE) / 10) * 10;
  const shareOfProfitTotal = round2(netProfit - remunerationTotal - taxProvision);

  const perPartner = partners.map((p) => ({
    name: p.name,
    interest: 0,
    remuneration: round2(remunerationTotal * p.sharePct),
    shareOfProfit: round2(shareOfProfitTotal * p.sharePct),
  }));

  return { netProfit, bookProfit, remunerationTotal, taxProvision, shareOfProfitTotal, perPartner };
}
