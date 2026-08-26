import type { WorkbookAnalysis, IcaiNoteCode } from "../types";

const LIABILITY_CODES: IcaiNoteCode[] = [
  "N4_RESERVES",
  "N5_LT_BORROWING",
  "N5_ST_BORROWING",
  "N6_DEFERRED_TAX",
  "N7_OTHER_LT_LIAB",
  "N8_PROVISION",
  "N9_TRADE_PAYABLE",
  "N10_OTHER_CURR_LIAB",
];

// Excludes N11 (fixed assets), which never appears in `accounts` - it's
// carried separately in analysis.fixedAssets and totalled from there.
const ASSET_CODES: IcaiNoteCode[] = [
  "N12_INVESTMENT",
  "N13_LOANS_ADVANCES",
  "N14_OTHER_NONCURR_ASSET",
  "N15_INVENTORY",
  "N16_TRADE_RECEIVABLE",
  "N17_CASH_BANK",
  "N18_OTHER_CURR_ASSET",
];

const money = (n: number) => `Rs. ${n.toLocaleString("en-IN", { maximumFractionDigits: 0 })}`;

// Section 17's headline check: Total Assets = Owners' Funds + Liabilities,
// for both years, plus a soft cross-check against the source Balance
// Sheet's own grand total (when one could be found). Returns human-readable
// messages meant to be appended to WorkbookAnalysis.warnings - this never
// blocks generation, it only tells the CA where to look.
export function crossCheckBalanceSheet(
  analysis: WorkbookAnalysis,
  bsTotalCurrent: number | null,
  bsTotalPrevious: number | null,
): string[] {
  const messages: string[] = [];

  (["current", "previous"] as const).forEach((yearKey) => {
    const isCurrent = yearKey === "current";
    const amtOf = (code: IcaiNoteCode) =>
      analysis.accounts
        .filter((a) => a.code === code)
        .reduce((s, a) => s + (isCurrent ? a.amountCurrent : a.amountPrevious), 0);

    const ownersFunds = analysis.owners.reduce((s, o) => s + (isCurrent ? o.closingBalance : o.closingBalancePrevious), 0);
    const liabilities = LIABILITY_CODES.reduce((s, c) => s + amtOf(c), 0);
    const fixedAssetsTotal = (isCurrent ? analysis.fixedAssets.current : analysis.fixedAssets.previous).reduce(
      (s, a) => s + a.closing,
      0,
    );
    const otherAssets = ASSET_CODES.reduce((s, c) => s + amtOf(c), 0);

    const totalLiabilitiesSide = ownersFunds + liabilities;
    const totalAssetsSide = fixedAssetsTotal + otherAssets;
    const diff = totalAssetsSide - totalLiabilitiesSide;
    const label = isCurrent ? analysis.currentYearLabel || "current year" : analysis.previousYearLabel || "previous year";

    if (Math.abs(diff) > 1) {
      messages.push(
        `Balance Sheet does not balance for ${label}: Total Assets ${money(totalAssetsSide)} vs Owners' Funds + Liabilities ${money(totalLiabilitiesSide)} (difference ${money(diff)}). Check the review items below for anything left unmapped.`,
      );
    }

    const sourceTotal = isCurrent ? bsTotalCurrent : bsTotalPrevious;
    if (sourceTotal != null && Math.abs(sourceTotal - totalLiabilitiesSide) > 1) {
      messages.push(
        `For ${label}, the recomputed total ${money(totalLiabilitiesSide)} does not match the source workbook's own Balance Sheet total ${money(sourceTotal)}.`,
      );
    }
  });

  return messages;
}
