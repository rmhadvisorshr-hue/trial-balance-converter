import type { WorkbookAnalysis, IcaiNoteCode, NormalizedAccount } from "../types";

export type YearKey = "amountCurrent" | "amountPrevious";

export function accountsFor(analysis: WorkbookAnalysis, code: IcaiNoteCode): NormalizedAccount[] {
  return analysis.accounts.filter((a) => a.code === code);
}

export function totalFor(analysis: WorkbookAnalysis, code: IcaiNoteCode, year: YearKey): number {
  return accountsFor(analysis, code).reduce((s, a) => s + a[year], 0);
}

export function totalMany(analysis: WorkbookAnalysis, codes: IcaiNoteCode[], year: YearKey): number {
  return codes.reduce((s, c) => s + totalFor(analysis, c, year), 0);
}

export function ownersFundsTotal(analysis: WorkbookAnalysis, year: "closingBalance" | "closingBalancePrevious"): number {
  return analysis.owners.reduce((s, o) => s + o[year], 0);
}

export function fixedAssetsTotal(analysis: WorkbookAnalysis, year: "current" | "previous"): number {
  return analysis.fixedAssets[year].reduce((s, a) => s + a.closing, 0);
}

// Depreciation for the year ties to Note 11's own roll-forward whenever a
// fixed asset schedule was found, rather than to however the P&L account
// happened to label its depreciation line that year (which can drift - e.g.
// "Depreciation" one year, "Depreciation account" the next - and would
// otherwise fail to merge across years). Falls back to the raw P&L-sourced
// figure only when no fixed asset schedule was parsed at all.
export function depreciationTotal(analysis: WorkbookAnalysis, year: YearKey): number {
  const hasFixedAssets = analysis.fixedAssets.current.length > 0 || analysis.fixedAssets.previous.length > 0;
  if (hasFixedAssets) {
    const rows = year === "amountCurrent" ? analysis.fixedAssets.current : analysis.fixedAssets.previous;
    return round2(rows.reduce((s, a) => s + a.depreciation, 0));
  }
  return totalFor(analysis, "N25_DEPRECIATION", year);
}

export const round2 = (n: number) => Math.round((n + Number.EPSILON) * 100) / 100;

// Cost of materials consumed is the ICAI-prescribed title; a business whose
// cost lines read like a construction/development trade (contract/labour/
// development charges) gets the more descriptive title the reference case
// used, derived entirely from the line names actually present - never from
// the entity name or any hardcoded business-type flag.
export function costNoteTitle(analysis: WorkbookAnalysis): string {
  const names = accountsFor(analysis, "N21_COST_OF_CONSTRUCTION").map((a) => a.name.toLowerCase());
  const isConstruction = names.some((n) => /contract|construction|labour|development/.test(n));
  return isConstruction ? "Cost of Construction and Development Expenses" : "Cost of Materials Consumed";
}
