import type { Grid } from "./grid";
import { parseTBlocks } from "./tAccountParser";
import { parseScheduleBlocks } from "./scheduleParser";
import { parseFixedAssetSchedule } from "./fixedAssetParser";

export type SheetKind =
  | "tAccount" // trading/P&L/appropriation/capital blocks - see blockRole.ts for per-block detail
  | "schedule" // party-wise Sr.No/Particulars/Amount listings
  | "fixedAssets" // WDV/gross-block roll-forward
  | "balanceSheet" // two-column Liabilities|Assets - cross-check only
  | "workingPaper" // not part of the financial statements (tax computation, etc.)
  | "unknown"; // no recognizable financial-statement shape - excluded, flagged for the CA

export interface SheetClassification {
  kind: SheetKind;
  reason: string;
}

// Deliberately shape-based, not name-based: every check below asks "does
// this sheet's *layout* match a financial-statement shape" by re-running the
// real parser for that shape, never "is the sheet named X". A sheet that
// doesn't match any recognized shape (a tax computation, an interest
// working, a unit-wise sales tracker, ...) simply falls through to
// "unknown" and is excluded by default - there is no negative keyword list
// to keep in sync with every possible working-paper sheet name a firm might
// use.
export function classifySheet(grid: Grid): SheetClassification {
  if (grid.rows.length === 0) return { kind: "workingPaper", reason: "Empty sheet." };

  const headText = grid.rows
    .slice(0, 15)
    .map((r) => r.cells.join(" "))
    .join(" ");

  if (parseFixedAssetSchedule(grid).length > 0) {
    return { kind: "fixedAssets", reason: "Rate/Opening/Depreciation/Closing roll-forward columns found." };
  }
  if (parseTBlocks(grid).some((b) => b.lines.length > 0)) {
    return { kind: "tAccount", reason: "Dr/Cr 'Particulars | Amount | Particulars | Amount' header found." };
  }
  if (parseScheduleBlocks(grid).some((b) => b.lines.length > 0)) {
    return { kind: "schedule", reason: "Sr.No/Particulars/Amount listing found." };
  }
  if (/balance\s*sheet/i.test(headText)) {
    return { kind: "balanceSheet", reason: "Heading contains 'Balance Sheet'." };
  }
  if (/computation of total income/i.test(headText)) {
    return { kind: "workingPaper", reason: "Income-tax computation working, not a financial statement." };
  }
  return { kind: "unknown", reason: "No recognizable financial-statement layout on this sheet." };
}
