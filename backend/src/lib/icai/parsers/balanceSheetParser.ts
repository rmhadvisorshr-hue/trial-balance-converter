import type { Grid } from "./grid";

// The Balance Sheet sheet is mostly a re-presentation of data already
// captured from the trading/capital/schedule/fixed-asset sheets, but it also
// commonly carries a few items that appear NOWHERE else (bank loans,
// deposits, a cash/bank break-up, GST/advance-tax balances). This file reads
// it as a two-column hierarchical listing - Liabilities on the left,
// Assets on the right, each with group header rows (no amount on that row)
// followed by leaf rows (an amount) until the next group header - without
// assuming which groups are "new" vs "already captured elsewhere"; that
// dedup happens by amount match in normalizeWorkbook.ts instead.

export interface BsLeaf {
  side: "liability" | "asset";
  group: string;
  label: string;
  amount: number;
}

function parseColumn(grid: Grid, labelCol: number, amtCol: number, side: "liability" | "asset"): BsLeaf[] {
  const out: BsLeaf[] = [];
  let group = "";
  for (const row of grid.rows) {
    const label = row.cells[labelCol] ?? "";
    if (!label) continue;
    const amount = row.nums[amtCol];
    if (amount == null) {
      group = label;
      continue;
    }
    out.push({ side, group: group || label, label, amount });
  }
  return out;
}

// Finds the "LIABILITIES | AMOUNT | ... | ASSETS | AMOUNT" header and reads
// both columns beneath it.
export function parseBalanceSheetLeaves(grid: Grid): BsLeaf[] {
  for (let i = 0; i < grid.rows.length; i++) {
    const row = grid.rows[i];
    const liabCol = row.cells.findIndex((t, c) => c > 0 && /^liabilities$/i.test(t));
    const assetCol = row.cells.findIndex((t, c) => c > 0 && /^assets$/i.test(t));
    if (liabCol < 0 || assetCol < 0) continue;
    const amountCols = row.cells
      .map((t, c) => (c > 0 && /^amount$/i.test(t) ? c : -1))
      .filter((c) => c > 0);
    const liabAmtCol = amountCols.find((c) => c > liabCol && c < assetCol) ?? liabCol + 1;
    const assetAmtCol = amountCols.find((c) => c > assetCol) ?? assetCol + 1;

    const subGrid: Grid = { rows: grid.rows.slice(i + 1) };
    return [
      ...parseColumn(subGrid, liabCol, liabAmtCol, "liability"),
      ...parseColumn(subGrid, assetCol, assetAmtCol, "asset"),
    ];
  }
  return [];
}

// Used only as a soft cross-check (validation/crossCheck.ts): a balanced
// sheet's own grand total row is - by definition - the row where the
// Liabilities total equals the Assets total, and in practice it's also the
// single largest number on the sheet.
export function findBalanceSheetGrandTotal(grid: Grid): number | null {
  let best: number | null = null;
  for (const row of grid.rows) {
    const vals = row.nums.filter((n): n is number => n != null && Math.abs(n) > 1000);
    for (let i = 0; i < vals.length; i++) {
      for (let j = i + 1; j < vals.length; j++) {
        if (Math.abs(vals[i] - vals[j]) < 1 && (best == null || Math.abs(vals[i]) > best)) {
          best = Math.abs(vals[i]);
        }
      }
    }
  }
  return best;
}
