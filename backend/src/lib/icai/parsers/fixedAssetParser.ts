import type { Grid } from "./grid";
import { GRAND_TOTAL_RE } from "./grid";

export interface RawFixedAssetRow {
  name: string;
  rate: number;
  opening: number;
  additions: number;
  depreciation: number;
  closing: number;
}

// Fixed asset schedules in practice are a WDV/tax-block roll-forward:
// Particulars | Rate | Opening | <one or more addition columns> | ... |
// Depreciation | Closing. The addition columns vary (a single "Additions
// during the year" column, or a pair split by the Income-tax Act's half-year
// cutoff date) - so additions are computed as "every numeric column between
// Opening and Depreciation that isn't itself a running Total", not by name.
export function parseFixedAssetSchedule(grid: Grid): RawFixedAssetRow[] {
  let headerIdx = -1;
  let particularsCol = -1;
  let rateCol = -1;
  let openingCol = -1;
  let depreciationCol = -1;
  let closingCol = -1;

  for (let i = 0; i < grid.rows.length; i++) {
    const row = grid.rows[i];
    const pCol = row.cells.findIndex((t, c) => c > 0 && /particulars|asset/i.test(t));
    const rCol = row.cells.findIndex((t, c) => c > 0 && /^rate$/i.test(t));
    const oCol = row.cells.findIndex((t, c) => c > 0 && /opening/i.test(t));
    const dCol = row.cells.findIndex((t, c) => c > 0 && /depreciation/i.test(t));
    const cCol = row.cells.findIndex((t, c) => c > 0 && /closing/i.test(t));
    if (pCol > 0 && rCol > 0 && oCol > 0 && dCol > 0 && cCol > 0) {
      headerIdx = i;
      particularsCol = pCol;
      rateCol = rCol;
      openingCol = oCol;
      depreciationCol = dCol;
      closingCol = cCol;
      break;
    }
  }
  if (headerIdx < 0) return [];

  const headerRow = grid.rows[headerIdx];
  const totalCol = headerRow.cells.findIndex((t, c) => c > 0 && /^total$/i.test(t));
  const additionCols: number[] = [];
  for (let c = openingCol + 1; c < depreciationCol; c++) {
    if (c === totalCol) continue;
    additionCols.push(c);
  }

  const out: RawFixedAssetRow[] = [];
  for (let r = headerIdx + 1; r < grid.rows.length; r++) {
    const row = grid.rows[r];
    const name = row.cells[particularsCol] ?? "";
    if (!name) continue;
    if (GRAND_TOTAL_RE.test(name) || /^total\s*-/i.test(name)) continue;
    const closing = row.nums[closingCol];
    if (closing == null) continue;
    const additions = additionCols.reduce((s, c) => s + (row.nums[c] ?? 0), 0);
    out.push({
      name,
      rate: row.nums[rateCol] ?? 0,
      opening: row.nums[openingCol] ?? 0,
      additions,
      depreciation: row.nums[depreciationCol] ?? 0,
      closing,
    });
  }
  return out;
}
