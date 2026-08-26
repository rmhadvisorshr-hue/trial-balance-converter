import type { Grid } from "./grid";
import { rowText, GRAND_TOTAL_RE } from "./grid";

export interface ScheduleLine {
  label: string;
  amount: number;
}

export interface ScheduleBlock {
  title: string; // e.g. "SUNDRY CREDITORS- SCH-I"
  lines: ScheduleLine[];
  total: number | null; // the sheet's own "Total" row, for cross-check
}

// A sheet can hold several single-column Sr.No/Particulars/Amount schedules
// stacked one after another (creditors, then customer advances, then loans
// to contractors, ...). Each is bounded by its own "Particulars"/"Amount"
// header above and a "Total" row below.
export function parseScheduleBlocks(grid: Grid): ScheduleBlock[] {
  const headers: { idx: number; labelCol: number; amtCol: number }[] = [];
  for (let i = 0; i < grid.rows.length; i++) {
    const row = grid.rows[i];
    const labelCol = row.cells.findIndex((t, c) => c > 0 && /^particulars$/i.test(t));
    if (labelCol < 0) continue;
    const amtCol = row.cells.findIndex((t, c) => c > labelCol && /^amount$/i.test(t));
    if (amtCol < 0) continue;
    headers.push({ idx: i, labelCol, amtCol });
  }

  const blocks: ScheduleBlock[] = [];
  for (let hi = 0; hi < headers.length; hi++) {
    const { idx, labelCol, amtCol } = headers[hi];
    const nextIdx = headers[hi + 1]?.idx ?? grid.rows.length;
    const titleRow = idx > 0 ? grid.rows[idx - 1] : undefined;
    const title = titleRow ? rowText(titleRow) : "";

    const lines: ScheduleLine[] = [];
    let total: number | null = null;
    for (let r = idx + 1; r < nextIdx; r++) {
      const row = grid.rows[r];
      const label = row.cells[labelCol] ?? "";
      const amount = row.nums[amtCol];
      if (!label) continue;
      if (GRAND_TOTAL_RE.test(label)) {
        if (amount != null) total = amount;
        break;
      }
      if (amount == null) continue;
      lines.push({ label, amount });
    }
    blocks.push({ title, lines, total });
  }
  return blocks;
}
