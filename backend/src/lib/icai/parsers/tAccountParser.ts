import type { Grid } from "./grid";
import { rowText, GRAND_TOTAL_RE } from "./grid";

export interface TLine {
  label: string;
  amount: number;
  side: "dr" | "cr"; // dr = the "To ..." column, cr = the "By ..." column
}

export interface TBlock {
  title: string; // heading line immediately above the block, e.g. "TRADING ACCOUNT FOR THE YEAR ENDED..."
  // Every line including brought/carried-forward balances ("To Balance c/d",
  // "By Opening Balance") - only a literal Total/Grand Total row is dropped
  // here. Whether a b/d or c/d line is noise (a Trading A/c's balancing
  // figure) or the actual figure needed (a capital or WIP account's opening/
  // closing balance) depends on what kind of account the block is, which
  // only the caller (the per-role interpreter) knows - so filtering that out
  // is the caller's job, via grid.ts's CARRY_BALANCE_RE, not this parser's.
  lines: TLine[];
}

function stripPrefix(label: string): string {
  return label.replace(/^(to|by)\s+/i, "").trim();
}

// A T-account particular is never just a bare number - some source sheets
// carry a "% of total" or a mirrored total figure in a stray cell that
// happens to line up with the particulars column (e.g. a gross-profit
// percentage sitting in the credit-label column on the same row as the
// column's grand-total mirror). Filtering those out by "is this text
// actually a number" is more robust than trying to special-case the row it
// happens to land on.
function isNumericLabel(label: string): boolean {
  const s = label.replace(/,/g, "").replace(/%$/, "").trim();
  return s !== "" && Number.isFinite(Number(s));
}

// Finds every "PARTICULARS ... AMOUNT ... PARTICULARS ... AMOUNT" T-account
// header in a sheet and parses the block that follows each one, up to the
// next header (or end of sheet). A single sheet commonly holds several
// T-accounts stacked one after another (e.g. Trading A/c, then a WIP A/c,
// then a Sales A/c all on the same "trad" sheet) - each gets its own block.
export function parseTBlocks(grid: Grid): TBlock[] {
  const headerIdx: { idx: number; labelCol: number; amtCol: number }[] = [];

  for (let i = 0; i < grid.rows.length; i++) {
    const row = grid.rows[i];
    const particularsCols = row.cells
      .map((t, c) => (c > 0 && /^particulars$/i.test(t) ? c : -1))
      .filter((c) => c > 0);
    if (particularsCols.length < 2) continue;
    const amountCols = row.cells
      .map((t, c) => (c > 0 && /^amount$/i.test(t) ? c : -1))
      .filter((c) => c > 0);
    for (const labelCol of particularsCols) {
      const amtCol = amountCols.find((c) => c > labelCol) ?? labelCol + 1;
      headerIdx.push({ idx: i, labelCol, amtCol });
    }
  }
  if (headerIdx.length === 0) return [];

  // Group header column-pairs that belong to the same header row into one block.
  const blocks: TBlock[] = [];
  const byRow = new Map<number, { labelCol: number; amtCol: number }[]>();
  for (const h of headerIdx) {
    const list = byRow.get(h.idx) ?? [];
    list.push({ labelCol: h.labelCol, amtCol: h.amtCol });
    byRow.set(h.idx, list);
  }
  const headerRowIdxs = [...byRow.keys()].sort((a, b) => a - b);

  for (let hi = 0; hi < headerRowIdxs.length; hi++) {
    const idx = headerRowIdxs[hi];
    const pairs = byRow.get(idx)!;
    const nextIdx = headerRowIdxs[hi + 1] ?? grid.rows.length;
    const titleRow = idx > 0 ? grid.rows[idx - 1] : undefined;
    const title = titleRow ? rowText(titleRow) : "";

    const lines: TLine[] = [];
    for (let r = idx + 1; r < nextIdx; r++) {
      const row = grid.rows[r];
      pairs.forEach((pair, pairIdx) => {
        const label = row.cells[pair.labelCol] ?? "";
        const amount = row.nums[pair.amtCol];
        if (!label || amount == null) return;
        if (GRAND_TOTAL_RE.test(label) || isNumericLabel(label)) return;
        lines.push({ label: stripPrefix(label), amount, side: pairIdx === 0 ? "dr" : "cr" });
      });
    }
    blocks.push({ title, lines });
  }

  return blocks;
}

export function blockTotal(block: TBlock, side: "dr" | "cr"): number {
  return block.lines.filter((l) => l.side === side).reduce((s, l) => s + l.amount, 0);
}
