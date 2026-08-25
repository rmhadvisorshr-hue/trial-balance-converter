import type ExcelJS from "exceljs";
import type { BuildCtx } from "../common";
import type { Statements } from "../../statements";
import { setCell, setFormula, applyColumnWidths, topBorder } from "../../excel/helpers";
import { round2 } from "../../classify";

const FACE_VALUE = 10; // Rs. per share - a TB has no share-count/face-value data, so this is
// the standard Schedule III default; flagged for the CA to confirm.
const MAX_SHAREHOLDER_ROWS = 6;

export interface ShareCapitalReservesResult {
  shareCapitalRow: number; // Note 2 total (issued capital) row
  reservesRow: number; // Note 3 total (closing reserves) row
}

// Note 2 (Share Capital) + Note 3 (Reserves & Surplus). Kept at a fixed row
// layout (shareholder list capped at MAX_SHAREHOLDER_ROWS, truncated with a
// "+N others" line) so P&L!D<netProfitRow> below can reference this sheet's
// Note 3 row via a plain constant instead of needing a dynamic lookup - the
// only genuinely variable-length sheet in this generator is NOTES.
export function writeShareCapitalReservesSheet(
  ws: ExcelJS.Worksheet,
  ctx: BuildCtx,
  netProfitFormulaRef: string, // e.g. "'P&L'!D41" - the P&L's fixed Net Profit cell
): ShareCapitalReservesResult {
  applyColumnWidths(ws, [46, 18, 18, 14]);
  const { meta, st, divisor } = ctx;
  const money = (v: number) => v / divisor;

  let r = 2;
  setCell(ws, r++, 1, meta.firmName || "COMPANY NAME", { bold: true });
  if (meta.cin) setCell(ws, r++, 1, `CIN : ${meta.cin}`, {});
  setCell(ws, r++, 1, `Notes to Financial Statements for the year ended ${meta.periodLabel || ""}`, {
    bold: true,
  });
  r++;

  setCell(ws, r++, 1, 'Note "2" : SHARE CAPITAL', { bold: true });
  setCell(ws, r++, 1, `As at ${meta.asAtLabel || meta.periodLabel || ""}`, { italic: true });
  r++;

  const capitalTotal = round2(st.total("CAPITAL"));
  const issuedShares = Math.round(capitalTotal / FACE_VALUE);

  setCell(ws, r++, 1, "Authorized Shares");
  setCell(ws, r, 1, `Equity Shares of Rs.${FACE_VALUE} each`);
  setCell(ws, r++, 3, 0, { money: false, align: "right" });
  setCell(ws, r++, 1, "(authorized share count is not present in a trial balance - enter manually)", {
    italic: true,
  });
  r++;

  setCell(ws, r++, 1, "Issued, Subscribed & Fully Paid up Shares");
  const issuedRow = r;
  setCell(ws, r, 1, `Equity Shares of Rs.${FACE_VALUE} each fully paid`);
  setCell(ws, r, 2, issuedShares, { align: "right" });
  setCell(ws, r++, 3, capitalTotal, { money: true });
  setCell(
    ws,
    r++,
    1,
    `(share count estimated as Capital / Rs.${FACE_VALUE} face value - confirm against the share register; the Rupee amount above is the actual TB balance)`,
    { italic: true },
  );
  r++;

  setCell(ws, r, 1, "Total (Note 2)", { bold: true });
  const shareCapitalRow = r;
  setFormula(ws, r, 3, `C${issuedRow}`, capitalTotal, { bold: true, money: true });
  topBorder(ws, r, 1, 4);
  r += 2;

  setCell(ws, r++, 1, "a. Reconciliation of shares outstanding at the beginning and end of the year");
  setCell(ws, r, 1, "Shares outstanding at the beginning of the year");
  setFormula(ws, r++, 3, `B${issuedRow}`, issuedShares, { align: "right" });
  setCell(ws, r, 1, "Shares issued during the year");
  setCell(ws, r++, 3, 0, { align: "right" });
  setCell(ws, r, 1, "Shares bought back during the year");
  setCell(ws, r++, 3, 0, { align: "right" });
  const reconStart = r - 3;
  const reconEnd = r - 1;
  setCell(ws, r, 1, "Shares outstanding at the end of the year", { bold: true });
  setFormula(ws, r, 3, `SUM(C${reconStart}:C${reconEnd})`, issuedShares, { bold: true, align: "right" });
  r += 2;

  setCell(ws, r++, 1, "b. Terms/Rights attached to Equity Shares");
  setCell(
    ws,
    r++,
    1,
    "The Company has only one class of equity shares. Each holder is entitled to one vote per share. In the event of liquidation, equity shareholders will be entitled to receive the remaining assets of the Company, in proportion to the number of shares held.",
    { italic: true },
  );
  r++;

  setCell(ws, r++, 1, "c. Shares held by holding/ultimate holding company and/or their subsidiaries - NIL");
  r++;

  setCell(ws, r++, 1, "d. Details of shareholders holding more than 5% shares in the company", {
    bold: true,
  });
  setCell(ws, r, 1, "Name of Shareholder", { bold: true });
  setCell(ws, r, 3, "No. of Shares", { bold: true, align: "right" });
  setCell(ws, r++, 4, "% Holding", { bold: true, align: "right" });

  const shareholders = capitalShareholders(st, issuedShares);
  for (let i = 0; i < MAX_SHAREHOLDER_ROWS; i++) {
    const sh = shareholders[i];
    if (sh) {
      setCell(ws, r, 1, sh.name);
      setCell(ws, r, 3, sh.shares, { align: "right" });
      setCell(ws, r++, 4, sh.pct, { align: "right" });
    } else if (i === 0 && shareholders.length === 0) {
      setCell(ws, r++, 1, "(refer share register - capital ledgers were not individually named)", {
        italic: true,
      });
    } else {
      r++;
    }
  }
  if (shareholders.length > MAX_SHAREHOLDER_ROWS) {
    setCell(
      ws,
      r++,
      1,
      `+ ${shareholders.length - MAX_SHAREHOLDER_ROWS} more - refer share register`,
      { italic: true },
    );
  }
  r++;

  setCell(
    ws,
    r++,
    1,
    "e. Aggregate number of bonus shares issued, shares issued for consideration other than cash and shares bought back during the 5 years preceding the reporting date - NIL",
  );
  r++;

  setCell(ws, r++, 1, "f. Shareholding of Promoters and % change during the year", { bold: true });
  setCell(
    ws,
    r++,
    1,
    "(assumes shareholders above are promoters - confirm and split non-promoter holders if any)",
    { italic: true },
  );
  r += 2;

  setCell(ws, r++, 1, 'Note "3" : RESERVES & SURPLUS', { bold: true });
  setCell(ws, r++, 1, "Surplus");
  // The TB's own accumulated "Profit & Loss A/c" / General Reserve ledger
  // (RESERVES code) already holds everything carried forward from prior
  // years - that IS the opening balance, not 0. Only the split between
  // "how much of that opening figure is prior-year vs. even-older years" is
  // unavailable (this app is single-period), which doesn't matter for the
  // total.
  const openingReserves = round2(st.total("RESERVES"));
  setCell(ws, r, 1, "Opening balance (per trial balance)");
  setCell(ws, r++, 3, money(openingReserves), { money: true });
  const openingRow = r - 1;

  const netProfit = st.netProfit;
  setCell(ws, r, 1, "(+) Net Profit/(Loss) for the current year");
  const netProfitRow = r;
  setFormula(ws, r++, 3, netProfitFormulaRef, money(netProfit), { money: true });
  setCell(ws, r, 1, "(+) Opening Balance Difference");
  setCell(ws, r++, 3, 0, { money: true });
  const diffRow = r - 1;

  r++;
  setCell(ws, r, 1, "Total (Note 3 - Closing Balance)", { bold: true });
  const reservesRow = r;
  setFormula(
    ws,
    r,
    3,
    `C${openingRow}+C${netProfitRow}+C${diffRow}`,
    money(round2(openingReserves + netProfit + 0)),
    { bold: true, money: true },
  );
  topBorder(ws, r, 1, 4);

  return { shareCapitalRow, reservesRow };
}

export function capitalShareholders(
  st: Statements,
  totalShares: number,
): { name: string; shares: number; pct: string }[] {
  const items = st.byCode.get("CAPITAL") ?? [];
  // Only treat this as a per-shareholder breakdown when every capital ledger
  // carries a distinguishable name (not a single lump "Share Capital" /
  // "Capital Account" ledger) - otherwise there's nothing to list.
  if (items.length < 2) return [];
  const total = items.reduce((s, i) => s + i.amount, 0);
  if (total <= 0) return [];
  return items
    .filter((i) => i.amount > 0)
    .sort((a, b) => b.amount - a.amount)
    .map((i) => ({
      name: i.name,
      shares: Math.round((i.amount / total) * totalShares),
      pct: `${((i.amount / total) * 100).toFixed(2)}%`,
    }));
}
