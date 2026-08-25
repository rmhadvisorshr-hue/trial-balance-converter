import type ExcelJS from "exceljs";
import type { BuildCtx } from "../common";
import { setCell, setFormula, applyColumnWidths, topBorder } from "../../excel/helpers";
import { round2 } from "../../classify";

export interface FixedAssetScheduleResult {
  closingTotalRow: number; // SCH-5 (Fixed Assets) Total row, column G - for BS
  depForYearRow: number; // same Total row, column F (depreciation) - for P&L
}

// SCH-5 Fixed Assets: asset-by-asset, same approach as the Pvt Ltd '11. FA'
// register - a trial balance only gives the closing net book value per
// asset ledger, never the WDV rate, opening balance, or pre/post-3-Oct
// addition split (the reference file's own rates/openings are typed-in
// figures, not TB-derived even there). Rate/Opening/Additions default to 0
// (manual); Closing is the known TB value per asset, and one "unallocated"
// row carries the real total depreciation from the TB's P&L Depreciation
// ledger, so the schedule's totals still tie to real figures.
export function writeFixedAssetSchedule(
  ws: ExcelJS.Worksheet,
  ctx: BuildCtx,
): FixedAssetScheduleResult {
  applyColumnWidths(ws, [34, 8, 14, 16, 16, 14, 16]);
  const { meta, st, divisor } = ctx;
  const money = (v: number) => v / divisor;
  const ppeItems = st.byCode.get("PPE") ?? [];

  let r = 2;
  setCell(ws, r++, 1, meta.firmName || "FIRM NAME", { bold: true });
  setCell(ws, r++, 1, meta.periodLabel || "", { italic: true });
  if (ctx.unitHeading) setCell(ws, r++, 1, ctx.unitHeading, { italic: true });
  r++;
  setCell(ws, r++, 1, "SCH-5  FIXED ASSETS", { bold: true });
  ["Particulars", "Rate", "Opening Balance", "Addition upto 3 Oct", "Addition after 3 Oct", "Dep / Round Off", "Closing Balance"].forEach(
    (h, i) => setCell(ws, r, i + 1, h, { bold: true, align: i === 0 ? "left" : "right", wrap: true }),
  );
  r++;

  const start = r;
  for (const a of ppeItems) {
    setCell(ws, r, 1, a.name);
    setCell(ws, r, 2, 0);
    setCell(ws, r, 3, 0);
    setCell(ws, r, 4, 0);
    setCell(ws, r, 5, 0);
    setCell(ws, r, 6, 0);
    setCell(ws, r, 7, money(round2(a.amount)), { money: true }); // known: TB closing balance
    r++;
  }
  setCell(ws, r, 1, "Depreciation per Trial Balance (unallocated across the assets above)", {
    italic: true,
  });
  setCell(ws, r, 2, 0);
  setCell(ws, r, 3, 0);
  setCell(ws, r, 4, 0);
  setCell(ws, r, 5, 0);
  setCell(ws, r, 6, money(round2(st.depreciation)), { money: true }); // known: TB P&L depreciation charge
  setCell(ws, r, 7, 0);
  r++;
  const end = r - 1;

  r++;
  setCell(ws, r, 1, "Total", { bold: true });
  const closingTotal = round2(ppeItems.reduce((s, a) => s + a.amount, 0));
  setFormula(ws, r, 3, `SUM(C${start}:C${end})`, 0, { money: true });
  setFormula(ws, r, 4, `SUM(D${start}:D${end})`, 0, { money: true });
  setFormula(ws, r, 5, `SUM(E${start}:E${end})`, 0, { money: true });
  setFormula(ws, r, 6, `SUM(F${start}:F${end})`, money(round2(st.depreciation)), { bold: true, money: true });
  setFormula(ws, r, 7, `SUM(G${start}:G${end})`, money(closingTotal), { bold: true, money: true });
  topBorder(ws, r, 1, 7);

  return { closingTotalRow: r, depForYearRow: r };
}
