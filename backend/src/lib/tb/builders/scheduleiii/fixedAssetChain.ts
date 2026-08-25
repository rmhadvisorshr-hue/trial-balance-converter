import type ExcelJS from "exceljs";
import type { BuildCtx } from "../common";
import { setCell, setFormula, applyColumnWidths, topBorder } from "../../excel/helpers";
import { round2 } from "../../classify";

const DEFAULT_TAX_RATE = 0.26; // 25% + 4% cess, the common corporate rate; editable.

export interface FixedAssetChainResult {
  faNetClosingRow: number; // '11. FA' Total row, column J (Net Block Closing) - for BS PPE
  faDepForYearRow: number; // '11. FA' Total row, column G (Depreciation for the year) - for P&L
  dtaNetExpenseRow: number; // 'DTA' Total row - for P&L "Deferred tax" and BS Notes 6/13
}

// 11. FA -> DTL -> DTA, formula-chained exactly as in the reference format.
// A trial balance only gives the *closing net book value* per fixed-asset
// ledger, never the Gross Block / Accumulated Depreciation split, opening
// balances, additions/deletions, or Income-Tax-Act depreciation - those
// columns are left as editable 0s. Depreciation as per Income Tax defaults
// to *equal* Depreciation as per Books (zero timing difference) rather than
// zero, so the default Deferred Tax figure is a clearly-flagged "no
// movement yet" placeholder instead of a silently wrong large number.
export function writeFixedAssetChain(
  faWs: ExcelJS.Worksheet,
  dtlWs: ExcelJS.Worksheet,
  dtaWs: ExcelJS.Worksheet,
  ctx: BuildCtx,
): FixedAssetChainResult {
  const { meta, st, divisor } = ctx;
  const money = (v: number) => v / divisor;
  const ppeItems = st.byCode.get("PPE") ?? [];

  // ---------------- 11. FA ----------------
  applyColumnWidths(faWs, [32, 13, 13, 13, 13, 13, 13, 13, 13, 14, 14]);
  let r = 2;
  setCell(faWs, r++, 1, meta.firmName || "COMPANY NAME", { bold: true });
  if (meta.cin) setCell(faWs, r++, 1, `CIN : ${meta.cin}`, {});
  setCell(faWs, r++, 1, `Notes to Financial Statements for the year ended ${meta.periodLabel || ""}`, {
    bold: true,
  });
  r++;
  setCell(faWs, r++, 1, 'Note "11" : PROPERTY, PLANT & EQUIPMENT AND INTANGIBLE ASSETS', { bold: true });
  if (ctx.unitHeading) setCell(faWs, r++, 1, ctx.unitHeading, { italic: true });
  r++;
  ["Particular", "Gross Block", "", "", "", "Accumulated Depreciation", "", "", "", "Net Block", ""].forEach(
    (h, i) => h && setCell(faWs, r, i + 1, h, { bold: true, align: "center" }),
  );
  r++;
  [
    "",
    "Opening",
    "Additions",
    "Deletions",
    "Closing",
    "Opening",
    "For the year",
    "On disposals",
    "Closing",
    "Closing (Net)",
    "Opening (Net, prior)",
  ].forEach((h, i) => setCell(faWs, r, i + 1, h, { bold: true, align: i === 0 ? "left" : "right", wrap: true }));
  r++;

  const faStart = r;
  for (const a of ppeItems) {
    setCell(faWs, r, 1, a.name);
    setCell(faWs, r, 2, 0);
    setCell(faWs, r, 3, 0);
    setCell(faWs, r, 4, 0);
    setFormula(faWs, r, 5, `B${r}+C${r}-D${r}`, 0, { money: true });
    setCell(faWs, r, 6, 0);
    setCell(faWs, r, 7, 0);
    setCell(faWs, r, 8, 0);
    setFormula(faWs, r, 9, `F${r}+G${r}-H${r}`, 0, { money: true });
    setCell(faWs, r, 10, money(round2(a.amount)), { money: true }); // known: TB net book value
    setFormula(faWs, r, 11, `B${r}-F${r}`, 0, { money: true });
    r++;
  }
  const depRow = r;
  setCell(faWs, r, 1, "Depreciation per Trial Balance (unallocated across the assets above)", {
    italic: true,
  });
  for (const c of [2, 3, 4, 5, 6, 8, 9, 10, 11]) setCell(faWs, r, c, 0);
  setCell(faWs, r, 7, money(round2(st.depreciation)), { money: true }); // known: TB P&L depreciation charge
  r++;
  const faEnd = r - 1;

  r++;
  setCell(faWs, r, 1, "Total", { bold: true });
  for (const c of [2, 3, 4, 5, 6, 7, 8, 9, 10, 11]) {
    const col = String.fromCharCode(64 + c);
    const known =
      c === 7
        ? round2(st.depreciation)
        : c === 10
          ? round2(ppeItems.reduce((s, a) => s + a.amount, 0))
          : 0;
    setFormula(faWs, r, c, `SUM(${col}${faStart}:${col}${faEnd})`, money(known), {
      bold: true,
      money: true,
    });
  }
  topBorder(faWs, r, 1, 11);
  const faTotalRow = r;

  // ---------------- DTL ----------------
  applyColumnWidths(dtlWs, [32, 13, 13, 13, 13, 13, 13, 13, 13, 14, 14]);
  r = 2;
  setCell(dtlWs, r++, 1, meta.firmName || "COMPANY NAME", { bold: true });
  if (meta.cin) setCell(dtlWs, r++, 1, `CIN : ${meta.cin}`, {});
  setCell(dtlWs, r++, 1, `Notes to Financial Statements for the year ended ${meta.periodLabel || ""}`, {
    bold: true,
  });
  r++;
  setCell(dtlWs, r++, 1, "AS PER BOOKS OF ACCOUNTS", { bold: true });
  setCell(dtlWs, r++, 1, 'Note "9" : PROPERTY, PLANT & EQUIPMENT AND INTANGIBLE ASSETS', { bold: true });
  r++;
  const dtlBooksStart = r;
  for (let i = 0; i < ppeItems.length; i++) {
    const srcRow = faStart + i;
    setCell(dtlWs, r, 1, ppeItems[i].name);
    for (const c of [2, 3, 4, 5, 6, 7, 8, 9, 10, 11]) {
      const col = String.fromCharCode(64 + c);
      const known = c === 10 ? money(round2(ppeItems[i].amount)) : 0;
      setFormula(dtlWs, r, c, `+'11. FA'!${col}${srcRow}`, known, { money: true });
    }
    r++;
  }
  setCell(dtlWs, r, 1, "Depreciation per Trial Balance (unallocated across the assets above)", {
    italic: true,
  });
  for (const c of [2, 3, 4, 5, 6, 8, 9, 10, 11]) {
    const col = String.fromCharCode(64 + c);
    setFormula(dtlWs, r, c, `+'11. FA'!${col}${depRow}`, 0, { money: true });
  }
  setFormula(dtlWs, r, 7, `+'11. FA'!G${depRow}`, money(round2(st.depreciation)), { money: true });
  r++;
  const dtlBooksEnd = r - 1;
  r++;
  setCell(dtlWs, r, 1, "Total", { bold: true });
  for (const c of [2, 3, 4, 5, 6, 7, 8, 9, 10, 11]) {
    const col = String.fromCharCode(64 + c);
    const known = c === 7 ? round2(st.depreciation) : c === 10 ? round2(ppeItems.reduce((s, a) => s + a.amount, 0)) : 0;
    setFormula(dtlWs, r, c, `SUM(${col}${dtlBooksStart}:${col}${dtlBooksEnd})`, money(known), {
      bold: true,
      money: true,
    });
  }
  topBorder(dtlWs, r, 1, 11);
  const dtlBooksTotalRow = r;
  r += 2;

  setCell(dtlWs, r++, 1, "AS PER INCOME TAX ACT", { bold: true });
  setCell(
    dtlWs,
    r++,
    1,
    "Income-tax depreciation (WDV method, block-wise per the Income Tax Rules) cannot be derived from a trial balance - it is not tracked block-wise in Tally. Enter the actual IT Act depreciation for the year below.",
    { italic: true },
  );
  setCell(dtlWs, r, 1, "Depreciation for the year (enter manually)");
  const dtlItActRow = r;
  // Defaults to the books figure (zero timing difference) rather than 0, so
  // the Deferred Tax computation below defaults to "no movement yet" instead
  // of a silently wrong large number - see DTA below.
  setFormula(dtlWs, r, 7, `G${dtlBooksTotalRow}`, money(round2(st.depreciation)), { money: true });
  r += 2;

  // ---------------- DTA ----------------
  applyColumnWidths(dtaWs, [4, 40, 16, 16]);
  r = 2;
  setCell(dtaWs, r++, 1, meta.firmName || "COMPANY NAME", { bold: true });
  setCell(dtaWs, r++, 2, "Deferred Tax Calculation", { bold: true });
  r++;
  setCell(dtaWs, r, 2, "Tax rate (edit if different)");
  const taxRateRow = r;
  setCell(dtaWs, r++, 3, DEFAULT_TAX_RATE, { align: "right" });
  r++;

  setCell(dtaWs, r, 1, "(A)");
  setCell(dtaWs, r++, 2, "Fixed Assets", { bold: true });
  setCell(dtaWs, r, 2, "Depreciation as per Books");
  const booksDepRefRow = r;
  setFormula(dtaWs, r++, 3, `+DTL!G${dtlBooksTotalRow}`, money(round2(st.depreciation)), { money: true });
  setCell(dtaWs, r, 2, "Depreciation as per Income Tax Act");
  const itActDepRefRow = r;
  setFormula(dtaWs, r++, 3, `+DTL!G${dtlItActRow}`, money(round2(st.depreciation)), { money: true });
  setCell(dtaWs, r, 2, "Difference");
  const diffRow = r;
  setFormula(dtaWs, r++, 3, `C${itActDepRefRow}-C${booksDepRefRow}`, 0, { money: true });
  setCell(dtaWs, r, 2, "Deferred Tax Liability / (Asset)", { bold: true });
  const dtlFromFaRow = r;
  setFormula(dtaWs, r++, 3, `ROUND(C${diffRow}*C${taxRateRow},0)`, 0, { bold: true, money: true });
  r++;

  setCell(dtaWs, r, 1, "(B)");
  setCell(dtaWs, r++, 2, "Disallowances under Sec. 43B (unpaid statutory dues, etc. - enter manually)", {
    bold: true,
  });
  const disallowStart = r;
  setCell(dtaWs, r, 2, "(enter each disallowance as a separate line)");
  setCell(dtaWs, r++, 3, 0, { money: true });
  const disallowEnd = r - 1;
  setCell(dtaWs, r, 2, "Total", { bold: true });
  const disallowTotalRow = r;
  setFormula(dtaWs, r++, 3, `SUM(C${disallowStart}:C${disallowEnd})`, 0, { bold: true, money: true });
  setCell(dtaWs, r, 2, "Deferred Tax Asset (43B)", { bold: true });
  const dtaFromDisallowRow = r;
  setFormula(dtaWs, r++, 3, `ROUND(C${disallowTotalRow}*C${taxRateRow},0)`, 0, { bold: true, money: true });
  r += 2;

  setCell(dtaWs, r, 2, "Net Deferred Tax Expense / (Income)", { bold: true });
  const netExpenseRow = r;
  setFormula(dtaWs, r, 3, `C${dtlFromFaRow}+C${dtaFromDisallowRow}`, 0, { bold: true, money: true });
  topBorder(dtaWs, r, 1, 4);

  return {
    faNetClosingRow: faTotalRow,
    faDepForYearRow: faTotalRow,
    dtaNetExpenseRow: netExpenseRow,
  };
}
