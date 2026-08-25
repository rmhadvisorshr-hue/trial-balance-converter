import type ExcelJS from "exceljs";
import type { BuildCtx } from "../common";
import { setCell, setFormula, applyColumnWidths, topBorder } from "../../excel/helpers";
import { round2 } from "../../classify";
import type { NotesResult } from "./notes";
import type { ShareCapitalReservesResult } from "./shareCapitalReserves";
import type { FixedAssetChainResult } from "./fixedAssetChain";

// Note 24 (Disclosure of Ratios), Schedule III (2021 amendment). Every ratio
// here is a live formula built from NOTES/Note 2,3-SC/11. FA/P&L totals -
// unlike Ageing or Notes 27-31, ratios genuinely are derivable from a single
// trial balance. The only real gap is "average" balances and the prior-year
// comparison, both of which need a second period; this uses the closing
// balance only in place of an average, flagged below rather than left blank.
export function writeRatiosSheet(
  ws: ExcelJS.Worksheet,
  ctx: BuildCtx,
  notes: NotesResult,
  shareCap: ShareCapitalReservesResult,
  fa: FixedAssetChainResult,
  netProfitFormulaRef: string, // e.g. "'P&L'!D41"
) {
  applyColumnWidths(ws, [4, 34, 30, 30, 16]);
  const { meta, st, divisor } = ctx;
  const noteVal = (n: number) => `NOTES!B${notes.byNote.get(n)!.row}`;
  const noteTotal = (n: number) => st.total(codeForNote(n));
  const scRef = (row: number) => `'Note 2,3- SC'!C${row}`;

  let r = 2;
  setCell(ws, r++, 1, meta.firmName || "COMPANY NAME", { bold: true });
  setCell(ws, r++, 1, 'Note "24" : Disclosure of Ratios', { bold: true });
  setCell(
    ws,
    r++,
    1,
    "Ratios use closing-balance figures in place of period averages (a trial balance has no prior-year data); confirm before filing.",
    { italic: true },
  );
  r++;
  setCell(ws, r, 1, "Sr", { bold: true });
  setCell(ws, r, 2, "Particulars", { bold: true });
  setCell(ws, r, 3, "Numerator", { bold: true });
  setCell(ws, r, 4, "Denominator", { bold: true });
  setCell(ws, r++, 5, `Year ended ${meta.asAtLabel || meta.periodLabel || ""}`, {
    bold: true,
    align: "right",
  });
  r++;

  const currentAssets = round2(
    noteTotal(14) + noteTotal(15) + noteTotal(16) + noteTotal(17) + noteTotal(18),
  );
  const currentLiab = round2(noteTotal(7) + noteTotal(8) + noteTotal(9) + noteTotal(10));
  const capitalAndReserves = round2(st.total("CAPITAL") + st.total("RESERVES") + st.netProfit);
  const borrowings = round2(st.total("LONG_TERM_BORROW") + st.total("SHORT_TERM_BORROW"));
  const ebitda = round2(st.grossProfit + st.otherIncome - st.employeeBenefit - st.otherExpense);
  const cogs = round2(
    st.total("PURCHASES") +
      st.total("DIRECT_EXP") +
      st.total("OPENING_STOCK") -
      st.total("CLOSING_STOCK_PL"),
  );
  const purchases = round2(st.total("PURCHASES") + st.total("DIRECT_EXP"));
  const workingCapital = round2(currentAssets - currentLiab);
  const capitalEmployed = round2(capitalAndReserves + st.total("LONG_TERM_BORROW"));

  // Matches the IFERROR(...,0) wrapped around every live formula below, so
  // the cached "known" value (shown before Excel recalculates) never shows
  // a nonsense large number from dividing by a fallback of 1.
  const safeDiv = (n: number, d: number) => (d === 0 ? 0 : round2(n / d));

  const CA = `(${noteVal(14)}+${noteVal(15)}+${noteVal(16)}+${noteVal(17)}+${noteVal(18)})`;
  const CL = `(${noteVal(7)}+${noteVal(8)}+${noteVal(9)}+${noteVal(10)})`;
  const equity = `(${scRef(shareCap.shareCapitalRow)}+${scRef(shareCap.reservesRow)})`;
  const EBITDA = `((${noteVal(19)}+${noteVal(20)})-(${noteVal(21)}+${noteVal(22)})-${noteVal(23)}-${noteVal(25)})`;

  const row = (
    sr: string,
    label: string,
    numeratorLabel: string,
    denominatorLabel: string,
    formula: string,
    known: number,
  ) => {
    setCell(ws, r, 1, sr);
    setCell(ws, r, 2, label);
    setCell(ws, r, 3, numeratorLabel);
    setCell(ws, r, 4, denominatorLabel);
    // IFERROR guards against a zero denominator (e.g. no trade payables) -
    // Excel would otherwise show #DIV/0! once fullCalcOnLoad recalculates.
    setFormula(ws, r, 5, `IFERROR(${formula},0)`, known);
    r++;
  };

  row(
    "a.",
    "Current Ratio (times)",
    "Current Assets",
    "Current Liabilities",
    `${CA}/${CL}`,
    safeDiv(currentAssets, currentLiab),
  );
  row(
    "b.",
    "Debt-Equity Ratio (times)",
    "Long + Short Term Borrowings",
    "Share Capital & Reserves",
    `(${noteVal(4)}+${noteVal(7)})/${equity}`,
    safeDiv(borrowings, capitalAndReserves),
  );
  row(
    "c.",
    "Debt Service Coverage Ratio (times)",
    "EBITDA",
    "Interest (principal repayment not available from a TB - interest only)",
    `${EBITDA}/${noteVal(24)}`,
    safeDiv(ebitda, st.financeCost),
  );
  row(
    "d.",
    "Return on Equity Ratio (%)",
    "Net Profit After Tax",
    "Shareholders' Equity (closing)",
    `${netProfitFormulaRef}/${equity}`,
    safeDiv(st.netProfit, capitalAndReserves),
  );
  row(
    "e.",
    "Inventory Turnover Ratio (times)",
    "Cost of Goods Sold",
    "Inventory (closing)",
    `(${noteVal(21)}+${noteVal(22)})/${noteVal(14)}`,
    safeDiv(cogs, noteTotal(14)),
  );
  row(
    "f.",
    "Trade Receivables Turnover Ratio (times)",
    "Revenue from Operations",
    "Trade Receivables (closing)",
    `${noteVal(19)}/${noteVal(15)}`,
    safeDiv(st.revenue, noteTotal(15)),
  );
  row(
    "g.",
    "Trade Payables Turnover Ratio (times)",
    "Purchases",
    "Trade Payables (closing)",
    `${noteVal(21)}/${noteVal(8)}`,
    safeDiv(purchases, noteTotal(8)),
  );
  row(
    "h.",
    "Net Capital Turnover Ratio (times)",
    "Revenue from Operations",
    "Working Capital",
    `${noteVal(19)}/(${CA}-${CL})`,
    safeDiv(st.revenue, workingCapital),
  );
  row(
    "i.",
    "Net Profit Ratio (%)",
    "Net Profit",
    "Revenue from Operations",
    `${netProfitFormulaRef}/${noteVal(19)}`,
    safeDiv(st.netProfit, st.revenue),
  );
  row(
    "j.",
    "Return on Capital Employed (%)",
    "EBIT (EBITDA - Depreciation)",
    "Capital Employed (Equity + Long Term Borrowings)",
    `(${EBITDA}-'11. FA'!G${fa.faDepForYearRow})/(${equity}+${noteVal(4)})`,
    safeDiv(round2(ebitda - st.depreciation), capitalEmployed),
  );
  row(
    "k.",
    "Return on Investment (%)",
    "Other Income (proxy - investment income is not separately tracked in a TB)",
    "Non-current Investments (closing)",
    `${noteVal(20)}/${noteVal(12)}`,
    safeDiv(st.otherIncome, noteTotal(12)),
  );

  topBorder(ws, r - 1, 1, 5);
}

function codeForNote(n: number) {
  const map: Record<number, Parameters<BuildCtx["st"]["total"]>[0]> = {
    4: "LONG_TERM_BORROW",
    7: "SHORT_TERM_BORROW",
    8: "TRADE_PAYABLES",
    9: "OTHER_CURR_LIAB",
    10: "SHORT_TERM_PROV",
    12: "NONCURR_INVEST",
    14: "INVENTORY",
    15: "TRADE_RECV",
    16: "CASH_BANK",
    17: "SHORT_TERM_LOANS_ADV",
    18: "OTHER_CURR_ASSETS",
  };
  return map[n];
}
