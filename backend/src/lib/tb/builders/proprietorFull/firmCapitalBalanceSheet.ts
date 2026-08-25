import type ExcelJS from "exceljs";
import type { BuildCtx } from "../common";
import { setCell, setFormula, applyColumnWidths, topBorder, signatureBlock } from "../../excel/helpers";
import { round2 } from "../../classify";
import type { SchResult } from "./schedules";
import type { FixedAssetScheduleResult } from "./fixedAssetSchedule";

export interface FirmCapBsResult {
  capitalClosingRow: number; // roll-forward closing balance row, column C - for BS and SELF CAP&BS
  capitalClosingAmount: number; // known cached value, in actual rupees
  netProfitRow: number; // 'FIRM CAP&BS' row holding the Net Profit credit - for SELF CAP&BS
}

// 'FIRM CAP&BS': the business Capital roll-forward + Balance Sheet. A
// Proprietor has exactly one capital account, so - unlike Partnership,
// which detects multiple named partners - every CAPITAL-coded ledger is
// simply summed into one figure (this is the concrete fix for the "multiple
// shareholder-style sub-accounts" symptom reported when a Company's trial
// balance was tested under the Proprietor entity type: that TB genuinely
// had multiple named capital ledgers, which is correct company behaviour but
// wrong for a Proprietor, where they must collapse into one).
// Opening = the TB's collapsed Capital + accumulated Reserves (the same
// "must include accumulated P&L, not just current year" fix applied to
// Pvt Ltd/Partnership); Closing is computed by adding Net Profit and the
// (0/manual) personal items - same known-opening/computed-closing pattern
// used throughout. Cash uses the TB's actual CASH_BANK balance rather than
// a forced plug, with a separate balancing-check line, consistent with
// every other entity type built so far.
export function writeFirmCapitalBalanceSheet(
  ws: ExcelJS.Worksheet,
  ctx: BuildCtx,
  netProfitFormulaRef: string, // e.g. "'P&L'!B36" - the Trading & P&L's Net Profit c/d cell
  sch: SchResult,
  fa: FixedAssetScheduleResult,
): FirmCapBsResult {
  applyColumnWidths(ws, [26, 18, 4, 26, 18, 4]);
  const { meta, st, divisor } = ctx;
  const money = (v: number) => v / divisor;

  const openingCapital = round2(st.total("CAPITAL") + st.total("RESERVES"));

  let r = 2;
  setCell(ws, r++, 1, meta.firmName || "FIRM NAME", { bold: true });
  setCell(ws, r++, 1, meta.periodLabel || "", { italic: true });
  if (ctx.unitHeading) setCell(ws, r++, 1, ctx.unitHeading, { italic: true });
  r++;
  setCell(ws, r++, 1, "Sch I : Capital Account", { bold: true });
  setCell(ws, r, 1, "PARTICULARS", { bold: true });
  setCell(ws, r, 3, "AMOUNT", { bold: true, align: "right" });
  setCell(ws, r, 4, "PARTICULARS", { bold: true });
  setCell(ws, r++, 6, "AMOUNT", { bold: true, align: "right" });

  const drawingsRow = r;
  setCell(ws, r, 1, "To Drawings");
  setCell(ws, r, 3, 0, { money: true });
  setCell(ws, r, 4, "By Balance b/d (Capital + Reserves per trial balance)");
  const openingRow = r;
  setCell(ws, r, 6, money(openingCapital), { money: true }); // known
  r++;

  const saTaxRow = r;
  setCell(ws, r, 1, "To Self-Assessment Tax");
  setCell(ws, r, 3, 0, { money: true });
  setCell(ws, r, 4, "By Net Profit");
  const netProfitRow = r;
  const netProfit = st.netProfit;
  setFormula(ws, r, 6, netProfitFormulaRef, money(netProfit), { money: true });
  r++;

  const insuranceRow = r;
  setCell(ws, r, 1, "To Personal Insurance Premium (paid from business funds)");
  setCell(ws, r, 3, 0, { money: true });
  setCell(ws, r, 4, "By Gifts Received (enter manually - exempt from specified relatives, must still be disclosed)");
  const giftRow = r;
  setCell(ws, r, 6, 0, { money: true });
  r++;

  setCell(ws, r, 1, "To Balance c/d", { bold: true });
  const closingRow = r;
  const closingAmount = round2(openingCapital + netProfit);
  setFormula(
    ws,
    r,
    3,
    `F${openingRow}+F${netProfitRow}+F${giftRow}-C${drawingsRow}-C${saTaxRow}-C${insuranceRow}`,
    money(closingAmount),
    { bold: true, money: true },
  );
  r++;

  r++;
  setCell(ws, r, 1, "Total", { bold: true });
  setFormula(ws, r, 3, `SUM(C${drawingsRow}:C${closingRow})`, money(closingAmount), { bold: true, money: true });
  setCell(ws, r, 4, "Total", { bold: true });
  setFormula(ws, r, 6, `SUM(F${openingRow}:F${giftRow})`, money(closingAmount), { bold: true, money: true });
  topBorder(ws, r, 1, 6);
  r += 2;

  // ---------------- Balance Sheet ----------------
  setCell(ws, r++, 1, `BALANCE SHEET AS AT ${(meta.asAtLabel || "").replace(/^as at\s*/i, "")}`.trim(), {
    bold: true,
  });
  r++;
  setCell(ws, r, 1, "LIABILITIES", { bold: true });
  setCell(ws, r, 3, "AMOUNT", { bold: true, align: "right" });
  setCell(ws, r, 4, "ASSETS", { bold: true });
  setCell(ws, r++, 6, "AMOUNT", { bold: true, align: "right" });
  r++;
  const contentStart = r;

  const liabRows: number[] = [];
  setCell(ws, r, 1, "Capital");
  setFormula(ws, r, 3, `C${closingRow}`, money(closingAmount), { money: true });
  liabRows.push(r);
  r++;

  const secured = round2(st.total("SHORT_TERM_BORROW"));
  if (secured) {
    setCell(ws, r, 1, "Secured Loans");
    setCell(ws, r, 3, money(secured), { money: true });
    liabRows.push(r);
    r++;
  }
  const sch1 = sch.schRow.get(1)!; // Unsecured Loans
  setCell(ws, r, 1, "Unsecured Loans (Sch 1)");
  setFormula(ws, r, 3, `+'FIRM SCH'!B${sch1}`, money(round2(st.total("LONG_TERM_BORROW"))), { money: true });
  liabRows.push(r);
  r++;

  const sch2 = sch.schRow.get(2)!; // Sundry Creditors
  setCell(ws, r, 1, "Sundry Creditors (Sch 2)");
  setFormula(ws, r, 3, `+'FIRM SCH'!B${sch2}`, money(round2(st.total("TRADE_PAYABLES"))), { money: true });
  liabRows.push(r);
  r++;

  const sch3 = sch.schRow.get(3)!; // Other Current Liabilities
  setCell(ws, r, 1, "Other Current Liabilities (Sch 3)");
  setFormula(
    ws,
    r,
    3,
    `+'FIRM SCH'!B${sch3}`,
    money(round2(st.total("OTHER_CURR_LIAB") + st.total("SHORT_TERM_PROV") + st.total("DEFERRED_TAX_LIAB"))),
    { money: true },
  );
  liabRows.push(r);
  r++;

  const liabFirst = liabRows[0];
  const liabLast = liabRows[liabRows.length - 1];
  r++;
  setCell(ws, r, 1, "Total", { bold: true });
  const totalLiabilities = round2(
    closingAmount +
      secured +
      st.total("LONG_TERM_BORROW") +
      st.total("TRADE_PAYABLES") +
      st.total("OTHER_CURR_LIAB") +
      st.total("SHORT_TERM_PROV") +
      st.total("DEFERRED_TAX_LIAB"),
  );
  setFormula(ws, r, 3, `SUM(C${liabFirst}:C${liabLast})`, money(totalLiabilities), { bold: true, money: true });
  const liabTotalRow = r;
  topBorder(ws, r, 1, 3);

  let ar = contentStart;
  const assetRows: number[] = [];
  setCell(ws, ar, 4, "Fixed Assets");
  setFormula(ws, ar, 6, `+'FIRM FA'!G${fa.closingTotalRow}`, money(round2(st.total("PPE"))), { money: true });
  assetRows.push(ar);
  ar++;

  if (sch.investmentsRow) {
    setCell(ws, ar, 4, "Investments and Advances");
    setFormula(ws, ar, 6, `+'FIRM SCH'!B${sch.investmentsRow}`, money(round2(st.total("NONCURR_INVEST"))), {
      money: true,
    });
    assetRows.push(ar);
    ar++;
  }

  const sch4 = sch.schRow.get(4)!; // Loans and Advances
  setCell(ws, ar, 4, "Loans & Advances (Sch 4)");
  setFormula(
    ws,
    ar,
    6,
    `+'FIRM SCH'!B${sch4}`,
    money(round2(st.total("SHORT_TERM_LOANS_ADV") + st.total("LONG_TERM_LOANS_ADV"))),
    { money: true },
  );
  assetRows.push(ar);
  ar++;

  const sch6 = sch.schRow.get(6)!; // Sundry Debtors
  setCell(ws, ar, 4, "Sundry Debtors (Sch 6)");
  setFormula(ws, ar, 6, `+'FIRM SCH'!B${sch6}`, money(round2(st.total("TRADE_RECV"))), { money: true });
  assetRows.push(ar);
  ar++;

  const sch5 = sch.schRow.get(5)!; // Other Current Assets
  setCell(ws, ar, 4, "Other Current Assets (Sch 5)");
  setFormula(
    ws,
    ar,
    6,
    `+'FIRM SCH'!B${sch5}`,
    money(round2(st.total("INVENTORY") + st.total("OTHER_CURR_ASSETS") + st.total("DEFERRED_TAX_ASSET"))),
    { money: true },
  );
  assetRows.push(ar);
  ar++;

  setCell(ws, ar, 4, "Advance Tax (enter manually)");
  setCell(ws, ar, 6, 0, { money: true });
  const advanceTaxRow = ar;
  assetRows.push(ar);
  ar++;

  setCell(ws, ar, 4, "Cash and Bank Balances");
  setCell(ws, ar, 6, money(round2(st.total("CASH_BANK"))), { money: true }); // known TB balance, not a plug
  assetRows.push(ar);
  ar++;

  const assetFirst = assetRows[0];
  const assetLast = assetRows[assetRows.length - 1];
  ar++;
  setCell(ws, ar, 4, "Total", { bold: true });
  const totalAssets = round2(st.totalAssets);
  setFormula(ws, ar, 6, `SUM(F${assetFirst}:F${assetLast})`, money(totalAssets), { bold: true, money: true });
  const assetTotalRow = ar;
  topBorder(ws, ar, 4, 6);

  const bottomRow = Math.max(liabTotalRow, assetTotalRow) + 2;
  setCell(ws, bottomRow, 1, "Balancing check (Total Assets - Total Liabilities, should be 0)", {
    italic: true,
  });
  setFormula(
    ws,
    bottomRow,
    3,
    `F${assetTotalRow}-C${liabTotalRow}`,
    money(round2(totalAssets - totalLiabilities)),
    { italic: true, money: true },
  );

  signatureBlock(ws, bottomRow + 2, meta, ctx.forFirmText, "Proprietor");

  void advanceTaxRow;
  return { capitalClosingRow: closingRow, capitalClosingAmount: closingAmount, netProfitRow };
}
