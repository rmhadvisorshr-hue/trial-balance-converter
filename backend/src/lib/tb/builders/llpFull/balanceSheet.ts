import type ExcelJS from "exceljs";
import type { BuildCtx } from "../common";
import { writeVerticalBS } from "../common";
import { setCell, setFormula, applyColumnWidths, topBorder, signatureBlock } from "../../excel/helpers";
import { round2 } from "../../classify";
import type { Partner } from "../partnershipFull/partners";
import type { SchResult } from "../partnershipFull/schedules";
import type { FixedAssetScheduleResult } from "../partnershipFull/fixedAssetSchedule";
import type { PlAppropriationResult } from "../partnershipFull/plAppropriation";
import type { CapitalCurrentResult } from "./capitalCurrentAccounts";

export interface LlpBalanceSheetResult {
  totalSourcesRow: number;
  totalApplicationRow: number;
}

// 'BS': Sources of Funds / Application of Funds - the LLP Rules Second
// Schedule layout, NOT a relabeled Schedule III Equity & Liabilities
// statement and NOT Partnership's simple Liabilities/Assets T-format. The
// structural difference that must hold: Net Current Assets nets on the
// Application side (Current Assets, Loans & Advances less Current
// Liabilities & Provisions), rather than showing gross Current Liabilities
// as a Sources line the way a T-format or Schedule III statement would.
// Cash uses the trial balance's actual CASH_BANK balance rather than a
// balancing plug, consistent with every other entity built so far; a
// separate balancing-check line surfaces any real discrepancy - it reduces
// to the same "does the P&L's net profit reconcile with the BS totals"
// quantity used by Partnership's own check, just derived through the
// Sources/Application structure instead of a T-format.
export function writeLlpBalanceSheet(
  bsWs: ExcelJS.Worksheet,
  bsVWs: ExcelJS.Worksheet,
  ctx: BuildCtx,
  partners: Partner[],
  cap: CapitalCurrentResult,
  sch: SchResult,
  fa: FixedAssetScheduleResult,
  plApp: PlAppropriationResult,
): LlpBalanceSheetResult {
  applyColumnWidths(bsWs, [34, 18, 4, 34, 18]);
  const { meta, st, divisor } = ctx;
  const money = (v: number) => v / divisor;

  let r = 2;
  setCell(bsWs, r++, 1, meta.firmName || "LLP NAME", { bold: true });
  setCell(bsWs, r++, 1, meta.periodLabel || "", { italic: true });
  if (ctx.unitHeading) setCell(bsWs, r++, 1, ctx.unitHeading, { italic: true });
  setCell(bsWs, r++, 1, `BALANCE SHEET AS AT ${(meta.asAtLabel || "").replace(/^as at\s*/i, "")}`, {
    bold: true,
  });
  r++;
  setCell(bsWs, r, 1, "SOURCES OF FUNDS", { bold: true });
  setCell(bsWs, r, 2, "AMOUNT", { bold: true, align: "right" });
  setCell(bsWs, r, 4, "APPLICATION OF FUNDS", { bold: true });
  setCell(bsWs, r++, 5, "AMOUNT", { bold: true, align: "right" });
  r++;
  const contentStart = r;

  // ---------------- Sources of Funds ----------------
  const sourceRows: number[] = [];
  setCell(bsWs, r++, 1, "1. Partners' Funds", { bold: true });
  setCell(bsWs, r++, 1, "    (a) Fixed Capital / Contribution");
  for (const p of partners) {
    const closingRow = cap.fixedClosingRowByPartner.get(p.name)!;
    const closingAmount = cap.fixedClosingAmountByPartner.get(p.name)!;
    setCell(bsWs, r, 1, `        ${p.name}`);
    setFormula(bsWs, r, 2, `+'Cap & Current'!B${closingRow}`, money(closingAmount), { money: true });
    sourceRows.push(r);
    r++;
  }
  setCell(bsWs, r++, 1, "    (b) Partners' Current Accounts");
  for (const p of partners) {
    const closingRow = cap.currentClosingRowByPartner.get(p.name)!;
    const closingAmount = cap.currentClosingAmountByPartner.get(p.name)!;
    setCell(bsWs, r, 1, `        ${p.name}`);
    setFormula(bsWs, r, 2, `+'Cap & Current'!B${closingRow}`, money(closingAmount), { money: true });
    sourceRows.push(r);
    r++;
  }
  setCell(bsWs, r, 1, "    (c) Reserves and Surplus (not already reflected in Current Accounts above)");
  setCell(bsWs, r, 2, 0, { money: true });
  sourceRows.push(r);
  r++;
  r++;

  setCell(bsWs, r++, 1, "2. Loan Funds", { bold: true });
  const securedItems = st.byCode.get("LONG_TERM_BORROW") ?? [];
  const unsecuredItems = st.byCode.get("SHORT_TERM_BORROW") ?? [];
  if (securedItems.length) {
    setCell(bsWs, r++, 1, "    (a) Secured Loans");
    for (const item of securedItems) {
      setCell(bsWs, r, 1, `        ${item.name}`);
      setCell(bsWs, r, 2, money(round2(item.amount)), { money: true });
      sourceRows.push(r);
      r++;
    }
  }
  if (unsecuredItems.length) {
    setCell(bsWs, r++, 1, "    (b) Unsecured Loans");
    for (const item of unsecuredItems) {
      setCell(bsWs, r, 1, `        ${item.name}`);
      setCell(bsWs, r, 2, money(round2(item.amount)), { money: true });
      sourceRows.push(r);
      r++;
    }
  }
  r++;

  const sourceFirst = sourceRows[0];
  const sourceLast = sourceRows[sourceRows.length - 1];
  setCell(bsWs, r, 1, "TOTAL SOURCES OF FUNDS", { bold: true });
  const fixedCapitalTotal = round2(
    partners.reduce((s, p) => s + cap.fixedClosingAmountByPartner.get(p.name)!, 0),
  );
  const currentAccountTotal = round2(
    partners.reduce((s, p) => s + cap.currentClosingAmountByPartner.get(p.name)!, 0),
  );
  const loanFundsTotal = round2(st.total("LONG_TERM_BORROW") + st.total("SHORT_TERM_BORROW"));
  const totalSources = round2(fixedCapitalTotal + currentAccountTotal + loanFundsTotal);
  setFormula(bsWs, r, 2, `SUM(B${sourceFirst}:B${sourceLast})`, money(totalSources), { bold: true, money: true });
  const sourcesTotalRow = r;
  topBorder(bsWs, r, 1, 2);

  // ---------------- Application of Funds (independent row cursor) ----------------
  let ar = contentStart;
  setCell(bsWs, ar, 4, "1. Fixed Assets (net block)");
  setFormula(bsWs, ar, 5, `+'FIXED ASSET SCH'!G${fa.closingTotalRow}`, money(round2(st.total("PPE"))), {
    money: true,
  });
  const fixedAssetsRow = ar;
  ar += 2;

  let investmentsRow: number | null = null;
  if (sch.investmentsRow) {
    setCell(bsWs, ar, 4, "2. Investments");
    setFormula(bsWs, ar, 5, `+SCH!B${sch.investmentsRow}`, money(round2(st.total("NONCURR_INVEST"))), {
      money: true,
    });
    investmentsRow = ar;
    ar += 2;
  }

  setCell(bsWs, ar++, 4, "3. Current Assets, Loans & Advances", { bold: true });
  const caStart = ar;
  const sch4 = sch.schRow.get(4)!;
  const sch5 = sch.schRow.get(5)!;
  const sch3 = sch.schRow.get(3)!;
  setCell(bsWs, ar, 4, "      Sundry Debtors (Sch 5)");
  setFormula(bsWs, ar, 5, `+SCH!B${sch5}`, money(round2(st.total("TRADE_RECV"))), { money: true });
  ar++;
  setCell(bsWs, ar, 4, "      Cash and Bank Balance");
  setCell(bsWs, ar, 5, money(round2(st.total("CASH_BANK"))), { money: true }); // known TB balance, not a plug
  ar++;
  setCell(bsWs, ar, 4, "      Loans and Advances (Sch 4)");
  setFormula(
    bsWs,
    ar,
    5,
    `+SCH!B${sch4}`,
    money(round2(st.total("SHORT_TERM_LOANS_ADV") + st.total("LONG_TERM_LOANS_ADV"))),
    { money: true },
  );
  ar++;
  setCell(bsWs, ar, 4, "      Other Current Assets (Sch 3)");
  setFormula(
    bsWs,
    ar,
    5,
    `+SCH!B${sch3}`,
    money(round2(st.total("INVENTORY") + st.total("OTHER_CURR_ASSETS") + st.total("DEFERRED_TAX_ASSET"))),
    { money: true },
  );
  ar++;
  const caEnd = ar - 1;
  setCell(bsWs, ar, 4, "      Total (A)", { bold: true });
  const currentAssetsGross = round2(
    st.total("TRADE_RECV") +
      st.total("CASH_BANK") +
      st.total("SHORT_TERM_LOANS_ADV") +
      st.total("LONG_TERM_LOANS_ADV") +
      st.total("INVENTORY") +
      st.total("OTHER_CURR_ASSETS") +
      st.total("DEFERRED_TAX_ASSET"),
  );
  setFormula(bsWs, ar, 5, `SUM(E${caStart}:E${caEnd})`, money(currentAssetsGross), { bold: true, money: true });
  const totalARow = ar;
  ar += 2;

  setCell(bsWs, ar++, 4, "Less: Current Liabilities & Provisions", { bold: true });
  const clStart = ar;
  const sch1 = sch.schRow.get(1)!;
  const sch2 = sch.schRow.get(2)!;
  setCell(bsWs, ar, 4, "      Sundry Creditors (Sch 1)");
  setFormula(bsWs, ar, 5, `+SCH!B${sch1}`, money(round2(st.total("TRADE_PAYABLES"))), { money: true });
  ar++;
  setCell(bsWs, ar, 4, "      Other Current Liabilities (Sch 2)");
  setFormula(
    bsWs,
    ar,
    5,
    `+SCH!B${sch2}`,
    money(round2(st.total("OTHER_CURR_LIAB") + st.total("SHORT_TERM_PROV") + st.total("DEFERRED_TAX_LIAB"))),
    { money: true },
  );
  ar++;
  setCell(bsWs, ar, 4, "      Provision for Taxation");
  setFormula(bsWs, ar, 5, `+'p&l app'!B${plApp.taxProvisionRow}`, money(plApp.taxProvisionAmount), { money: true });
  ar++;
  const clEnd = ar - 1;
  setCell(bsWs, ar, 4, "      Total (B)", { bold: true });
  const currentLiabTotal = round2(
    st.total("TRADE_PAYABLES") +
      st.total("OTHER_CURR_LIAB") +
      st.total("SHORT_TERM_PROV") +
      st.total("DEFERRED_TAX_LIAB") +
      plApp.taxProvisionAmount,
  );
  setFormula(bsWs, ar, 5, `SUM(E${clStart}:E${clEnd})`, money(currentLiabTotal), { bold: true, money: true });
  const totalBRow = ar;
  ar += 2;

  setCell(bsWs, ar, 4, "Net Current Assets (A - B)", { bold: true });
  const netCurrentAssets = round2(currentAssetsGross - currentLiabTotal);
  setFormula(bsWs, ar, 5, `E${totalARow}-E${totalBRow}`, money(netCurrentAssets), { bold: true, money: true });
  const netCurrentAssetsRow = ar;
  ar += 2;

  const applicationRefs = [fixedAssetsRow, ...(investmentsRow ? [investmentsRow] : []), netCurrentAssetsRow];
  setCell(bsWs, ar, 4, "TOTAL APPLICATION OF FUNDS", { bold: true });
  const investmentsTotal = sch.investmentsRow ? round2(st.total("NONCURR_INVEST")) : 0;
  const totalApplication = round2(round2(st.total("PPE")) + investmentsTotal + netCurrentAssets);
  setFormula(
    bsWs,
    ar,
    5,
    applicationRefs.map((row) => `E${row}`).join("+"),
    money(totalApplication),
    { bold: true, money: true },
  );
  const applicationTotalRow = ar;
  topBorder(bsWs, ar, 4, 5);

  const bottomRow = Math.max(sourcesTotalRow, applicationTotalRow) + 2;
  setCell(bsWs, bottomRow, 1, "Balancing check (Total Sources - Total Application, should be 0)", {
    italic: true,
  });
  setFormula(
    bsWs,
    bottomRow,
    2,
    `B${sourcesTotalRow}-E${applicationTotalRow}`,
    money(round2(totalSources - totalApplication)),
    { italic: true, money: true },
  );

  signatureBlock(bsWs, bottomRow + 2, meta, ctx.forFirmText, ctx.designation);

  writeVerticalBS(bsVWs, ctx, "Balance Sheet");

  return { totalSourcesRow: sourcesTotalRow, totalApplicationRow: applicationTotalRow };
}
