import type ExcelJS from "exceljs";
import type { BuildCtx } from "../common";
import { writeVerticalBS } from "../common";
import { setCell, setFormula, applyColumnWidths, topBorder, signatureBlock } from "../../excel/helpers";
import { round2 } from "../../classify";
import type { Partner } from "./partners";
import type { SchResult } from "./schedules";
import type { CapitalAccountsResult } from "./capitalAccounts";
import type { FixedAssetScheduleResult } from "./fixedAssetSchedule";
import type { PlAppropriationResult } from "./plAppropriation";

// 'BS': T-format (Liabilities/Assets) Balance Sheet, every line a formula
// into Cap/SCH/'FIXED ASSET SCH'. Cash uses the trial balance's actual
// CASH_BANK ledger balance rather than the reference format's balancing-
// plug convention (=TotalLiabilities-SUM(other assets)) - consistent with
// every other figure in this app favoring known TB data; a separate
// balancing-check line surfaces any real discrepancy instead of hiding it
// inside Cash. 'BS (V)' is the vertical restatement, reusing the existing
// writeVerticalBS the same way tradingPL.ts reuses writeVerticalPL.
export function writePartnershipBalanceSheet(
  bsWs: ExcelJS.Worksheet,
  bsVWs: ExcelJS.Worksheet,
  ctx: BuildCtx,
  partners: Partner[],
  cap: CapitalAccountsResult,
  sch: SchResult,
  fa: FixedAssetScheduleResult,
  plApp: PlAppropriationResult,
) {
  applyColumnWidths(bsWs, [30, 18, 4, 30, 18]);
  const { meta, st, divisor } = ctx;
  const money = (v: number) => v / divisor;

  let r = 2;
  setCell(bsWs, r++, 1, meta.firmName || "FIRM NAME", { bold: true });
  setCell(bsWs, r++, 1, meta.periodLabel || "", { italic: true });
  if (ctx.unitHeading) setCell(bsWs, r++, 1, ctx.unitHeading, { italic: true });
  setCell(bsWs, r++, 1, `BALANCE SHEET AS AT ${(meta.asAtLabel || "").replace(/^as at\s*/i, "")}`, {
    bold: true,
  });
  r++;
  setCell(bsWs, r, 1, "LIABILITIES", { bold: true });
  setCell(bsWs, r, 2, "AMOUNT", { bold: true, align: "right" });
  setCell(bsWs, r, 4, "ASSETS", { bold: true });
  setCell(bsWs, r++, 5, "AMOUNT", { bold: true, align: "right" });
  r++;
  const contentStart = r; // Assets uses its own cursor from here, independent of how many liability rows there are.

  const liabRows: number[] = [];
  setCell(bsWs, r++, 1, "Capital Account", { bold: true });
  for (const p of partners) {
    const closingRow = cap.closingRowByPartner.get(p.name)!;
    const closingAmount = cap.closingAmountByPartner.get(p.name)!;
    setCell(bsWs, r, 1, `      ${p.name}`);
    setFormula(bsWs, r, 2, `+Cap!B${closingRow}`, money(closingAmount), { money: true });
    liabRows.push(r);
    r++;
  }
  r++;

  const secured = [...(st.byCode.get("LONG_TERM_BORROW") ?? []), ...(st.byCode.get("SHORT_TERM_BORROW") ?? [])];
  if (secured.length) {
    setCell(bsWs, r++, 1, "Secured / Unsecured Loans", { bold: true });
    for (const item of secured) {
      setCell(bsWs, r, 1, `      ${item.name}`);
      setCell(bsWs, r, 2, money(round2(item.amount)), { money: true });
      liabRows.push(r);
      r++;
    }
    r++;
  }

  setCell(bsWs, r++, 1, "Current Liabilities", { bold: true });
  const sch1 = sch.schRow.get(1)!;
  const sch2 = sch.schRow.get(2)!;
  setCell(bsWs, r, 1, "      Sundry Creditors (Sch 1)");
  setFormula(bsWs, r, 2, `+SCH!B${sch1}`, money(round2(st.total("TRADE_PAYABLES"))), { money: true });
  liabRows.push(r);
  r++;
  setCell(bsWs, r, 1, "      Other Current Liabilities (Sch 2)");
  setFormula(
    bsWs,
    r,
    2,
    `+SCH!B${sch2}`,
    money(round2(st.total("OTHER_CURR_LIAB") + st.total("SHORT_TERM_PROV") + st.total("DEFERRED_TAX_LIAB"))),
    { money: true },
  );
  liabRows.push(r);
  r++;
  // The tax provision computed in the P&L Appropriation account reduces
  // partners' Share in Net Profit, so it must appear here as a liability -
  // otherwise the accounting equation doesn't balance.
  setCell(bsWs, r, 1, "      Provision for Taxation");
  setFormula(bsWs, r, 2, `+'p&l app'!B${plApp.taxProvisionRow}`, money(plApp.taxProvisionAmount), {
    money: true,
  });
  liabRows.push(r);
  r++;

  const liabFirst = liabRows[0];
  const liabLast = liabRows[liabRows.length - 1];
  r++;
  setCell(bsWs, r, 1, "TOTAL", { bold: true });
  const totalLiabilities = round2(
    partners.reduce((s, p) => s + cap.closingAmountByPartner.get(p.name)!, 0) +
      secured.reduce((s, i) => s + i.amount, 0) +
      st.total("TRADE_PAYABLES") +
      st.total("OTHER_CURR_LIAB") +
      st.total("SHORT_TERM_PROV") +
      st.total("DEFERRED_TAX_LIAB") +
      plApp.taxProvisionAmount,
  );
  setFormula(bsWs, r, 2, `SUM(B${liabFirst}:B${liabLast})`, money(totalLiabilities), { bold: true, money: true });
  const liabTotalRow = r;
  topBorder(bsWs, r, 1, 2);

  // ---------------- Assets (independent row cursor, same starting row as Liabilities) ----------------
  let ar = contentStart;
  const assetRows: number[] = [];
  setCell(bsWs, ar, 4, "Fixed Assets (Sch 5)");
  setFormula(bsWs, ar, 5, `+'FIXED ASSET SCH'!G${fa.closingTotalRow}`, money(round2(st.total("PPE"))), {
    money: true,
  });
  assetRows.push(ar);
  ar++;
  ar++;

  if (sch.investmentsRow) {
    setCell(bsWs, ar, 4, "Investments");
    setFormula(bsWs, ar, 5, `+SCH!B${sch.investmentsRow}`, money(round2(st.total("NONCURR_INVEST"))), {
      money: true,
    });
    assetRows.push(ar);
    ar++;
    ar++;
  }

  setCell(bsWs, ar++, 4, "Current Assets", { bold: true });
  const sch4 = sch.schRow.get(4)!;
  const sch5 = sch.schRow.get(5)!;
  const sch3 = sch.schRow.get(3)!;
  setCell(bsWs, ar, 4, "      Loans and Advances (Sch 4)");
  setFormula(
    bsWs,
    ar,
    5,
    `+SCH!B${sch4}`,
    money(round2(st.total("SHORT_TERM_LOANS_ADV") + st.total("LONG_TERM_LOANS_ADV"))),
    { money: true },
  );
  assetRows.push(ar);
  ar++;
  setCell(bsWs, ar, 4, "      Sundry Debtors (Sch 5)");
  setFormula(bsWs, ar, 5, `+SCH!B${sch5}`, money(round2(st.total("TRADE_RECV"))), { money: true });
  assetRows.push(ar);
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
  assetRows.push(ar);
  ar++;
  setCell(bsWs, ar, 4, "      Cash and Bank Balance");
  setCell(bsWs, ar, 5, money(round2(st.total("CASH_BANK"))), { money: true }); // known TB balance, not a plug
  assetRows.push(ar);
  ar++;

  const assetFirst = assetRows[0];
  const assetLast = assetRows[assetRows.length - 1];
  ar++;
  setCell(bsWs, ar, 4, "TOTAL", { bold: true });
  const totalAssets = round2(st.totalAssets);
  setFormula(bsWs, ar, 5, `SUM(E${assetFirst}:E${assetLast})`, money(totalAssets), { bold: true, money: true });
  const assetTotalRow = ar;
  topBorder(bsWs, ar, 4, 5);

  const bottomRow = Math.max(liabTotalRow, assetTotalRow) + 2;
  setCell(bsWs, bottomRow, 1, "Balancing check (Total Assets - Total Liabilities, should be 0)", {
    italic: true,
  });
  setFormula(
    bsWs,
    bottomRow,
    2,
    `E${assetTotalRow}-B${liabTotalRow}`,
    money(round2(totalAssets - totalLiabilities)),
    { italic: true, money: true },
  );

  const designation = "Partner";
  signatureBlock(bsWs, bottomRow + 2, meta, ctx.forFirmText, designation);

  writeVerticalBS(bsVWs, ctx, "Balance Sheet");
}
