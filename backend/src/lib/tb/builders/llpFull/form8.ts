import type ExcelJS from "exceljs";
import type { BuildCtx } from "../common";
import { setCell, setFormula, applyColumnWidths, topBorder } from "../../excel/helpers";
import { round2 } from "../../classify";
import type { Partner } from "../partnershipFull/partners";
import type { CapitalCurrentResult } from "./capitalCurrentAccounts";
import type { LlpBalanceSheetResult } from "./balanceSheet";

// 'FORM 8': MCA Form 8 - Statement of Account & Solvency. Part A is the
// solvency declaration plus a restatement of the BS's own Sources/
// Application grand totals (formula-linked to 'BS', not re-derived). Part B
// independently lists and SUMs each P&L line-item group under Turnover/
// Other Income/Expenditure so every figure here is backed by a real SUM
// formula over named ledgers, right there on the sheet - the same "list
// real items, SUM locally" pattern used by SCH - rather than reaching into
// the 'P&L' T-format's row layout, which has no single "Total Sales" cell
// to point to (Sales/Closing Stock are individually listed credit items,
// see tradingPL.ts). LLPIN comes from `ctx.meta.llpin` (pre-filled from the
// same UI field that feeds DETAILS - see routes/index.tsx), not a
// cross-sheet formula. The audit-applicability flag is a literal Excel IF
// formula against the statutory Rs.40 lakh turnover threshold; the Rs.25
// lakh contribution threshold is noted alongside since a single TB total
// can't reliably distinguish "contribution" from other capital movements.
export function writeForm8Sheet(
  ws: ExcelJS.Worksheet,
  ctx: BuildCtx,
  partners: Partner[],
  cap: CapitalCurrentResult,
  bsResult: LlpBalanceSheetResult,
) {
  applyColumnWidths(ws, [50, 20]);
  const { meta, st, divisor } = ctx;
  const money = (v: number) => v / divisor;

  let r = 2;
  setCell(ws, r++, 1, meta.firmName || "LLP NAME", { bold: true });
  if (meta.llpin) setCell(ws, r++, 1, `LLPIN : ${meta.llpin}`, {});
  setCell(ws, r++, 1, "FORM 8 - STATEMENT OF ACCOUNT & SOLVENCY", { bold: true });
  setCell(ws, r++, 1, meta.periodLabel || "", { italic: true });
  r++;

  setCell(ws, r++, 1, "PART A - STATEMENT OF SOLVENCY", { bold: true, underline: true });
  setCell(
    ws,
    r++,
    1,
    "We, the Designated Partners of the above named Limited Liability Partnership, do hereby solemnly " +
      "declare that the LLP was able to pay its debts as they fell due in the normal course of its business " +
      "during the financial year ended, as at the close of the year, and that the LLP is not likely to be " +
      "rendered insolvent within the period of one year immediately following the said date.",
    { italic: true, wrap: true },
  );
  r += 2;

  setCell(ws, r++, 1, "STATEMENT OF ASSETS AND LIABILITIES", { bold: true });
  setCell(ws, r, 1, "Total Sources of Funds (Partners' Funds + Loan Funds)");
  setFormula(ws, r++, 2, `+BS!B${bsResult.totalSourcesRow}`, money(0), { money: true });
  setCell(ws, r, 1, "Total Application of Funds (Fixed Assets + Investments + Net Current Assets)");
  setFormula(ws, r++, 2, `+BS!E${bsResult.totalApplicationRow}`, money(0), { money: true });
  r++;

  setCell(
    ws,
    r++,
    1,
    "Contribution received during the year (disclosed separately from routine Current Account movement)",
    { bold: true },
  );
  const contribStart = r;
  for (const p of partners) {
    const row = cap.contributionDuringYearRowByPartner.get(p.name)!;
    setCell(ws, r, 1, `      ${p.name}`);
    setFormula(ws, r, 2, `+'Cap & Current'!E${row}`, money(0), { money: true });
    r++;
  }
  const contribEnd = r - 1;
  setCell(ws, r, 1, "Total Contribution Received", { bold: true });
  setFormula(ws, r, 2, `SUM(B${contribStart}:B${contribEnd})`, money(0), { bold: true, money: true });
  topBorder(ws, r, 1, 2);
  r += 2;

  setCell(ws, r++, 1, "SIGNATORIES (Designated Partners)", { bold: true });
  const signers = partners.slice(0, 2);
  for (const p of signers) {
    setCell(ws, r, 1, p.name);
    setCell(ws, r++, 2, "DPIN: (enter manually - see DETAILS)");
  }
  if (signers.length < 2) {
    setCell(
      ws,
      r++,
      1,
      "(fewer than 2 designated partners detected from the trial balance - enter the second signatory manually)",
      { italic: true },
    );
  }
  r += 2;

  setCell(ws, r++, 1, "PART B - STATEMENT OF INCOME & EXPENDITURE", { bold: true, underline: true });
  r++;

  const revItems = st.byCode.get("REV_OPS") ?? [];
  setCell(ws, r++, 1, "Turnover", { bold: true });
  const turnoverStart = r;
  for (const it of revItems) {
    setCell(ws, r, 1, `      ${it.name}`);
    setCell(ws, r++, 2, money(round2(it.amount)), { money: true });
  }
  if (r === turnoverStart) {
    setCell(ws, r, 1, "      -");
    setCell(ws, r++, 2, 0, { money: true });
  }
  const turnoverEnd = r - 1;
  const turnoverRow = r;
  setFormula(ws, r++, 2, `SUM(B${turnoverStart}:B${turnoverEnd})`, money(round2(st.revenue)), {
    bold: true,
    money: true,
  });
  r++;

  const otherIncomeItems = st.byCode.get("OTHER_INCOME") ?? [];
  setCell(ws, r++, 1, "Other Income", { bold: true });
  const oiStart = r;
  for (const it of otherIncomeItems) {
    setCell(ws, r, 1, `      ${it.name}`);
    setCell(ws, r++, 2, money(round2(it.amount)), { money: true });
  }
  if (r === oiStart) {
    setCell(ws, r, 1, "      -");
    setCell(ws, r++, 2, 0, { money: true });
  }
  const oiEnd = r - 1;
  const otherIncomeRow = r;
  setFormula(ws, r++, 2, `SUM(B${oiStart}:B${oiEnd})`, money(round2(st.otherIncome)), { bold: true, money: true });
  r++;

  setCell(ws, r, 1, "Total Income", { bold: true });
  const totalIncomeRow = r;
  const totalIncome = round2(st.revenue + st.otherIncome);
  setFormula(ws, r++, 2, `B${turnoverRow}+B${otherIncomeRow}`, money(totalIncome), { bold: true, money: true });
  r++;

  setCell(ws, r++, 1, "Expenditure", { bold: true });
  const expGroups: { title: string; codes: ("OPENING_STOCK" | "PURCHASES" | "DIRECT_EXP" | "EMP_BENEFIT" | "FINANCE_COST" | "DEPRECIATION" | "OTHER_EXP")[] }[] = [
    { title: "Cost of Sales (Opening Stock + Purchases + Direct Expenses, less Closing Stock)", codes: ["OPENING_STOCK", "PURCHASES", "DIRECT_EXP"] },
    { title: "Employee Benefit Expenses", codes: ["EMP_BENEFIT"] },
    { title: "Finance Costs", codes: ["FINANCE_COST"] },
    { title: "Depreciation", codes: ["DEPRECIATION"] },
    { title: "Other Expenses", codes: ["OTHER_EXP"] },
  ];
  const expGroupRows: number[] = [];
  let expTotalCached = 0;
  for (const g of expGroups) {
    const items = g.codes.flatMap((c) => st.byCode.get(c) ?? []);
    const isCostOfSales = g.codes.includes("PURCHASES");
    setCell(ws, r++, 1, `      ${g.title}`);
    const gStart = r;
    for (const it of items) {
      setCell(ws, r, 1, `            ${it.name}`);
      setCell(ws, r++, 2, money(round2(it.amount)), { money: true });
    }
    const closingItems = isCostOfSales ? st.byCode.get("CLOSING_STOCK_PL") ?? [] : [];
    for (const it of closingItems) {
      setCell(ws, r, 1, `            Less: ${it.name}`);
      setCell(ws, r++, 2, money(-round2(it.amount)), { money: true });
    }
    if (r === gStart) {
      setCell(ws, r, 1, "            -");
      setCell(ws, r++, 2, 0, { money: true });
    }
    const gEnd = r - 1;
    const groupTotal = round2(
      g.codes.reduce((s, c) => s + st.total(c), 0) - (isCostOfSales ? st.total("CLOSING_STOCK_PL") : 0),
    );
    setCell(ws, r, 1, "      Sub-total", { italic: true });
    setFormula(ws, r, 2, `SUM(B${gStart}:B${gEnd})`, money(groupTotal), { italic: true, money: true });
    expGroupRows.push(r);
    expTotalCached = round2(expTotalCached + groupTotal);
    r++;
  }
  r++;

  setCell(ws, r, 1, "Total Expenditure", { bold: true });
  const totalExpRow = r;
  setFormula(ws, r++, 2, expGroupRows.map((row) => `B${row}`).join("+"), money(expTotalCached), {
    bold: true,
    money: true,
  });
  r++;

  setCell(ws, r, 1, "Net Profit (Total Income - Total Expenditure)", { bold: true });
  const netProfitFigure = round2(totalIncome - expTotalCached);
  setFormula(ws, r, 2, `B${totalIncomeRow}-B${totalExpRow}`, money(netProfitFigure), { bold: true, money: true });
  topBorder(ws, r, 1, 2);
  r += 2;

  setCell(
    ws,
    r++,
    1,
    "AUDIT APPLICABILITY (Section 34(4), LLP Act 2008 - Rs.40 lakh turnover / Rs.25 lakh contribution)",
    { bold: true },
  );
  setCell(ws, r, 1, "Turnover exceeds Rs.40,00,000?");
  setFormula(
    ws,
    r++,
    2,
    `IF(B${turnoverRow}>4000000,"Audit Required","Audit Not Mandatory")`,
    st.revenue > 4000000 ? "Audit Required" : "Audit Not Mandatory",
    {},
  );
  setCell(
    ws,
    r++,
    1,
    "(the Rs.25 lakh contribution threshold must be cross-checked manually against the Fixed Capital/Contribution total on 'Cap & Current' - a single trial balance total cannot reliably distinguish formal contribution from other capital movement)",
    { italic: true, wrap: true },
  );
}
