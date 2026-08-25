import type ExcelJS from "exceljs";
import type { ClassifiedLedger, StatementCode, StatementMeta } from "../types";
import type { Statements, LineItem } from "../statements";
import { setCell, signatureBlock, applyColumnWidths, topBorder } from "../excel/helpers";

function items(st: Statements, ...codes: StatementCode[]): LineItem[] {
  const out: LineItem[] = [];
  for (const c of codes) out.push(...(st.byCode.get(c) ?? []));
  return out;
}

export interface BuildCtx {
  meta: StatementMeta;
  st: Statements;
  ledgers: ClassifiedLedger[];
  forFirmText: string; // e.g. "For M/s Orbit Corporation"
  designation: string; // "Partner" | "Proprietor"
  divisor: number; // scales displayed/exported amounts (1 / 1,000 / 100,000)
  unitHeading: string | null; // e.g. "Figures in Thousands ('000)", null when actual
}

// ---------------- Vertical P&L (the "(V)" sheet) ----------------
export function writeVerticalPL(ws: ExcelJS.Worksheet, ctx: BuildCtx, heading: string) {
  applyColumnWidths(ws, [44, 20, 20]);
  const { meta, st, divisor } = ctx;
  const money = (v: number) => v / divisor;
  let r = 2;
  setCell(ws, r++, 1, meta.firmName || "FIRM NAME", { bold: true });
  setCell(ws, r++, 1, heading, { bold: true });
  setCell(ws, r++, 1, meta.periodLabel, { italic: true });
  if (ctx.unitHeading) setCell(ws, r++, 1, ctx.unitHeading, { italic: true });
  setCell(ws, r++, 2, meta.firmName, { align: "center", bold: true });
  setCell(ws, r, 1, "Particulars", { bold: true });
  setCell(ws, r++, 2, meta.periodLabel, { bold: true, align: "right" });

  const section = (
    title: string,
    total: number,
    list: LineItem[],
    prefixes?: Record<number, string>,
  ) => {
    setCell(ws, r, 1, title, { bold: true });
    setCell(ws, r, 3, money(total), { money: true, bold: true });
    r++;
    list.forEach((it, idx) => {
      const label = prefixes?.[idx] ? `${prefixes[idx]}${it.name}` : it.name;
      setCell(ws, r, 1, label);
      setCell(ws, r, 2, money(it.amount), { money: true });
      r++;
    });
  };

  setCell(ws, r++, 1, "Trading Account:", { bold: true });
  section("Sales Accounts", st.revenue, items(st, "REV_OPS"));

  const purchases = items(st, "OPENING_STOCK", "PURCHASES", "CLOSING_STOCK_PL");
  const costTotal =
    st.total("PURCHASES") + st.total("OPENING_STOCK") - st.total("CLOSING_STOCK_PL");
  section("Cost of Sales :", Math.round(costTotal * 100) / 100, purchases);

  const directList = items(st, "DIRECT_EXP");
  if (directList.length) section("Direct Expenses", st.total("DIRECT_EXP"), directList);

  setCell(ws, r, 1, "Gross Profit :", { bold: true });
  setCell(ws, r, 3, money(st.grossProfit), { money: true, bold: true });
  r += 2;

  setCell(ws, r++, 1, "Income Statement:", { bold: true });
  section("Indirect Incomes", st.otherIncome, items(st, "OTHER_INCOME"));
  section(
    "Indirect Expenses",
    st.totalIndirectExpense,
    items(st, "EMP_BENEFIT", "FINANCE_COST", "DEPRECIATION", "OTHER_EXP"),
  );

  setCell(ws, r, 1, "Net Profit :", { bold: true });
  setCell(ws, r, 3, money(st.netProfit), { money: true, bold: true });
  topBorder(ws, r, 1, 3);
  r += 1;

  signatureBlock(ws, r, meta, ctx.forFirmText, ctx.designation);
}

// ---------------- Vertical Balance Sheet (the "(V)" sheet) ----------------
export function writeVerticalBS(ws: ExcelJS.Worksheet, ctx: BuildCtx, heading: string) {
  applyColumnWidths(ws, [44, 20, 20]);
  const { meta, st, divisor } = ctx;
  const money = (v: number) => v / divisor;
  let r = 2;
  setCell(ws, r++, 1, meta.firmName || "FIRM NAME", { bold: true });
  setCell(ws, r++, 1, heading, { bold: true });
  setCell(ws, r++, 1, meta.periodLabel, { italic: true });
  if (ctx.unitHeading) setCell(ws, r++, 1, ctx.unitHeading, { italic: true });
  setCell(ws, r++, 2, meta.firmName, { align: "center", bold: true });
  setCell(ws, r++, 2, meta.asAtLabel, { align: "center", italic: true });

  const section = (title: string, total: number, list: LineItem[]) => {
    setCell(ws, r, 1, title, { bold: true });
    setCell(ws, r, 3, money(total), { money: true, bold: true });
    r++;
    for (const it of list) {
      setCell(ws, r, 1, it.name);
      setCell(ws, r, 2, money(it.amount), { money: true });
      r++;
    }
  };

  // Liabilities
  setCell(ws, r++, 1, "Liabilities :", { bold: true });
  const capList = [
    ...items(st, "CAPITAL", "RESERVES"),
    { name: "Net Profit", amount: st.netProfit },
  ];
  section(
    "Capital Account",
    Math.round((st.totalMany(["CAPITAL", "RESERVES"]) + st.netProfit) * 100) / 100,
    capList,
  );
  section(
    "Loans (Liability)",
    st.totalMany(["LONG_TERM_BORROW", "SHORT_TERM_BORROW"]),
    items(st, "LONG_TERM_BORROW", "SHORT_TERM_BORROW"),
  );
  section(
    "Current Liabilities",
    st.totalMany(["TRADE_PAYABLES", "OTHER_CURR_LIAB", "SHORT_TERM_PROV", "DEFERRED_TAX_LIAB"]),
    items(st, "TRADE_PAYABLES", "OTHER_CURR_LIAB", "SHORT_TERM_PROV", "DEFERRED_TAX_LIAB"),
  );
  setCell(ws, r, 1, "Total", { bold: true });
  setCell(ws, r, 3, money(st.totalLiabilities), { money: true, bold: true });
  topBorder(ws, r, 1, 3);
  r += 1;

  // Assets
  setCell(ws, r++, 1, "Assets :", { bold: true });
  section("Fixed Assets", st.total("PPE"), items(st, "PPE"));
  section(
    "Investments",
    st.totalMany(["NONCURR_INVEST", "LONG_TERM_LOANS_ADV", "DEFERRED_TAX_ASSET"]),
    items(st, "NONCURR_INVEST", "LONG_TERM_LOANS_ADV", "DEFERRED_TAX_ASSET"),
  );
  section(
    "Current Assets",
    st.totalMany([
      "INVENTORY",
      "TRADE_RECV",
      "CASH_BANK",
      "SHORT_TERM_LOANS_ADV",
      "OTHER_CURR_ASSETS",
    ]),
    items(st, "INVENTORY", "TRADE_RECV", "CASH_BANK", "SHORT_TERM_LOANS_ADV", "OTHER_CURR_ASSETS"),
  );
  setCell(ws, r, 1, "Total", { bold: true });
  setCell(ws, r, 3, money(st.totalAssets), { money: true, bold: true });
  topBorder(ws, r, 1, 3);
  r += 1;

  if (Math.abs(st.balanceDifference) > 0.5) {
    r += 1;
    setCell(
      ws,
      r++,
      1,
      `Note: Assets - Liabilities difference = ${st.balanceDifference.toFixed(2)} (review mapping / enter missing items).`,
      {
        italic: true,
      },
    );
  }

  signatureBlock(ws, r, meta, ctx.forFirmText, ctx.designation);
}

// ---------------- Schedules sheet (SCH) ----------------
export function writeScheduleSheet(ws: ExcelJS.Worksheet, ctx: BuildCtx) {
  applyColumnWidths(ws, [44, 20, 4, 8]);
  const { meta, ledgers, divisor } = ctx;
  const money = (v: number) => v / divisor;
  let r = 1;
  setCell(ws, r++, 1, meta.firmName || "FIRM NAME", { bold: true });
  setCell(ws, r++, 1, meta.asAtLabel || meta.periodLabel, { italic: true });
  if (ctx.unitHeading) setCell(ws, r++, 1, ctx.unitHeading, { italic: true });
  r++;

  const schedule = (no: number, title: string, codes: StatementCode[]) => {
    const list = ctx.ledgers.filter((l) => codes.includes(l.code));
    if (!list.length) return;
    setCell(ws, r++, 1, `SCH-${no}  ${title}`, { bold: true });
    setCell(ws, r, 1, "PARTICULARS", { bold: true });
    setCell(ws, r++, 2, "AMOUNT", { bold: true, align: "right" });
    let tot = 0;
    for (const l of list) {
      const amt = l.code === "PPE" || isAsset(l.code) ? l.signed : -l.signed;
      setCell(ws, r, 1, l.name);
      setCell(ws, r, 2, money(Math.round(amt * 100) / 100), { money: true });
      tot += amt;
      r++;
    }
    setCell(ws, r, 1, "TOTAL", { bold: true });
    setCell(ws, r, 2, money(Math.round(tot * 100) / 100), { money: true, bold: true });
    topBorder(ws, r, 1, 2);
    r += 2;
  };

  void ledgers;
  schedule(1, "SUNDRY CREDITORS", ["TRADE_PAYABLES"]);
  schedule(2, "OTHER CURRENT LIABILITIES", ["OTHER_CURR_LIAB", "SHORT_TERM_PROV"]);
  schedule(3, "OTHER CURRENT ASSETS", ["OTHER_CURR_ASSETS", "INVENTORY"]);
  schedule(4, "LOANS AND ADVANCES", ["SHORT_TERM_LOANS_ADV", "LONG_TERM_LOANS_ADV"]);
  schedule(5, "SUNDRY DEBTORS", ["TRADE_RECV"]);
}

function isAsset(code: StatementCode): boolean {
  return [
    "PPE",
    "NONCURR_INVEST",
    "LONG_TERM_LOANS_ADV",
    "DEFERRED_TAX_ASSET",
    "INVENTORY",
    "TRADE_RECV",
    "CASH_BANK",
    "SHORT_TERM_LOANS_ADV",
    "OTHER_CURR_ASSETS",
  ].includes(code);
}

// ---------------- Fixed Asset Schedule (template using PPE closing balances) ----------------
export function writeFixedAssetSchedule(ws: ExcelJS.Worksheet, ctx: BuildCtx) {
  applyColumnWidths(ws, [34, 8, 16, 16, 16, 16, 16]);
  const { meta, st, divisor } = ctx;
  const money = (v: number) => v / divisor;
  let r = 1;
  setCell(ws, r++, 1, meta.firmName || "FIRM NAME", { bold: true });
  setCell(ws, r++, 1, meta.asAtLabel || meta.periodLabel, { italic: true });
  if (ctx.unitHeading) setCell(ws, r++, 1, ctx.unitHeading, { italic: true });
  r++;
  setCell(ws, r++, 1, "SCH - FIXED ASSETS", { bold: true });
  const headers = [
    "PARTICULARS",
    "RATE",
    "OPENING BALANCE",
    "ADDITION (UPTO 03 OCT)",
    "ADDITION (AFTER 03 OCT)",
    "DEP / ROUND OFF",
    "CLOSING BALANCE",
  ];
  headers.forEach((h, i) =>
    setCell(ws, r, i + 1, h, { bold: true, align: i === 0 ? "left" : "right", wrap: true }),
  );
  r++;

  const ppe = st.byCode.get("PPE") ?? [];
  let totClosing = 0;
  for (const a of ppe) {
    setCell(ws, r, 1, a.name);
    // Rate / opening / additions / depreciation are not in the TB -> left blank for the user.
    setCell(ws, r, 7, money(a.amount), { money: true }); // closing = current book balance
    totClosing += a.amount;
    r++;
  }
  setCell(ws, r, 1, "TOTAL", { bold: true });
  setCell(ws, r, 7, money(Math.round(totClosing * 100) / 100), { money: true, bold: true });
  topBorder(ws, r, 1, 7);
  r += 2;
  setCell(
    ws,
    r++,
    1,
    "Note: Rate, opening balance, additions and depreciation must be entered manually (not available in the trial balance).",
    {
      italic: true,
    },
  );
}
