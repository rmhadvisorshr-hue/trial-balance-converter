import type { ClassifiedLedger, StatementCode } from "./types";
import { round2 } from "./classify";

// Codes whose natural balance is a credit (income / liability / equity).
const CREDIT_POSITIVE = new Set<StatementCode>([
  "REV_OPS",
  "OTHER_INCOME",
  "CLOSING_STOCK_PL",
  "CAPITAL",
  "RESERVES",
  "LONG_TERM_BORROW",
  "SHORT_TERM_BORROW",
  "DEFERRED_TAX_LIAB",
  "TRADE_PAYABLES",
  "OTHER_CURR_LIAB",
  "SHORT_TERM_PROV",
]);

export interface LineItem {
  name: string;
  amount: number; // display magnitude (positive for a natural balance)
}

export function displayAmount(led: ClassifiedLedger): number {
  return CREDIT_POSITIVE.has(led.code) ? round2(-led.signed) : round2(led.signed);
}

export interface Statements {
  byCode: Map<StatementCode, LineItem[]>;
  total: (code: StatementCode) => number;
  totalMany: (codes: StatementCode[]) => number;
  // P&L
  revenue: number;
  otherIncome: number;
  costOfSales: number;
  grossProfit: number;
  employeeBenefit: number;
  financeCost: number;
  depreciation: number;
  otherExpense: number;
  totalIndirectExpense: number;
  netProfit: number;
  // BS
  totalLiabilities: number;
  totalAssets: number;
  balanceDifference: number; // assets - (liabilities incl. net profit)
}

export function computeStatements(ledgers: ClassifiedLedger[]): Statements {
  const byCode = new Map<StatementCode, LineItem[]>();
  for (const led of ledgers) {
    if (led.code === "UNMAPPED") continue;
    const list = byCode.get(led.code) ?? [];
    list.push({ name: led.name, amount: displayAmount(led) });
    byCode.set(led.code, list);
  }

  const total = (code: StatementCode) =>
    round2((byCode.get(code) ?? []).reduce((s, i) => s + i.amount, 0));
  const totalMany = (codes: StatementCode[]) => round2(codes.reduce((s, c) => s + total(c), 0));

  const revenue = total("REV_OPS");
  const otherIncome = total("OTHER_INCOME");
  const purchases = total("PURCHASES");
  const directExp = total("DIRECT_EXP");
  const openingStock = total("OPENING_STOCK");
  const closingStock = total("CLOSING_STOCK_PL");
  const costOfSales = round2(purchases + directExp + openingStock - closingStock);
  const grossProfit = round2(revenue - costOfSales);

  const employeeBenefit = total("EMP_BENEFIT");
  const financeCost = total("FINANCE_COST");
  const depreciation = total("DEPRECIATION");
  const otherExpense = total("OTHER_EXP");
  const totalIndirectExpense = round2(employeeBenefit + financeCost + depreciation + otherExpense);

  const netProfit = round2(grossProfit + otherIncome - totalIndirectExpense);

  // Balance sheet: capital includes reserves + current year profit.
  const liabilityCodes: StatementCode[] = [
    "CAPITAL",
    "RESERVES",
    "LONG_TERM_BORROW",
    "SHORT_TERM_BORROW",
    "DEFERRED_TAX_LIAB",
    "TRADE_PAYABLES",
    "OTHER_CURR_LIAB",
    "SHORT_TERM_PROV",
  ];
  const assetCodes: StatementCode[] = [
    "PPE",
    "NONCURR_INVEST",
    "LONG_TERM_LOANS_ADV",
    "DEFERRED_TAX_ASSET",
    "INVENTORY",
    "TRADE_RECV",
    "CASH_BANK",
    "SHORT_TERM_LOANS_ADV",
    "OTHER_CURR_ASSETS",
  ];

  const totalLiabilities = round2(totalMany(liabilityCodes) + netProfit);
  const totalAssets = totalMany(assetCodes);
  const balanceDifference = round2(totalAssets - totalLiabilities);

  return {
    byCode,
    total,
    totalMany,
    revenue,
    otherIncome,
    costOfSales,
    grossProfit,
    employeeBenefit,
    financeCost,
    depreciation,
    otherExpense,
    totalIndirectExpense,
    netProfit,
    totalLiabilities,
    totalAssets,
    balanceDifference,
  };
}
