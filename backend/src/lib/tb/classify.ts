import type { TBRow, ClassifiedLedger, StatementCode, EntityType } from "./types";
import { entityFamily } from "./types";

const norm = (s: string) => s.toLowerCase().replace(/\s+/g, " ").trim();
const has = (s: string, ...words: string[]) => words.some((w) => s.includes(w));

// Map a Tally group name to a coarse bucket.
function groupBucket(group: string): string {
  const g = norm(group);
  if (has(g, "capital")) return "capital";
  if (has(g, "reserve", "surplus")) return "reserves";
  if (has(g, "loan") && has(g, "liab")) return "loans_liab";
  if (has(g, "secured loan", "unsecured loan", "bank od", "bank o/d", "term loan"))
    return "loans_liab";
  if (has(g, "current liab", "provision", "duties", "sundry credit")) return "curr_liab";
  if (has(g, "fixed asset")) return "fixed_assets";
  if (has(g, "investment")) return "investments";
  if (
    has(g, "current asset", "bank account", "cash", "deposit", "loans & advances", "sundry debtor")
  )
    return "curr_assets";
  // NOTE: check "indirect" BEFORE "direct" because "indirect" contains the substring "direct".
  if (has(g, "indirect inc")) return "indirect_income";
  if (has(g, "indirect exp")) return "indirect_exp";
  if (has(g, "sales", "direct inc")) return "sales";
  if (has(g, "purchase")) return "purchases";
  if (has(g, "direct exp")) return "direct_exp";
  if (has(g, "expenses")) return "indirect_exp";
  if (has(g, "stock")) return "stock";
  if (has(g, "profit & loss", "profit and loss")) return "pl_account";
  if (has(g, "misc")) return "misc";
  return "unknown";
}

function classifyExpense(name: string): StatementCode {
  const n = norm(name);
  if (has(n, "depreciation", "amortis")) return "DEPRECIATION";
  if (
    has(
      n,
      "interest",
      "od cc",
      "o.d",
      "processing fee",
      "bank guarantee",
      "loan processing",
      "od interest",
    )
  )
    return "FINANCE_COST";
  if (
    has(
      n,
      "salary",
      "wages",
      "bonus",
      "staff welfare",
      "staff",
      "pf ",
      "provident",
      "esic",
      "gratuity",
      "stipend",
      "remuneration",
      "incentive",
    )
  )
    return "EMP_BENEFIT";
  return "OTHER_EXP";
}

function classifyLeaf(row: TBRow, entity: EntityType): { code: StatementCode; confident: boolean } {
  const bucket = groupBucket(row.parentGroup || row.name);
  const n = norm(row.name);

  switch (bucket) {
    case "capital":
      return { code: "CAPITAL", confident: true };
    case "reserves":
      return { code: "RESERVES", confident: true };
    case "pl_account":
      return { code: "RESERVES", confident: false };
    case "loans_liab": {
      if (has(n, "od", "o/d", "overdraft", "cc ", "cash credit"))
        return { code: "SHORT_TERM_BORROW", confident: true };
      return { code: "LONG_TERM_BORROW", confident: true };
    }
    case "curr_liab": {
      if (has(n, "provision")) return { code: "SHORT_TERM_PROV", confident: true };
      if (has(n, "sundry credit", "creditor", "bills payable", "trade payable"))
        return { code: "TRADE_PAYABLES", confident: true };
      if (has(n, "deferred tax")) return { code: "DEFERRED_TAX_LIAB", confident: true };
      return { code: "OTHER_CURR_LIAB", confident: true };
    }
    case "fixed_assets":
      return { code: "PPE", confident: true };
    case "investments":
      return { code: "NONCURR_INVEST", confident: true };
    case "curr_assets": {
      if (has(n, "sundry debtor", "debtor", "receivable") && !has(n, "tds"))
        return { code: "TRADE_RECV", confident: true };
      if (has(n, "cash", "bank")) return { code: "CASH_BANK", confident: true };
      if (
        has(
          n,
          "deposit",
          "loans & advances",
          "loan & advance",
          "advance recoverable",
          "security deposit",
        )
      )
        return { code: "SHORT_TERM_LOANS_ADV", confident: true };
      if (has(n, "tds", "advance tax", "prepaid", "accrued", "gst input", "input credit"))
        return { code: "OTHER_CURR_ASSETS", confident: true };
      if (has(n, "closing stock", "inventory", "stock"))
        return { code: "INVENTORY", confident: true };
      return { code: "OTHER_CURR_ASSETS", confident: false };
    }
    case "sales":
      return { code: "REV_OPS", confident: true };
    case "indirect_income":
      return { code: "OTHER_INCOME", confident: true };
    case "purchases":
      return { code: "PURCHASES", confident: true };
    case "direct_exp":
      return { code: "DIRECT_EXP", confident: true };
    case "indirect_exp":
      return { code: classifyExpense(row.name), confident: true };
    case "stock":
      return { code: "INVENTORY", confident: false };
    default:
      break;
  }

  // Fallback: try to classify by the ledger name's own signedness later, but flag for review.
  // Heuristics on name when group is unknown.
  if (has(n, "depreciation")) return { code: "DEPRECIATION", confident: false };
  if (has(n, "salary", "wages")) return { code: "EMP_BENEFIT", confident: false };
  if (has(n, "sales", "revenue", "fees", "commission") && row.credit > row.debit)
    return { code: "REV_OPS", confident: false };
  if (has(n, "capital")) return { code: "CAPITAL", confident: false };
  if (has(n, "creditor")) return { code: "TRADE_PAYABLES", confident: false };
  if (has(n, "debtor")) return { code: "TRADE_RECV", confident: false };

  // Last resort: classify by balance nature so totals are not wildly off, but flag it.
  const credit = row.credit > row.debit;
  void entity;
  return { code: credit ? "OTHER_CURR_LIAB" : "OTHER_CURR_ASSETS", confident: false };
}

export function classifyLedgers(leaves: TBRow[], entity: EntityType): ClassifiedLedger[] {
  void entityFamily(entity);
  return leaves.map((row) => {
    const { code, confident } = classifyLeaf(row, entity);
    return {
      name: row.name,
      parentGroup: row.parentGroup,
      debit: row.debit,
      credit: row.credit,
      signed: round2(row.debit - row.credit),
      code,
      autoCode: code,
      confident,
    };
  });
}

export function round2(n: number): number {
  return Math.round((n + Number.EPSILON) * 100) / 100;
}
