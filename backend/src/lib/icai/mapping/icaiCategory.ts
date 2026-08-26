import type { IcaiNoteCode } from "../types";

export interface CategoryResult {
  code: IcaiNoteCode;
  confident: boolean;
  reason: string;
}

// Same style of heuristic as lib/tb/classify.ts's classifyExpense(), reused
// here for the indirect-P&L expense lines this pipeline reads off a
// "profit and loss account" T-block instead of a Tally trial balance.
export function categorizeExpenseLine(name: string): CategoryResult {
  const n = name.toLowerCase();
  const has = (...words: string[]) => words.some((w) => n.includes(w));

  if (has("depreciation", "amortis")) {
    return { code: "N25_DEPRECIATION", confident: true, reason: "Matches depreciation/amortisation." };
  }
  if (has("interest", "od cc", "o.d", "processing fee", "bank guarantee", "loan processing", "od interest")) {
    return { code: "N24_FINANCE_COST", confident: true, reason: "Matches an interest/borrowing-cost keyword." };
  }
  if (
    has(
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
      "incentive",
    )
  ) {
    return { code: "N23_EMPLOYEE_BENEFIT", confident: true, reason: "Matches an employee-cost keyword." };
  }
  return { code: "N26_OTHER_EXPENSE", confident: true, reason: "No specific expense keyword matched; defaulted to Other Expenses." };
}

// Categorizes a Balance Sheet leaf item that wasn't already captured by the
// trading/capital/schedule/fixed-asset parsers (typically bank loans,
// deposits, and the cash/bank/GST/advance-tax break-up, which in practice
// only ever show up on the Balance Sheet sheet itself). Returns null for
// anything not confidently recognizable, rather than guessing - callers
// should surface a null result as a review item.
export function categorizeBalanceSheetLine(
  leaf: { group: string; label: string; side: "liability" | "asset" },
): CategoryResult | null {
  const text = `${leaf.group} ${leaf.label}`.toLowerCase();
  const has = (...words: string[]) => words.some((w) => text.includes(w));

  if (leaf.side === "liability") {
    if (has("secured loan", "unsecured loan", "term loan", "bank loan")) {
      return {
        code: "N5_LT_BORROWING",
        confident: false,
        reason:
          "Bank/related-party loan defaulted to long-term borrowing because the source doesn't separately identify a current-maturity portion - confirm with loan terms.",
      };
    }
    if (has("tds payable")) return { code: "N10_OTHER_CURR_LIAB", confident: true, reason: "TDS payable." };
    if (has("gst") && has("payable")) return { code: "N10_OTHER_CURR_LIAB", confident: true, reason: "GST payable." };
    return null;
  }

  if (has("deposit")) {
    return {
      code: "N14_OTHER_NONCURR_ASSET",
      confident: false,
      reason: "Classified as a non-current deposit - confirm this is a long-term/refundable deposit.",
    };
  }
  if (has("advance tax")) return { code: "N13_LOANS_ADVANCES", confident: true, reason: "Advance tax paid." };
  if (has("tds receivable")) return { code: "N13_LOANS_ADVANCES", confident: true, reason: "TDS receivable." };
  if (has("gst")) return { code: "N18_OTHER_CURR_ASSET", confident: true, reason: "GST input/balance." };
  if (has("bank", "cash")) return { code: "N17_CASH_BANK", confident: true, reason: "Cash/bank balance." };
  return null;
}

// Groups a fixed asset's own description into one of the ICAI Note 11
// columns. Falls back to "Others (specify nature)" - itself an ICAI-
// recognised column, not an invented one - rather than guessing wrong.
export function categorizeFixedAsset(name: string): string {
  const n = name.toLowerCase();
  const has = (...words: string[]) => words.some((w) => n.includes(w));

  if (has("land")) return "Land";
  if (has("building")) return "Buildings";
  if (has("vehicle", "car", "bike", "scooter", "truck")) return "Vehicles";
  if (has("furniture", "sofa", "chair", "table", "fixture")) return "Furniture & Fixtures";
  if (
    has(
      "computer",
      "software",
      "printer",
      "camera",
      "telephone",
      "mobile",
      "led",
      "split a/c",
      "ac ",
      "counter",
      "coffee machine",
      "lock machine",
      "mouse",
      "office equip",
    )
  ) {
    return "Office Equipment";
  }
  if (has("plant", "machinery", "transformer", "solar", "generator", "equipment")) {
    return "Plant and Equipment";
  }
  return "Others (specify nature)";
}
