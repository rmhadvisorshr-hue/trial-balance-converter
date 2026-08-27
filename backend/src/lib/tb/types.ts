// Shared types for the Trial Balance -> Financial Statements pipeline.

export type EntityType = "pvtltd" | "partnership" | "llp" | "proprietor";

export const ENTITY_LABELS: Record<EntityType, string> = {
  pvtltd: "Private Limited Company",
  partnership: "Partnership Firm",
  llp: "Limited Liability Partnership (LLP)",
  proprietor: "Proprietor",
};

// The four statement layout families - one per entity type. LLP has its own
// family (Sources/Application of Funds BS, Capital+Current split, Form 8) -
// it is not a Partnership variant despite sharing the appropriation-account
// tax engine.
export type EntityFamily = "corporate" | "noncorporate" | "proprietor" | "llp";

export function entityFamily(entity: EntityType): EntityFamily {
  if (entity === "pvtltd") return "corporate";
  if (entity === "proprietor") return "proprietor";
  if (entity === "llp") return "llp";
  return "noncorporate"; // partnership
}

// Scale applied to displayed/exported statement figures. The underlying trial
// balance data and calculations are always done in actual rupees; this only
// affects how amounts are shown (UI + Excel export).
export type FiguresUnit = "actual" | "thousands" | "lakhs";

export const FIGURES_UNIT_LABELS: Record<FiguresUnit, string> = {
  actual: "Actual",
  thousands: "Thousands ('000)",
  lakhs: "Lakhs",
};

export const FIGURES_UNIT_DIVISOR: Record<FiguresUnit, number> = {
  actual: 1,
  thousands: 1_000,
  lakhs: 100_000,
};

// Heading shown on each statement when a non-actual unit is selected.
export const FIGURES_UNIT_HEADING: Record<FiguresUnit, string | null> = {
  actual: null,
  thousands: "Figures in Thousands ('000)",
  lakhs: "Figures in Lakhs",
};

// A single raw row parsed from the Tally trial balance.
export interface TBRow {
  name: string;
  parentGroup: string; // the Tally group this ledger sits under ("" for top-level groups)
  isGroup: boolean; // true if this is a group header row (has children)
  debit: number;
  credit: number;
  level: number; // indentation depth (0 = group, 1+ = ledger)
}

// Canonical statement line codes. Each maps a ledger into a place in the P&L or BS.
export type StatementCode =
  // ---- Profit & Loss ----
  | "REV_OPS" // Revenue from operations / Sales
  | "OTHER_INCOME" // Indirect / other income
  | "PURCHASES" // Purchases / cost of materials
  | "DIRECT_EXP" // Direct expenses (trading)
  | "OPENING_STOCK"
  | "CLOSING_STOCK_PL" // closing stock credited in trading a/c
  | "EMP_BENEFIT" // salaries, wages, staff welfare, PF/ESIC
  | "FINANCE_COST" // interest, OD/CC interest, processing fees, bank guarantee
  | "DEPRECIATION"
  | "OTHER_EXP" // all remaining indirect expenses
  // ---- Balance Sheet : Equity & Liabilities ----
  | "CAPITAL" // partners'/proprietor capital or share capital
  | "RESERVES" // reserves & surplus / accumulated P&L
  | "LONG_TERM_BORROW" // secured/unsecured term loans
  | "SHORT_TERM_BORROW" // bank OD / CC / short term loans
  | "DEFERRED_TAX_LIAB"
  | "TRADE_PAYABLES" // sundry creditors / bills payable
  | "OTHER_CURR_LIAB" // duties & taxes, outstanding expenses
  | "SHORT_TERM_PROV" // provision for tax etc.
  // ---- Balance Sheet : Assets ----
  | "PPE" // fixed assets
  | "NONCURR_INVEST" // investments
  | "LONG_TERM_LOANS_ADV" // deposits, capital advances
  | "DEFERRED_TAX_ASSET"
  | "INVENTORY" // closing stock (balance sheet)
  | "TRADE_RECV" // sundry debtors
  | "CASH_BANK" // cash in hand + bank balances
  | "SHORT_TERM_LOANS_ADV" // loans & advances (asset)
  | "OTHER_CURR_ASSETS" // TDS receivable, advance tax, prepaid, accrued income
  // ---- Special ----
  | "UNMAPPED"; // could not be classified - flagged for the user

export type StatementSide = "pl" | "bs" | "none";

export interface StatementCodeMeta {
  code: StatementCode;
  label: string;
  side: StatementSide;
  // For the BS: is this a debit-natured (asset/expense) or credit-natured (liab/income) line.
  // Used only for display grouping, not for math (math uses signed amounts).
  group: string;
}

export const STATEMENT_CODES: StatementCodeMeta[] = [
  { code: "REV_OPS", label: "Revenue from operations", side: "pl", group: "Income" },
  { code: "OTHER_INCOME", label: "Other income", side: "pl", group: "Income" },
  { code: "PURCHASES", label: "Purchases / Cost of materials", side: "pl", group: "Cost of sales" },
  { code: "DIRECT_EXP", label: "Direct expenses", side: "pl", group: "Cost of sales" },
  { code: "OPENING_STOCK", label: "Opening stock", side: "pl", group: "Cost of sales" },
  {
    code: "CLOSING_STOCK_PL",
    label: "Closing stock (trading)",
    side: "pl",
    group: "Cost of sales",
  },
  { code: "EMP_BENEFIT", label: "Employee benefit expenses", side: "pl", group: "Expenses" },
  { code: "FINANCE_COST", label: "Finance costs", side: "pl", group: "Expenses" },
  { code: "DEPRECIATION", label: "Depreciation", side: "pl", group: "Expenses" },
  { code: "OTHER_EXP", label: "Other expenses", side: "pl", group: "Expenses" },

  { code: "CAPITAL", label: "Capital account", side: "bs", group: "Equity & Liabilities" },
  { code: "RESERVES", label: "Reserves & surplus", side: "bs", group: "Equity & Liabilities" },
  {
    code: "LONG_TERM_BORROW",
    label: "Long term borrowings",
    side: "bs",
    group: "Equity & Liabilities",
  },
  {
    code: "SHORT_TERM_BORROW",
    label: "Short term borrowings",
    side: "bs",
    group: "Equity & Liabilities",
  },
  {
    code: "DEFERRED_TAX_LIAB",
    label: "Deferred tax liability",
    side: "bs",
    group: "Equity & Liabilities",
  },
  {
    code: "TRADE_PAYABLES",
    label: "Trade payables (creditors)",
    side: "bs",
    group: "Equity & Liabilities",
  },
  {
    code: "OTHER_CURR_LIAB",
    label: "Other current liabilities",
    side: "bs",
    group: "Equity & Liabilities",
  },
  {
    code: "SHORT_TERM_PROV",
    label: "Short term provisions",
    side: "bs",
    group: "Equity & Liabilities",
  },

  { code: "PPE", label: "Property, plant & equipment", side: "bs", group: "Assets" },
  { code: "NONCURR_INVEST", label: "Non-current investments", side: "bs", group: "Assets" },
  { code: "LONG_TERM_LOANS_ADV", label: "Long term loans & advances", side: "bs", group: "Assets" },
  { code: "DEFERRED_TAX_ASSET", label: "Deferred tax asset", side: "bs", group: "Assets" },
  { code: "INVENTORY", label: "Inventories (closing stock)", side: "bs", group: "Assets" },
  { code: "TRADE_RECV", label: "Trade receivables (debtors)", side: "bs", group: "Assets" },
  { code: "CASH_BANK", label: "Cash & bank balances", side: "bs", group: "Assets" },
  {
    code: "SHORT_TERM_LOANS_ADV",
    label: "Short term loans & advances",
    side: "bs",
    group: "Assets",
  },
  { code: "OTHER_CURR_ASSETS", label: "Other current assets", side: "bs", group: "Assets" },

  { code: "UNMAPPED", label: "Unmapped (needs review)", side: "none", group: "Review" },
];

export const STATEMENT_CODE_LABEL: Record<StatementCode, string> = STATEMENT_CODES.reduce(
  (acc, m) => {
    acc[m.code] = m.label;
    return acc;
  },
  {} as Record<StatementCode, string>,
);

// A ledger after classification, including the (editable) mapped statement code.
export interface ClassifiedLedger {
  name: string;
  parentGroup: string;
  debit: number;
  credit: number;
  signed: number; // debit - credit
  code: StatementCode;
  autoCode: StatementCode; // the engine's original suggestion (before user edits)
  confident: boolean; // false => flagged in the review UI
}

export interface ParseResult {
  entity: EntityType;
  meta: StatementMeta;
  ledgers: ClassifiedLedger[];
  warnings: string[];
}

export interface StatementMeta {
  firmName: string; // e.g. "M/S ORBIT CORPORATION"
  periodLabel: string; // e.g. "1-Apr-24 to 31-Mar-25"
  asAtLabel: string; // e.g. "as at 31-Mar-25"
  cin?: string; // Company Identification Number - Schedule III (pvtltd) only
  llpin?: string; // LLP Identification Number - LLP only
  caName?: string; // e.g. "CA Namrata Prakash Sharma" - from the selected CA profile
  caFirmName?: string; // the CA firm's name, e.g. "Namrata Prakash Sharma"
  caFirmType?: string; // e.g. "Chartered Accountants"
  caDesignation?: string; // the CA's own designation, e.g. "Proprietor" - not the client entity's
  caMembershipNo?: string;
  caFirmRegNo?: string; // FRN - optional
  udin?: string;
  place?: string;
  date?: string;
  figuresUnit: FiguresUnit; // display scale for generated statements (default "actual")
}

// The generated financial statement's file format.
export type OutputFormat = "excel" | "pdf";

export const OUTPUT_FORMAT_LABELS: Record<OutputFormat, string> = {
  excel: "Excel (.xlsx)",
  pdf: "PDF (.pdf)",
};

// Which layout to generate. Pvt Ltd companies always get "scheduleIII" (the
// only output offered for that entity - see entityFamily/family === "corporate"
// in builders/index.ts); partnership/LLP/proprietor entities always get
// "statutory" (the vertical BS/P&L layout) - there is no user-facing choice.
export type StatementStyle = "statutory" | "scheduleIII";

export const STATEMENT_STYLE_LABELS: Record<StatementStyle, string> = {
  statutory: "Full Financial Statements (Balance Sheet, P&L, Notes)",
  scheduleIII: "Full Schedule III Financials (BS/P&L linked to Notes, DTL/DTA)",
};

// Payload sent from the client review screen to /api/tbconvert.
export interface ConvertPayload {
  entity: EntityType;
  meta: StatementMeta;
  ledgers: ClassifiedLedger[];
  outputFormat: OutputFormat;
  statementStyle: StatementStyle;
}
