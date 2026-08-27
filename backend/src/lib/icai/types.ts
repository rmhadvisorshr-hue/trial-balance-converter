// Shared types for the Accounting Workbook -> ICAI Financial Statements pipeline.
// Parallel to lib/tb/types.ts but for a structurally different input: one or
// two already-part-prepared accounting workbooks (Trading/P&L/Capital/Balance
// Sheet/Schedules/Fixed Assets sheets), not a flat Tally trial balance.

export type IcaiEntityKind = "partnership" | "proprietor" | "llp";

// Every line item lands on exactly one of these. Numbered ones map 1:1 to the
// ICAI Guidance Note's own note numbers; the two non-numbered codes are face-
// of-P&L lines that never get their own note in the ICAI format.
export type IcaiNoteCode =
  | "N3_CAPITAL"
  | "N4_RESERVES"
  | "N5_LT_BORROWING"
  | "N5_ST_BORROWING"
  | "N6_DEFERRED_TAX"
  | "N7_OTHER_LT_LIAB"
  | "N8_PROVISION"
  | "N9_TRADE_PAYABLE"
  | "N10_OTHER_CURR_LIAB"
  | "N11_FIXED_ASSET"
  | "N12_INVESTMENT"
  | "N13_LOANS_ADVANCES"
  | "N14_OTHER_NONCURR_ASSET"
  | "N15_INVENTORY"
  | "N16_TRADE_RECEIVABLE"
  | "N17_CASH_BANK"
  | "N18_OTHER_CURR_ASSET"
  | "N19_REVENUE"
  | "N20_OTHER_INCOME"
  | "N21_COST_OF_CONSTRUCTION"
  | "N22_INVENTORY_CHANGE"
  | "N23_EMPLOYEE_BENEFIT"
  | "N24_FINANCE_COST"
  | "N25_DEPRECIATION"
  | "N26_OTHER_EXPENSE"
  | "PARTNERS_REMUNERATION"
  | "UNMAPPED";

export interface IcaiNoteMeta {
  code: IcaiNoteCode;
  noteNo: string; // "3", "5", "19" etc - shared by the LT/ST split codes
  title: string;
  side: "bs" | "pl" | "face";
}

export const ICAI_NOTES: IcaiNoteMeta[] = [
  { code: "N3_CAPITAL", noteNo: "3", title: "Owners'/Partners' Capital Account", side: "bs" },
  { code: "N4_RESERVES", noteNo: "4", title: "Reserves and Surplus", side: "bs" },
  { code: "N5_LT_BORROWING", noteNo: "5", title: "Long-term Borrowings", side: "bs" },
  { code: "N5_ST_BORROWING", noteNo: "5", title: "Short-term Borrowings", side: "bs" },
  { code: "N6_DEFERRED_TAX", noteNo: "6", title: "Deferred Tax Liabilities/(Assets) (Net)", side: "bs" },
  { code: "N7_OTHER_LT_LIAB", noteNo: "7", title: "Other Long-term Liabilities", side: "bs" },
  { code: "N8_PROVISION", noteNo: "8", title: "Provisions", side: "bs" },
  { code: "N9_TRADE_PAYABLE", noteNo: "9", title: "Trade Payables", side: "bs" },
  { code: "N10_OTHER_CURR_LIAB", noteNo: "10", title: "Other Current Liabilities", side: "bs" },
  { code: "N11_FIXED_ASSET", noteNo: "11", title: "Property, Plant and Equipment", side: "bs" },
  { code: "N12_INVESTMENT", noteNo: "12", title: "Investments", side: "bs" },
  { code: "N13_LOANS_ADVANCES", noteNo: "13", title: "Loans and Advances", side: "bs" },
  { code: "N14_OTHER_NONCURR_ASSET", noteNo: "14", title: "Other Non-current Assets", side: "bs" },
  { code: "N15_INVENTORY", noteNo: "15", title: "Inventories", side: "bs" },
  { code: "N16_TRADE_RECEIVABLE", noteNo: "16", title: "Trade Receivables", side: "bs" },
  { code: "N17_CASH_BANK", noteNo: "17", title: "Cash and Bank Balances", side: "bs" },
  { code: "N18_OTHER_CURR_ASSET", noteNo: "18", title: "Other Current Assets", side: "bs" },
  { code: "N19_REVENUE", noteNo: "19", title: "Revenue from Operations", side: "pl" },
  { code: "N20_OTHER_INCOME", noteNo: "20", title: "Other Income", side: "pl" },
  { code: "N21_COST_OF_CONSTRUCTION", noteNo: "21", title: "Cost of Construction and Development Expenses", side: "pl" },
  { code: "N22_INVENTORY_CHANGE", noteNo: "22", title: "Changes in Inventories of Work-in-progress", side: "pl" },
  { code: "N23_EMPLOYEE_BENEFIT", noteNo: "23", title: "Employee Benefits Expense", side: "pl" },
  { code: "N24_FINANCE_COST", noteNo: "24", title: "Finance Costs", side: "pl" },
  { code: "N25_DEPRECIATION", noteNo: "25", title: "Depreciation and Amortization Expense", side: "pl" },
  { code: "N26_OTHER_EXPENSE", noteNo: "26", title: "Other Expenses", side: "pl" },
  { code: "PARTNERS_REMUNERATION", noteNo: "", title: "Partners' Remuneration", side: "face" },
  { code: "UNMAPPED", noteNo: "", title: "Unmapped (needs review)", side: "face" },
];

export const ICAI_NOTE_LABEL: Record<IcaiNoteCode, string> = ICAI_NOTES.reduce(
  (acc, n) => {
    acc[n.code] = n.title;
    return acc;
  },
  {} as Record<IcaiNoteCode, string>,
);

// One line item, after parsing + normalizing across whichever source
// sheet(s)/year(s) it was found in, but before final ICAI classification.
export interface NormalizedAccount {
  name: string;
  amountCurrent: number; // 0 if this account has no current-year figure
  amountPrevious: number; // 0 if this account has no comparative figure
  sourceSheet: string; // for CA traceability ("trad", "cap - Amit Gharat", ...)
  code: IcaiNoteCode;
  autoCode: IcaiNoteCode; // engine's original suggestion, before CA edits
  confident: boolean;
  reason: string; // short human-readable justification, shown in the review UI
}

// A partner/proprietor capital account, fully reconstructed for Note 3(a).
export interface OwnerCapitalAccount {
  name: string;
  pan?: string;
  sharePercent?: number;
  openingBalance: number;
  openingBalancePrevious: number;
  capitalIntroduced: number;
  capitalIntroducedPrevious: number;
  remuneration: number;
  remunerationPrevious: number;
  interest: number;
  interestPrevious: number;
  shareOfProfit: number;
  shareOfProfitPrevious: number;
  otherCredits: number;
  otherCreditsPrevious: number;
  withdrawals: number;
  withdrawalsPrevious: number;
  withdrawalBreakup: { label: string; amount: number }[]; // current year only, for the footnote
  closingBalance: number;
  closingBalancePrevious: number;
}

// One party-wise line in a supporting schedule (Sch A/B/C).
export interface SchedulePartyLine {
  name: string;
  amountCurrent: number;
  amountPrevious: number;
}

export interface ScheduleSection {
  heading: string; // the source block's own title, preserved verbatim
  lines: SchedulePartyLine[];
  subtotalCurrent: number;
  subtotalPrevious: number;
}

export interface SupportingSchedule {
  letter: string; // "A", "B", "C", ... assigned in output order
  title: string;
  supportsNote: string; // e.g. "9" - which ICAI note this schedule backs
  sections: ScheduleSection[];
}

// One asset's movement for one year. The reference format presents Note 11
// as two separate year-by-year movement tables rather than a single merged
// row per asset, so current/previous years are kept as two separate arrays
// (see WorkbookAnalysis.fixedAssets) instead of paired fields here.
export interface FixedAssetLine {
  name: string;
  category: string; // ICAI Note 11 column: "Plant and Equipment", "Office equipment", ...
  rate: number;
  opening: number;
  additions: number;
  depreciation: number;
  closing: number;
}

// A review item surfaced to the CA: either an unmapped account, or a section
// of a source workbook the engine could not interpret at all.
export interface ReviewItem {
  sourceSheet: string;
  label: string;
  amountCurrent: number;
  amountPrevious: number;
  suggestedCode: IcaiNoteCode;
  confidence: number; // 0-1
  reason: string;
}

export interface WorkbookAnalysis {
  entityName: string;
  entityKind: IcaiEntityKind;
  currentYearLabel: string; // e.g. "31st March, 2026"
  previousYearLabel: string; // e.g. "31st March, 2025"
  // Set only when the CA picks a non-actual display unit at generate time
  // (see excel/scale.ts) - null for the analysis the review screen shows.
  unitHeading?: string | null;
  // Report Details + selected CA - set by the frontend only at generate time
  // (see IcaiWorkflow.tsx), never present on the analysis the review screen
  // shows. Mirrors StatementMeta's equivalent fields in lib/tb/types.ts -
  // both satisfy excel/helpers.ts's shared SignatureMeta interface.
  place?: string;
  date?: string;
  udin?: string;
  caName?: string;
  caFirmName?: string;
  caFirmType?: string;
  caDesignation?: string;
  caMembershipNo?: string;
  caFirmRegNo?: string;
  accounts: NormalizedAccount[];
  owners: OwnerCapitalAccount[];
  fixedAssets: { current: FixedAssetLine[]; previous: FixedAssetLine[] };
  schedules: SupportingSchedule[];
  reviewItems: ReviewItem[];
  warnings: string[];
}

// Payload sent from the review screen to the generate endpoint - the CA may
// have edited `accounts[].code` for any UNMAPPED / !confident line.
export interface IcaiGeneratePayload {
  analysis: WorkbookAnalysis;
}
