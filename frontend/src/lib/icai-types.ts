// Wire-contract types for the Accounting Workbook -> ICAI Financial
// Statements flow. Deliberately duplicated from backend/src/lib/icai/types.ts
// - see frontend/src/lib/types.ts for why (frontend and backend are two
// independent npm projects, not a shared package).

export type IcaiEntityKind = "partnership" | "proprietor" | "llp";

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

export const ICAI_NOTE_LABEL: Record<IcaiNoteCode, string> = {
  N3_CAPITAL: "Owners'/Partners' Capital Account",
  N4_RESERVES: "Reserves and Surplus",
  N5_LT_BORROWING: "Long-term Borrowings",
  N5_ST_BORROWING: "Short-term Borrowings",
  N6_DEFERRED_TAX: "Deferred Tax Liabilities/(Assets) (Net)",
  N7_OTHER_LT_LIAB: "Other Long-term Liabilities",
  N8_PROVISION: "Provisions",
  N9_TRADE_PAYABLE: "Trade Payables",
  N10_OTHER_CURR_LIAB: "Other Current Liabilities",
  N11_FIXED_ASSET: "Property, Plant and Equipment",
  N12_INVESTMENT: "Investments",
  N13_LOANS_ADVANCES: "Loans and Advances",
  N14_OTHER_NONCURR_ASSET: "Other Non-current Assets",
  N15_INVENTORY: "Inventories",
  N16_TRADE_RECEIVABLE: "Trade Receivables",
  N17_CASH_BANK: "Cash and Bank Balances",
  N18_OTHER_CURR_ASSET: "Other Current Assets",
  N19_REVENUE: "Revenue from Operations",
  N20_OTHER_INCOME: "Other Income",
  N21_COST_OF_CONSTRUCTION: "Cost of Materials/Construction",
  N22_INVENTORY_CHANGE: "Changes in Inventories",
  N23_EMPLOYEE_BENEFIT: "Employee Benefits Expense",
  N24_FINANCE_COST: "Finance Costs",
  N25_DEPRECIATION: "Depreciation and Amortization",
  N26_OTHER_EXPENSE: "Other Expenses",
  PARTNERS_REMUNERATION: "Partners' Remuneration",
  UNMAPPED: "Unmapped (needs review)",
};

export const ICAI_NOTE_CODES = Object.keys(ICAI_NOTE_LABEL) as IcaiNoteCode[];

export interface NormalizedAccount {
  name: string;
  amountCurrent: number;
  amountPrevious: number;
  sourceSheet: string;
  code: IcaiNoteCode;
  autoCode: IcaiNoteCode;
  confident: boolean;
  reason: string;
}

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
  withdrawalBreakup: { label: string; amount: number }[];
  closingBalance: number;
  closingBalancePrevious: number;
}

export interface FixedAssetLine {
  name: string;
  category: string;
  rate: number;
  opening: number;
  additions: number;
  depreciation: number;
  closing: number;
}

export interface SchedulePartyLine {
  name: string;
  amountCurrent: number;
  amountPrevious: number;
}

export interface ScheduleSection {
  heading: string;
  lines: SchedulePartyLine[];
  subtotalCurrent: number;
  subtotalPrevious: number;
}

export interface SupportingSchedule {
  letter: string;
  title: string;
  supportsNote: string;
  sections: ScheduleSection[];
}

export interface ReviewItem {
  sourceSheet: string;
  label: string;
  amountCurrent: number;
  amountPrevious: number;
  suggestedCode: IcaiNoteCode;
  confidence: number;
  reason: string;
}

export interface WorkbookAnalysis {
  entityName: string;
  entityKind: IcaiEntityKind;
  currentYearLabel: string;
  previousYearLabel: string;
  accounts: NormalizedAccount[];
  owners: OwnerCapitalAccount[];
  fixedAssets: { current: FixedAssetLine[]; previous: FixedAssetLine[] };
  schedules: SupportingSchedule[];
  reviewItems: ReviewItem[];
  warnings: string[];
  unitHeading?: string | null;
  // Report Details + selected CA - populated only at generate time, see
  // IcaiWorkflow.tsx. Mirrors StatementMeta's equivalent fields in lib/types.ts.
  place?: string;
  date?: string;
  udin?: string;
  caName?: string;
  caFirmName?: string;
  caFirmType?: string;
  caDesignation?: string;
  caMembershipNo?: string;
  caFirmRegNo?: string;
}

export type FileRole = "current" | "previous" | "exclude";

export interface FileRoleDetection {
  fileName: string;
  entityName: string;
  periodEndLabel: string;
  detectedYear: number | null;
  role: FileRole;
  confidence: number;
  reason: string;
}

export type AnalyzeOutcome =
  | { status: "resolved"; analysis: WorkbookAnalysis }
  | { status: "needs-review"; reason: string; detections: FileRoleDetection[] };
