// Wire-contract types for the centralized CA Profile store, shared by both
// the Trial Balance and Accounting Workbook workflows. Deliberately
// duplicated from backend/src/lib/caProfiles.ts - see lib/types.ts for why
// (frontend and backend are two independent npm projects, not a shared package).

export interface CAProfile {
  id: string;
  caName: string;
  designation: string;
  firmName: string;
  membershipNo: string;
  frn?: string;
  place: string;
  firmType: string;
  createdAt: string;
  updatedAt: string;
}

export type CAProfileInput = Omit<CAProfile, "id" | "createdAt" | "updatedAt">;

export const DEFAULT_CA_FIRM_TYPE = "Chartered Accountants";

export const CA_DESIGNATION_OPTIONS = [
  "Proprietor",
  "Partner",
  "Designated Partner",
  "Authorized Signatory",
];

export function formatCAProfileLabel(p: CAProfile): string {
  return `${p.caName} — ${p.designation}`;
}
