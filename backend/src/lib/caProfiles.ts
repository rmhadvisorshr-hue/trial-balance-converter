// Centralized CA (Chartered Accountant) profile store, shared by both the
// Trial Balance and Accounting Workbook pipelines - see excel/helpers.ts's
// signatureBlock(), which both pipelines call with the fields defined here.
// Plain JSON file, not a database: this project has no DB dependency at all,
// and a handful of office CA profiles don't warrant adding one (same
// data/-folder-in-project-root convention as ITRBankStatementAnalyzer/backend/db.py).
import { randomUUID } from "node:crypto";
import { mkdir, readFile, writeFile } from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";

export interface CAProfile {
  id: string;
  caName: string; // e.g. "CA Namrata Prakash Sharma"
  designation: string; // "Proprietor" | "Partner" | "Designated Partner" | "Authorized Signatory"
  firmName: string; // the CA firm's name, e.g. "Namrata Prakash Sharma"
  membershipNo: string;
  frn?: string; // optional - not every firm type has one
  place: string; // default office place - report Place is a separate, editable field
  firmType: string; // e.g. "Chartered Accountants"
  createdAt: string;
  updatedAt: string;
}

export type CAProfileInput = Omit<CAProfile, "id" | "createdAt" | "updatedAt">;

const dirname = path.dirname(fileURLToPath(import.meta.url));
const DATA_PATH =
  process.env.CA_PROFILES_PATH || path.join(dirname, "..", "..", "data", "ca-profiles.json");

async function readAll(): Promise<CAProfile[]> {
  try {
    const raw = await readFile(DATA_PATH, "utf8");
    return JSON.parse(raw) as CAProfile[];
  } catch (err) {
    if ((err as NodeJS.ErrnoException).code === "ENOENT") return [];
    throw err;
  }
}

async function writeAll(profiles: CAProfile[]): Promise<void> {
  await mkdir(path.dirname(DATA_PATH), { recursive: true });
  await writeFile(DATA_PATH, JSON.stringify(profiles, null, 2), "utf8");
}

export function validateCAProfileInput(input: Partial<CAProfileInput>): string | null {
  if (!input.caName?.trim()) return "CA Name is required.";
  if (!input.designation?.trim()) return "Designation is required.";
  if (!input.firmName?.trim()) return "Firm Name is required.";
  if (!input.membershipNo?.trim()) return "Membership Number is required.";
  if (!input.place?.trim()) return "Place is required.";
  return null;
}

export function listCAProfiles(): Promise<CAProfile[]> {
  return readAll();
}

export async function createCAProfile(input: CAProfileInput): Promise<CAProfile> {
  const profiles = await readAll();
  const now = new Date().toISOString();
  const profile: CAProfile = {
    id: randomUUID(),
    caName: input.caName.trim(),
    designation: input.designation.trim(),
    firmName: input.firmName.trim(),
    membershipNo: input.membershipNo.trim(),
    frn: input.frn?.trim() || undefined,
    place: input.place.trim(),
    firmType: input.firmType?.trim() || "Chartered Accountants",
    createdAt: now,
    updatedAt: now,
  };
  profiles.push(profile);
  await writeAll(profiles);
  return profile;
}

export async function updateCAProfile(id: string, input: CAProfileInput): Promise<CAProfile | null> {
  const profiles = await readAll();
  const idx = profiles.findIndex((p) => p.id === id);
  if (idx === -1) return null;
  const updated: CAProfile = {
    ...profiles[idx],
    caName: input.caName.trim(),
    designation: input.designation.trim(),
    firmName: input.firmName.trim(),
    membershipNo: input.membershipNo.trim(),
    frn: input.frn?.trim() || undefined,
    place: input.place.trim(),
    firmType: input.firmType?.trim() || "Chartered Accountants",
    updatedAt: new Date().toISOString(),
  };
  profiles[idx] = updated;
  await writeAll(profiles);
  return updated;
}

export async function deleteCAProfile(id: string): Promise<boolean> {
  const profiles = await readAll();
  const next = profiles.filter((p) => p.id !== id);
  if (next.length === profiles.length) return false;
  await writeAll(next);
  return true;
}

// Shared gate used by both /api/tbconvert (Excel path only) and
// /api/icai/generate before a workbook is built - see server.ts. There is no
// fallback CA identity: a report cannot be generated without one selected.
export function hasRequiredCaFields(x: {
  caName?: string;
  caFirmName?: string;
  caDesignation?: string;
  caMembershipNo?: string;
}): boolean {
  return Boolean(
    x.caName?.trim() && x.caFirmName?.trim() && x.caDesignation?.trim() && x.caMembershipNo?.trim(),
  );
}

export const CA_REQUIRED_MESSAGE = "Please select a CA / Signing Authority before generating the report.";
