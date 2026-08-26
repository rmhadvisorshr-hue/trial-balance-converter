import type { WorkbookScan } from "./workbookLoader";

export type FileRole = "current" | "previous" | "exclude";

export interface FileRoleDetection {
  fileName: string;
  entityName: string;
  periodEndLabel: string;
  detectedYear: number | null;
  role: FileRole;
  confidence: number; // 0-1
  reason: string;
}

export interface RoleDetectionResult {
  resolved: boolean; // true => detections can be used as-is; false => ask the CA to confirm/correct
  detections: FileRoleDetection[];
  reason: string;
}

// A period-end label is free text taken verbatim from the workbook ("31ST
// MARCH 2026", "31st March, 2025", ...) - the only thing every format shares
// is a 4-digit year, which is enough to order workbooks chronologically
// without needing to fully parse the date.
function extractYear(periodEndLabel: string): number | null {
  const m = periodEndLabel.match(/(\d{4})/);
  return m ? Number(m[1]) : null;
}

// Assigns each uploaded workbook a role (current/previous/exclude) from its
// own detected entity name + period end - never from filenames or upload
// order. Resolves automatically only when the answer is unambiguous;
// otherwise returns resolved:false so the caller can ask the CA to confirm.
export function detectFileRoles(scans: WorkbookScan[]): RoleDetectionResult {
  const withYear = scans.map((s) => ({ scan: s, year: extractYear(s.periodEndLabel) }));

  if (scans.length === 1) {
    const { scan, year } = withYear[0];
    return {
      resolved: true,
      reason: "Single workbook uploaded - treated as the current year.",
      detections: [
        {
          fileName: scan.fileName,
          entityName: scan.entityName,
          periodEndLabel: scan.periodEndLabel,
          detectedYear: year,
          role: "current",
          confidence: year ? 0.9 : 0.5,
          reason: year
            ? `Only workbook uploaded; detected period end ${scan.periodEndLabel}.`
            : "Only workbook uploaded; no period-end date could be detected - confirm this is correct.",
        },
      ],
    };
  }

  const years = withYear.map((w) => w.year);
  const allDetected = years.every((y) => y != null);
  const distinctYears = new Set(years).size;

  if (scans.length === 2 && allDetected && distinctYears === 2) {
    const sorted = [...withYear].sort((a, b) => (b.year as number) - (a.year as number));
    const [newer, older] = sorted;
    return {
      resolved: true,
      reason: "Two workbooks with distinct detected financial years - assigned by chronological order.",
      detections: [
        {
          fileName: newer.scan.fileName,
          entityName: newer.scan.entityName,
          periodEndLabel: newer.scan.periodEndLabel,
          detectedYear: newer.year,
          role: "current",
          confidence: 0.9,
          reason: `Later period end (${newer.scan.periodEndLabel}) - treated as the current year.`,
        },
        {
          fileName: older.scan.fileName,
          entityName: older.scan.entityName,
          periodEndLabel: older.scan.periodEndLabel,
          detectedYear: older.year,
          role: "previous",
          confidence: 0.9,
          reason: `Earlier period end (${older.scan.periodEndLabel}) - treated as the comparative year.`,
        },
      ],
    };
  }

  // Everything else needs a human: >2 files, missing period-end detection,
  // or two files that appear to be the same year.
  let reason = "Could not confidently determine the current/previous year assignment - please confirm.";
  if (scans.length > 2) reason = `${scans.length} workbooks uploaded - confirm which years each one represents.`;
  else if (!allDetected) reason = "Could not detect a financial year in one or more workbooks - confirm manually.";
  else if (distinctYears < scans.length) reason = "Two or more workbooks appear to be for the same financial year - confirm which is current.";

  return {
    resolved: false,
    reason,
    detections: withYear.map(({ scan, year }) => ({
      fileName: scan.fileName,
      entityName: scan.entityName,
      periodEndLabel: scan.periodEndLabel,
      detectedYear: year,
      role: "exclude",
      confidence: 0.3,
      reason: "Role not auto-assigned - please select Current Year, Previous Year, or Exclude.",
    })),
  };
}
