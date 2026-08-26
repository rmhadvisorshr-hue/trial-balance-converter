import { loadWorkbook } from "./parsers/workbookLoader";
import { detectFileRoles, type FileRole } from "./parsers/fileRoleDetector";
import { normalizeWorkbook } from "./normalize/normalizeWorkbook";
import { mergeYears } from "./normalize/mergeYears";
import { crossCheckBalanceSheet } from "./validation/crossCheck";
import type { IcaiEntityKind, WorkbookAnalysis } from "./types";

export interface UploadedFile {
  buffer: Buffer | ArrayBuffer;
  fileName: string;
}

export type AnalyzeOutcome =
  | { status: "resolved"; analysis: WorkbookAnalysis }
  | { status: "needs-review"; reason: string; detections: ReturnType<typeof detectFileRoles>["detections"] };

// One or more accounting workbooks in - role detection decides which is the
// current year and which (if any) is the comparative year, purely from each
// workbook's own detected entity/period-end content, never from upload order
// or filenames. `roleOverrides` (fileName -> role) lets the caller (the CA,
// via the review step) correct an ambiguous auto-detection.
export async function analyzeWorkbookFiles(
  files: UploadedFile[],
  roleOverrides?: Record<string, FileRole>,
  entityKindOverride?: IcaiEntityKind,
): Promise<AnalyzeOutcome> {
  const scans = await Promise.all(files.map((f) => loadWorkbook(f.buffer, f.fileName)));

  const detection = roleOverrides
    ? {
        resolved: true,
        reason: "Roles confirmed by the CA.",
        detections: scans.map((s) => ({
          fileName: s.fileName,
          entityName: s.entityName,
          periodEndLabel: s.periodEndLabel,
          detectedYear: null,
          role: roleOverrides[s.fileName] ?? "exclude",
          confidence: 1,
          reason: "Confirmed by the CA.",
        })),
      }
    : detectFileRoles(scans);

  if (!detection.resolved) {
    return { status: "needs-review", reason: detection.reason, detections: detection.detections };
  }

  const currentScan = scans.find((s) => detection.detections.find((d) => d.fileName === s.fileName)?.role === "current");
  const previousScan = scans.find((s) => detection.detections.find((d) => d.fileName === s.fileName)?.role === "previous");
  if (!currentScan) {
    return {
      status: "needs-review",
      reason: "No workbook was assigned the Current Year role - please confirm.",
      detections: detection.detections,
    };
  }

  const current = normalizeWorkbook(currentScan);
  const previous = previousScan ? normalizeWorkbook(previousScan) : null;

  const analysis = mergeYears(current, previous, entityKindOverride);

  const balanceMessages = crossCheckBalanceSheet(analysis, current.balanceSheetTotal, previous?.balanceSheetTotal ?? null);
  analysis.warnings.push(...balanceMessages);

  return { status: "resolved", analysis };
}

export { buildIcaiWorkbook } from "./excel/buildIcaiWorkbook";
export { scaleAnalysis } from "./excel/scale";
