import type { AnalyzeOutcome, FileRole, IcaiEntityKind, WorkbookAnalysis } from "./icai-types";
import type { FiguresUnit } from "./types";
import { triggerDownload } from "./download";

const API_BASE = import.meta.env.VITE_API_BASE_URL ?? "";

export async function analyzeIcaiWorkbooks(
  files: File[],
  roleOverrides?: Record<string, FileRole>,
  entityKind?: IcaiEntityKind,
): Promise<AnalyzeOutcome> {
  const form = new FormData();
  for (const f of files) form.append("files", f, f.name);
  if (roleOverrides) form.append("roleOverrides", JSON.stringify(roleOverrides));
  if (entityKind) form.append("entityKind", entityKind);

  const res = await fetch(`${API_BASE}/api/icai/analyze`, { method: "POST", body: form });
  if (!res.ok) {
    const payload = (await res.json().catch(() => null)) as { message?: string } | null;
    throw new Error(payload?.message || "Could not read the accounting workbook(s).");
  }
  return (await res.json()) as AnalyzeOutcome;
}

export async function generateIcaiWorkbook(analysis: WorkbookAnalysis, figuresUnit: FiguresUnit): Promise<void> {
  const res = await fetch(`${API_BASE}/api/icai/generate`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ analysis, figuresUnit }),
  });
  if (!res.ok) {
    const data = (await res.json().catch(() => null)) as { message?: string } | null;
    throw new Error(data?.message || "Could not generate the financial statements.");
  }
  const blob = await res.blob();
  triggerDownload(blob, res.headers.get("content-disposition"), `${analysis.entityName || "financials"}.xlsx`);
}
