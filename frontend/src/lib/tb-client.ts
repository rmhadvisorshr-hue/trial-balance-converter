import type { ConvertPayload, EntityType, ParseResult } from "./types";

// Empty by default (bare /api/*), which is correct for standalone dev
// (proxied by vite.config.ts) and standalone production (the backend serves
// its own frontend + /api/* at the same origin). When embedded behind the
// unified staff portal, build-unified-tools.mjs sets this to "/tools/gst-xml"
// so requests reach unified-preview.mjs's proxy for this tool -- see that
// file's SIDECAR_TOOLS entry for "gst-xml".
const API_BASE = import.meta.env.VITE_API_BASE_URL ?? "";

export async function parseTrialBalanceFile(file: File, entity: EntityType): Promise<ParseResult> {
  const form = new FormData();
  form.append("file", file, file.name);
  form.append("entity", entity);
  const res = await fetch(`${API_BASE}/api/tbparse`, { method: "POST", body: form });
  if (!res.ok) {
    const payload = (await res.json().catch(() => null)) as { message?: string } | null;
    throw new Error(payload?.message || "Could not read the trial balance.");
  }
  return (await res.json()) as ParseResult;
}

export async function convertToWorkbook(payload: ConvertPayload): Promise<void> {
  const res = await fetch(`${API_BASE}/api/tbconvert`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(payload),
  });
  if (!res.ok) {
    const data = (await res.json().catch(() => null)) as { message?: string } | null;
    throw new Error(data?.message || "Could not generate the workbook.");
  }
  const blob = await res.blob();
  const disposition = res.headers.get("content-disposition") ?? "";
  const match = disposition.match(/filename="([^"]+)"/i);
  const fileName = match?.[1] || `${payload.meta.firmName || "financials"}.xlsx`;

  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = fileName;
  document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);
}
