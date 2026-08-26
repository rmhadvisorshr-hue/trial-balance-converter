import type ExcelJS from "exceljs";

// A worksheet flattened into plain text/number cells - every parser in this
// module works off this instead of touching ExcelJS cell objects directly.
export interface Grid {
  rows: GridRow[];
}

export interface GridRow {
  r: number;
  cells: string[]; // 1-indexed: cells[0] is unused, cells[1] is column A
  nums: (number | null)[]; // same indexing; null where the cell isn't numeric
}

function cellText(value: unknown): string {
  if (value == null) return "";
  if (typeof value === "object") {
    const v = value as { result?: unknown; text?: unknown; richText?: Array<{ text: string }> };
    if (Array.isArray(v.richText)) return v.richText.map((r) => r.text).join("").trim();
    if (v.text != null) return String(v.text).trim();
    if (v.result != null) return String(v.result).trim();
    return "";
  }
  return String(value).trim();
}

function cellNumber(value: unknown): number | null {
  if (value == null) return null;
  if (typeof value === "number") return Number.isFinite(value) ? value : null;
  if (typeof value === "object") {
    const v = value as { result?: unknown };
    if (v.result != null) return cellNumber(v.result);
    return null;
  }
  const s = String(value)
    .replace(/[,\s₹]/g, "")
    .replace(/^\((.*)\)$/, "-$1");
  if (!s || s === "-" || /^-*$/.test(s)) return null;
  const n = Number(s);
  return Number.isFinite(n) ? n : null;
}

export function readGrid(ws: ExcelJS.Worksheet): Grid {
  const rows: GridRow[] = [];
  for (let r = 1; r <= ws.rowCount; r++) {
    const row = ws.getRow(r);
    const cells: string[] = [""];
    const nums: (number | null)[] = [null];
    let hasContent = false;
    let lastText = "";
    for (let c = 1; c <= ws.columnCount; c++) {
      const raw = row.getCell(c).value;
      let t = cellText(raw);
      // A merged cell reports its value on every cell in the range, not just
      // the top-left one - collapse consecutive repeats so a title/heading
      // merged across many columns doesn't look like it repeats N times.
      if (t && t === lastText) t = "";
      else if (t) lastText = t;
      if (t) hasContent = true;
      cells.push(t);
      nums.push(cellNumber(raw));
    }
    if (hasContent) rows.push({ r, cells, nums });
  }
  return { rows };
}

export function rowText(row: GridRow): string {
  return row.cells.filter(Boolean).join(" ");
}

// True for lines that represent a carried/brought balance rather than a real
// transaction - "To Balance c/d", "By Balance b/f" etc. Universal Indian
// bookkeeping notation, not specific to any one firm's books, so this is safe
// to use as a generic exclusion rule in any T-account.
export const CARRY_BALANCE_RE = /\b(b\/d|c\/d|b\/f|c\/f)\b|balance\s*(brought|carried)\s*(down|forward)?/i;

export const GRAND_TOTAL_RE = /^(grand\s*)?total\b|^total\s*\(|^total$/i;
