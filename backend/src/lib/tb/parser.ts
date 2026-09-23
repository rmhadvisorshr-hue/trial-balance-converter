import ExcelJS from "exceljs";
import type { TBRow, StatementMeta } from "./types";
import { loadExcelWorkbook } from "../excelCompat";

export interface RawTrialBalance {
  meta: StatementMeta;
  leaves: TBRow[]; // leaf ledgers only (groups excluded)
  groups: TBRow[]; // group rows (for diagnostics)
  warnings: string[];
}

const SKIP_RE =
  /^(grand\s*total|carried\s*over|brought\s*forward|c\s*a\s*r\s*r\s*i\s*e\s*d|b\s*r\s*o\s*u\s*g\s*h\s*t|continued|page\b|trial\s*balance|particulars|closing\s*balance|opening\s*balance|--\s*\d)/i;

function toNumber(value: unknown): number {
  if (value == null) return 0;
  if (typeof value === "number") return Number.isFinite(value) ? value : 0;
  if (typeof value === "object") {
    const v = value as { result?: unknown; value?: unknown };
    if (v.result != null) return toNumber(v.result);
    if (v.value != null) return toNumber(v.value);
    return 0;
  }
  const s = String(value)
    .replace(/[,\s₹]/g, "")
    .replace(/[()]/g, (m) => (m === "(" ? "-" : ""));
  const n = Number(s);
  return Number.isFinite(n) ? n : 0;
}

function cellText(value: unknown): string {
  if (value == null) return "";
  if (typeof value === "string") return value.trim();
  if (typeof value === "number") return String(value);
  if (typeof value === "object") {
    const v = value as { result?: unknown; text?: unknown; richText?: Array<{ text: string }> };
    if (Array.isArray(v.richText))
      return v.richText
        .map((r) => r.text)
        .join("")
        .trim();
    if (v.text != null) return String(v.text).trim();
    if (v.result != null) return String(v.result).trim();
  }
  return String(value).trim();
}

export async function parseTrialBalanceXlsx(
  buffer: ArrayBuffer | Buffer,
): Promise<RawTrialBalance> {
  const wb = await loadExcelWorkbook(buffer);
  const ws = wb.worksheets[0];
  const warnings: string[] = [];

  if (!ws) {
    return {
      meta: emptyMeta(),
      leaves: [],
      groups: [],
      warnings: ["No worksheet found in the file."],
    };
  }

  // Locate the header row containing Debit / Credit and the amount columns.
  let headerRow = 0;
  let debitCol = 0;
  let creditCol = 0;
  let nameCol = 1;

  for (let r = 1; r <= Math.min(ws.rowCount, 30); r++) {
    const row = ws.getRow(r);
    let foundDebit = 0;
    let foundCredit = 0;
    for (let c = 1; c <= ws.columnCount; c++) {
      const t = cellText(row.getCell(c).value).toLowerCase();
      if (t === "debit") foundDebit = c;
      if (t === "credit") foundCredit = c;
    }
    if (foundDebit && foundCredit) {
      headerRow = r;
      debitCol = foundDebit;
      creditCol = foundCredit;
      break;
    }
  }

  if (!headerRow) {
    // Fallback: assume columns 2 (debit) & 3 (credit), data from row 1.
    headerRow = 0;
    debitCol = 2;
    creditCol = 3;
    warnings.push("Could not find Debit/Credit headers; assumed columns B and C.");
  }

  const meta = extractMeta(ws, headerRow || 6);

  // Determine the name column: first column whose header row cell is empty but data has text.
  // Tally puts ledger names in column A.
  nameCol = 1;

  const leaves: TBRow[] = [];
  const groups: TBRow[] = [];

  const startRow = (headerRow || 0) + 1;
  // Precompute indent for each candidate row.
  type Parsed = { r: number; name: string; indent: number; debit: number; credit: number };
  const parsed: Parsed[] = [];
  for (let r = startRow; r <= ws.rowCount; r++) {
    const row = ws.getRow(r);
    const nameCell = row.getCell(nameCol);
    const name = cellText(nameCell.value);
    if (!name) continue;
    if (SKIP_RE.test(name)) continue;
    const indent = Number(nameCell.alignment?.indent ?? 0) || 0;
    const debit = toNumber(row.getCell(debitCol).value);
    const credit = toNumber(row.getCell(creditCol).value);
    parsed.push({ r, name, indent, debit, credit });
  }

  // A row is a GROUP only if the contiguous block of following rows with a
  // GREATER indent actually nets to this row's balance. This is robust against
  // Tally exports that use slightly inconsistent indent values for sibling
  // ledgers (e.g. some at indent 3, some at indent 4), which would otherwise
  // cause genuine leaf ledgers to be mistaken for group headers.
  const isGroupRow = (i: number): boolean => {
    const cur = parsed[i];
    let childSigned = 0;
    let childCount = 0;
    for (let j = i + 1; j < parsed.length; j++) {
      if (parsed[j].indent <= cur.indent) break;
      childSigned += parsed[j].debit - parsed[j].credit;
      childCount++;
    }
    if (childCount === 0) return false;
    const curSigned = cur.debit - cur.credit;
    return Math.abs(childSigned - curSigned) <= 1;
  };

  const ancestors: { name: string; indent: number }[] = [];
  for (let i = 0; i < parsed.length; i++) {
    const cur = parsed[i];
    const isGroup = isGroupRow(i);

    // Pop ancestors at same-or-deeper indent.
    while (ancestors.length && ancestors[ancestors.length - 1].indent >= cur.indent) {
      ancestors.pop();
    }
    const topGroup = ancestors.length ? ancestors[0].name : cur.name;

    const rowObj: TBRow = {
      name: cur.name,
      parentGroup: ancestors.length ? topGroup : "",
      isGroup,
      debit: cur.debit,
      credit: cur.credit,
      level: cur.indent,
    };

    if (isGroup) {
      groups.push(rowObj);
      ancestors.push({ name: cur.name, indent: cur.indent });
    } else {
      // leaf ledger
      rowObj.parentGroup = ancestors.length ? ancestors[0].name : cur.name;
      leaves.push(rowObj);
    }
  }

  if (leaves.length === 0) {
    warnings.push(
      "No ledger rows were detected. The file may not be a standard Tally trial balance.",
    );
  }

  return { meta, leaves, groups, warnings };
}

function extractMeta(ws: ExcelJS.Worksheet, headerRow: number): StatementMeta {
  const meta = emptyMeta();
  const lines: string[] = [];
  for (let r = 1; r < Math.max(headerRow, 2); r++) {
    const row = ws.getRow(r);
    for (let c = 1; c <= ws.columnCount; c++) {
      const t = cellText(row.getCell(c).value);
      if (t) {
        lines.push(t);
        break;
      }
    }
  }
  if (lines[0]) meta.firmName = cleanFirmName(lines[0]);
  const period = lines.find((l) => /\d.*\bto\b.*\d/i.test(l));
  if (period) {
    meta.periodLabel = period;
    const endMatch = period.match(/to\s*(.+)$/i);
    if (endMatch) meta.asAtLabel = `as at ${endMatch[1].trim()}`;
  }
  return meta;
}

function cleanFirmName(raw: string): string {
  // "EDUGUIDE OVERSEAS STUDIES PVT LTD - 25-26 - (from 1-Apr-25)" -> strip year/period suffix.
  return raw
    .replace(/\s*-\s*\d{2}-\d{2}.*/i, "")
    .replace(/\s*\(from.*\)\s*$/i, "")
    .trim();
}

function emptyMeta(): StatementMeta {
  return { firmName: "", periodLabel: "", asAtLabel: "", figuresUnit: "actual" };
}

// -------- PDF (secondary path) --------

export async function parseTrialBalancePdf(buffer: ArrayBuffer | Buffer): Promise<RawTrialBalance> {
  const warnings: string[] = [];
  // Lazy import to keep cold path light.
  const pdfjs = (await import("pdfjs-dist/legacy/build/pdf.mjs")) as typeof import("pdfjs-dist");
  const data = buffer instanceof ArrayBuffer ? new Uint8Array(buffer) : new Uint8Array(buffer);
  const doc = await pdfjs.getDocument({ data, useSystemFonts: true }).promise;

  const meta = emptyMeta();
  type Item = { str: string; x: number; y: number; w: number };
  const allLines: { y: number; items: Item[] }[] = [];
  let pageWidth = 0;

  for (let p = 1; p <= doc.numPages; p++) {
    const page = await doc.getPage(p);
    const viewport = page.getViewport({ scale: 1 });
    pageWidth = Math.max(pageWidth, viewport.width);
    const content = await page.getTextContent();
    const rows = new Map<number, Item[]>();
    for (const it of content.items as Array<{ str: string; transform: number[]; width: number }>) {
      if (!it.str.trim()) continue;
      const x = it.transform[4];
      const y = Math.round(it.transform[5]);
      const key = y;
      if (!rows.has(key)) rows.set(key, []);
      rows.get(key)!.push({ str: it.str, x, y, w: it.width });
    }
    for (const [y, items] of rows) {
      allLines.push({ y, items: items.sort((a, b) => a.x - b.x) });
    }
  }

  const isNumberToken = (str: string) =>
    /^[()\d,.-]+$/.test(str.replace(/\s/g, "")) && /\d/.test(str);

  // Tally right-aligns amounts within each column, so a token's *left* x
  // shifts with its digit count and can't reliably separate Debit from
  // Credit (a short Debit value and a long Credit value can start at the
  // same x). The right edge (x + width) stays put per column regardless of
  // digit count, so the two columns show up as two tight clusters of right
  // edges with a wide gap between them; splitting at that gap is layout-
  // agnostic, unlike a fixed fraction of the page width.
  function findColumnSplit(): number | null {
    const rightEdges: number[] = [];
    for (const line of allLines) {
      for (const it of line.items) {
        if (isNumberToken(it.str)) rightEdges.push(it.x + it.w);
      }
    }
    if (rightEdges.length < 2) return null;
    rightEdges.sort((a, b) => a - b);
    let bestGap = -1;
    let bestIdx = -1;
    for (let i = 1; i < rightEdges.length; i++) {
      const gap = rightEdges[i] - rightEdges[i - 1];
      if (gap > bestGap) {
        bestGap = gap;
        bestIdx = i;
      }
    }
    // A real two-column split has a wide gap between the clusters; a small
    // largest-gap means the numbers likely form one column, not two.
    if (bestGap < 15) return null;
    return (rightEdges[bestIdx - 1] + rightEdges[bestIdx]) / 2;
  }

  const columnSplit = findColumnSplit();

  // Reconstruct meta from the first few lines.
  const textLines = allLines.map((l) =>
    l.items
      .map((i) => i.str)
      .join(" ")
      .trim(),
  );
  if (textLines[0]) meta.firmName = cleanFirmName(textLines[0]);
  const period = textLines.find((l) => /\d.*\bto\b.*\d/i.test(l));
  if (period) {
    meta.periodLabel = period.replace(/page\s*\d+/i, "").trim();
    const endMatch = meta.periodLabel.match(/to\s*([0-9A-Za-z-]+)/i);
    if (endMatch) meta.asAtLabel = `as at ${endMatch[1].trim()}`;
  }

  // Fallback for the rare case findColumnSplit() couldn't find a clean two-
  // column gap (e.g. a single-column list): assume amounts left of ~72% of
  // the page width are Credit, matching Tally's Debit-then-Credit layout.
  const fallbackMid = pageWidth * 0.72;
  const leaves: TBRow[] = [];

  for (const line of allLines) {
    const joined = line.items
      .map((i) => i.str)
      .join(" ")
      .trim();
    if (!joined || SKIP_RE.test(joined)) continue;

    // Split into name part and numeric tokens (with their x positions).
    const numberItems = line.items.filter((i) => isNumberToken(i.str));
    if (numberItems.length === 0) continue;
    const nameItems = line.items.filter((i) => !numberItems.includes(i));
    const name = nameItems
      .map((i) => i.str)
      .join(" ")
      .trim();
    if (!name || /^[\d.,()\s-]+$/.test(name)) continue;

    let debit = 0;
    let credit = 0;
    for (const ni of numberItems) {
      const val = toNumber(ni.str);
      // Debit is the left column, Credit the right (Tally's standard order,
      // matching the "Debit  Credit" header on the trial balance).
      const isDebit =
        columnSplit !== null ? ni.x + ni.w <= columnSplit : ni.x >= fallbackMid;
      if (isDebit) debit += val;
      else credit += val;
    }
    leaves.push({ name, parentGroup: "", isGroup: false, debit, credit, level: 1 });
  }

  warnings.push(
    "PDF parsing is best-effort: ledger grouping is not available from PDF, so review and correct the mapping carefully.",
  );

  return { meta, leaves, groups: [], warnings };
}

export async function parseTrialBalance(
  buffer: ArrayBuffer | Buffer,
  fileName: string,
): Promise<RawTrialBalance> {
  if (/\.pdf$/i.test(fileName)) {
    return parseTrialBalancePdf(buffer);
  }
  return parseTrialBalanceXlsx(buffer);
}
