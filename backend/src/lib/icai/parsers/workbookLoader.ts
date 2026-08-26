import ExcelJS from "exceljs";
import { readGrid, type Grid } from "./grid";
import { classifySheet, type SheetKind } from "./sheetClassifier";

export interface ScannedSheet {
  name: string;
  grid: Grid;
  kind: SheetKind;
  reason: string;
}

export interface WorkbookScan {
  fileName: string;
  entityName: string;
  periodEndLabel: string; // e.g. "31ST MARCH 2026", as found in the source - not reformatted here
  sheets: ScannedSheet[];
}

function mostCommon(values: string[]): string {
  const counts = new Map<string, number>();
  for (const v of values) if (v) counts.set(v, (counts.get(v) ?? 0) + 1);
  let best = "";
  let bestCount = 0;
  for (const [v, c] of counts) {
    if (c > bestCount) {
      best = v;
      bestCount = c;
    }
  }
  return best;
}

function detectPeriodEnd(sheets: ScannedSheet[]): string {
  for (const s of sheets) {
    for (const row of s.grid.rows.slice(0, 10)) {
      const text = row.cells.join(" ");
      const m = text.match(/P\.?\s*Y\.?\s*ENDED\s+(.+)/i) ?? text.match(/BALANCE SHEET AS ON\s+(.+)/i);
      if (m) return m[1].trim();
    }
  }
  return "";
}

export async function loadWorkbook(buffer: Buffer | ArrayBuffer, fileName: string): Promise<WorkbookScan> {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(buffer as ArrayBuffer);

  const sheets: ScannedSheet[] = wb.worksheets.map((ws) => {
    const grid = readGrid(ws);
    const { kind, reason } = classifySheet(grid);
    return { name: ws.name, grid, kind, reason };
  });

  const entityName = mostCommon(sheets.map((s) => s.grid.rows[0]?.cells.find((c) => c) ?? ""));
  const periodEndLabel = detectPeriodEnd(sheets);

  return { fileName, entityName, periodEndLabel, sheets };
}
