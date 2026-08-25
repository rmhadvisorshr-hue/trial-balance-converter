import type ExcelJS from "exceljs";
import type { StatementMeta } from "../types";

export const INR_FMT = "#,##,##0.00;(#,##,##0.00)";

export function setCell(
  ws: ExcelJS.Worksheet,
  row: number,
  col: number,
  value: ExcelJS.CellValue,
  opts: {
    bold?: boolean;
    italic?: boolean;
    align?: "left" | "center" | "right";
    money?: boolean;
    size?: number;
    underline?: boolean;
    wrap?: boolean;
  } = {},
) {
  const cell = ws.getCell(row, col);
  cell.value = value;
  cell.font = {
    bold: opts.bold,
    italic: opts.italic,
    size: opts.size ?? 10,
    underline: opts.underline,
  };
  cell.alignment = {
    horizontal: opts.align ?? (opts.money ? "right" : "left"),
    vertical: "middle",
    wrapText: opts.wrap,
  };
  if (opts.money) cell.numFmt = INR_FMT;
  return cell;
}

// Writes a live Excel formula with a cached result. ExcelJS never evaluates
// formulas itself, so without a cached result Excel/LibreOffice shows a
// blank cell until the user forces a recalculation - `result` must be kept
// in sync with what the formula actually computes to.
export function setFormula(
  ws: ExcelJS.Worksheet,
  row: number,
  col: number,
  formula: string,
  result: number | string,
  opts: {
    bold?: boolean;
    italic?: boolean;
    align?: "left" | "center" | "right";
    money?: boolean;
  } = {},
) {
  const cell = ws.getCell(row, col);
  cell.value = { formula, result } as ExcelJS.CellFormulaValue;
  cell.font = { bold: opts.bold, italic: opts.italic, size: 10 };
  cell.alignment = { horizontal: opts.align ?? "right", vertical: "middle" };
  if (opts.money) cell.numFmt = INR_FMT;
  return cell;
}

export function mergeTitle(
  ws: ExcelJS.Worksheet,
  row: number,
  fromCol: number,
  toCol: number,
  value: string,
  opts: { bold?: boolean; size?: number; italic?: boolean } = {},
) {
  ws.mergeCells(row, fromCol, row, toCol);
  const cell = ws.getCell(row, fromCol);
  cell.value = value;
  cell.font = { bold: opts.bold ?? true, size: opts.size ?? 11, italic: opts.italic };
  cell.alignment = { horizontal: "center", vertical: "middle" };
  return cell;
}

export function topBorder(ws: ExcelJS.Worksheet, row: number, fromCol: number, toCol: number) {
  for (let c = fromCol; c <= toCol; c++) {
    const cell = ws.getCell(row, c);
    cell.border = { ...cell.border, top: { style: "thin" } };
  }
}

export function boxBorder(ws: ExcelJS.Worksheet, r1: number, c1: number, r2: number, c2: number) {
  for (let r = r1; r <= r2; r++) {
    for (let c = c1; c <= c2; c++) {
      const cell = ws.getCell(r, c);
      cell.border = {
        top: r === r1 ? { style: "thin" } : cell.border?.top,
        bottom: r === r2 ? { style: "thin" } : cell.border?.bottom,
        left: c === c1 ? { style: "thin" } : cell.border?.left,
        right: c === c2 ? { style: "thin" } : cell.border?.right,
      };
    }
  }
}

// The CA signature / UDIN footer block used in the (V) sheets.
export function signatureBlock(
  ws: ExcelJS.Worksheet,
  startRow: number,
  meta: StatementMeta,
  forFirmText: string,
  designation: string,
): number {
  let r = startRow + 2;
  setCell(ws, r, 1, `Place : ${meta.place ?? "Vasai"}`);
  setCell(ws, r, 3, forFirmText, { bold: true });
  r += 2;
  setCell(ws, r, 1, `Date : ${meta.date ?? ""}`);
  r += 2;
  setCell(ws, r, 1, `UDIN : ${meta.udin ?? ""}`);
  r += 2;
  setCell(ws, r, 3, designation);
  r += 2;
  setCell(ws, r, 3, "As per our report on even date");
  r += 1;
  setCell(ws, r, 3, `For ${meta.caName ?? "Namrata Prakash Sharma"}`);
  r += 1;
  setCell(ws, r, 3, "(Chartered Accountants)");
  r += 3;
  setCell(ws, r, 3, `Proprietor : CA ${meta.caName ?? "Namrata Prakash Sharma"}`);
  r += 1;
  setCell(ws, r, 3, `M No : ${meta.caMembershipNo ?? "177309"}`);
  r += 1;
  setCell(ws, r, 3, `FRN No : ${meta.caFirmRegNo ?? "144860W"}`);
  return r;
}

export function applyColumnWidths(ws: ExcelJS.Worksheet, widths: number[]) {
  widths.forEach((w, i) => {
    ws.getColumn(i + 1).width = w;
  });
}
