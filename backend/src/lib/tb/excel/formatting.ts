import type ExcelJS from "exceljs";
import { INR_FMT } from "./helpers";

// Presentation-only pass over a finished workbook: fonts, fills, borders,
// merges, freeze panes, print setup. Never touches a cell's value/formula,
// never renames a sheet, never inserts/deletes rows or columns. Runs once,
// after every builder has finished writing (see builders/index.ts), so it
// works purely from what's already on the sheet - no per-sheet row-number
// bookkeeping to keep in sync with the generation code.
//
// Every sheet in this app is written through the same handful of helpers
// (setCell/setFormula/topBorder in excel/helpers.ts), which gives a few
// reliable, content-based signals to classify rows without ever hardcoding
// a row number:
//  - a "column header" row has bold text cells, one of them right-aligned,
//    and *no* money-formatted cell (e.g. "PARTICULARS" / "AMOUNT");
//  - a "total" row has at least one bold money cell (the codebase's own
//    convention: every subtotal/grand-total formula is written bold);
//  - a "section" row has bold text in every populated cell and no money
//    cell (e.g. "SCH-1 SUNDRY CREDITORS", "1. Partners' Funds") - this
//    excludes label/value form rows (DETAILS) where the value cell isn't
//    bold;
//  - the banner block (firm name / statement title / period) is the
//    sheet's leading run of single-cell rows - see computeBannerEnd.

const NAVY = "FF1F3864";
const SLATE = "FFD9E2F3";
const GREY = "FFF2F2F2";
const RED = "FFFF0000";
const WHITE = "FFFFFFFF";
const GREEN_FILL = "FFC6EFCE";
const GREEN_TEXT = "FF006100";
const BAD_FILL = "FFFFC7CE";
const BAD_TEXT = "FF9C0006";

const GRAND_TOTAL_KEYWORDS =
  /total sources of funds|total application of funds|total liabilities|total assets|net profit|grand total|profit for the year/i;

const DARK_BORDER: Partial<ExcelJS.Border> = { style: "medium", color: { argb: NAVY } };
const LIGHT_BORDER: Partial<ExcelJS.Border> = { style: "thin", color: { argb: "FFBFBFBF" } };

const AUTOFIT_MIN = 10;
const AUTOFIT_MAX = 45;
const AUTOFIT_PADDING = 2;

function solidFill(argb: string): ExcelJS.Fill {
  return { type: "pattern", pattern: "solid", fgColor: { argb } };
}

function columnLetter(col: number): string {
  let n = col;
  let s = "";
  while (n > 0) {
    const rem = (n - 1) % 26;
    s = String.fromCharCode(65 + rem) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

function isMoneyCell(cell: ExcelJS.Cell): boolean {
  return cell.numFmt === INR_FMT;
}

function cellNumericValue(cell: ExcelJS.Cell): number | null {
  const v = cell.value;
  if (typeof v === "number") return v;
  if (v && typeof v === "object" && "result" in v) {
    const result = (v as { result?: unknown }).result;
    if (typeof result === "number") return result;
  }
  return null;
}

// Indian-style grouped, 2-decimal, parenthesised-negative display width - an
// estimate (not pixel-exact) good enough to size a column so it doesn't clip.
function formattedMoneyLength(n: number): number {
  const intDigits = Math.max(1, Math.floor(Math.abs(n)).toString().length);
  const groups = Math.max(0, Math.ceil((intDigits - 3) / 2));
  return intDigits + groups + 3 + (n < 0 ? 2 : 0);
}

// A merged non-master cell's `.value`/`.style` getters echo the master cell
// (that's how ExcelJS renders merges), so scanning one naively double-counts
// the master's text into every column it spans - this excludes them.
function isMergeSlave(cell: ExcelJS.Cell): boolean {
  return cell.isMerged && cell.master !== cell;
}

function displayLength(cell: ExcelJS.Cell): number {
  if (isMergeSlave(cell)) return 0;
  const v = cell.value;
  if (typeof v === "string") return v.length;
  if (typeof v === "number") return isMoneyCell(cell) ? formattedMoneyLength(v) : String(v).length;
  if (v && typeof v === "object" && "result" in v) {
    const result = (v as { result?: unknown }).result;
    if (typeof result === "number")
      return isMoneyCell(cell) ? formattedMoneyLength(result) : String(result).length;
    if (typeof result === "string") return result.length;
  }
  return 0;
}

interface RowScan {
  cells: { col: number; cell: ExcelJS.Cell }[];
  minCol: number;
  maxCol: number;
  hasMoney: boolean;
  hasBoldMoney: boolean;
  boldTextCols: number[];
  boldRightTextCols: number[];
  text: string;
}

function scanRow(row: ExcelJS.Row): RowScan | null {
  const cells: { col: number; cell: ExcelJS.Cell }[] = [];
  row.eachCell({ includeEmpty: false }, (cell, col) => {
    cells.push({ col, cell });
  });
  if (!cells.length) return null;

  let hasMoney = false;
  let hasBoldMoney = false;
  const boldTextCols: number[] = [];
  const boldRightTextCols: number[] = [];
  const textParts: string[] = [];

  for (const { col, cell } of cells) {
    const money = isMoneyCell(cell);
    if (money) hasMoney = true;
    if (money && cell.font?.bold) hasBoldMoney = true;
    if (typeof cell.value === "string") {
      textParts.push(cell.value);
      if (cell.font?.bold) {
        boldTextCols.push(col);
        if (cell.alignment?.horizontal === "right") boldRightTextCols.push(col);
      }
    }
  }

  const cols = cells.map((c) => c.col);
  return {
    cells,
    minCol: Math.min(...cols),
    maxCol: Math.max(...cols),
    hasMoney,
    hasBoldMoney,
    boldTextCols,
    boldRightTextCols,
    text: textParts.join(" ").toLowerCase(),
  };
}

type RowKind = "columnHeader" | "total" | "section" | "normal";

function classify(scan: RowScan): RowKind {
  if (
    !scan.hasMoney &&
    scan.boldRightTextCols.length > 0 &&
    scan.boldTextCols.length > scan.boldRightTextCols.length
  ) {
    return "columnHeader";
  }
  if (scan.hasBoldMoney) return "total";
  if (
    !scan.hasMoney &&
    scan.boldTextCols.length > 0 &&
    scan.boldTextCols.length === scan.cells.length
  ) {
    return "section";
  }
  return "normal";
}

// The banner block (firm name / CIN / statement title / period / unit
// heading) is always a leading run of single-cell rows, immediately at the
// top of the sheet. It's detected by position, not by content, so it can't
// be confused with a later single-cell section title (e.g. each partner's
// account heading in 'Cap & Current', or "Liabilities :" in the vertical BS)
// - those are never part of the sheet's leading run because either a blank
// spacer row or the row-count cap below separates them from it.
function computeBannerEnd(ws: ExcelJS.Worksheet): number {
  const maxRow = Math.min(20, ws.rowCount);
  let end = 0;
  let count = 0;
  let started = false;
  for (let r = 1; r <= maxRow; r++) {
    const scan = scanRow(ws.getRow(r));
    if (!scan) {
      if (started) break; // a blank row after content began closes the banner
      continue; // leading blank rows (e.g. row 1) don't count against it
    }
    if (scan.cells.length !== 1) break; // real tabular content
    started = true;
    end = r;
    if (++count >= 6) break; // real banners in this app are never longer than this
  }
  return end;
}

function findSignatureRow(ws: ExcelJS.Worksheet): number {
  const maxRow = ws.rowCount;
  for (let r = 1; r <= maxRow; r++) {
    let found = false;
    ws.getRow(r).eachCell({ includeEmpty: false }, (cell) => {
      if (typeof cell.value === "string" && /^place\s*:/i.test(cell.value.trim())) found = true;
    });
    if (found) return r;
  }
  return Infinity;
}

function findHeaderRow(ws: ExcelJS.Worksheet): number | null {
  const maxRow = Math.min(30, ws.rowCount);
  for (let r = 1; r <= maxRow; r++) {
    const scan = scanRow(ws.getRow(r));
    if (scan && classify(scan) === "columnHeader") return r;
  }
  return null;
}

// The width of the sheet's actual statement table - deliberately bounded to
// rows above the signature block, so a stray footer cell (e.g. "Director"/
// "DIN :" a few columns further right than the table itself) doesn't widen
// the header band / table box past where the real table ends.
function computeTableWidth(ws: ExcelJS.Worksheet, signatureRow: number): number {
  const limit = Math.min(signatureRow - 1, ws.rowCount);
  let max = 2;
  for (let r = 1; r <= limit; r++) {
    const scan = scanRow(ws.getRow(r));
    if (scan) max = Math.max(max, scan.maxCol);
  }
  return max;
}

function styleBanner(
  ws: ExcelJS.Worksheet,
  r: number,
  scan: RowScan,
  tableWidth: number,
  isFirstContentRow: boolean,
) {
  const from = scan.minCol;
  const to = Math.max(from, tableWidth);
  if (to > from) {
    try {
      ws.mergeCells(r, from, r, to);
    } catch {
      // already merged - safe to ignore
    }
  }
  const master = ws.getCell(r, from);
  const bold = !!master.font?.bold;
  master.font = {
    ...master.font,
    size: isFirstContentRow ? 14 : bold ? 12 : (master.font?.size ?? 10),
  };
  master.alignment = { ...master.alignment, horizontal: "center", vertical: "middle" };
}

function styleColumnHeader(ws: ExcelJS.Worksheet, r: number, tableWidth: number) {
  for (let c = 1; c <= tableWidth; c++) {
    const cell = ws.getCell(r, c);
    cell.fill = solidFill(NAVY);
    cell.font = { ...cell.font, bold: true, color: { argb: WHITE } };
    cell.alignment = { ...cell.alignment, horizontal: "center", vertical: "middle" };
    cell.border = { ...cell.border, bottom: { style: "medium", color: { argb: NAVY } } };
  }
}

function styleSectionRow(ws: ExcelJS.Worksheet, r: number, scan: RowScan) {
  for (let c = scan.minCol; c <= scan.maxCol; c++) {
    const cell = ws.getCell(r, c);
    cell.fill = solidFill(SLATE);
    cell.font = { ...cell.font, bold: true, color: { argb: NAVY } };
  }
}

function styleTotalRow(ws: ExcelJS.Worksheet, r: number, scan: RowScan) {
  const grand = GRAND_TOTAL_KEYWORDS.test(scan.text);
  for (let c = scan.minCol; c <= scan.maxCol; c++) {
    const cell = ws.getCell(r, c);
    cell.fill = solidFill(GREY);
    cell.font = { ...cell.font, bold: true };
    if (grand) {
      cell.border = { ...cell.border, top: { style: "double" } };
    }
  }
}

function drawOuterBorder(ws: ExcelJS.Worksheet, r1: number, c1: number, r2: number, c2: number) {
  const setSide = (r: number, c: number, side: "top" | "bottom" | "left" | "right") => {
    const cell = ws.getCell(r, c);
    cell.border = { ...cell.border, [side]: DARK_BORDER };
  };
  for (let c = c1; c <= c2; c++) {
    setSide(r1, c, "top");
    setSide(r2, c, "bottom");
  }
  for (let r = r1; r <= r2; r++) {
    setSide(r, c1, "left");
    setSide(r, c2, "right");
  }
}

// Fills in a light-grey thin border on any side that doesn't already have a
// style - run after drawOuterBorder, so the perimeter's dark sides are
// already in place and simply win (this only ever adds a side that was
// still blank, never replaces one).
function addInteriorBorders(ws: ExcelJS.Worksheet, r1: number, c1: number, r2: number, c2: number) {
  for (let r = r1; r <= r2; r++) {
    for (let c = c1; c <= c2; c++) {
      const cell = ws.getCell(r, c);
      const existing = cell.border ?? {};
      cell.border = {
        left: existing.left?.style ? existing.left : LIGHT_BORDER,
        right: existing.right?.style ? existing.right : LIGHT_BORDER,
        top: existing.top?.style ? existing.top : LIGHT_BORDER,
        bottom: existing.bottom?.style ? existing.bottom : LIGHT_BORDER,
      };
    }
  }
}

// A "table" is a column-header row through the last total row seen before
// either the next column-header row or the signature block - see the
// worked examples in the formatting addendum this implements (each
// partner's account block in 'Cap & Current' closes at its own "TOTAL",
// not at the first bold-money row like "To Balance c/d" that appears
// partway through it). Each span gets the dark outer perimeter plus a
// light-grey fill-in on every still-blank interior side, so line items
// read as a ruled table instead of a hollow box.
function applyTableBoxes(ws: ExcelJS.Worksheet, tableWidth: number, signatureRow: number) {
  const limit = Math.min(signatureRow - 1, ws.rowCount);
  let openHeaderRow: number | null = null;
  let lastTotalRow: number | null = null;

  const closeSpan = () => {
    if (openHeaderRow !== null && lastTotalRow !== null && lastTotalRow > openHeaderRow) {
      drawOuterBorder(ws, openHeaderRow, 1, lastTotalRow, tableWidth);
      addInteriorBorders(ws, openHeaderRow, 1, lastTotalRow, tableWidth);
    }
    openHeaderRow = null;
    lastTotalRow = null;
  };

  for (let r = 1; r <= limit; r++) {
    const scan = scanRow(ws.getRow(r));
    if (!scan) continue;
    const kind = classify(scan);
    if (kind === "columnHeader") {
      closeSpan();
      openHeaderRow = r;
    } else if (kind === "total" && openHeaderRow !== null) {
      lastTotalRow = r;
    }
  }
  closeSpan();
}

function applyBalancingCheckCF(ws: ExcelJS.Worksheet, r: number, scan: RowScan) {
  const moneyCells = scan.cells.filter(({ cell }) => isMoneyCell(cell));
  if (!moneyCells.length) return;
  const target = moneyCells[moneyCells.length - 1];
  const ref = `${columnLetter(target.col)}${r}`;
  ws.addConditionalFormatting({
    ref,
    rules: [
      {
        type: "cellIs",
        operator: "between",
        formulae: [-0.5, 0.5],
        priority: 1,
        style: { fill: solidFill(GREEN_FILL), font: { color: { argb: GREEN_TEXT } } },
      },
      {
        type: "cellIs",
        operator: "greaterThan",
        formulae: [0.5],
        priority: 2,
        style: { fill: solidFill(BAD_FILL), font: { color: { argb: BAD_TEXT } } },
      },
      {
        type: "cellIs",
        operator: "lessThan",
        formulae: [-0.5],
        priority: 3,
        style: { fill: solidFill(BAD_FILL), font: { color: { argb: BAD_TEXT } } },
      },
    ],
  });
}

function applyNegativeRed(ws: ExcelJS.Worksheet) {
  const maxRow = ws.rowCount;
  for (let r = 1; r <= maxRow; r++) {
    ws.getRow(r).eachCell({ includeEmpty: false }, (cell) => {
      if (!isMoneyCell(cell)) return;
      const num = cellNumericValue(cell);
      if (num !== null && num < 0) {
        cell.font = { ...cell.font, color: { argb: RED } };
      }
    });
  }
}

// Sizes every column to its widest cell (capped), then wraps + heightens any
// row that still overflows the cap (e.g. the Form 8 solvency paragraph,
// a long related-party note) instead of letting it clip.
function autofitColumns(ws: ExcelJS.Worksheet) {
  const maxLen = new Map<number, number>();
  for (let r = 1; r <= ws.rowCount; r++) {
    ws.getRow(r).eachCell({ includeEmpty: false }, (cell, col) => {
      maxLen.set(col, Math.max(maxLen.get(col) ?? 0, displayLength(cell)));
    });
  }
  maxLen.forEach((len, col) => {
    ws.getColumn(col).width = Math.min(Math.max(len + AUTOFIT_PADDING, AUTOFIT_MIN), AUTOFIT_MAX);
  });

  const usableWidth = AUTOFIT_MAX - AUTOFIT_PADDING;
  for (let r = 1; r <= ws.rowCount; r++) {
    const row = ws.getRow(r);
    let maxLines = 1;
    let overflow = false;
    row.eachCell({ includeEmpty: false }, (cell) => {
      const len = displayLength(cell);
      if (len > usableWidth) {
        cell.alignment = {
          ...cell.alignment,
          wrapText: true,
          vertical: cell.alignment?.vertical ?? "top",
        };
        overflow = true;
        maxLines = Math.max(maxLines, Math.ceil(len / usableWidth));
      }
    });
    if (overflow) row.height = 15 * maxLines;
  }
}

// Page Break Preview greys out everything outside the print area
// automatically once it's set - the tight bound is what makes that grey-out
// line up with the sheet's real content instead of the default full sheet.
function applyPrintArea(ws: ExcelJS.Worksheet) {
  const lastRow = ws.rowCount;
  if (lastRow < 1) return;
  const lastCol = Math.max(ws.columnCount, ws.actualColumnCount, 2);
  ws.pageSetup.printArea = `A1:${columnLetter(lastCol)}${lastRow}`;
}

function formatSheet(ws: ExcelJS.Worksheet): void {
  const lastRow = ws.rowCount;
  const signatureRow = findSignatureRow(ws);
  const tableWidth = computeTableWidth(ws, signatureRow);
  const headerRow = findHeaderRow(ws);
  const bannerEnd = computeBannerEnd(ws);

  // Runs before the styling loop below deliberately: it re-derives the same
  // "columnHeader"/"total" classification from live cell state, and the
  // styling loop re-centers header cells (overwriting the right-alignment
  // that classification depends on) - scanning first keeps this reading
  // the original, untouched layout.
  applyTableBoxes(ws, tableWidth, signatureRow);

  let firstContentRow: number | null = null;

  for (let r = 1; r <= lastRow; r++) {
    const scan = scanRow(ws.getRow(r));
    if (!scan) continue;
    if (firstContentRow === null) firstContentRow = r;

    if (scan.text.includes("balancing check")) applyBalancingCheckCF(ws, r, scan);

    if (r >= signatureRow) continue; // leave the CA sign-off block untouched

    if (r <= bannerEnd) {
      styleBanner(ws, r, scan, tableWidth, r === firstContentRow);
      continue;
    }
    const kind = classify(scan);
    if (kind === "columnHeader") styleColumnHeader(ws, r, tableWidth);
    else if (kind === "total") styleTotalRow(ws, r, scan);
    else if (kind === "section") styleSectionRow(ws, r, scan);
  }

  applyNegativeRed(ws);
  autofitColumns(ws);

  ws.pageSetup.orientation = "landscape";
  ws.pageSetup.fitToPage = true;
  ws.pageSetup.fitToWidth = 1;
  ws.pageSetup.fitToHeight = 0;
  ws.pageSetup.margins = { left: 0.4, right: 0.4, top: 0.5, bottom: 0.5, header: 0.2, footer: 0.2 };
  if (headerRow) ws.pageSetup.printTitlesRow = `1:${headerRow}`;
  applyPrintArea(ws);

  ws.views = headerRow
    ? [{ state: "frozen", xSplit: 0, ySplit: headerRow, showGridLines: false }]
    : [{ showGridLines: false }];
}

export function applyWorkbookFormatting(wb: ExcelJS.Workbook): void {
  wb.worksheets.forEach((ws) => formatSheet(ws));
}
