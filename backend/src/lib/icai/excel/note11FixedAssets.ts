import type ExcelJS from "exceljs";
import type { WorkbookAnalysis, FixedAssetLine } from "../types";
import { setCell, setFormula, topBorder, applyColumnWidths } from "../../tb/excel/helpers";
import { round2 } from "./common";

// The source records fixed assets on a WDV/tax-block basis (Opening +
// Additions - Depreciation = Closing), not the Schedule III-style gross-
// block/accumulated-depreciation roll-forward the ICAI template illustrates
// - so this mirrors the simpler presentation the records actually support
// rather than fabricating a gross-cost split that isn't in the data.
function writeMovementTable(ws: ExcelJS.Worksheet, startRow: number, yearLabel: string, rows: FixedAssetLine[]): number {
  let r = startRow;
  setCell(ws, r++, 1, `Movement for the year ended ${yearLabel}`, { bold: true });
  const headers = ["Asset", "Rate", "Opening WDV", "Additions during the year", "Depreciation for the year", "Closing WDV"];
  headers.forEach((h, i) => setCell(ws, r, i + 1, h, { bold: true, align: i === 0 ? "left" : "right", wrap: true }));
  r++;

  const byCategory = new Map<string, FixedAssetLine[]>();
  for (const row of rows) {
    const list = byCategory.get(row.category) ?? [];
    list.push(row);
    byCategory.set(row.category, list);
  }

  let grandOpening = 0;
  let grandAdditions = 0;
  let grandDep = 0;
  let grandClosing = 0;

  for (const [category, assets] of byCategory) {
    setCell(ws, r++, 1, category, { bold: true });
    const catStart = r;
    for (const a of assets) {
      setCell(ws, r, 1, a.name);
      setCell(ws, r, 2, a.rate, { align: "right" });
      setCell(ws, r, 3, round2(a.opening), { money: true });
      setCell(ws, r, 4, round2(a.additions), { money: true });
      setCell(ws, r, 5, round2(a.depreciation), { money: true });
      setCell(ws, r++, 6, round2(a.closing), { money: true });
    }
    const catEnd = r - 1;
    const catTotal = (key: keyof FixedAssetLine) => round2(assets.reduce((s, a) => s + (a[key] as number), 0));
    setCell(ws, r, 1, `Total - ${category}`, { bold: true });
    setFormula(ws, r, 3, `SUM(C${catStart}:C${catEnd})`, catTotal("opening"), { bold: true, money: true });
    setFormula(ws, r, 4, `SUM(D${catStart}:D${catEnd})`, catTotal("additions"), { bold: true, money: true });
    setFormula(ws, r, 5, `SUM(E${catStart}:E${catEnd})`, catTotal("depreciation"), { bold: true, money: true });
    setFormula(ws, r, 6, `SUM(F${catStart}:F${catEnd})`, catTotal("closing"), { bold: true, money: true });
    r += 2;
    grandOpening += catTotal("opening");
    grandAdditions += catTotal("additions");
    grandDep += catTotal("depreciation");
    grandClosing += catTotal("closing");
  }

  if (rows.length === 0) {
    setCell(ws, r++, 1, "NIL");
    return r + 1;
  }

  setCell(ws, r, 1, "Total Property, Plant and Equipment", { bold: true });
  setCell(ws, r, 3, round2(grandOpening), { money: true, bold: true });
  setCell(ws, r, 4, round2(grandAdditions), { money: true, bold: true });
  setCell(ws, r, 5, round2(grandDep), { money: true, bold: true });
  setCell(ws, r, 6, round2(grandClosing), { money: true, bold: true });
  topBorder(ws, r, 1, 6);
  r += 2;
  return r;
}

export function writeNote11FixedAssets(ws: ExcelJS.Worksheet, analysis: WorkbookAnalysis) {
  applyColumnWidths(ws, [34, 8, 18, 20, 18, 16]);
  let r = 1;
  setCell(ws, r++, 1, analysis.entityName || "ENTITY NAME", { bold: true });
  setCell(ws, r++, 1, `Notes forming part of the Financial Statements for the year ended ${analysis.currentYearLabel}`, { bold: true });
  r++;
  setCell(ws, r++, 1, "Note 11: Property, Plant and Equipment (owned assets)", { bold: true });
  r++;

  r = writeMovementTable(ws, r, analysis.currentYearLabel, analysis.fixedAssets.current);
  if (analysis.previousYearLabel) {
    writeMovementTable(ws, r, analysis.previousYearLabel, analysis.fixedAssets.previous);
  }
}
