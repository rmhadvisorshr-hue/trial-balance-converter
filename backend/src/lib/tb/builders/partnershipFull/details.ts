import type ExcelJS from "exceljs";
import type { BuildCtx } from "../common";
import { setCell, applyColumnWidths } from "../../excel/helpers";

// 'DETAILS': client master data (PAN, address, bank details, TAN/PTEC/PTRC,
// IT/TDS file numbers). None of this is in a trial balance - it's manually
// maintained CA-firm client records, same category as CIN for Pvt Ltd.
// Firm name is pre-filled from existing meta; everything else is a labeled
// blank for the CA to fill in once.
export function writeDetailsSheet(ws: ExcelJS.Worksheet, ctx: BuildCtx) {
  applyColumnWidths(ws, [22, 40]);
  const { meta } = ctx;

  let r = 2;
  setCell(ws, r++, 2, "DETAILS OF CLIENT", { bold: true });
  r++;

  const field = (label: string, value = "") => {
    setCell(ws, r, 1, label, { bold: true });
    setCell(ws, r, 2, value);
    r += 2;
  };

  field("NAME", meta.firmName || "");
  field("BUSINESS NAME", meta.firmName || "");
  field("NATURE OF BUSINESS");
  field("ADDRESS", meta.place || "");
  field("PAN");
  field("WARD");
  field("BANK DETAILS");
  field("VAT NUMBER");
  field("SERVICE TAX / GST NUMBER");
  field("TAN NUMBER");
  field("PTEC NUMBER");
  field("PTRC NUMBER");
  field("FILE NO. IT");
  field("FILE NO. TDS");
}
