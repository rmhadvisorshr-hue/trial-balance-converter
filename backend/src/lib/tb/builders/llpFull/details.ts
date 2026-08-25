import type ExcelJS from "exceljs";
import type { BuildCtx } from "../common";
import { setCell, applyColumnWidths } from "../../excel/helpers";
import type { Partner } from "../partnershipFull/partners";

// 'DETAILS': Partnership's client-master-data fields, plus the two
// disclosures unique to an LLP - LLPIN (the entity's registration number,
// pre-filled from meta.llpin the same way meta.cin pre-fills Pvt Ltd's CIN)
// and one DPIN per designated partner. LLPIN lives on `ctx.meta`, so every
// other sheet (BS, FORM 8) reads the same value directly rather than
// cross-referencing this sheet - the "single source of truth" is the shared
// BuildCtx, the same pattern already used for firmName/periodLabel
// everywhere. DPIN is genuinely only known here (a TB carries no DPIN data);
// each detected partner gets a labeled blank row for the CA to fill in.
export function writeDetailsSheet(ws: ExcelJS.Worksheet, ctx: BuildCtx, partners: Partner[]) {
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
  field("LLP NAME", meta.firmName || "");
  field("LLPIN", meta.llpin || "");
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

  r++;
  setCell(ws, r++, 2, "DESIGNATED PARTNERS - DPIN", { bold: true });
  for (const p of partners) {
    field(p.name);
  }
}
