import ExcelJS from "exceljs";
import type { ConvertPayload } from "../types";
import { entityFamily, FIGURES_UNIT_DIVISOR, FIGURES_UNIT_HEADING } from "../types";
import { computeStatements } from "../statements";
import type { BuildCtx } from "./common";
import { buildPartnershipFull } from "./partnershipFull";
import { buildProprietorFull } from "./proprietorFull";
import { buildScheduleIII } from "./scheduleiii";
import { buildLlpFull } from "./llpFull";
import { applyWorkbookFormatting } from "../excel/formatting";

export async function buildWorkbook(payload: ConvertPayload): Promise<Buffer> {
  // Statements are always computed on actual (unscaled) figures; the unit
  // only affects how amounts are written into the sheets below.
  const st = computeStatements(payload.ledgers);
  const family = entityFamily(payload.entity);
  const designation =
    payload.entity === "proprietor" ? "Proprietor" : payload.entity === "llp" ? "Designated Partner" : "Partner";
  const figuresUnit = payload.meta.figuresUnit || "actual";
  const ctx: BuildCtx = {
    meta: payload.meta,
    st,
    ledgers: payload.ledgers,
    forFirmText: `For ${payload.meta.firmName || "the entity"}`,
    designation,
    divisor: FIGURES_UNIT_DIVISOR[figuresUnit],
    unitHeading: FIGURES_UNIT_HEADING[figuresUnit],
  };

  const wb = new ExcelJS.Workbook();
  wb.creator = "RMH Advisors - Trial Balance Converter";
  wb.created = new Date();

  // Pvt Ltd companies always get the Schedule III package; Partnership
  // always gets the full partner-wise package (single Cap account, P&L
  // Appropriation, Section 40(b), STAT); LLP always gets its own package
  // (Fixed Capital/Current Account split, Sources/Application BS, FORM 8);
  // Proprietor always gets the Business + Personal split package (FIRM
  // CAP&BS/SELF CAP&BS, COMP).
  if (family === "corporate") {
    buildScheduleIII(wb, ctx);
  } else if (family === "proprietor") {
    buildProprietorFull(wb, ctx);
  } else if (family === "llp") {
    buildLlpFull(wb, ctx);
  } else {
    buildPartnershipFull(wb, ctx);
  }

  // Presentation-only pass: fonts/fills/borders/merges/freeze panes/print
  // setup. Runs after every builder above has finished writing values and
  // formulas - see excel/formatting.ts for why this can be a single,
  // content-driven pass instead of per-sheet styling calls.
  applyWorkbookFormatting(wb);

  const buf = await wb.xlsx.writeBuffer();
  return Buffer.from(buf);
}
