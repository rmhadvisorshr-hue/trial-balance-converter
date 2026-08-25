import type ExcelJS from "exceljs";
import type { BuildCtx } from "../common";
import { writeSchedulesSheet } from "./schedules";
import { writeFixedAssetSchedule } from "./fixedAssetSchedule";
import { writeTradingProfitLoss } from "./tradingPL";
import { writeFirmCapitalBalanceSheet } from "./firmCapitalBalanceSheet";
import { writeSelfSchedule } from "./selfSchedule";
import { writeSelfCapitalBalanceSheet } from "./selfCapitalBalanceSheet";
import { writeCompSheet } from "./comp";

// Build order: FIRM SCH/FIRM FA/P&L (pure TB data) -> FIRM CAP&BS (needs Net
// Profit + schedules + FA) -> SELF SCH (standalone template) -> SELF
// CAP&BS(V)/SELF CAP&BS (needs FIRM CAP&BS's Net Profit reference) -> COMP
// (needs Net Profit). Same dependency-ordered-write pattern as
// scheduleiii/index.ts and partnershipFull/index.ts.
export function buildProprietorFull(wb: ExcelJS.Workbook, ctx: BuildCtx) {
  wb.calcProperties.fullCalcOnLoad = true;

  const compWs = wb.addWorksheet("COMP");
  const plVWs = wb.addWorksheet("P&L(V)");
  const plWs = wb.addWorksheet("P&L");
  const firmCapWs = wb.addWorksheet("FIRM CAP&BS");
  const firmSchWs = wb.addWorksheet("FIRM SCH");
  const firmFaWs = wb.addWorksheet("FIRM FA");
  const selfCapVWs = wb.addWorksheet("SELF CAP&BS(V)");
  const selfSchWs = wb.addWorksheet("SELF SCH");
  const selfCapWs = wb.addWorksheet("SELF CAP&BS");

  const sch = writeSchedulesSheet(firmSchWs, ctx);
  const fa = writeFixedAssetSchedule(firmFaWs, ctx);
  const tradingPL = writeTradingProfitLoss(plWs, plVWs, ctx);

  // writeTSection places a profit's balancing figure on the debit side
  // (column B, "To Net Profit c/d") - matches the common case (a going
  // concern); a trading loss would instead land in column E, same
  // limitation noted in tradingPL.ts and the Partnership build.
  const firmCap = writeFirmCapitalBalanceSheet(
    firmCapWs,
    ctx,
    `'P&L'!B${tradingPL.netProfitRow}`,
    sch,
    fa,
  );

  const selfSch = writeSelfSchedule(selfSchWs, ctx);
  writeSelfCapitalBalanceSheet(
    selfCapVWs,
    selfCapWs,
    ctx,
    `'FIRM CAP&BS'!F${firmCap.netProfitRow}`,
    ctx.st.netProfit,
    selfSch,
  );

  writeCompSheet(compWs, ctx, `'P&L'!B${tradingPL.netProfitRow}`, ctx.st.netProfit);
}
