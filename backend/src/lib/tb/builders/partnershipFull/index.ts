import type ExcelJS from "exceljs";
import type { BuildCtx } from "../common";
import { detectPartners } from "./partners";
import { writeSchedulesSheet } from "./schedules";
import { writeFixedAssetSchedule } from "./fixedAssetSchedule";
import { writeTradingProfitLoss } from "./tradingPL";
import { writePlAppropriation } from "./plAppropriation";
import { writeCapitalAccounts } from "./capitalAccounts";
import { writePartnershipBalanceSheet } from "./balanceSheet";
import { writeStatSheet } from "./stat";
import { writeInterestOnCapital } from "./interestOnCapital";
import { writeDetailsSheet } from "./details";

// Build order: SCH/Fixed Assets/Trading P&L (pure TB data, no cross-sheet
// dependencies) -> p&l app (needs Net Profit from P&L) -> Cap (needs p&l
// app's partner rows) -> BS/BS(V) (needs Cap+SCH+Fixed Assets) -> STAT
// (needs p&l app's share-of-profit range) -> Int on Capital/DETAILS
// (standalone). Same dependency-ordered-write, dynamic-row-tracking pattern
// as builders/scheduleiii/index.ts.
export function buildPartnershipFull(wb: ExcelJS.Workbook, ctx: BuildCtx) {
  wb.calcProperties.fullCalcOnLoad = true;

  const detailsWs = wb.addWorksheet("DETAILS");
  const statWs = wb.addWorksheet("STAT");
  const plWs = wb.addWorksheet("P&L");
  const plVWs = wb.addWorksheet("PL (V)");
  const plAppWs = wb.addWorksheet("p&l app");
  const capWs = wb.addWorksheet("Cap");
  const bsWs = wb.addWorksheet("BS");
  const bsVWs = wb.addWorksheet("BS (V)");
  const schWs = wb.addWorksheet("SCH");
  const faWs = wb.addWorksheet("FIXED ASSET SCH");
  const intWs = wb.addWorksheet("Int on Capital");

  const partners = detectPartners(ctx.st);

  const sch = writeSchedulesSheet(schWs, ctx);
  const fa = writeFixedAssetSchedule(faWs, ctx);
  const tradingPL = writeTradingProfitLoss(plWs, plVWs, ctx);

  // writeTSection places a profit's balancing figure on the debit side
  // (column B, "To Net Profit c/d") - matches the common case (a going
  // concern); a trading loss would instead land in column E and this
  // reference would need updating, same limitation noted in tradingPL.ts.
  const plApp = writePlAppropriation(plAppWs, ctx, partners, `'P&L'!B${tradingPL.netProfitRow}`);
  const cap = writeCapitalAccounts(capWs, ctx, partners, plApp);
  writePartnershipBalanceSheet(bsWs, bsVWs, ctx, partners, cap, sch, fa, plApp);
  writeStatSheet(statWs, ctx, plApp);
  writeInterestOnCapital(intWs, ctx, partners);
  writeDetailsSheet(detailsWs, ctx);
}
