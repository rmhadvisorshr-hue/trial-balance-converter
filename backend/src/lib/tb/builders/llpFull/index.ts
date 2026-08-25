import type ExcelJS from "exceljs";
import type { BuildCtx } from "../common";
import { detectPartners } from "../partnershipFull/partners";
import { writeSchedulesSheet } from "../partnershipFull/schedules";
import { writeFixedAssetSchedule } from "../partnershipFull/fixedAssetSchedule";
import { writeTradingProfitLoss } from "../partnershipFull/tradingPL";
import { writePlAppropriation } from "../partnershipFull/plAppropriation";
import { writeInterestOnCapital } from "../partnershipFull/interestOnCapital";
import { writeCapitalCurrentAccounts } from "./capitalCurrentAccounts";
import { writeLlpBalanceSheet } from "./balanceSheet";
import { writeStatSheet } from "./stat";
import { writeAgingSheet } from "./agingSchedule";
import { writeStatutoryNotesSheet } from "./statutoryNotes";
import { writeForm8Sheet } from "./form8";
import { writeDetailsSheet } from "./details";

// LLP gets its own family/builder (not a Partnership variant) because its
// statutory presentation genuinely differs: two accounts per partner
// (Fixed Capital/Contribution + Current, not one Cap account), a Sources of
// Funds/Application of Funds Balance Sheet (not Partnership's T-format),
// and LLP-only disclosures (LLPIN/DPIN, FORM 8). The tax/appropriation
// engine, schedules, fixed asset register and trading P&L are reused
// unchanged from partnershipFull/ - see each imported module's own file for
// why that reuse is safe (they're already entity-agnostic).
//
// Build order (dependency-ordered, same discipline as scheduleiii/index.ts
// and partnershipFull/index.ts): SCH/FIXED ASSET SCH/P&L (pure TB data) ->
// p&l app (needs Net Profit) -> Cap & Current (needs p&l app's partner
// rows) -> BS/BS (V) (needs Cap & Current + SCH + FIXED ASSET SCH) -> STAT
// (needs p&l app's share-of-profit range) -> FORM 8 (needs BS's totals) ->
// AGING SC (needs SCH) -> STATUTORY NOTES (needs partners) -> Int on
// Capital/DETAILS (standalone).
export function buildLlpFull(wb: ExcelJS.Workbook, ctx: BuildCtx) {
  wb.calcProperties.fullCalcOnLoad = true;

  const detailsWs = wb.addWorksheet("DETAILS");
  const statWs = wb.addWorksheet("STAT");
  const plWs = wb.addWorksheet("P&L");
  const plVWs = wb.addWorksheet("PL (V)");
  const plAppWs = wb.addWorksheet("p&l app");
  const capWs = wb.addWorksheet("Cap & Current");
  const bsWs = wb.addWorksheet("BS");
  const bsVWs = wb.addWorksheet("BS (V)");
  const schWs = wb.addWorksheet("SCH");
  const faWs = wb.addWorksheet("FIXED ASSET SCH");
  const intWs = wb.addWorksheet("Int on Capital");
  const agingWs = wb.addWorksheet("AGING SC");
  const notesWs = wb.addWorksheet("STATUTORY NOTES");
  const form8Ws = wb.addWorksheet("FORM 8");

  const partners = detectPartners(ctx.st);

  const sch = writeSchedulesSheet(schWs, ctx);
  const fa = writeFixedAssetSchedule(faWs, ctx);
  const tradingPL = writeTradingProfitLoss(plWs, plVWs, ctx);

  // writeTSection places a profit's balancing figure on the debit side
  // (column B, "To Net Profit c/d") - matches the common case (a going
  // concern); a trading loss would instead land in column E, same
  // limitation noted in tradingPL.ts and the Partnership build.
  const plApp = writePlAppropriation(plAppWs, ctx, partners, `'P&L'!B${tradingPL.netProfitRow}`);
  const cap = writeCapitalCurrentAccounts(capWs, ctx, partners, plApp);
  const bsResult = writeLlpBalanceSheet(bsWs, bsVWs, ctx, partners, cap, sch, fa, plApp);
  writeStatSheet(statWs, ctx, plApp);
  writeForm8Sheet(form8Ws, ctx, partners, cap, bsResult);
  writeAgingSheet(agingWs, ctx, sch);
  writeStatutoryNotesSheet(notesWs, ctx, partners);
  writeInterestOnCapital(intWs, ctx, partners);
  writeDetailsSheet(detailsWs, ctx, partners);
}
