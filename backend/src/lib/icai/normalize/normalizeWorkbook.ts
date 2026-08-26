import type { WorkbookScan } from "../parsers/workbookLoader";
import type { FixedAssetLine, IcaiNoteCode } from "../types";
import { interpretTAccounts, type RawLine } from "./interpretTAccounts";
import { parseCapitalAccounts, type RawOwnerCapital } from "../parsers/capitalParser";
import { parseFixedAssetSchedule } from "../parsers/fixedAssetParser";
import { parseScheduleBlocks } from "../parsers/scheduleParser";
import { parseBalanceSheetLeaves, findBalanceSheetGrandTotal } from "../parsers/balanceSheetParser";
import { categorizeFixedAsset, categorizeBalanceSheetLine } from "../mapping/icaiCategory";

export type ScheduleTarget = "payable" | "otherCurrLiab" | "loansAdvances" | "unknown";

export interface ScheduleGroup {
  heading: string;
  lines: { label: string; amount: number }[];
}

export interface ClassifiedLine {
  name: string;
  amount: number;
  sourceSheet: string;
  code: IcaiNoteCode;
  confident: boolean;
  reason: string;
}

export interface YearData {
  fileName: string;
  entityName: string;
  periodEndLabel: string;
  revenueLines: RawLine[];
  costLines: RawLine[];
  otherIncomeLines: RawLine[];
  expenseLines: RawLine[];
  taxProvision: number;
  closingWip: number | null;
  owners: RawOwnerCapital[];
  fixedAssets: FixedAssetLine[];
  scheduleGroups: Record<ScheduleTarget, ScheduleGroup[]>;
  bsSupplementLines: ClassifiedLine[];
  balanceSheetTotal: number | null;
  reviewItems: RawLine[];
  flags: string[];
}

function classifyScheduleTarget(title: string): ScheduleTarget {
  if (/creditor/i.test(title)) return "payable";
  if (/book(ing)?/i.test(title)) return "otherCurrLiab";
  if (/loan|advance/i.test(title)) return "loansAdvances";
  return "unknown";
}

export function normalizeWorkbook(scan: WorkbookScan): YearData {
  const tAccounts = interpretTAccounts(scan.sheets);

  const owners = scan.sheets
    .filter((s) => s.kind === "tAccount")
    .flatMap((s) => parseCapitalAccounts(s.grid));

  const fixedAssets: FixedAssetLine[] = scan.sheets
    .filter((s) => s.kind === "fixedAssets")
    .flatMap((s) => parseFixedAssetSchedule(s.grid))
    .map((row) => ({ ...row, category: categorizeFixedAsset(row.name) }));

  const scheduleGroups: Record<ScheduleTarget, ScheduleGroup[]> = {
    payable: [],
    otherCurrLiab: [],
    loansAdvances: [],
    unknown: [],
  };
  const reviewItems: RawLine[] = [...tAccounts.reviewItems];

  for (const sheet of scan.sheets.filter((s) => s.kind === "schedule")) {
    for (const block of parseScheduleBlocks(sheet.grid)) {
      if (block.lines.length === 0) continue;
      const target = classifyScheduleTarget(block.title);
      if (target === "unknown") {
        for (const l of block.lines) {
          reviewItems.push({ label: l.label, amount: l.amount, sourceSheet: `${sheet.name} / ${block.title}` });
        }
        continue;
      }
      scheduleGroups[target].push({
        heading: block.title,
        lines: block.lines.map((l) => ({ label: l.label, amount: l.amount })),
      });
    }
  }

  const balanceSheetSheets = scan.sheets.filter((s) => s.kind === "balanceSheet");
  const balanceSheetTotal = balanceSheetSheets
    .map((s) => findBalanceSheetGrandTotal(s.grid))
    .find((t): t is number => t != null) ?? null;

  // A Balance Sheet leaf is "new" data worth extracting only if its amount
  // doesn't already match something captured from a more specific sheet -
  // matching by value, not by label, is what lets this stay generic instead
  // of hardcoding which BS line labels happen to duplicate which schedule.
  const knownAmounts = new Set<number>();
  const remember = (n: number) => knownAmounts.add(Math.round(n));
  for (const l of [...tAccounts.revenueLines, ...tAccounts.costLines, ...tAccounts.otherIncomeLines, ...tAccounts.expenseLines]) remember(l.amount);
  for (const o of owners) remember(o.closingBalance);
  remember(owners.reduce((s, o) => s + o.closingBalance, 0));
  for (const fa of fixedAssets) remember(fa.closing);
  remember(fixedAssets.reduce((s, fa) => s + fa.closing, 0));
  for (const groups of Object.values(scheduleGroups)) {
    for (const g of groups) {
      for (const l of g.lines) remember(l.amount);
      remember(g.lines.reduce((s, l) => s + l.amount, 0));
    }
  }
  if (tAccounts.closingWip != null) remember(tAccounts.closingWip);
  remember(tAccounts.taxProvision);

  const bsSupplementLines: ClassifiedLine[] = [];
  for (const sheet of balanceSheetSheets) {
    for (const leaf of parseBalanceSheetLeaves(sheet.grid)) {
      if (Math.abs(leaf.amount) < 0.01 || knownAmounts.has(Math.round(leaf.amount))) continue;
      const result = categorizeBalanceSheetLine(leaf);
      if (!result) {
        reviewItems.push({ label: `${leaf.group} - ${leaf.label}`, amount: leaf.amount, sourceSheet: sheet.name });
        continue;
      }
      bsSupplementLines.push({
        name: leaf.label,
        amount: leaf.amount,
        sourceSheet: sheet.name,
        code: result.code,
        confident: result.confident,
        reason: result.reason,
      });
    }
  }

  return {
    fileName: scan.fileName,
    entityName: scan.entityName,
    periodEndLabel: scan.periodEndLabel,
    revenueLines: tAccounts.revenueLines,
    costLines: tAccounts.costLines,
    otherIncomeLines: tAccounts.otherIncomeLines,
    expenseLines: tAccounts.expenseLines,
    taxProvision: tAccounts.taxProvision,
    closingWip: tAccounts.closingWip,
    owners,
    fixedAssets,
    scheduleGroups,
    bsSupplementLines,
    balanceSheetTotal,
    reviewItems,
    flags: tAccounts.flags,
  };
}
