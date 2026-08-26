import type {
  WorkbookAnalysis,
  NormalizedAccount,
  OwnerCapitalAccount,
  SupportingSchedule,
  ScheduleSection,
  ReviewItem,
  IcaiEntityKind,
} from "../types";
import type { YearData, ScheduleTarget, ClassifiedLine } from "./normalizeWorkbook";
import type { RawLine } from "./interpretTAccounts";
import type { RawOwnerCapital } from "../parsers/capitalParser";
import { categorizeExpenseLine } from "../mapping/icaiCategory";
import type { IcaiNoteCode } from "../types";

function norm(name: string): string {
  return name.trim().toLowerCase().replace(/\s+/g, " ");
}

interface MergedLine {
  name: string;
  amountCurrent: number;
  amountPrevious: number;
  sourceSheet: string;
}

function mergeLines(current: RawLine[], previous: RawLine[]): MergedLine[] {
  const byKey = new Map<string, MergedLine>();
  for (const l of current) {
    const key = norm(l.label);
    const existing = byKey.get(key);
    if (existing) existing.amountCurrent += l.amount;
    else byKey.set(key, { name: l.label, amountCurrent: l.amount, amountPrevious: 0, sourceSheet: l.sourceSheet });
  }
  for (const l of previous) {
    const key = norm(l.label);
    const existing = byKey.get(key);
    if (existing) existing.amountPrevious += l.amount;
    else byKey.set(key, { name: l.label, amountCurrent: 0, amountPrevious: l.amount, sourceSheet: l.sourceSheet });
  }
  return [...byKey.values()];
}

function toAccounts(
  lines: MergedLine[],
  categorize: (name: string) => { code: IcaiNoteCode; confident: boolean; reason: string },
): NormalizedAccount[] {
  return lines.map((l) => {
    const { code, confident, reason } = categorize(l.name);
    return {
      name: l.name,
      amountCurrent: l.amountCurrent,
      amountPrevious: l.amountPrevious,
      sourceSheet: l.sourceSheet,
      code,
      autoCode: code,
      confident,
      reason,
    };
  });
}

function mergeOwners(current: RawOwnerCapital[], previous: RawOwnerCapital[]): OwnerCapitalAccount[] {
  const prevByName = new Map(previous.map((o) => [norm(o.name), o]));
  const seen = new Set<string>();
  const out: OwnerCapitalAccount[] = current.map((o) => {
    const key = norm(o.name);
    seen.add(key);
    const p = prevByName.get(key);
    return {
      name: o.name,
      pan: o.pan,
      sharePercent: o.sharePercent,
      openingBalance: o.openingBalance,
      openingBalancePrevious: p?.openingBalance ?? 0,
      capitalIntroduced: o.capitalIntroduced,
      capitalIntroducedPrevious: p?.capitalIntroduced ?? 0,
      remuneration: o.remuneration,
      remunerationPrevious: p?.remuneration ?? 0,
      interest: o.interest,
      interestPrevious: p?.interest ?? 0,
      shareOfProfit: o.shareOfProfit,
      shareOfProfitPrevious: p?.shareOfProfit ?? 0,
      otherCredits: o.otherCredits,
      otherCreditsPrevious: p?.otherCredits ?? 0,
      withdrawals: o.withdrawals,
      withdrawalsPrevious: p?.withdrawals ?? 0,
      withdrawalBreakup: o.withdrawalBreakup,
      closingBalance: o.closingBalance,
      closingBalancePrevious: p?.closingBalance ?? 0,
    };
  });
  for (const p of previous) {
    if (seen.has(norm(p.name))) continue;
    out.push({
      name: p.name,
      pan: p.pan,
      sharePercent: p.sharePercent,
      openingBalance: 0,
      openingBalancePrevious: p.openingBalance,
      capitalIntroduced: 0,
      capitalIntroducedPrevious: p.capitalIntroduced,
      remuneration: 0,
      remunerationPrevious: p.remuneration,
      interest: 0,
      interestPrevious: p.interest,
      shareOfProfit: 0,
      shareOfProfitPrevious: p.shareOfProfit,
      otherCredits: 0,
      otherCreditsPrevious: p.otherCredits,
      withdrawals: 0,
      withdrawalsPrevious: p.withdrawals,
      withdrawalBreakup: [],
      closingBalance: 0,
      closingBalancePrevious: p.closingBalance,
    });
  }
  return out;
}

const SCHEDULE_META: Record<Exclude<ScheduleTarget, "unknown">, { title: string; note: string; letter: string }> = {
  payable: { title: "Trade Payables (Sundry Creditors)", note: "9", letter: "A" },
  otherCurrLiab: { title: "Advance Against Booking / Customer Advances", note: "10", letter: "B" },
  loansAdvances: { title: "Loans & Advances", note: "13", letter: "C" },
};

function mergeSchedules(current: YearData, previous: YearData | null): SupportingSchedule[] {
  const out: SupportingSchedule[] = [];
  const targets: Exclude<ScheduleTarget, "unknown">[] = ["payable", "otherCurrLiab", "loansAdvances"];

  for (const target of targets) {
    const curGroups = current.scheduleGroups[target];
    const prevGroups = previous?.scheduleGroups[target] ?? [];
    if (curGroups.length === 0 && prevGroups.length === 0) continue;

    const prevByName = new Map<string, number>();
    for (const g of prevGroups) for (const l of g.lines) prevByName.set(norm(l.label), (prevByName.get(norm(l.label)) ?? 0) + l.amount);
    const matchedPrevKeys = new Set<string>();

    const sections: ScheduleSection[] = curGroups.map((g) => {
      const lines = g.lines.map((l) => {
        const key = norm(l.label);
        matchedPrevKeys.add(key);
        return { name: l.label, amountCurrent: l.amount, amountPrevious: prevByName.get(key) ?? 0 };
      });
      return {
        heading: g.heading,
        lines,
        subtotalCurrent: lines.reduce((s, l) => s + l.amountCurrent, 0),
        subtotalPrevious: lines.reduce((s, l) => s + l.amountPrevious, 0),
      };
    });

    // Parties that existed only in the previous year's schedule (fully paid
    // off / no longer applicable this year) still belong in the comparative
    // column, grouped under the previous year's own section headings.
    const leftoverByHeading = new Map<string, { name: string; amountPrevious: number }[]>();
    for (const g of prevGroups) {
      for (const l of g.lines) {
        const key = norm(l.label);
        if (matchedPrevKeys.has(key)) continue;
        const list = leftoverByHeading.get(g.heading) ?? [];
        list.push({ name: l.label, amountPrevious: l.amount });
        leftoverByHeading.set(g.heading, list);
      }
    }
    for (const [heading, lines] of leftoverByHeading) {
      const existing = sections.find((s) => s.heading === heading);
      const mapped = lines.map((l) => ({ name: l.name, amountCurrent: 0, amountPrevious: l.amountPrevious }));
      if (existing) {
        existing.lines.push(...mapped);
        existing.subtotalPrevious += lines.reduce((s, l) => s + l.amountPrevious, 0);
      } else {
        sections.push({
          heading,
          lines: mapped,
          subtotalCurrent: 0,
          subtotalPrevious: lines.reduce((s, l) => s + l.amountPrevious, 0),
        });
      }
    }

    const meta = SCHEDULE_META[target];
    out.push({ letter: meta.letter, title: meta.title, supportsNote: meta.note, sections });
  }
  return out;
}

function mergeClassifiedLines(current: ClassifiedLine[], previous: ClassifiedLine[]): NormalizedAccount[] {
  const byKey = new Map<string, NormalizedAccount>();
  for (const l of current) {
    const key = norm(l.name);
    byKey.set(key, {
      name: l.name,
      amountCurrent: l.amount,
      amountPrevious: 0,
      sourceSheet: l.sourceSheet,
      code: l.code,
      autoCode: l.code,
      confident: l.confident,
      reason: l.reason,
    });
  }
  for (const l of previous) {
    const key = norm(l.name);
    const existing = byKey.get(key);
    if (existing) existing.amountPrevious = l.amount;
    else
      byKey.set(key, {
        name: l.name,
        amountCurrent: 0,
        amountPrevious: l.amount,
        sourceSheet: l.sourceSheet,
        code: l.code,
        autoCode: l.code,
        confident: l.confident,
        reason: l.reason,
      });
  }
  return [...byKey.values()];
}

function reviewItemsFrom(items: RawLine[], sign: "current" | "previous"): ReviewItem[] {
  return items.map((i) => ({
    sourceSheet: i.sourceSheet,
    label: i.label,
    amountCurrent: sign === "current" ? i.amount : 0,
    amountPrevious: sign === "previous" ? i.amount : 0,
    suggestedCode: "UNMAPPED",
    confidence: 0,
    reason: "No recognizable ICAI category matched this line - please classify it manually.",
  }));
}

export function mergeYears(
  current: YearData,
  previous: YearData | null,
  entityKindOverride?: IcaiEntityKind,
): WorkbookAnalysis {
  const accounts: NormalizedAccount[] = [
    ...toAccounts(mergeLines(current.revenueLines, previous?.revenueLines ?? []), () => ({
      code: "N19_REVENUE",
      confident: true,
      reason: "Credit side of the Trading Account.",
    })),
    ...toAccounts(mergeLines(current.costLines, previous?.costLines ?? []), () => ({
      code: "N21_COST_OF_CONSTRUCTION",
      confident: true,
      reason: "Debit side of the Trading Account.",
    })),
    ...toAccounts(mergeLines(current.otherIncomeLines, previous?.otherIncomeLines ?? []), () => ({
      code: "N20_OTHER_INCOME",
      confident: true,
      reason: "Credit side of the Profit & Loss Account.",
    })),
    ...toAccounts(mergeLines(current.expenseLines, previous?.expenseLines ?? []), categorizeExpenseLine),
    ...mergeClassifiedLines(current.bsSupplementLines, previous?.bsSupplementLines ?? []),
  ];

  if (current.taxProvision || previous?.taxProvision) {
    accounts.push({
      name: "Provision for income-tax",
      amountCurrent: current.taxProvision,
      amountPrevious: previous?.taxProvision ?? 0,
      sourceSheet: "P&L Appropriation",
      code: "N8_PROVISION",
      autoCode: "N8_PROVISION",
      confident: true,
      reason: "Provision for tax debited in the Profit & Loss Appropriation Account.",
    });
  }

  if (current.closingWip != null || previous?.closingWip != null) {
    accounts.push({
      name: "Work-in-progress",
      amountCurrent: current.closingWip ?? 0,
      amountPrevious: previous?.closingWip ?? 0,
      sourceSheet: "Work in Progress Account",
      code: "N15_INVENTORY",
      autoCode: "N15_INVENTORY",
      confident: true,
      reason: "Closing balance of the Work-in-Progress account.",
    });
  }

  const owners = mergeOwners(current.owners, previous?.owners ?? []);
  const totalRemuneration = owners.reduce((s, o) => s + o.remuneration, 0);
  const totalRemunerationPrevious = owners.reduce((s, o) => s + o.remunerationPrevious, 0);
  if (totalRemuneration || totalRemunerationPrevious) {
    accounts.push({
      name: "Partners' remuneration",
      amountCurrent: totalRemuneration,
      amountPrevious: totalRemunerationPrevious,
      sourceSheet: "Partners' Capital Accounts",
      code: "PARTNERS_REMUNERATION",
      autoCode: "PARTNERS_REMUNERATION",
      confident: true,
      reason: "Sum of remuneration credited across all partners' capital accounts.",
    });
  }

  const schedules = mergeSchedules(current, previous);
  for (const sch of schedules) {
    const amountCurrent = sch.sections.reduce((s, sec) => s + sec.subtotalCurrent, 0);
    const amountPrevious = sch.sections.reduce((s, sec) => s + sec.subtotalPrevious, 0);
    const code: IcaiNoteCode =
      sch.supportsNote === "9" ? "N9_TRADE_PAYABLE" : sch.supportsNote === "10" ? "N10_OTHER_CURR_LIAB" : "N13_LOANS_ADVANCES";
    accounts.push({
      name: sch.title,
      amountCurrent,
      amountPrevious,
      sourceSheet: `Schedule ${sch.letter}`,
      code,
      autoCode: code,
      confident: true,
      reason: `Total of Schedule ${sch.letter} (${sch.title}).`,
    });
  }

  const reviewItems = [
    ...reviewItemsFrom(current.reviewItems, "current"),
    ...reviewItemsFrom(previous?.reviewItems ?? [], "previous"),
  ];

  const warnings = [...new Set([...current.flags, ...(previous?.flags ?? [])])];
  if (!previous) {
    warnings.push("No comparative-year workbook was provided - previous-year figures are shown as 0.");
  }

  const entityKind: IcaiEntityKind =
    entityKindOverride ??
    (/\bllp\b/i.test(current.entityName) ? "llp" : owners.length <= 1 ? "proprietor" : "partnership");

  return {
    entityName: current.entityName || previous?.entityName || "",
    entityKind,
    currentYearLabel: current.periodEndLabel,
    previousYearLabel: previous?.periodEndLabel ?? "",
    accounts,
    owners,
    fixedAssets: { current: current.fixedAssets, previous: previous?.fixedAssets ?? [] },
    schedules,
    reviewItems,
    warnings,
  };
}
