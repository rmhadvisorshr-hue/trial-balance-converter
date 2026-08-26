import type { ScannedSheet } from "../parsers/workbookLoader";
import { parseTBlocks, blockTotal } from "../parsers/tAccountParser";
import { classifyBlockRole } from "../parsers/blockRole";
import { deriveBalance } from "../parsers/balanceAlgebra";
import { CARRY_BALANCE_RE } from "../parsers/grid";

export interface RawLine {
  label: string;
  amount: number;
  sourceSheet: string;
}

export interface TAccountInterpretation {
  revenueLines: RawLine[];
  costLines: RawLine[];
  otherIncomeLines: RawLine[];
  expenseLines: RawLine[];
  taxProvision: number;
  closingWip: number | null;
  openingWip: number | null;
  reviewItems: RawLine[]; // blocks the engine couldn't interpret at all
  flags: string[]; // generic accounting-judgement notices for the Basis of Preparation sheet
}

const TAX_PROVISION_RE = /provision.*tax|tax.*provision/i;

// Interprets every trading/P&L/appropriation/WIP block across all tAccount
// sheets in one workbook. Genuinely reusable across any firm's books: it
// never keys off a specific sheet or firm name, only off each block's own
// heading text (classifyBlockRole) and the universal b/d-c/d carry-forward
// notation (CARRY_BALANCE_RE / deriveBalance).
export function interpretTAccounts(sheets: ScannedSheet[]): TAccountInterpretation {
  const result: TAccountInterpretation = {
    revenueLines: [],
    costLines: [],
    otherIncomeLines: [],
    expenseLines: [],
    taxProvision: 0,
    closingWip: null,
    openingWip: null,
    reviewItems: [],
    flags: [],
  };

  let tradingGrossProfit: number | null = null;

  for (const sheet of sheets.filter((s) => s.kind === "tAccount")) {
    const blocks = parseTBlocks(sheet.grid);
    for (const block of blocks) {
      const role = classifyBlockRole(block.title);

      if (role === "trading") {
        const genuine = block.lines.filter((l) => !CARRY_BALANCE_RE.test(l.label));
        for (const l of genuine) {
          const line = { label: l.label, amount: l.amount, sourceSheet: sheet.name };
          if (l.side === "cr") result.revenueLines.push(line);
          else result.costLines.push(line);
        }
        tradingGrossProfit = blockTotal({ ...block, lines: genuine }, "cr") - blockTotal({ ...block, lines: genuine }, "dr");
        continue;
      }

      if (role === "profitLoss") {
        const genuine = block.lines.filter((l) => !CARRY_BALANCE_RE.test(l.label));
        for (const l of genuine) {
          // The credit side of an indirect P&L account carries the trading
          // account's gross profit forward under whatever label the books
          // happen to use for it ("Gross Profit b/d", "Profit from Loading",
          // ...) - matched by value against the trading block, not by label,
          // since that carry-in is never itself "other income".
          if (
            l.side === "cr" &&
            tradingGrossProfit != null &&
            Math.abs(l.amount - tradingGrossProfit) < 1
          ) {
            continue;
          }
          const line = { label: l.label, amount: l.amount, sourceSheet: sheet.name };
          if (l.side === "cr") result.otherIncomeLines.push(line);
          else result.expenseLines.push(line);
        }
        continue;
      }

      if (role === "appropriation") {
        const taxLine = block.lines.find((l) => l.side === "dr" && TAX_PROVISION_RE.test(l.label));
        if (taxLine) result.taxProvision += taxLine.amount;
        // Interest on capital, remuneration and share of profit are read
        // from each owner's own capital account (parsers/capitalParser.ts)
        // instead of re-parsed here, since they're already itemized there
        // per partner with less ambiguous column layout.
        continue;
      }

      if (role === "workInProgress") {
        const { opening, closing } = deriveBalance(block, "debit");
        result.closingWip = (result.closingWip ?? 0) + closing;
        result.openingWip = (result.openingWip ?? 0) + opening;
        result.flags.push(
          `Revenue appears to be recognised into a Work-in-Progress roll-forward account ('${sheet.name}') rather than only on sale/registration - verify this is the intended accounting policy (see AS 2/AS 7/ICDS III) rather than accepting it as a formatting choice.`,
        );
        continue;
      }

      if (role === "capital") continue; // handled by parsers/capitalParser.ts

      // Unknown block shape/role: don't guess - list whatever real money
      // moved through it as a review item instead of silently dropping it
      // or silently mixing it into another note.
      for (const l of block.lines) {
        if (CARRY_BALANCE_RE.test(l.label) || Math.abs(l.amount) < 0.01) continue;
        result.reviewItems.push({ label: l.label, amount: l.amount, sourceSheet: `${sheet.name} / ${block.title || "(untitled block)"}` });
      }
    }
  }

  return result;
}
