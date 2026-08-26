import type { TBlock, TLine } from "./tAccountParser";

// Matches an opening/closing balance line regardless of phrasing ("Opening
// Balance", "Balance b/d", "Balance C/f", a bare "Balance", ...) and
// regardless of which side of the T-account it landed on - a debit- or
// credit-normal account carries its opening/closing balance on whichever
// side keeps the account in balance, so the side alone can't tell you which
// line this is. A bare "Balance" is deliberately included: in a two-column
// T-account, any line whose only description is some form of "balance" is
// the carried-forward figure, never a real transaction.
const BALANCE_RE = /\bbalance\b/i;

function contribution(line: TLine, normalSide: "debit" | "credit"): number {
  const onNormalSide = (normalSide === "debit" && line.side === "dr") || (normalSide === "credit" && line.side === "cr");
  return onNormalSide ? line.amount : -line.amount;
}

export interface DerivedBalance {
  opening: number; // signed so that a positive value always means "normal" (e.g. a real asset, a firm's actual capital)
  closing: number;
  genuineLines: TLine[]; // every line except the opening/closing balance entries
}

// Derives a running account's closing balance the same way the books
// themselves derive it - closing = opening +/- every other movement -
// instead of trusting whichever side the source sheet happened to place the
// closing figure on (which flips between debit-normal asset accounts and
// credit-normal equity accounts, and even flips within a credit-normal
// account whenever a partner ends the year overdrawn).
export function deriveBalance(block: TBlock, normalSide: "debit" | "credit"): DerivedBalance {
  let opening = 0;
  let openingSeen = false;
  const genuineLines: TLine[] = [];
  for (const line of block.lines) {
    if (BALANCE_RE.test(line.label)) {
      if (!openingSeen) {
        opening = contribution(line, normalSide);
        openingSeen = true;
      }
      // A second balance-shaped line is the source's own derived closing
      // figure - redundant with what this function computes, so it's
      // dropped rather than trusted (the source's WDV/rounding conventions
      // can differ slightly from a from-scratch recomputation).
      continue;
    }
    genuineLines.push(line);
  }
  const closing = opening + genuineLines.reduce((s, l) => s + contribution(l, normalSide), 0);
  return { opening, closing, genuineLines };
}

export function signedContribution(line: TLine, normalSide: "debit" | "credit"): number {
  return contribution(line, normalSide);
}
