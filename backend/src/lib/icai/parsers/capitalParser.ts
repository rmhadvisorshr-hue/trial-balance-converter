import type { Grid } from "./grid";
import { parseTBlocks, type TBlock } from "./tAccountParser";
import { deriveBalance, signedContribution } from "./balanceAlgebra";

export interface RawOwnerCapital {
  name: string;
  pan?: string;
  sharePercent?: number;
  openingBalance: number; // signed: negative if the partner was in debit/overdrawn
  capitalIntroduced: number;
  remuneration: number;
  interest: number;
  shareOfProfit: number;
  otherCredits: number;
  withdrawals: number;
  withdrawalBreakup: { label: string; amount: number }[];
  closingBalance: number;
}

const INTRODUCED_RE = /capital\s*intr?o/i;
const REMUNERATION_RE = /remuneration|salary/i;
const INTEREST_RE = /interest/i;
const PROFIT_SHARE_RE = /share.*(profit|loss)|(profit|loss).*share/i;

function classifyBucket(label: string): "introduced" | "remuneration" | "interest" | "profit" | "other" {
  if (INTRODUCED_RE.test(label)) return "introduced";
  if (REMUNERATION_RE.test(label)) return "remuneration";
  if (INTEREST_RE.test(label)) return "interest";
  if (PROFIT_SHARE_RE.test(label)) return "profit";
  return "other";
}

// A capital account is credit-normal (a partner's capital is a liability
// from the firm's point of view), but an overdrawn partner's account can
// flip debit-normal for a year - deriveBalance() handles that; this function
// only has to sort the *genuine* (non-balance) lines into the Note 3(a)
// columns.
function parseCapitalBlock(block: TBlock): RawOwnerCapital {
  const m = block.title.match(/^(.*?)-?\s*capital\s*account/i);
  const name = (m?.[1] ?? block.title).trim().replace(/[-–]\s*$/, "").trim();
  const pan = block.title.match(/PAN\s*:\s*([A-Z0-9]+)/i)?.[1];
  const shareStr = block.title.match(/SHARE\s*:\s*([\d.]+)\s*%/i)?.[1];
  const sharePercent = shareStr ? Number(shareStr) / 100 : undefined;

  const { opening, closing, genuineLines } = deriveBalance(block, "credit");

  let capitalIntroduced = 0;
  let remuneration = 0;
  let interest = 0;
  let shareOfProfit = 0;
  let otherCredits = 0;
  let withdrawals = 0;
  const withdrawalBreakup: { label: string; amount: number }[] = [];

  for (const line of genuineLines) {
    const signed = signedContribution(line, "credit");
    const bucket = classifyBucket(line.label);
    if (bucket === "introduced") capitalIntroduced += signed;
    else if (bucket === "remuneration") remuneration += signed;
    else if (bucket === "interest") interest += signed;
    else if (bucket === "profit") shareOfProfit += signed;
    else if (line.side === "cr") otherCredits += line.amount;
    else {
      withdrawals += line.amount;
      withdrawalBreakup.push({ label: line.label, amount: line.amount });
    }
  }

  return {
    name,
    pan,
    sharePercent,
    openingBalance: opening,
    capitalIntroduced,
    remuneration,
    interest,
    shareOfProfit,
    otherCredits,
    withdrawals,
    withdrawalBreakup,
    closingBalance: closing,
  };
}

export function parseCapitalAccounts(grid: Grid): RawOwnerCapital[] {
  const blocks = parseTBlocks(grid).filter((b) => /capital\s*account/i.test(b.title));
  return blocks.map(parseCapitalBlock);
}
