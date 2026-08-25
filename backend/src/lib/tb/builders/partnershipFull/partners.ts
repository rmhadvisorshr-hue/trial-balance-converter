import type { Statements } from "../../statements";
import { round2 } from "../../classify";

export interface Partner {
  name: string;
  capital: number; // TB closing capital balance for this partner
  sharePct: number; // 0-1, defaults from relative capital balance - not a real profit-sharing ratio
  detected: boolean; // false when the TB has no individually-named capital ledgers
}

// A trial balance has no concept of a "profit-sharing ratio" - only ledger
// balances, and its Capital group can hold two very different shapes: (a)
// one ledger per partner ("Sameer Gonsalves Capital", "Ulhas Gonsalves
// Capital", ...) or (b) a single person's capital account broken into
// transaction-type sub-ledgers ("Drawing", "LIC", "School Fees", "Advance
// Tax", ...) that are NOT different partners at all. Only (a) should be
// split into multiple partners - distinguished here by requiring the name
// to actually read like a capital ledger ("capital" in the name, or the
// literal ledger "Capital"), which transaction-type sub-ledgers won't
// match. When capital is a single lump ledger (or doesn't look
// partner-wise), there's no way to tell how many partners there are, so
// this returns one placeholder partner holding the full balance.
//
// The TB's accumulated "Profit & Loss A/c" / General Reserve balance
// (RESERVES code) isn't attributable to a specific partner from the TB
// alone, so it's folded into each partner's capital proportionally by
// their share % - same principle as splitting "Share in Net Profit", and
// the same bug class as the Pvt Ltd Reserves & Surplus fix (a trial
// balance's accumulated reserves must be counted, not dropped).
export function detectPartners(st: Statements): Partner[] {
  const items = st.byCode.get("CAPITAL") ?? [];
  // "contribution" is included alongside "capital" so LLP trial balances
  // using that naming (LLP Agreements call it Contribution, not Capital)
  // are still detected - partnership ledgers won't incidentally match it.
  const capitalLedgers = items.filter((i) => i.amount !== 0 && /capital|contribution/i.test(i.name));
  const reserves = round2(st.total("RESERVES"));

  if (capitalLedgers.length >= 2) {
    const total = capitalLedgers.reduce((s, i) => s + i.amount, 0);
    return capitalLedgers
      .sort((a, b) => b.amount - a.amount)
      .map((i) => {
        const sharePct = total !== 0 ? i.amount / total : 1 / capitalLedgers.length;
        return {
          name: i.name,
          capital: round2(i.amount + reserves * sharePct),
          sharePct,
          detected: true,
        };
      });
  }

  return [
    {
      name: "Partner (not individually named in trial balance - enter manually)",
      capital: round2(st.total("CAPITAL") + reserves),
      sharePct: 1,
      detected: false,
    },
  ];
}
