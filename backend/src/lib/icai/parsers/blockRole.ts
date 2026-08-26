export type BlockRole =
  | "trading"
  | "profitLoss"
  | "appropriation"
  | "workInProgress"
  | "capital"
  | "unknown";

// Classifies one T-account block by its own heading text, independent of
// which sheet it lives on - a sheet commonly stacks several such blocks
// (e.g. a Trading Account followed by a Work-in-Progress account on the same
// "trad" sheet).
export function classifyBlockRole(title: string): BlockRole {
  const t = title.toLowerCase();
  if (/appropriation/.test(t)) return "appropriation";
  if (/capital\s*account/.test(t)) return "capital";
  if (/work[\s-]*in[\s-]*progress|\bwip\b/.test(t)) return "workInProgress";
  if (/trading\s*account/.test(t)) return "trading";
  if (/profit\s*(and|&)\s*loss\s*account/.test(t)) return "profitLoss";
  return "unknown";
}
