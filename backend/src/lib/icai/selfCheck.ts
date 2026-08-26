// Runnable self-check for the ICAI pipeline: parses the real Space Home
// reference workbooks and asserts the aggregate figures against the known-
// correct reference output ("Space Home (A.G) - Financial Statements ...").
// Run with: npx tsx src/lib/icai/selfCheck.ts   (from backend/)
import assert from "node:assert/strict";
import { readFile, writeFile } from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import ExcelJS from "exceljs";
import { analyzeWorkbookFiles, buildIcaiWorkbook } from "./pipeline";

const dataDir = path.resolve(import.meta.dirname, "../../../../data");

async function main() {
  const current = await readFile(path.join(dataDir, "SPACE HOME 26-25.xlsx"));
  const previous = await readFile(path.join(dataDir, "SPACE HOME 24-25.xlsx"));

  // Two files, uploaded together as a collection (no current/previous
  // slots) - role detection must tell them apart on its own.
  const outcome = await analyzeWorkbookFiles([
    { buffer: current, fileName: "SPACE HOME 26-25.xlsx" },
    { buffer: previous, fileName: "SPACE HOME 24-25.xlsx" },
  ]);
  assert.equal(outcome.status, "resolved", `Expected auto-resolved roles, got: ${"reason" in outcome ? outcome.reason : ""}`);
  if (outcome.status !== "resolved") throw new Error("unreachable");
  const analysis = outcome.analysis;
  assert.equal(analysis.currentYearLabel, "31ST MARCH 2026", "Current year should be the later-dated file");
  assert.equal(analysis.previousYearLabel, "31ST MARCH 2025", "Previous year should be the earlier-dated file");

  const sum = (code: string, key: "amountCurrent" | "amountPrevious") =>
    analysis.accounts.filter((a) => a.code === code).reduce((s, a) => s + a[key], 0);

  console.log("Entity:", analysis.entityName, "|", analysis.entityKind);
  console.log("Current year label:", analysis.currentYearLabel);
  console.log("Previous year label:", analysis.previousYearLabel);
  console.log("Owners:", analysis.owners.map((o) => `${o.name} (${o.closingBalance.toFixed(2)} / ${o.closingBalancePrevious.toFixed(2)})`));
  console.log("Review items:", analysis.reviewItems.length, analysis.reviewItems);
  console.log("Warnings:", analysis.warnings);

  const close = (actual: number, expected: number, label: string) => {
    const diff = Math.abs(actual - expected);
    assert.ok(diff < 2, `${label}: expected ${expected}, got ${actual} (diff ${diff})`);
  };

  close(sum("N19_REVENUE", "amountCurrent"), 60998603.57, "Revenue (current)");
  close(sum("N19_REVENUE", "amountPrevious"), 184010369, "Revenue (previous)");
  close(sum("N21_COST_OF_CONSTRUCTION", "amountCurrent"), 34852874.57, "Cost of construction (current)");
  close(sum("N21_COST_OF_CONSTRUCTION", "amountPrevious"), 146039975, "Cost of construction (previous)");
  close(sum("N20_OTHER_INCOME", "amountCurrent"), 8965, "Other income (current)");
  close(sum("N20_OTHER_INCOME", "amountPrevious"), 10752, "Other income (previous)");
  close(sum("N23_EMPLOYEE_BENEFIT", "amountCurrent"), 1176910, "Employee benefit (current)");
  close(sum("N24_FINANCE_COST", "amountCurrent"), 260589, "Finance cost (current)");
  close(sum("N25_DEPRECIATION", "amountCurrent"), 958001, "Depreciation (current)");
  close(sum("N26_OTHER_EXPENSE", "amountCurrent"), 13120012, "Other expenses (current)");
  close(sum("PARTNERS_REMUNERATION", "amountCurrent"), 6563509, "Partners' remuneration (current)");
  close(sum("PARTNERS_REMUNERATION", "amountPrevious"), 13406605, "Partners' remuneration (previous)");
  close(sum("N8_PROVISION", "amountCurrent"), 1271610, "Tax provision (current)");
  close(sum("N9_TRADE_PAYABLE", "amountCurrent"), 1721447, "Trade payables (current)");
  close(sum("N9_TRADE_PAYABLE", "amountPrevious"), 1714372, "Trade payables (previous)");
  close(sum("N10_OTHER_CURR_LIAB", "amountCurrent"), 294310191, "Other current liabilities (current)");
  close(sum("N13_LOANS_ADVANCES", "amountCurrent"), 96825255.86, "Loans & advances incl. advance tax/TDS (current)");
  close(sum("N15_INVENTORY", "amountCurrent"), 295846724.7, "Inventory/WIP (current)");
  close(sum("N15_INVENTORY", "amountPrevious"), 234848121.13, "Inventory/WIP (previous)");
  close(sum("N17_CASH_BANK", "amountCurrent"), 7619374.1, "Cash & bank (current)");
  close(sum("N14_OTHER_NONCURR_ASSET", "amountCurrent"), 72000, "Deposits (current)");
  close(sum("N18_OTHER_CURR_ASSET", "amountCurrent"), 30931, "Other current assets/GST (current)");

  const ownersFundsCurrent = analysis.owners.reduce((s, o) => s + o.closingBalance, 0);
  close(ownersFundsCurrent, 102885809.55, "Owners' funds (current)");
  const ownersFundsPrevious = analysis.owners.reduce((s, o) => s + o.closingBalancePrevious, 0);
  close(ownersFundsPrevious, 102007463.61, "Owners' funds (previous)");

  const ltBorrowing = sum("N5_LT_BORROWING", "amountCurrent");
  close(ltBorrowing, 5404263.11, "Long-term borrowings (current)");

  const faClosingCurrent = analysis.fixedAssets.current.reduce((s, a) => s + a.closing, 0);
  close(faClosingCurrent, 5199035, "Fixed assets closing (current)");
  const faClosingPrevious = analysis.fixedAssets.previous.reduce((s, a) => s + a.closing, 0);
  close(faClosingPrevious, 6015576.36, "Fixed assets closing (previous)");

  const balanceWarnings = analysis.warnings.filter((w) => w.includes("does not balance") || w.includes("does not match"));
  assert.equal(balanceWarnings.length, 0, `Balance sheet cross-check failed: ${balanceWarnings.join(" | ")}`);

  const buffer = await buildIcaiWorkbook(analysis);
  const outPath = path.join(os.tmpdir(), "space-home-icai-selfcheck.xlsx");
  await writeFile(outPath, buffer);
  const reread = new ExcelJS.Workbook();
  await reread.xlsx.load(buffer as unknown as ArrayBuffer);
  console.log("\nGenerated workbook sheets:", reread.worksheets.map((w) => w.name));
  console.log("Saved to:", outPath);
  assert.ok(reread.worksheets.length >= 10, "Expected at least 10 sheets in the generated workbook.");

  // Single-file upload: role detection must resolve on its own (no CA review needed).
  const singleOutcome = await analyzeWorkbookFiles([{ buffer: current, fileName: "SPACE HOME 26-25.xlsx" }]);
  assert.equal(singleOutcome.status, "resolved", "A single uploaded workbook should resolve without review.");

  // Same file uploaded twice: two workbooks, indistinguishable financial
  // years - must NOT silently pick one, must ask the CA to confirm.
  const ambiguousOutcome = await analyzeWorkbookFiles([
    { buffer: current, fileName: "A.xlsx" },
    { buffer: current, fileName: "B.xlsx" },
  ]);
  assert.equal(ambiguousOutcome.status, "needs-review", "Two same-year workbooks should be flagged for CA review, not auto-resolved.");

  // CA-confirmed roles for that same ambiguous pair must then resolve.
  const confirmedOutcome = await analyzeWorkbookFiles(
    [
      { buffer: current, fileName: "A.xlsx" },
      { buffer: previous, fileName: "B.xlsx" },
    ],
    { "A.xlsx": "current", "B.xlsx": "previous" },
  );
  assert.equal(confirmedOutcome.status, "resolved", "CA-confirmed role overrides should resolve.");

  console.log("\nAll self-checks passed.");
}

main().catch((err) => {
  console.error(err);
  process.exitCode = 1;
});
