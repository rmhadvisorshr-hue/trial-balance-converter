import type ExcelJS from "exceljs";
import type { WorkbookAnalysis, OwnerCapitalAccount } from "../types";
import { setCell, setFormula, topBorder, applyColumnWidths } from "../../tb/excel/helpers";
import { round2 } from "./common";

const CAPITAL_HEADERS = [
  "Sr. No.",
  "Name of Partner/Proprietor",
  "Share of Profit/(Loss) (%)",
  "Opening Balance",
  "Capital Introduced during the year",
  "Remuneration for the year",
  "Interest for the year",
  "Withdrawals during the year",
  "Share of Profit/(Loss) for the year",
  "Closing Balance",
];

function writeCapitalTable(
  ws: ExcelJS.Worksheet,
  startRow: number,
  yearLabel: string,
  owners: OwnerCapitalAccount[],
  which: "current" | "previous",
): number {
  let r = startRow;
  setCell(ws, r++, 1, `Partners' Capital Account - year ended ${yearLabel}`, { bold: true });
  CAPITAL_HEADERS.forEach((h, i) => setCell(ws, r, i + 1, h, { bold: true, align: i <= 1 ? "left" : "right", wrap: true }));
  r++;

  const startData = r;
  owners.forEach((o, i) => {
    const opening = which === "current" ? o.openingBalance : o.openingBalancePrevious;
    const introduced = which === "current" ? o.capitalIntroduced : o.capitalIntroducedPrevious;
    const remuneration = which === "current" ? o.remuneration : o.remunerationPrevious;
    const interest = which === "current" ? o.interest : o.interestPrevious;
    const withdrawals = which === "current" ? o.withdrawals : o.withdrawalsPrevious;
    const shareOfProfit = which === "current" ? o.shareOfProfit : o.shareOfProfitPrevious;
    const closing = which === "current" ? o.closingBalance : o.closingBalancePrevious;

    setCell(ws, r, 1, i + 1);
    setCell(ws, r, 2, o.name);
    setCell(ws, r, 3, o.sharePercent ?? "", { align: "right" });
    setCell(ws, r, 4, round2(opening), { money: true });
    setCell(ws, r, 5, round2(introduced), { money: true });
    setCell(ws, r, 6, round2(remuneration), { money: true });
    setCell(ws, r, 7, round2(interest), { money: true });
    setCell(ws, r, 8, round2(withdrawals), { money: true });
    setCell(ws, r, 9, round2(shareOfProfit), { money: true });
    setCell(ws, r++, 10, round2(closing), { money: true });
  });
  const endData = r - 1;

  setCell(ws, r, 2, "Total", { bold: true });
  for (const col of [4, 5, 6, 7, 8, 9, 10]) {
    const letter = String.fromCharCode(64 + col);
    const known = round2(owners.reduce((s, o) => {
      const key = (
        {
          4: which === "current" ? "openingBalance" : "openingBalancePrevious",
          5: which === "current" ? "capitalIntroduced" : "capitalIntroducedPrevious",
          6: which === "current" ? "remuneration" : "remunerationPrevious",
          7: which === "current" ? "interest" : "interestPrevious",
          8: which === "current" ? "withdrawals" : "withdrawalsPrevious",
          9: which === "current" ? "shareOfProfit" : "shareOfProfitPrevious",
          10: which === "current" ? "closingBalance" : "closingBalancePrevious",
        } as const
      )[col as 4 | 5 | 6 | 7 | 8 | 9 | 10];
      return s + (o[key] as number);
    }, 0));
    setFormula(ws, r, col, `SUM(${letter}${startData}:${letter}${endData})`, known, { bold: true, money: true });
  }
  topBorder(ws, r, 1, 10);
  r += 2;

  if (which === "current") {
    const withBreakup = owners.filter((o) => o.withdrawalBreakup.length > 1);
    if (withBreakup.length > 0) {
      setCell(ws, r++, 1, "* Withdrawals during the year include the following:", { italic: true });
      setCell(ws, r, 2, "Name of Partner", { bold: true });
      const labels = [...new Set(withBreakup.flatMap((o) => o.withdrawalBreakup.map((w) => w.label)))];
      labels.forEach((l, i) => setCell(ws, r, 3 + i, l, { bold: true, align: "right" }));
      r++;
      for (const o of owners) {
        setCell(ws, r, 2, o.name);
        labels.forEach((l, i) => {
          const amt = o.withdrawalBreakup.find((w) => w.label === l)?.amount ?? 0;
          setCell(ws, r, 3 + i, round2(amt), { money: true });
        });
        r++;
      }
      r++;
    }
  }

  return r;
}

export function writeNotes1to3(ws: ExcelJS.Worksheet, analysis: WorkbookAnalysis) {
  applyColumnWidths(ws, [46, 20, 16, 16, 16, 16, 16, 16, 16, 16]);
  let r = 1;
  setCell(ws, r++, 1, analysis.entityName || "ENTITY NAME", { bold: true });
  setCell(ws, r++, 1, `Notes forming part of the Financial Statements for the year ended ${analysis.currentYearLabel}`, { bold: true });
  r++;

  setCell(ws, r++, 1, "Note 1: Brief about the entity", { bold: true });
  setCell(
    ws,
    r++,
    1,
    `[Insert a brief description of ${analysis.entityName || "the entity"}'s constitution, registration and principal business activity - not derivable from the accounting workbook alone.]`,
    { italic: true },
  );
  r++;

  setCell(ws, r++, 1, "Note 2: Significant Accounting Policies", { bold: true });
  setCell(ws, r++, 1, "(a) Basis of preparation", { bold: true });
  setCell(
    ws,
    r++,
    1,
    "The financial statements have been prepared on accrual basis of accounting under the historical cost convention, in accordance with the generally accepted accounting principles in India and presented in the format recommended in the ICAI's Guidance Note on Financial Statements of Non-Corporate Entities.",
  );
  r++;
  if (analysis.fixedAssets.current.length > 0 || analysis.fixedAssets.previous.length > 0) {
    setCell(ws, r++, 1, "(b) Property, Plant and Equipment and Depreciation", { bold: true });
    setCell(
      ws,
      r++,
      1,
      "Fixed assets are stated at cost of acquisition less accumulated depreciation. Depreciation is provided on the Written Down Value (WDV) method at the rates recorded in the entity's own fixed asset schedule.",
    );
    r++;
  }
  const hasWip = analysis.accounts.some((a) => a.code === "N15_INVENTORY" && a.sourceSheet.toLowerCase().includes("work in progress"));
  setCell(ws, r++, 1, "(c) Inventories", { bold: true });
  setCell(
    ws,
    r++,
    1,
    hasWip
      ? "Work-in-progress is carried at the value determined by management as recorded in the books - see the Basis of Preparation sheet for the accounting-policy matter this raises."
      : "[Describe the entity's inventory valuation policy.]",
    hasWip ? {} : { italic: true },
  );
  r++;
  if (analysis.owners.some((o) => o.remuneration || o.remunerationPrevious)) {
    setCell(ws, r++, 1, "(d) Partners' remuneration and interest on capital", { bold: true });
    setCell(
      ws,
      r++,
      1,
      "Remuneration to working partners is provided in the books in accordance with the partnership deed and section 40(b) of the Income-tax Act, 1961, and is presented in the Statement of Profit and Loss as an appropriation of profit.",
    );
    r++;
  }
  setCell(ws, r++, 1, "(e) Taxes on income", { bold: true });
  setCell(ws, r++, 1, "Provision for current tax is made on the taxable income for the year, computed in accordance with the Income-tax Act, 1961.");
  r++;
  setCell(ws, r++, 1, "(f) Cash and cash equivalents", { bold: true });
  setCell(ws, r++, 1, "Cash and cash equivalents comprise cash on hand and balances with banks in current/savings accounts.");
  r += 2;

  setCell(ws, r++, 1, "Note 3: Owners'/Partners' Capital Account", { bold: true });
  r = writeCapitalTable(ws, r, analysis.currentYearLabel, analysis.owners, "current");
  if (analysis.previousYearLabel) {
    r = writeCapitalTable(ws, r, analysis.previousYearLabel, analysis.owners, "previous");
  }

  setCell(ws, r++, 1, "Note 3(b): Owners'/Partners' Current Account", { bold: true });
  setCell(
    ws,
    r++,
    1,
    "NIL - all capital-related transactions of the owners are routed through the Capital Account (Note 3(a)) above.",
    { italic: true },
  );
}
