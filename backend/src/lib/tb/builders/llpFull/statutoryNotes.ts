import type ExcelJS from "exceljs";
import type { BuildCtx } from "../common";
import { setCell, applyColumnWidths, topBorder } from "../../excel/helpers";
import type { Partner } from "../partnershipFull/partners";

// Notes 27-31 equivalent, kept per spec ("still apply, unchanged"). Notes
// 29-31 are standard statutory boilerplate that doesn't depend on the trial
// balance at all, so they're written out in full (adapted to LLP wording -
// "Board" -> "Designated Partners", "Company" -> "LLP"). Notes 27 (Related
// Party) and 28 (MSME) genuinely need data a TB doesn't carry (who counts
// as related, actual transactions, which creditors are MSME-registered) -
// seeded with the detected partners as the likely related-party list
// (the LLP equivalent of Company's capitalShareholders seed) and left for
// the CA to complete.
export function writeStatutoryNotesSheet(ws: ExcelJS.Worksheet, ctx: BuildCtx, partners: Partner[]) {
  applyColumnWidths(ws, [30, 24, 20, 20, 16, 16]);
  const { meta } = ctx;

  let r = 2;
  setCell(ws, r++, 1, meta.firmName || "LLP NAME", { bold: true });
  setCell(ws, r++, 1, `Notes to Financial Statements for the year ended ${meta.periodLabel || ""}`, {
    bold: true,
  });
  r++;

  setCell(ws, r++, 1, "27 - Related Party Disclosures", { bold: true });
  setCell(
    ws,
    r++,
    1,
    "Names of related parties and description of relationship (seeded from detected partners below - confirm and add key management personnel, their relatives, and entities they control):",
    { italic: true },
  );
  const named = partners.filter((p) => p.detected);
  if (named.length) {
    for (const p of named) {
      setCell(ws, r, 1, p.name);
      setCell(ws, r++, 2, "Designated Partner / Partner (confirm)");
    }
  } else {
    setCell(ws, r++, 1, "(no individually-named partner capital ledgers found - enter related parties manually)", {
      italic: true,
    });
  }
  r++;
  setCell(ws, r, 1, "Name of the Related Party", { bold: true });
  setCell(ws, r, 2, "Transaction", { bold: true });
  setCell(ws, r, 3, `Year ended ${meta.periodLabel || ""}`, { bold: true, align: "right" });
  setCell(ws, r++, 4, "Closing Balance", { bold: true, align: "right" });
  setCell(ws, r++, 1, "(enter each related-party transaction as a row - rent, remuneration, sales, purchases, loans, etc.)", {
    italic: true,
  });
  r += 2;

  setCell(ws, r++, 1, "28 - Micro, Small and Medium Enterprises Development Act, 2006 (MSME) Disclosure", {
    bold: true,
  });
  setCell(
    ws,
    r++,
    1,
    "This app cannot determine MSME-registration status of individual creditors from a trial balance - the amounts below default to 0 and must be entered from the MSME register / vendor confirmations.",
    { italic: true },
  );
  const msmeItems = [
    "(a) Principal amount and interest due thereon remaining unpaid to any supplier as at year end",
    "(b) Interest paid under Section 16, along with the amount paid beyond the appointed day",
    "(c) Interest due and payable for delay in payment (beyond the appointed day) without adding interest under this Act",
    "(d) Interest accrued and remaining unpaid at year end",
    "(e) Interest due and payable even in succeeding years, until such interest is actually paid",
  ];
  for (const item of msmeItems) {
    setCell(ws, r, 1, item);
    setCell(ws, r++, 3, 0, { money: true });
  }
  r += 2;

  setCell(ws, r++, 1, "29 - Realizability of Current Assets", { bold: true });
  setCell(
    ws,
    r++,
    1,
    "In the opinion of the Designated Partners, the current assets, loans and advances have a value on realization in the ordinary course of business at least equal to the amount at which they are stated in the Balance Sheet, unless otherwise stated, and provision for all known liabilities has been made.",
  );
  r += 2;

  setCell(ws, r++, 1, "30 - Other Statutory Information", { bold: true });
  const checklist = [
    "(i) The LLP does not have any benami property, and no proceedings have been initiated or pending against the LLP for holding any benami property.",
    "(ii) The LLP has not traded or invested in Crypto currency or Virtual Currency during the financial year.",
    "(iii) The LLP has not advanced or loaned or invested funds to any other person(s) or entity(ies), including foreign entities (Intermediaries), with the understanding that the Intermediary shall lend or invest in party identified by or on behalf of the LLP (Ultimate Beneficiaries).",
    "(iv) The LLP has not received any fund from any person(s) or entity(ies), including foreign entities (Funding Party), with the understanding that the LLP shall lend or invest in party identified by or on behalf of the Funding Party (Ultimate Beneficiaries).",
    "(v) The LLP does not have any such transaction which is not recorded in the books of accounts that has been surrendered or disclosed as income during the year in the tax assessments under the Income Tax Act, 1961.",
    "(vi) The LLP has not been declared a wilful defaulter by any bank or financial institution or other lender.",
    "(vii) No registration or satisfaction of charges is pending to be filed with the Registrar of Companies beyond the statutory period.",
  ];
  for (const item of checklist) {
    setCell(ws, r++, 1, item);
  }
  r += 2;

  setCell(ws, r++, 1, "31 - Regrouping / Reclassification", { bold: true });
  setCell(
    ws,
    r++,
    1,
    "Previous year's figures have been regrouped / reclassified wherever necessary to conform to the current year's presentation.",
  );

  topBorder(ws, r, 1, 6);
}
