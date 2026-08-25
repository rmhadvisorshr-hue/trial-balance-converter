import type ExcelJS from "exceljs";
import type { StatementCode } from "../../types";
import type { BuildCtx } from "../common";
import { setCell, setFormula, applyColumnWidths, topBorder } from "../../excel/helpers";
import { round2 } from "../../classify";

// Schedule III Notes 4-10 & 12-26. Notes 2 (Share Capital) and 3 (Reserves)
// live on their own sheet (shareCapitalReserves.ts); Note 11 (PPE) lives on
// 11. FA. Note 5 (Other Long Term Liabilities) has no matching StatementCode
// in this app's taxonomy (it only distinguishes long-term borrowings, not
// other long-term liabilities/provisions), so it's emitted empty.
export interface NoteRef {
  row: number; // the Total row for this note on the NOTES sheet
}

export interface NotesResult {
  byNote: Map<number, NoteRef>;
}

export interface NoteDef {
  no: number;
  title: string;
  codes: StatementCode[];
  // Cost of materials / changes in inventory need the codes summed instead
  // of listed as separate line items (opening - closing stock nets out).
  net?: boolean;
}

// Exported so the PDF renderer can reuse the exact same note-to-code
// mapping instead of maintaining a second, divergence-prone copy.
export const LIABILITY_NOTES: NoteDef[] = [
  { no: 4, title: "LONG-TERM BORROWINGS", codes: ["LONG_TERM_BORROW"] },
  { no: 5, title: "OTHER LONG TERM LIABILITIES", codes: [] },
  { no: 6, title: "DEFERRED TAX LIABILITIES", codes: ["DEFERRED_TAX_LIAB"] },
  { no: 7, title: "SHORT TERM BORROWINGS", codes: ["SHORT_TERM_BORROW"] },
  { no: 8, title: "TRADE PAYABLES", codes: ["TRADE_PAYABLES"] },
  { no: 9, title: "OTHER CURRENT LIABILITIES", codes: ["OTHER_CURR_LIAB"] },
  { no: 10, title: "SHORT TERM PROVISIONS", codes: ["SHORT_TERM_PROV"] },
];

export const ASSET_NOTES: NoteDef[] = [
  { no: 12, title: "NON CURRENT INVESTMENTS", codes: ["NONCURR_INVEST"] },
  { no: 13, title: "DEFERRED TAX ASSET", codes: ["DEFERRED_TAX_ASSET"] },
  { no: 14, title: "INVENTORY", codes: ["INVENTORY"] },
  { no: 15, title: "TRADE RECEIVABLES", codes: ["TRADE_RECV"] },
  { no: 16, title: "CASH & BANK BALANCES", codes: ["CASH_BANK"] },
  { no: 17, title: "SHORT TERM LOANS AND ADVANCES", codes: ["SHORT_TERM_LOANS_ADV"] },
  { no: 18, title: "OTHER CURRENT ASSETS", codes: ["OTHER_CURR_ASSETS"] },
];

export const PL_NOTES: NoteDef[] = [
  { no: 19, title: "REVENUE FROM OPERATIONS", codes: ["REV_OPS"] },
  { no: 20, title: "OTHER INCOME", codes: ["OTHER_INCOME"] },
  { no: 21, title: "COST OF MATERIALS CONSUMED", codes: ["PURCHASES", "DIRECT_EXP"] },
  {
    no: 22,
    title: "CHANGES IN INVENTORY",
    codes: ["OPENING_STOCK", "CLOSING_STOCK_PL"],
    net: true,
  },
  { no: 23, title: "EMPLOYEE BENEFIT EXPENSES", codes: ["EMP_BENEFIT"] },
  { no: 24, title: "FINANCE COSTS", codes: ["FINANCE_COST"] },
  { no: 25, title: "OTHER EXPENSES", codes: ["OTHER_EXP"] },
];

export function writeNotesSheet(ws: ExcelJS.Worksheet, ctx: BuildCtx): NotesResult {
  applyColumnWidths(ws, [46, 20, 20]);
  const { meta, st, divisor } = ctx;
  const money = (v: number) => v / divisor;
  const byNote = new Map<number, NoteRef>();
  let r = 2;

  setCell(ws, r++, 1, meta.firmName || "COMPANY NAME", { bold: true });
  if (meta.cin) setCell(ws, r++, 1, `CIN : ${meta.cin}`, {});
  setCell(ws, r++, 1, `Notes to Financial Statements for the year ended ${meta.periodLabel || ""}`, {
    bold: true,
  });
  if (ctx.unitHeading) setCell(ws, r++, 1, ctx.unitHeading, { italic: true });
  r++;

  function writeNote(def: NoteDef) {
    setCell(ws, r++, 1, `Note "${def.no}" : ${def.title}`, { bold: true });
    setCell(ws, r, 1, "Particulars", { bold: true });
    setCell(ws, r, 2, `As at ${meta.asAtLabel || meta.periodLabel || ""}`, {
      bold: true,
      align: "right",
    });
    setCell(ws, r++, 3, "As at (prior year)", { bold: true, align: "right" });

    if (def.codes.length === 0) {
      setCell(ws, r, 1, "-");
      setCell(ws, r, 2, money(0), { money: true });
      const totalRow = r;
      byNote.set(def.no, { row: totalRow });
      r += 2;
      return;
    }

    const startRow = r;
    if (def.net) {
      // Opening stock (debit) less Closing stock (credit-natured in the trading a/c).
      const opening = st.total("OPENING_STOCK");
      const closing = st.total("CLOSING_STOCK_PL");
      setCell(ws, r, 1, "Opening stock");
      setCell(ws, r++, 2, money(opening), { money: true });
      setCell(ws, r, 1, "Less: Closing stock");
      setCell(ws, r++, 2, money(-closing), { money: true });
    } else {
      for (const code of def.codes) {
        for (const item of st.byCode.get(code) ?? []) {
          setCell(ws, r, 1, item.name);
          setCell(ws, r++, 2, money(item.amount), { money: true });
        }
      }
    }
    // Guarantee at least one backing row inside the SUM range before adding
    // any cosmetic rows below it, so the range's end never drifts onto a
    // text-only row (e.g. the MSME disclaimer).
    if (r - 1 < startRow) {
      setCell(ws, r, 1, "-");
      setCell(ws, r++, 2, money(0), { money: true });
    }
    const endRow = r - 1;

    if (def.no === 8) {
      // Trade Payables: this app doesn't determine MSME status from a TB -
      // the whole balance defaults to "Others" and is flagged for the CA to
      // reclassify, matching the format's (i)/(ii) split.
      setCell(ws, r, 1, "(reclassify above into MSME / Others per the MSME Act - defaults to Others)", {
        italic: true,
      });
      r++;
    }

    r++;
    setCell(ws, r, 1, "Total", { bold: true });
    const total = def.net
      ? round2(st.total("OPENING_STOCK") - st.total("CLOSING_STOCK_PL"))
      : round2(def.codes.reduce((s, c) => s + st.total(c), 0));
    setFormula(ws, r, 2, `SUM(B${startRow}:B${endRow})`, money(total), {
      bold: true,
      money: true,
    });
    topBorder(ws, r, 1, 3);
    byNote.set(def.no, { row: r });
    r += 2;
  }

  for (const def of LIABILITY_NOTES) writeNote(def);
  for (const def of ASSET_NOTES) writeNote(def);
  for (const def of PL_NOTES) writeNote(def);

  return { byNote };
}
