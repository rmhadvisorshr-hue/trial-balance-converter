import type ExcelJS from "exceljs";
import type { StatementCode } from "../../types";
import type { BuildCtx } from "../common";
import { setCell, setFormula, applyColumnWidths, topBorder } from "../../excel/helpers";
import { round2 } from "../../classify";

export interface SchResult {
  schRow: Map<number, number>; // schedule number -> its Total row on FIRM SCH
  investmentsRow: number | null; // Investments section Total row, or null if empty
}

interface SchDef {
  no: number;
  title: string;
  codes: StatementCode[];
}

// Secured Loans (SHORT_TERM_BORROW, typically a bank OD/CC) is shown
// directly on the Balance Sheet, not through a schedule - matches the
// reference file's own convention and the Pvt Ltd/Partnership precedent.
const SCHEDULES: SchDef[] = [
  { no: 1, title: "UNSECURED LOANS", codes: ["LONG_TERM_BORROW"] },
  { no: 2, title: "SUNDRY CREDITORS", codes: ["TRADE_PAYABLES"] },
  { no: 3, title: "OTHER CURRENT LIABILITIES", codes: ["OTHER_CURR_LIAB", "SHORT_TERM_PROV", "DEFERRED_TAX_LIAB"] },
  { no: 4, title: "LOANS AND ADVANCES", codes: ["SHORT_TERM_LOANS_ADV", "LONG_TERM_LOANS_ADV"] },
  { no: 5, title: "OTHER CURRENT ASSETS", codes: ["INVENTORY", "OTHER_CURR_ASSETS", "DEFERRED_TAX_ASSET"] },
  { no: 6, title: "SUNDRY DEBTORS", codes: ["TRADE_RECV"] },
];

// 'FIRM SCH': the business schedules referenced by 'FIRM CAP&BS' - same
// StatementCode-mapped, dynamic-row-tracking pattern as
// partnershipFull/schedules.ts (this app doesn't split Bank Balances into
// its own schedule the way the reference does - Cash and Bank stay one
// known-TB-balance BS line, consistent with every other entity type here).
export function writeSchedulesSheet(ws: ExcelJS.Worksheet, ctx: BuildCtx): SchResult {
  applyColumnWidths(ws, [44, 20]);
  const { meta, st, divisor } = ctx;
  const money = (v: number) => v / divisor;
  const schRow = new Map<number, number>();
  let r = 2;

  setCell(ws, r++, 1, meta.firmName || "FIRM NAME", { bold: true });
  setCell(ws, r++, 1, meta.periodLabel || "", { italic: true });
  if (ctx.unitHeading) setCell(ws, r++, 1, ctx.unitHeading, { italic: true });
  r++;

  for (const def of SCHEDULES) {
    setCell(ws, r++, 1, `Sch ${def.no}  ${def.title}`, { bold: true });
    setCell(ws, r, 1, "Particulars", { bold: true });
    setCell(ws, r++, 2, "Amount", { bold: true, align: "right" });

    const startRow = r;
    for (const code of def.codes) {
      for (const item of st.byCode.get(code) ?? []) {
        setCell(ws, r, 1, item.name);
        setCell(ws, r++, 2, money(item.amount), { money: true });
      }
    }
    if (r === startRow) {
      setCell(ws, r, 1, "-");
      setCell(ws, r++, 2, money(0), { money: true });
    }
    const endRow = r - 1;

    r++;
    setCell(ws, r, 1, "Total", { bold: true });
    const total = round2(def.codes.reduce((s, c) => s + st.total(c), 0));
    setFormula(ws, r, 2, `SUM(B${startRow}:B${endRow})`, money(total), { bold: true, money: true });
    topBorder(ws, r, 1, 2);
    schRow.set(def.no, r);
    r += 2;
  }

  let investmentsRow: number | null = null;
  const investItems = st.byCode.get("NONCURR_INVEST") ?? [];
  if (investItems.length) {
    setCell(ws, r++, 1, "INVESTMENTS AND ADVANCES", { bold: true });
    setCell(ws, r, 1, "Particulars", { bold: true });
    setCell(ws, r++, 2, "Amount", { bold: true, align: "right" });
    const startRow = r;
    for (const item of investItems) {
      setCell(ws, r, 1, item.name);
      setCell(ws, r++, 2, money(item.amount), { money: true });
    }
    const endRow = r - 1;
    r++;
    setCell(ws, r, 1, "Total", { bold: true });
    setFormula(ws, r, 2, `SUM(B${startRow}:B${endRow})`, money(round2(st.total("NONCURR_INVEST"))), {
      bold: true,
      money: true,
    });
    topBorder(ws, r, 1, 2);
    investmentsRow = r;
  }

  return { schRow, investmentsRow };
}
