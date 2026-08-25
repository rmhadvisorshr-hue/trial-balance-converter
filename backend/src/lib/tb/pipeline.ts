import { parseTrialBalance } from "./parser";
import { classifyLedgers } from "./classify";
import type { EntityType, ParseResult } from "./types";

export async function parseAndClassify(
  buffer: ArrayBuffer | Buffer,
  fileName: string,
  entity: EntityType,
): Promise<ParseResult> {
  const raw = await parseTrialBalance(buffer, fileName);
  const ledgers = classifyLedgers(raw.leaves, entity);
  return {
    entity,
    meta: raw.meta,
    ledgers,
    warnings: raw.warnings,
  };
}

export { buildWorkbook } from "./builders";
export { buildFinancialStatementsPdf } from "./builders/pdf";
