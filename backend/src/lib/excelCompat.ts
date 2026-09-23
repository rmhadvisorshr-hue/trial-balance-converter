import ExcelJS from "exceljs";
import * as XLSX from "xlsx";

const OLE2_MAGIC = Buffer.from([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1]);

function isLegacyXls(buffer: Buffer): boolean {
  return buffer.length >= 8 && buffer.subarray(0, 8).equals(OLE2_MAGIC);
}

// Tally and old Excel installs still export the binary "Excel 97-2003" (.xls)
// format - ExcelJS only reads OOXML (.xlsx), so it 400s on those with an
// opaque "central directory" error. Detect by magic bytes (not extension -
// uploads get renamed) and re-encode through SheetJS, which reads both.
export async function loadExcelWorkbook(buffer: Buffer | ArrayBuffer): Promise<ExcelJS.Workbook> {
  const buf = Buffer.isBuffer(buffer) ? buffer : Buffer.from(buffer);
  const wb = new ExcelJS.Workbook();
  if (isLegacyXls(buf)) {
    const legacy = XLSX.read(buf, { type: "buffer" });
    const converted = XLSX.write(legacy, { type: "buffer", bookType: "xlsx" }) as Buffer;
    await wb.xlsx.load(converted as unknown as ArrayBuffer);
  } else {
    await wb.xlsx.load(buf as unknown as ArrayBuffer);
  }
  return wb;
}
