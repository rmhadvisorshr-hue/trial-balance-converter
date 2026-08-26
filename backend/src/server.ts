import path from "node:path";
import { fileURLToPath } from "node:url";
import express, { type ErrorRequestHandler } from "express";
import multer, { MulterError } from "multer";

import { parseAndClassify, buildWorkbook, buildFinancialStatementsPdf } from "./lib/tb/pipeline";
import type { ConvertPayload, EntityType } from "./lib/tb/types";
import { analyzeWorkbookFiles, buildIcaiWorkbook, scaleAnalysis } from "./lib/icai/pipeline";
import type { FileRole } from "./lib/icai/parsers/fileRoleDetector";
import type { FiguresUnit } from "./lib/icai/excel/scale";
import type { IcaiEntityKind, WorkbookAnalysis } from "./lib/icai/types";

const PORT = Number(process.env.PORT) || 8097;
const MAX_UPLOAD_BYTES = 25 * 1024 * 1024;
const dirname = path.dirname(fileURLToPath(import.meta.url));

const upload = multer({ storage: multer.memoryStorage(), limits: { fileSize: MAX_UPLOAD_BYTES } });

function errorPayload(error: unknown) {
  const e = error as Partial<Error & { code?: string }>;
  return {
    code: e.code || "CONVERSION_FAILED",
    message:
      typeof e.message === "string" && e.message.trim()
        ? e.message
        : "Could not process this trial balance.",
  };
}

function sanitizeOutputName(name: string, extension: "xlsx" | "pdf") {
  const base = name
    .replace(/\.[^.]+$/, "")
    .replace(/[<>:"/\\|?*\r\n]/g, "")
    .trim();
  return `${base || "financials"} - Financials.${extension}`;
}

// Unlike sanitizeOutputName above, the name here is an entity name (e.g.
// "M/S SPACE HOME (A.G)"), not an uploaded filename - it has no extension to
// strip, and its embedded periods must not be mistaken for one.
function sanitizeEntityFileName(name: string, extension: "xlsx") {
  const base = name.replace(/[<>:"/\\|?*\r\n]/g, "").trim();
  return `${base || "financials"} - Financials.${extension}`;
}

const app = express();
app.use(express.json({ limit: "5mb" }));

app.post("/api/tbparse", upload.single("file"), async (req, res) => {
  try {
    if (!req.file) {
      res.status(400).json({ code: "FILE_REQUIRED", message: "Upload a trial balance file." });
      return;
    }
    const entity = String(req.body.entity || "partnership") as EntityType;
    const result = await parseAndClassify(req.file.buffer, req.file.originalname, entity);
    res.json(result);
  } catch (error) {
    console.error(error);
    res.status(400).json(errorPayload(error));
  }
});

app.post("/api/tbconvert", async (req, res) => {
  try {
    const payload = req.body as ConvertPayload;
    if (!payload || !Array.isArray(payload.ledgers)) {
      res.status(400).json({ code: "BAD_PAYLOAD", message: "Invalid conversion payload." });
      return;
    }
    const firmName = payload.meta?.firmName || "financials";

    if (payload.outputFormat === "pdf") {
      const buffer = await buildFinancialStatementsPdf(payload);
      res.setHeader("Content-Type", "application/pdf");
      res.setHeader(
        "Content-Disposition",
        `attachment; filename="${sanitizeOutputName(firmName, "pdf")}"`,
      );
      res.send(buffer);
      return;
    }

    const buffer = await buildWorkbook(payload);
    res.setHeader(
      "Content-Type",
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    );
    res.setHeader(
      "Content-Disposition",
      `attachment; filename="${sanitizeOutputName(firmName, "xlsx")}"`,
    );
    res.send(buffer);
  } catch (error) {
    console.error(error);
    res.status(400).json(errorPayload(error));
  }
});

app.post("/api/icai/analyze", upload.array("files"), async (req, res) => {
  try {
    const uploaded = (req.files as Express.Multer.File[] | undefined) ?? [];
    if (uploaded.length === 0) {
      res.status(400).json({ code: "FILE_REQUIRED", message: "Upload one or more accounting workbooks." });
      return;
    }
    // Both optional: sent by the review step's "Confirm & Continue" once the
    // CA has resolved an ambiguous role-detection / corrected the entity kind.
    const roleOverrides = req.body.roleOverrides
      ? (JSON.parse(req.body.roleOverrides) as Record<string, FileRole>)
      : undefined;
    const entityKindOverride = req.body.entityKind ? (req.body.entityKind as IcaiEntityKind) : undefined;

    const outcome = await analyzeWorkbookFiles(
      uploaded.map((f) => ({ buffer: f.buffer, fileName: f.originalname })),
      roleOverrides,
      entityKindOverride,
    );
    res.json(outcome);
  } catch (error) {
    console.error(error);
    res.status(400).json(errorPayload(error));
  }
});

app.post("/api/icai/generate", async (req, res) => {
  try {
    const body = req.body as { analysis: WorkbookAnalysis; figuresUnit?: FiguresUnit };
    const analysis = body?.analysis;
    if (!analysis || !Array.isArray(analysis.accounts)) {
      res.status(400).json({ code: "BAD_PAYLOAD", message: "Invalid analysis payload." });
      return;
    }
    const scaled = scaleAnalysis(analysis, body.figuresUnit || "actual");
    const buffer = await buildIcaiWorkbook(scaled);
    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    res.setHeader(
      "Content-Disposition",
      `attachment; filename="${sanitizeEntityFileName(analysis.entityName || "financials", "xlsx")}"`,
    );
    res.send(buffer);
  } catch (error) {
    console.error(error);
    res.status(400).json(errorPayload(error));
  }
});

// When embedded in the unified staff portal (see
// CA-StaffPortal/scripts/unified-preview.mjs), this backend also serves the
// built frontend as static files -- same pattern as the 2A-2B and
// GSTR2B-vs-Books sidecars in that repo.
app.use(express.static(path.join(dirname, "..", "..", "frontend", "dist")));

const handleUploadError: ErrorRequestHandler = (err, _req, res, next) => {
  if (err instanceof MulterError && err.code === "LIMIT_FILE_SIZE") {
    res.status(400).json({ code: "FILE_TOO_LARGE", message: "Files must be 25 MB or smaller." });
    return;
  }
  next(err);
};
app.use(handleUploadError);

app.listen(PORT, "0.0.0.0", () => {
  console.log(`[tb-converter backend] listening on ${PORT}`);
});
