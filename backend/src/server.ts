import path from "node:path";
import { fileURLToPath } from "node:url";
import express, { type ErrorRequestHandler } from "express";
import multer, { MulterError } from "multer";
import cors from "cors";

import { parseAndClassify, buildWorkbook, buildFinancialStatementsPdf } from "./lib/tb/pipeline";
import type { ConvertPayload, EntityType } from "./lib/tb/types";
import { analyzeWorkbookFiles, buildIcaiWorkbook, scaleAnalysis } from "./lib/icai/pipeline";
import type { FileRole } from "./lib/icai/parsers/fileRoleDetector";
import type { FiguresUnit } from "./lib/icai/excel/scale";
import type { IcaiEntityKind, WorkbookAnalysis } from "./lib/icai/types";
import {
  listCAProfiles,
  createCAProfile,
  updateCAProfile,
  deleteCAProfile,
  validateCAProfileInput,
  hasRequiredCaFields,
  CA_REQUIRED_MESSAGE,
  type CAProfileInput,
} from "./lib/caProfiles";

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

// Only enabled when the frontend is deployed on a different origin (e.g. a
// Vercel-hosted SPA calling this backend on its own domain) - unset in the
// embedded/same-origin deployment (see the static-file serving below), which
// needs no CORS headers at all. Comma-separated so both a production and a
// Vercel preview-deployment origin can be allowed at once.
const allowedOrigins = (process.env.FRONTEND_ORIGIN ?? "")
  .split(",")
  .map((s) => s.trim())
  .filter(Boolean);
if (allowedOrigins.length > 0) {
  app.use(cors({ origin: allowedOrigins }));
}

app.use(express.json({ limit: "5mb" }));

app.get("/health", (_req, res) => {
  res.json({ status: "ok" });
});

// Centralized CA profile store - shared by both the Trial Balance and
// Accounting Workbook frontends (see lib/caProfiles.ts).
app.get("/api/ca-profiles", async (_req, res) => {
  res.json(await listCAProfiles());
});

app.post("/api/ca-profiles", async (req, res) => {
  const input = req.body as Partial<CAProfileInput>;
  const validationError = validateCAProfileInput(input);
  if (validationError) {
    res.status(400).json({ code: "VALIDATION_ERROR", message: validationError });
    return;
  }
  const profile = await createCAProfile(input as CAProfileInput);
  res.status(201).json(profile);
});

app.put("/api/ca-profiles/:id", async (req, res) => {
  const input = req.body as Partial<CAProfileInput>;
  const validationError = validateCAProfileInput(input);
  if (validationError) {
    res.status(400).json({ code: "VALIDATION_ERROR", message: validationError });
    return;
  }
  const profile = await updateCAProfile(req.params.id, input as CAProfileInput);
  if (!profile) {
    res.status(404).json({ code: "NOT_FOUND", message: "CA profile not found." });
    return;
  }
  res.json(profile);
});

app.delete("/api/ca-profiles/:id", async (req, res) => {
  const deleted = await deleteCAProfile(req.params.id);
  if (!deleted) {
    res.status(404).json({ code: "NOT_FOUND", message: "CA profile not found." });
    return;
  }
  res.status(204).end();
});

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

    // Excel output only: a CA / Signing Authority must be selected so the
    // generated signature block never falls back to another CA's details -
    // see lib/caProfiles.ts and excel/helpers.ts's signatureBlock.
    if (!hasRequiredCaFields(payload.meta || {})) {
      res.status(400).json({ code: "CA_REQUIRED", message: CA_REQUIRED_MESSAGE });
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
    // A CA / Signing Authority must be selected - same gate as /api/tbconvert.
    if (!hasRequiredCaFields(analysis)) {
      res.status(400).json({ code: "CA_REQUIRED", message: CA_REQUIRED_MESSAGE });
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
