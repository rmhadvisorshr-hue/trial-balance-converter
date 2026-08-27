import { useMemo, useState } from "react";
import { FileSpreadsheet, Upload, Play, Loader2, X } from "lucide-react";
import type { EntityType, FiguresUnit } from "./lib/types";
import {
  ICAI_NOTE_CODES,
  ICAI_NOTE_LABEL,
  type FileRole,
  type FileRoleDetection,
  type IcaiEntityKind,
  type IcaiNoteCode,
  type WorkbookAnalysis,
} from "./lib/icai-types";
import type { CAProfile } from "./lib/ca-types";
import { analyzeIcaiWorkbooks, generateIcaiWorkbook } from "./lib/icai-client";
import ErrorBanner from "./components/ErrorBanner";
import EntityTypeSelect from "./components/EntityTypeSelect";
import FiguresUnitSelect from "./components/FiguresUnitSelect";
import CASelect from "./components/CASelect";
import TextField from "./components/TextField";

const ICAI_ENTITY_OPTIONS: EntityType[] = ["partnership", "llp", "proprietor"];
const ROLE_LABELS: Record<FileRole, string> = { current: "Current Year", previous: "Previous Year", exclude: "Exclude" };

const CA_REQUIRED_MESSAGE = "Please select a CA / Signing Authority before generating the report.";

interface Props {
  entity: EntityType;
  onEntityChange: (v: EntityType) => void;
  figuresUnit: FiguresUnit;
  onFiguresUnitChange: (v: FiguresUnit) => void;
  caProfiles: CAProfile[];
}

export default function IcaiWorkflow({ entity, onEntityChange, figuresUnit, onFiguresUnitChange, caProfiles }: Props) {
  const [files, setFiles] = useState<File[]>([]);
  const [pendingReview, setPendingReview] = useState<{ reason: string; detections: FileRoleDetection[] } | null>(null);
  const [roleChoices, setRoleChoices] = useState<Record<string, FileRole>>({});
  const [analysis, setAnalysis] = useState<WorkbookAnalysis | null>(null);
  const [busy, setBusy] = useState<"idle" | "analyzing" | "generating">("idle");
  const [error, setError] = useState<string | null>(null);

  // Report Details - this pipeline had no such fields before; they exist now
  // purely to carry the report-specific place/date/UDIN and selected CA into
  // the signature block the generated Excel now includes (see
  // backend/src/lib/icai/excel/{balanceSheet,profitAndLoss}.ts).
  const [reportPlace, setReportPlace] = useState("");
  const [reportDate, setReportDate] = useState("");
  const [reportUdin, setReportUdin] = useState("");
  const [selectedCaId, setSelectedCaId] = useState("");
  const [placeTouched, setPlaceTouched] = useState(false);

  function handleSelectCa(id: string) {
    setSelectedCaId(id);
    const profile = caProfiles.find((p) => p.id === id);
    if (profile && !placeTouched) {
      setReportPlace(profile.place);
    }
  }

  // ICAI's Guidance Note only covers non-corporate entities; if the shared
  // entity-type value is still at the Trial Balance tab's "pvtltd" default,
  // show a sensible entity-appropriate default here without touching that
  // shared value until the CA actually picks something in this tab.
  const icaiEntity: EntityType = entity === "pvtltd" ? "partnership" : entity;

  const reviewCount = useMemo(
    () => (analysis ? analysis.accounts.filter((a) => !a.confident || a.code === "UNMAPPED").length : 0),
    [analysis],
  );

  function addFiles(list: FileList | null) {
    if (!list) return;
    const picked = Array.from(list);
    setFiles((prev) => {
      const merged = [...prev];
      for (const f of picked) {
        if (!merged.some((m) => m.name === f.name && m.size === f.size)) merged.push(f);
      }
      return merged;
    });
  }

  function removeFile(name: string) {
    setFiles((prev) => prev.filter((f) => f.name !== name));
  }

  async function handleAnalyze() {
    if (files.length === 0) {
      setError("Choose one or more accounting workbooks first.");
      return;
    }
    setError(null);
    setAnalysis(null);
    setPendingReview(null);
    setBusy("analyzing");
    try {
      const entityKind: IcaiEntityKind = icaiEntity === "pvtltd" ? "partnership" : icaiEntity;
      const outcome = await analyzeIcaiWorkbooks(files, undefined, entityKind);
      if (outcome.status === "needs-review") {
        setPendingReview({ reason: outcome.reason, detections: outcome.detections });
        setRoleChoices(Object.fromEntries(outcome.detections.map((d) => [d.fileName, d.role])));
      } else {
        setAnalysis(outcome.analysis);
      }
    } catch (e) {
      setError(e instanceof Error ? e.message : "Failed to read the accounting workbook(s).");
    } finally {
      setBusy("idle");
    }
  }

  async function handleConfirmRoles() {
    setError(null);
    setBusy("analyzing");
    try {
      const entityKind: IcaiEntityKind = icaiEntity === "pvtltd" ? "partnership" : icaiEntity;
      const outcome = await analyzeIcaiWorkbooks(files, roleChoices, entityKind);
      if (outcome.status === "needs-review") {
        setPendingReview({ reason: outcome.reason, detections: outcome.detections });
      } else {
        setPendingReview(null);
        setAnalysis(outcome.analysis);
      }
    } catch (e) {
      setError(e instanceof Error ? e.message : "Failed to read the accounting workbook(s).");
    } finally {
      setBusy("idle");
    }
  }

  async function handleGenerate() {
    if (!analysis) return;
    setError(null);
    const selectedCa = caProfiles.find((p) => p.id === selectedCaId);
    if (!selectedCa) {
      setError(CA_REQUIRED_MESSAGE);
      return;
    }
    setBusy("generating");
    try {
      const withReportDetails: WorkbookAnalysis = {
        ...analysis,
        place: reportPlace,
        date: reportDate,
        udin: reportUdin,
        caName: selectedCa.caName,
        caFirmName: selectedCa.firmName,
        caFirmType: selectedCa.firmType,
        caDesignation: selectedCa.designation,
        caMembershipNo: selectedCa.membershipNo,
        caFirmRegNo: selectedCa.frn,
      };
      await generateIcaiWorkbook(withReportDetails, figuresUnit);
    } catch (e) {
      setError(e instanceof Error ? e.message : "Failed to generate the financial statements.");
    } finally {
      setBusy("idle");
    }
  }

  function updateAccountCode(index: number, code: IcaiNoteCode) {
    if (!analysis) return;
    setAnalysis({
      ...analysis,
      accounts: analysis.accounts.map((a, i) => (i === index ? { ...a, code, confident: true } : a)),
    });
  }

  const hasCurrentRoleChosen = Object.values(roleChoices).includes("current");

  return (
    <div className="space-y-6">
      <ErrorBanner message={error} />

      <section className="rounded-2xl border bg-card p-6">
        <h2 className="text-sm font-semibold">1. Upload &amp; entity type</h2>
        <div className="mt-4 grid gap-4 sm:grid-cols-2 lg:grid-cols-3">
          <EntityTypeSelect value={icaiEntity} onChange={onEntityChange} options={ICAI_ENTITY_OPTIONS} />
          <FiguresUnitSelect value={figuresUnit} onChange={onFiguresUnitChange} />
        </div>

        <div className="mt-4">
          <label className="text-xs font-medium text-muted-foreground">Accounting Workbook(s)</label>
          <p className="mt-1 text-xs text-muted-foreground">
            Upload one or more Excel workbooks containing accounting data for the current and comparative financial
            years - the firm's own trading/P&amp;L, capital accounts, balance sheet, schedules and fixed assets, not a
            Tally trial balance. The system will detect the current and comparative years automatically.
          </p>
          <input
            type="file"
            multiple
            accept=".xlsx,.xls"
            onChange={(e) => {
              addFiles(e.target.files);
              e.target.value = "";
            }}
            className="mt-2 block w-full text-sm file:mr-3 file:rounded-md file:border-0 file:bg-secondary file:px-3 file:py-2 file:text-sm"
          />
          {files.length > 0 && (
            <ul className="mt-3 space-y-1">
              {files.map((f) => (
                <li key={f.name} className="flex items-center justify-between rounded-lg border bg-secondary/30 px-3 py-1.5 text-sm">
                  <span>✓ {f.name}</span>
                  <button onClick={() => removeFile(f.name)} className="text-muted-foreground hover:text-destructive" aria-label={`Remove ${f.name}`}>
                    <X className="h-4 w-4" />
                  </button>
                </li>
              ))}
            </ul>
          )}
        </div>

        <button
          onClick={handleAnalyze}
          disabled={busy !== "idle" || files.length === 0}
          className="mt-4 inline-flex items-center gap-2 rounded-lg bg-primary px-4 py-2 text-sm font-semibold text-primary-foreground disabled:opacity-50"
        >
          {busy === "analyzing" ? <Loader2 className="h-4 w-4 animate-spin" /> : <Upload className="h-4 w-4" />}
          Analyze Workbooks
        </button>
      </section>

      {pendingReview && (
        <section className="rounded-2xl border border-amber-400/50 bg-amber-50 p-6">
          <h2 className="text-sm font-semibold text-amber-800">Confirm financial year for each workbook</h2>
          <p className="mt-1 text-xs text-amber-800">{pendingReview.reason}</p>
          <div className="mt-4 overflow-auto rounded-lg border bg-white">
            <table className="w-full text-sm">
              <thead className="bg-secondary text-xs">
                <tr>
                  <th className="px-3 py-2 text-left">File</th>
                  <th className="px-3 py-2 text-left">Detected FY</th>
                  <th className="px-3 py-2 text-left">Role</th>
                </tr>
              </thead>
              <tbody>
                {pendingReview.detections.map((d) => (
                  <tr key={d.fileName}>
                    <td className="px-3 py-1.5">{d.fileName}</td>
                    <td className="px-3 py-1.5 text-xs text-muted-foreground">{d.detectedYear ?? (d.periodEndLabel || "(not detected)")}</td>
                    <td className="px-3 py-1.5">
                      <select
                        value={roleChoices[d.fileName] ?? "exclude"}
                        onChange={(e) => setRoleChoices((prev) => ({ ...prev, [d.fileName]: e.target.value as FileRole }))}
                        className="w-full rounded border bg-background px-2 py-1 text-xs"
                      >
                        {(Object.keys(ROLE_LABELS) as FileRole[]).map((r) => (
                          <option key={r} value={r}>
                            {ROLE_LABELS[r]}
                          </option>
                        ))}
                      </select>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
          <button
            onClick={handleConfirmRoles}
            disabled={busy !== "idle" || !hasCurrentRoleChosen}
            className="mt-4 inline-flex items-center gap-2 rounded-lg bg-primary px-4 py-2 text-sm font-semibold text-primary-foreground disabled:opacity-50"
          >
            {busy === "analyzing" ? <Loader2 className="h-4 w-4 animate-spin" /> : null}
            Confirm &amp; Continue
          </button>
          {!hasCurrentRoleChosen && <p className="mt-2 text-xs text-amber-700">Select at least one workbook as the Current Year.</p>}
        </section>
      )}

      {analysis && (
        <>
          <section className="rounded-2xl border bg-card p-6">
            <h2 className="text-sm font-semibold">2. Detected entity</h2>
            <div className="mt-3 grid gap-3 sm:grid-cols-3 text-sm">
              <div>
                <div className="text-xs text-muted-foreground">Entity</div>
                <div className="font-medium">{analysis.entityName || "(not detected)"}</div>
              </div>
              <div>
                <div className="text-xs text-muted-foreground">Current Year</div>
                <div className="font-medium">{analysis.currentYearLabel || "(not detected)"}</div>
              </div>
              <div>
                <div className="text-xs text-muted-foreground">Comparative Year</div>
                <div className="font-medium">{analysis.previousYearLabel || "(none provided)"}</div>
              </div>
            </div>
            {analysis.warnings.length > 0 && (
              <ul className="mt-4 space-y-1 text-xs text-amber-700">
                {analysis.warnings.map((w, i) => (
                  <li key={i}>• {w}</li>
                ))}
              </ul>
            )}
          </section>

          <section className="rounded-2xl border bg-card p-6">
            <h2 className="text-sm font-semibold">3. Owners' Capital Accounts ({analysis.owners.length})</h2>
            <div className="mt-4 overflow-auto rounded-lg border">
              <table className="w-full text-sm">
                <thead className="bg-secondary text-xs">
                  <tr>
                    <th className="px-3 py-2 text-left">Partner</th>
                    <th className="px-3 py-2 text-right">Opening</th>
                    <th className="px-3 py-2 text-right">Introduced</th>
                    <th className="px-3 py-2 text-right">Remuneration</th>
                    <th className="px-3 py-2 text-right">Withdrawals</th>
                    <th className="px-3 py-2 text-right">Share of Profit</th>
                    <th className="px-3 py-2 text-right">Closing</th>
                  </tr>
                </thead>
                <tbody>
                  {analysis.owners.map((o, i) => (
                    <tr key={i} className={i % 2 ? "bg-background" : "bg-card"}>
                      <td className="px-3 py-1.5">{o.name}</td>
                      <td className="px-3 py-1.5 text-right tabular-nums">{o.openingBalance.toLocaleString("en-IN")}</td>
                      <td className="px-3 py-1.5 text-right tabular-nums">{o.capitalIntroduced.toLocaleString("en-IN")}</td>
                      <td className="px-3 py-1.5 text-right tabular-nums">{o.remuneration.toLocaleString("en-IN")}</td>
                      <td className="px-3 py-1.5 text-right tabular-nums">{o.withdrawals.toLocaleString("en-IN")}</td>
                      <td className="px-3 py-1.5 text-right tabular-nums">{o.shareOfProfit.toLocaleString("en-IN")}</td>
                      <td className="px-3 py-1.5 text-right tabular-nums font-medium">{o.closingBalance.toLocaleString("en-IN")}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </section>

          <section className="rounded-2xl border bg-card p-6">
            <div className="flex items-center justify-between">
              <h2 className="text-sm font-semibold">4. Review mapping ({analysis.accounts.length} line items)</h2>
              {reviewCount > 0 && (
                <span className="rounded-full bg-amber-100 px-2.5 py-1 text-xs font-medium text-amber-700">
                  {reviewCount} need review
                </span>
              )}
            </div>
            <div className="mt-4 max-h-[480px] overflow-auto rounded-lg border">
              <table className="w-full text-sm">
                <thead className="sticky top-0 bg-secondary text-xs">
                  <tr>
                    <th className="px-3 py-2 text-left">Line item</th>
                    <th className="px-3 py-2 text-left">Source</th>
                    <th className="px-3 py-2 text-right">Current</th>
                    <th className="px-3 py-2 text-right">Previous</th>
                    <th className="px-3 py-2 text-left">Mapped to</th>
                  </tr>
                </thead>
                <tbody>
                  {analysis.accounts.map((a, i) => (
                    <tr key={i} className={!a.confident ? "bg-amber-50" : i % 2 ? "bg-background" : "bg-card"}>
                      <td className="px-3 py-1.5">
                        {a.name}
                        {!a.confident && <div className="text-[11px] text-amber-700">{a.reason}</div>}
                      </td>
                      <td className="px-3 py-1.5 text-xs text-muted-foreground">{a.sourceSheet}</td>
                      <td className="px-3 py-1.5 text-right tabular-nums">{a.amountCurrent ? a.amountCurrent.toLocaleString("en-IN") : ""}</td>
                      <td className="px-3 py-1.5 text-right tabular-nums">{a.amountPrevious ? a.amountPrevious.toLocaleString("en-IN") : ""}</td>
                      <td className="px-3 py-1.5">
                        <select
                          value={a.code}
                          onChange={(e) => updateAccountCode(i, e.target.value as IcaiNoteCode)}
                          className="w-full rounded border bg-background px-2 py-1 text-xs"
                        >
                          {ICAI_NOTE_CODES.map((c) => (
                            <option key={c} value={c}>
                              {ICAI_NOTE_LABEL[c]}
                            </option>
                          ))}
                        </select>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </section>

          {analysis.reviewItems.length > 0 && (
            <section className="rounded-2xl border border-amber-400/50 bg-amber-50 p-6">
              <h2 className="text-sm font-semibold text-amber-800">
                Unmapped items ({analysis.reviewItems.length}) - not included in the statements below
              </h2>
              <ul className="mt-3 space-y-1 text-xs text-amber-800">
                {analysis.reviewItems.map((r, i) => (
                  <li key={i}>
                    • {r.label} ({r.sourceSheet}): {(r.amountCurrent || r.amountPrevious).toLocaleString("en-IN")} - {r.reason}
                  </li>
                ))}
              </ul>
            </section>
          )}

          <section className="rounded-2xl border bg-card p-6">
            <h2 className="text-sm font-semibold">5. Report details</h2>
            <div className="mt-4 grid gap-4 sm:grid-cols-2 lg:grid-cols-3">
              <TextField
                label="Place"
                value={reportPlace}
                onChange={(v) => {
                  setPlaceTouched(true);
                  setReportPlace(v);
                }}
              />
              <TextField label="Date" value={reportDate} onChange={setReportDate} />
              <TextField label="UDIN" value={reportUdin} onChange={setReportUdin} />
              <CASelect profiles={caProfiles} value={selectedCaId} onChange={handleSelectCa} />
            </div>
          </section>

          <section className="rounded-2xl border bg-card p-6">
            <h2 className="text-sm font-semibold">6. Generate</h2>
            <p className="mt-2 text-xs text-muted-foreground">
              Output: ICAI-format Balance Sheet, Statement of P&amp;L, Notes 1-26, supporting schedules and a Basis of
              Preparation sheet flagging any judgement calls for your review.
            </p>
            <button
              onClick={handleGenerate}
              disabled={busy !== "idle"}
              className="mt-4 inline-flex items-center gap-2 rounded-lg bg-gradient-to-r from-primary to-accent px-5 py-2.5 text-sm font-semibold text-primary-foreground disabled:opacity-50"
            >
              {busy === "generating" ? <Loader2 className="h-4 w-4 animate-spin" /> : <Play className="h-4 w-4" />}
              Generate &amp; download Excel
            </button>
          </section>
        </>
      )}

      {!analysis && !pendingReview && (
        <div className="flex items-center gap-2 rounded-lg border bg-secondary/40 px-4 py-3 text-xs text-muted-foreground">
          <FileSpreadsheet className="h-4 w-4 shrink-0" />
          Upload one or more accounting workbooks and click Analyze Workbooks to continue.
        </div>
      )}
    </div>
  );
}
