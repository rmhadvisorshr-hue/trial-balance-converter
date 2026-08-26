import { useMemo, useRef, useState } from "react";
import { FileSpreadsheet, Upload, Play, Loader2 } from "lucide-react";
import {
  ENTITY_LABELS,
  OUTPUT_FORMAT_LABELS,
  STATEMENT_CODES,
  type ClassifiedLedger,
  type EntityType,
  type FiguresUnit,
  type OutputFormat,
  type ParseResult,
  type StatementCode,
  type StatementMeta,
  type StatementStyle,
} from "./lib/types";
import { parseTrialBalanceFile, convertToWorkbook } from "./lib/tb-client";
import IcaiWorkflow from "./IcaiWorkflow";
import ErrorBanner from "./components/ErrorBanner";
import EntityTypeSelect from "./components/EntityTypeSelect";
import FiguresUnitSelect from "./components/FiguresUnitSelect";

const ENTITY_ORDER: EntityType[] = ["pvtltd", "partnership", "llp", "proprietor"];
const OUTPUT_FORMAT_ORDER: OutputFormat[] = ["excel", "pdf"];

type InputType = "trialBalance" | "accountingWorkbook";

export default function App() {
  const [inputType, setInputType] = useState<InputType>("trialBalance");
  const [entity, setEntity] = useState<EntityType>("pvtltd");
  const [file, setFile] = useState<File | null>(null);
  const [parsed, setParsed] = useState<ParseResult | null>(null);
  const [ledgers, setLedgers] = useState<ClassifiedLedger[]>([]);
  const [meta, setMeta] = useState<StatementMeta | null>(null);
  // Independent of `meta`/`parsed`: a display preference, not data extracted
  // from the trial balance, so it must be choosable before a file is even
  // uploaded and must survive re-parsing a different file in the same session.
  const [figuresUnit, setFiguresUnit] = useState<FiguresUnit>("actual");
  const [outputFormat, setOutputFormat] = useState<OutputFormat>("excel");
  const [busy, setBusy] = useState<"idle" | "parsing" | "generating">("idle");
  const [error, setError] = useState<string | null>(null);
  const fileInput = useRef<HTMLInputElement>(null);

  const unmappedCount = useMemo(
    () => ledgers.filter((l) => l.code === "UNMAPPED" || !l.confident).length,
    [ledgers],
  );

  // No user-facing choice: Pvt Ltd companies always get the Schedule III
  // package (Excel or PDF); every other entity type always gets the
  // vertical BS/P&L layout - see the matching invariants in
  // builders/index.ts and builders/pdf.ts.
  const statementStyle: StatementStyle = entity === "pvtltd" ? "scheduleIII" : "statutory";

  async function handleParse() {
    if (!file) {
      setError("Choose a trial balance file first.");
      return;
    }
    setError(null);
    setBusy("parsing");
    try {
      const res = await parseTrialBalanceFile(file, entity);
      setParsed(res);
      setLedgers(res.ledgers);
      setMeta(res.meta);
    } catch (e) {
      setError(e instanceof Error ? e.message : "Failed to read the trial balance.");
    } finally {
      setBusy("idle");
    }
  }

  async function handleGenerate() {
    if (!meta || !outputFormat) return;
    setError(null);
    setBusy("generating");
    try {
      await convertToWorkbook({
        entity,
        meta: { ...meta, figuresUnit },
        ledgers,
        outputFormat,
        statementStyle,
      });
    } catch (e) {
      setError(e instanceof Error ? e.message : "Failed to generate the workbook.");
    } finally {
      setBusy("idle");
    }
  }

  function updateLedger(index: number, code: StatementCode) {
    setLedgers((prev) => prev.map((l, i) => (i === index ? { ...l, code, confident: true } : l)));
  }

  return (
    <div className="min-h-screen bg-background text-foreground">
      <header className="border-b bg-card">
        <div className="mx-auto flex max-w-6xl items-center justify-between gap-3 px-6 py-4">
          <div className="flex items-center gap-3">
            <div className="grid h-10 w-10 place-items-center rounded-lg bg-gradient-to-br from-primary to-accent text-primary-foreground">
              <FileSpreadsheet className="h-5 w-5" />
            </div>
            <div>
              <h1 className="text-lg font-bold">Financial Statement Generator</h1>
              <p className="text-xs text-muted-foreground">
                Convert a Tally trial balance or a firm's own accounting workbook into CA format financial
                statements.
              </p>
            </div>
          </div>
          <a href="/dashboard" className="shrink-0 text-sm text-muted-foreground hover:text-foreground transition">
            &larr; Back to Dashboard
          </a>
        </div>
      </header>

      <main className="mx-auto max-w-6xl space-y-6 px-6 py-8">
        <div className="rounded-2xl border bg-card p-2">
          <div className="grid grid-cols-2 gap-2">
            <button
              onClick={() => setInputType("trialBalance")}
              className={
                "rounded-lg px-4 py-2.5 text-sm font-semibold transition " +
                (inputType === "trialBalance" ? "bg-primary text-primary-foreground" : "text-muted-foreground hover:bg-secondary")
              }
            >
              Trial Balance
            </button>
            <button
              onClick={() => setInputType("accountingWorkbook")}
              className={
                "rounded-lg px-4 py-2.5 text-sm font-semibold transition " +
                (inputType === "accountingWorkbook" ? "bg-primary text-primary-foreground" : "text-muted-foreground hover:bg-secondary")
              }
            >
              Accounting Workbook (ICAI Format)
            </button>
          </div>
        </div>

        {inputType === "accountingWorkbook" ? (
          <IcaiWorkflow
            entity={entity}
            onEntityChange={setEntity}
            figuresUnit={figuresUnit}
            onFiguresUnitChange={setFiguresUnit}
          />
        ) : (
          <>
        <ErrorBanner message={error} />

        {/* Step 1: inputs */}
        <section className="rounded-2xl border bg-card p-6">
          <h2 className="text-sm font-semibold">1. Upload & entity type</h2>
          <div className="mt-4 grid gap-4 sm:grid-cols-2 lg:grid-cols-3">
            <EntityTypeSelect value={entity} onChange={setEntity} options={ENTITY_ORDER} />
            <div>
              <label className="text-xs font-medium text-muted-foreground">
                Trial Balance (.xlsx, .xls, .pdf)
              </label>
              <div className="mt-1 flex items-center gap-2">
                <input
                  ref={fileInput}
                  type="file"
                  accept=".xlsx,.xls,.pdf"
                  onChange={(e) => setFile(e.target.files?.[0] ?? null)}
                  className="block w-full text-sm file:mr-3 file:rounded-md file:border-0 file:bg-secondary file:px-3 file:py-2 file:text-sm"
                />
              </div>
            </div>
            <FiguresUnitSelect value={figuresUnit} onChange={setFiguresUnit} />
          </div>
          <div className="mt-4 flex items-center gap-3">
            <button
              onClick={handleParse}
              disabled={busy !== "idle" || !file}
              className="inline-flex items-center gap-2 rounded-lg bg-primary px-4 py-2 text-sm font-semibold text-primary-foreground disabled:opacity-50"
            >
              {busy === "parsing" ? (
                <Loader2 className="h-4 w-4 animate-spin" />
              ) : (
                <Upload className="h-4 w-4" />
              )}
              Read trial balance
            </button>
            {file && <span className="text-xs text-muted-foreground">{file.name}</span>}
          </div>
        </section>

        {parsed && meta && (
          <>
            {/* Meta */}
            <section className="rounded-2xl border bg-card p-6">
              <h2 className="text-sm font-semibold">2. Statement details</h2>
              <div className="mt-4 grid gap-4 sm:grid-cols-2 lg:grid-cols-3">
                <MetaField
                  label="Firm / company name"
                  value={meta.firmName}
                  onChange={(v) => setMeta({ ...meta, firmName: v })}
                />
                {entity === "pvtltd" && (
                  <MetaField
                    label="CIN"
                    value={meta.cin ?? ""}
                    onChange={(v) => setMeta({ ...meta, cin: v })}
                  />
                )}
                {entity === "llp" && (
                  <MetaField
                    label="LLPIN"
                    value={meta.llpin ?? ""}
                    onChange={(v) => setMeta({ ...meta, llpin: v })}
                  />
                )}
                <MetaField
                  label="Period"
                  value={meta.periodLabel}
                  onChange={(v) => setMeta({ ...meta, periodLabel: v })}
                />
                <MetaField
                  label="As at"
                  value={meta.asAtLabel}
                  onChange={(v) => setMeta({ ...meta, asAtLabel: v })}
                />
                <MetaField
                  label="Place"
                  value={meta.place ?? ""}
                  onChange={(v) => setMeta({ ...meta, place: v })}
                />
                <MetaField
                  label="Date"
                  value={meta.date ?? ""}
                  onChange={(v) => setMeta({ ...meta, date: v })}
                />
                <MetaField
                  label="UDIN"
                  value={meta.udin ?? ""}
                  onChange={(v) => setMeta({ ...meta, udin: v })}
                />
              </div>
              {parsed.warnings.length > 0 && (
                <ul className="mt-4 space-y-1 text-xs text-amber-600">
                  {parsed.warnings.map((w, i) => (
                    <li key={i}>• {w}</li>
                  ))}
                </ul>
              )}
            </section>

            {/* Review table */}
            <section className="rounded-2xl border bg-card p-6">
              <div className="flex items-center justify-between">
                <h2 className="text-sm font-semibold">
                  3. Review mapping ({ledgers.length} ledgers)
                </h2>
                {unmappedCount > 0 && (
                  <span className="rounded-full bg-amber-100 px-2.5 py-1 text-xs font-medium text-amber-700">
                    {unmappedCount} need review
                  </span>
                )}
              </div>
              <div className="mt-4 max-h-[480px] overflow-auto rounded-lg border">
                <table className="w-full text-sm">
                  <thead className="sticky top-0 bg-secondary text-xs">
                    <tr>
                      <th className="px-3 py-2 text-left">Ledger</th>
                      <th className="px-3 py-2 text-left">Tally group</th>
                      <th className="px-3 py-2 text-right">Debit</th>
                      <th className="px-3 py-2 text-right">Credit</th>
                      <th className="px-3 py-2 text-left">Mapped to</th>
                    </tr>
                  </thead>
                  <tbody>
                    {ledgers.map((l, i) => (
                      <tr
                        key={i}
                        className={
                          !l.confident ? "bg-amber-50" : i % 2 ? "bg-background" : "bg-card"
                        }
                      >
                        <td className="px-3 py-1.5">{l.name}</td>
                        <td className="px-3 py-1.5 text-xs text-muted-foreground">
                          {l.parentGroup}
                        </td>
                        <td className="px-3 py-1.5 text-right tabular-nums">
                          {l.debit ? l.debit.toLocaleString("en-IN") : ""}
                        </td>
                        <td className="px-3 py-1.5 text-right tabular-nums">
                          {l.credit ? l.credit.toLocaleString("en-IN") : ""}
                        </td>
                        <td className="px-3 py-1.5">
                          <select
                            value={l.code}
                            onChange={(e) => updateLedger(i, e.target.value as StatementCode)}
                            className="w-full rounded border bg-background px-2 py-1 text-xs"
                          >
                            {STATEMENT_CODES.map((c) => (
                              <option key={c.code} value={c.code}>
                                {c.label}
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

            {/* Generate */}
            <section className="rounded-2xl border bg-card p-6">
              <h2 className="text-sm font-semibold">4. Generate</h2>

              <div
                className={
                  "mt-3 rounded-lg border-2 px-4 py-3 text-sm " +
                  (entity === "pvtltd"
                    ? "border-primary/40 bg-primary/5"
                    : "border-amber-400/50 bg-amber-50")
                }
              >
                <div className="font-semibold">
                  {entity === "pvtltd"
                    ? "Output: Full Schedule III Financials (11 sheets: BS, P&L, NOTES, Note 2,3-SC, 11.FA, DTL, DTA, Ratio, Notes 27-31, AGING SC, Current Liability Maturity)"
                    : `Output: ${ENTITY_LABELS[entity]} standard layout (PL (V), BS (V) and supporting schedules) - not Schedule III.`}
                </div>
                {entity !== "pvtltd" && (
                  <div className="mt-1 text-xs text-amber-700">
                    Schedule III only applies to Private Limited Companies. Change{" "}
                    <span className="font-medium">Entity type</span> in Step 1 above to "Private
                    Limited Company" (and re-parse) if that's what you need.
                  </div>
                )}
              </div>

              <p className="mt-3 text-xs text-muted-foreground">
                {entity === "pvtltd"
                  ? "BS/P&L are linked to Notes, 11.FA, DTL and DTA. Figures a trial balance can't provide (opening asset balances, IT Act depreciation, MSME status) are left as editable 0s."
                  : "Schedules that require data not present in the trial balance (depreciation rates, prior-year figures, MSME numbers) are emitted as editable templates."}
              </p>
              <div className="mt-4 flex flex-wrap gap-4">
                <div className="max-w-xs flex-1 rounded-lg border bg-secondary/40 p-3">
                  <FiguresUnitSelect value={figuresUnit} onChange={setFiguresUnit} />
                  <p className="mt-1 text-[11px] text-muted-foreground">
                    Applies to the generated Profit &amp; Loss and Balance Sheet.
                  </p>
                </div>
                <div className="max-w-xs flex-1 rounded-lg border bg-secondary/40 p-3">
                  <label className="text-xs font-medium text-muted-foreground">
                    Output Format
                  </label>
                  <select
                    value={outputFormat}
                    onChange={(e) => setOutputFormat(e.target.value as OutputFormat)}
                    className="mt-1 w-full rounded-lg border bg-background px-3 py-2 text-sm"
                  >
                    {OUTPUT_FORMAT_ORDER.map((f) => (
                      <option key={f} value={f}>
                        {OUTPUT_FORMAT_LABELS[f]}
                      </option>
                    ))}
                  </select>
                  <p className="mt-1 text-[11px] text-muted-foreground">
                    File type for the downloaded financial statements.
                  </p>
                </div>
              </div>
              <button
                onClick={handleGenerate}
                disabled={busy !== "idle" || !file || !meta || !outputFormat}
                className="mt-4 inline-flex items-center gap-2 rounded-lg bg-gradient-to-r from-primary to-accent px-5 py-2.5 text-sm font-semibold text-primary-foreground disabled:opacity-50"
              >
                {busy === "generating" ? (
                  <Loader2 className="h-4 w-4 animate-spin" />
                ) : (
                  <Play className="h-4 w-4" />
                )}
                Generate &amp; download {OUTPUT_FORMAT_LABELS[outputFormat]}
              </button>
            </section>
          </>
        )}
          </>
        )}
      </main>
    </div>
  );
}

function MetaField({
  label,
  value,
  onChange,
}: {
  label: string;
  value: string;
  onChange: (v: string) => void;
}) {
  return (
    <div>
      <label className="text-xs font-medium text-muted-foreground">{label}</label>
      <input
        value={value}
        onChange={(e) => onChange(e.target.value)}
        className="mt-1 w-full rounded-lg border bg-background px-3 py-2 text-sm"
      />
    </div>
  );
}
