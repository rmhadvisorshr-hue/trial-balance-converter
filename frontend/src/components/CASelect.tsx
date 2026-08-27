import type { CAProfile } from "../lib/ca-types";
import { formatCAProfileLabel } from "../lib/ca-types";

// Shared by both the Trial Balance and Accounting Workbook workflows - reads
// from the same centralized CA profile list (see lib/ca-client.ts). No CA
// details are ever hardcoded or defaulted here: an empty selection stays
// empty until the user picks one.
export default function CASelect({
  profiles,
  value,
  onChange,
}: {
  profiles: CAProfile[];
  value: string;
  onChange: (id: string) => void;
}) {
  return (
    <div>
      <label className="text-xs font-medium text-muted-foreground">Select CA / Signing Authority</label>
      <select
        value={value}
        onChange={(e) => onChange(e.target.value)}
        className="mt-1 w-full rounded-lg border bg-background px-3 py-2 text-sm"
      >
        <option value="">{profiles.length ? "Select CA..." : "No CA profiles yet"}</option>
        {profiles.map((p) => (
          <option key={p.id} value={p.id}>
            {formatCAProfileLabel(p)}
          </option>
        ))}
      </select>
      {!value && (
        <p className="mt-1 text-[11px] text-amber-600">Required before generating the report.</p>
      )}
    </div>
  );
}
