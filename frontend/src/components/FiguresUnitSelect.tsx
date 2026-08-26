import { FIGURES_UNIT_LABELS, type FiguresUnit } from "../lib/types";

const FIGURES_UNIT_ORDER: FiguresUnit[] = ["actual", "thousands", "lakhs"];

export default function FiguresUnitSelect({
  value,
  onChange,
  label = "Figures Display Unit",
}: {
  value: FiguresUnit;
  onChange: (v: FiguresUnit) => void;
  label?: string;
}) {
  return (
    <div>
      <label className="text-xs font-medium text-muted-foreground">{label}</label>
      <select
        value={value}
        onChange={(e) => onChange(e.target.value as FiguresUnit)}
        className="mt-1 w-full rounded-lg border bg-background px-3 py-2 text-sm"
      >
        {FIGURES_UNIT_ORDER.map((u) => (
          <option key={u} value={u}>
            {FIGURES_UNIT_LABELS[u]}
          </option>
        ))}
      </select>
    </div>
  );
}
