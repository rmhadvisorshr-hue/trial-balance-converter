import { ENTITY_LABELS, type EntityType } from "../lib/types";

export default function EntityTypeSelect({
  value,
  onChange,
  options,
}: {
  value: EntityType;
  onChange: (v: EntityType) => void;
  options: EntityType[];
}) {
  return (
    <div>
      <label className="text-xs font-medium text-muted-foreground">Entity type</label>
      <select
        value={value}
        onChange={(e) => onChange(e.target.value as EntityType)}
        className="mt-1 w-full rounded-lg border bg-background px-3 py-2 text-sm"
      >
        {options.map((e) => (
          <option key={e} value={e}>
            {ENTITY_LABELS[e]}
          </option>
        ))}
      </select>
    </div>
  );
}
