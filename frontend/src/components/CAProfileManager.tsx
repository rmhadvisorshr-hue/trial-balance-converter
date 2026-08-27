import { useEffect, useState } from "react";
import { Plus, Pencil, Trash2, Loader2 } from "lucide-react";
import type { CAProfile, CAProfileInput } from "../lib/ca-types";
import { CA_DESIGNATION_OPTIONS, DEFAULT_CA_FIRM_TYPE } from "../lib/ca-types";
import { listCAProfiles, createCAProfile, updateCAProfile, deleteCAProfile } from "../lib/ca-client";
import ErrorBanner from "./ErrorBanner";
import TextField from "./TextField";

const EMPTY_FORM: CAProfileInput = {
  caName: "",
  designation: CA_DESIGNATION_OPTIONS[0],
  firmName: "",
  membershipNo: "",
  frn: "",
  place: "",
  firmType: DEFAULT_CA_FIRM_TYPE,
};

// Self-contained: owns its own fetch + mutations against the centralized
// /api/ca-profiles store. Trial Balance and Accounting Workbook each fetch
// their own copy of the same list independently (see App.tsx) - there is no
// shared in-memory state to keep in sync, only the one backend store.
export default function CAProfileManager() {
  const [profiles, setProfiles] = useState<CAProfile[]>([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);
  const [editingId, setEditingId] = useState<string | null>(null); // null = closed, "new" = creating, else editing that id
  const [form, setForm] = useState<CAProfileInput>(EMPTY_FORM);
  const [saving, setSaving] = useState(false);

  async function load() {
    setLoading(true);
    setError(null);
    try {
      setProfiles(await listCAProfiles());
    } catch (e) {
      setError(e instanceof Error ? e.message : "Failed to load CA profiles.");
    } finally {
      setLoading(false);
    }
  }

  useEffect(() => {
    load();
  }, []);

  function startCreate() {
    setForm(EMPTY_FORM);
    setEditingId("new");
  }

  function startEdit(p: CAProfile) {
    setForm({
      caName: p.caName,
      designation: p.designation,
      firmName: p.firmName,
      membershipNo: p.membershipNo,
      frn: p.frn ?? "",
      place: p.place,
      firmType: p.firmType,
    });
    setEditingId(p.id);
  }

  function cancelEdit() {
    setEditingId(null);
    setForm(EMPTY_FORM);
  }

  async function save() {
    setError(null);
    if (
      !form.caName.trim() ||
      !form.designation.trim() ||
      !form.firmName.trim() ||
      !form.membershipNo.trim() ||
      !form.place.trim()
    ) {
      setError("CA Name, Designation, Firm Name, Membership Number and Place are required.");
      return;
    }
    setSaving(true);
    try {
      const input: CAProfileInput = { ...form, frn: form.frn?.trim() || undefined };
      if (editingId === "new") {
        await createCAProfile(input);
      } else if (editingId) {
        await updateCAProfile(editingId, input);
      }
      cancelEdit();
      await load();
    } catch (e) {
      setError(e instanceof Error ? e.message : "Failed to save CA profile.");
    } finally {
      setSaving(false);
    }
  }

  async function remove(p: CAProfile) {
    if (!window.confirm(`Are you sure you want to delete this CA profile?\n\n${p.caName}`)) return;
    setError(null);
    try {
      await deleteCAProfile(p.id);
      await load();
    } catch (e) {
      setError(e instanceof Error ? e.message : "Failed to delete CA profile.");
    }
  }

  return (
    <div className="space-y-6">
      <ErrorBanner message={error} />

      <section className="rounded-2xl border bg-card p-6">
        <div className="flex items-center justify-between">
          <div>
            <h2 className="text-sm font-semibold">CA Profiles</h2>
            <p className="mt-1 text-xs text-muted-foreground">
              Shared by both Trial Balance and Accounting Workbook generation.
            </p>
          </div>
          <button
            onClick={startCreate}
            className="inline-flex items-center gap-2 rounded-lg bg-primary px-3 py-1.5 text-xs font-semibold text-primary-foreground"
          >
            <Plus className="h-3.5 w-3.5" /> Add CA
          </button>
        </div>

        {loading ? (
          <div className="mt-4 flex items-center gap-2 text-sm text-muted-foreground">
            <Loader2 className="h-4 w-4 animate-spin" /> Loading...
          </div>
        ) : profiles.length === 0 ? (
          <p className="mt-4 text-sm text-muted-foreground">
            No CA profiles yet. Add one to select it when generating a Trial Balance or Accounting
            Workbook.
          </p>
        ) : (
          <div className="mt-4 overflow-auto rounded-lg border">
            <table className="w-full text-sm">
              <thead className="bg-secondary text-xs">
                <tr>
                  <th className="px-3 py-2 text-left">CA Name</th>
                  <th className="px-3 py-2 text-left">Designation</th>
                  <th className="px-3 py-2 text-left">Firm Name</th>
                  <th className="px-3 py-2 text-left">M. No.</th>
                  <th className="px-3 py-2 text-left">FRN No.</th>
                  <th className="px-3 py-2 text-left">Place</th>
                  <th className="px-3 py-2 text-right">Actions</th>
                </tr>
              </thead>
              <tbody>
                {profiles.map((p, i) => (
                  <tr key={p.id} className={i % 2 ? "bg-background" : "bg-card"}>
                    <td className="px-3 py-2 font-medium">{p.caName}</td>
                    <td className="px-3 py-2">{p.designation}</td>
                    <td className="px-3 py-2">{p.firmName}</td>
                    <td className="px-3 py-2">{p.membershipNo}</td>
                    <td className="px-3 py-2">{p.frn || "—"}</td>
                    <td className="px-3 py-2">{p.place}</td>
                    <td className="px-3 py-2 text-right">
                      <button
                        onClick={() => startEdit(p)}
                        className="mr-3 text-muted-foreground hover:text-foreground"
                        aria-label={`Edit ${p.caName}`}
                      >
                        <Pencil className="h-4 w-4" />
                      </button>
                      <button
                        onClick={() => remove(p)}
                        className="text-muted-foreground hover:text-destructive"
                        aria-label={`Delete ${p.caName}`}
                      >
                        <Trash2 className="h-4 w-4" />
                      </button>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}
      </section>

      {editingId && (
        <section className="rounded-2xl border bg-card p-6">
          <h2 className="text-sm font-semibold">{editingId === "new" ? "Add CA Profile" : "Edit CA Profile"}</h2>
          <div className="mt-4 grid gap-4 sm:grid-cols-2 lg:grid-cols-3">
            <TextField
              label="CA Name"
              value={form.caName}
              onChange={(v) => setForm({ ...form, caName: v })}
              placeholder="CA Namrata Prakash Sharma"
            />
            <div>
              <label className="text-xs font-medium text-muted-foreground">Designation</label>
              <select
                value={form.designation}
                onChange={(e) => setForm({ ...form, designation: e.target.value })}
                className="mt-1 w-full rounded-lg border bg-background px-3 py-2 text-sm"
              >
                {CA_DESIGNATION_OPTIONS.map((d) => (
                  <option key={d} value={d}>
                    {d}
                  </option>
                ))}
              </select>
            </div>
            <TextField
              label="Firm Name"
              value={form.firmName}
              onChange={(v) => setForm({ ...form, firmName: v })}
              placeholder="Namrata Prakash Sharma"
            />
            <TextField
              label="Membership Number (M. No.)"
              value={form.membershipNo}
              onChange={(v) => setForm({ ...form, membershipNo: v })}
              placeholder="177309"
            />
            <TextField
              label="FRN Number (optional)"
              value={form.frn ?? ""}
              onChange={(v) => setForm({ ...form, frn: v })}
              placeholder="144860W"
            />
            <TextField
              label="Place"
              value={form.place}
              onChange={(v) => setForm({ ...form, place: v })}
              placeholder="Vasai"
            />
            <TextField
              label="Firm Type"
              value={form.firmType}
              onChange={(v) => setForm({ ...form, firmType: v })}
              placeholder="Chartered Accountants"
            />
          </div>
          <div className="mt-4 flex gap-3">
            <button
              onClick={save}
              disabled={saving}
              className="inline-flex items-center gap-2 rounded-lg bg-primary px-4 py-2 text-sm font-semibold text-primary-foreground disabled:opacity-50"
            >
              {saving ? <Loader2 className="h-4 w-4 animate-spin" /> : null}
              Save
            </button>
            <button onClick={cancelEdit} className="rounded-lg border px-4 py-2 text-sm font-semibold">
              Cancel
            </button>
          </div>
        </section>
      )}
    </div>
  );
}
