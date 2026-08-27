import type { CAProfile, CAProfileInput } from "./ca-types";

const API_BASE = import.meta.env.VITE_API_BASE_URL ?? "";

async function handle<T>(res: Response): Promise<T> {
  if (!res.ok) {
    const data = (await res.json().catch(() => null)) as { message?: string } | null;
    throw new Error(data?.message || "CA profile request failed.");
  }
  if (res.status === 204) return undefined as T;
  return (await res.json()) as T;
}

export async function listCAProfiles(): Promise<CAProfile[]> {
  return handle(await fetch(`${API_BASE}/api/ca-profiles`));
}

export async function createCAProfile(input: CAProfileInput): Promise<CAProfile> {
  return handle(
    await fetch(`${API_BASE}/api/ca-profiles`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(input),
    }),
  );
}

export async function updateCAProfile(id: string, input: CAProfileInput): Promise<CAProfile> {
  return handle(
    await fetch(`${API_BASE}/api/ca-profiles/${id}`, {
      method: "PUT",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(input),
    }),
  );
}

export async function deleteCAProfile(id: string): Promise<void> {
  return handle(await fetch(`${API_BASE}/api/ca-profiles/${id}`, { method: "DELETE" }));
}
