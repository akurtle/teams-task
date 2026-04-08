import { CreateTaskPayload, ServiceHealth, SyncedTaskRecord } from "./types";

const apiBaseUrl = (import.meta.env.VITE_API_BASE_URL ?? "http://localhost:3978").replace(
  /\/$/,
  ""
);

export async function fetchHealth() {
  return fetchJson<ServiceHealth>(`${apiBaseUrl}/health`);
}

export async function fetchTasks() {
  return fetchJson<SyncedTaskRecord[]>(`${apiBaseUrl}/api/tasks`);
}

export async function createTask(payload: CreateTaskPayload) {
  return fetchJson<SyncedTaskRecord>(`${apiBaseUrl}/api/tasks`, {
    method: "POST",
    headers: {
      "Content-Type": "application/json"
    },
    body: JSON.stringify(payload)
  });
}

async function fetchJson<T>(input: RequestInfo | URL, init?: RequestInit): Promise<T> {
  const response = await fetch(input, init);

  if (!response.ok) {
    const detail = await response.text();
    throw new Error(detail || `Request failed with status ${response.status}.`);
  }

  return (await response.json()) as T;
}
