// Typed client for the FastAPI service (see api/app/schemas.py).

export interface DatasetSummary {
  id: string;
  name: string;
  n_records: number;
  columns: string[];
  time_fields: Record<string, string>;
  interval: string;
  grouped: boolean;
  group_names: string[];
  group_counts: Record<string, number> | null;
  no_data: number;
}

export interface ModelTypeSpec {
  label: string;
  min_x: number;
  max_x: number;
}

export interface FitResponse {
  id: string;
  dataset_id: string;
  model: Record<string, unknown> & { model_type: string; equation: string; stats: Record<string, number | string | null> };
  summary: string;
  points: { x?: (number | null)[]; y: (number | null)[]; row: number[] };
  curve: { x: number[]; y: (number | null)[] } | null;
}

export interface SavingsResponse {
  savings: Record<string, number | string | null>;
  report: string;
  autocorrelated?: Record<string, number | null>;
  series: { timestamp: string[]; actual: (number | null)[]; adjusted_baseline: (number | null)[] };
}

async function request<T>(path: string, init?: RequestInit): Promise<T> {
  const res = await fetch(path, init);
  if (!res.ok) {
    let detail = res.statusText;
    try {
      const body = await res.json();
      detail = typeof body.detail === "string" ? body.detail : JSON.stringify(body.detail);
    } catch {
      /* not JSON */
    }
    throw new Error(detail);
  }
  return res.json() as Promise<T>;
}

const json = (body: unknown): RequestInit => ({
  method: "POST",
  headers: { "Content-Type": "application/json" },
  body: JSON.stringify(body),
});

export const api = {
  samples: () => request<string[]>("/api/datasets/samples"),
  loadSample: (name: string) => request<DatasetSummary>(`/api/datasets/samples/${encodeURIComponent(name)}`, { method: "POST" }),
  upload: (data: File, names?: File | null, noData = -99) => {
    const form = new FormData();
    form.append("data", data);
    if (names) form.append("names", names);
    form.append("no_data", String(noData));
    return request<DatasetSummary>("/api/datasets", { method: "POST", body: form });
  },
  group: (id: string, body: { mode: "date" | "dow"; start?: string; end?: string; names?: [string, string] }) =>
    request<DatasetSummary>(`/api/datasets/${id}/groups`, json(body)),
  modelTypes: () => request<Record<string, ModelTypeSpec>>("/api/model-types"),
  fit: (id: string, body: { model_type: string; y: string; x: string[]; group?: string | number | null }) =>
    request<FitResponse>(`/api/datasets/${id}/models`, json(body)),
  savings: (modelId: string, body: { measurement_error_pct: number; post_group?: string | number; autocorrelation_corrected?: boolean }) =>
    request<SavingsResponse>(`/api/models/${modelId}/savings`, json(body)),
};

export interface HistogramSeries {
  group: string;
  n: number;
  counts: number[];
  probability: number[];
  percent: number[];
}

export interface Histogram {
  variable: string;
  n_bins: number;
  xmin: number;
  xmax: number;
  edges: number[];
  midpoints: number[];
  series: HistogramSeries[];
}

export interface ProfilePolyline {
  date: string;
  points: ([number, number] | null)[];
}

export interface HourlyProfile {
  variable: string;
  group: string;
  dates: string[];
  hours: number[];
  values: (number | null)[][];
  scale: number;
  n: number;
  polylines: ProfilePolyline[];
}

export const analysisApi = {
  histogram: (
    id: string,
    body: { variable: string; n_bins: number; xmin?: number | null; xmax?: number | null; group?: string | null; by_group: boolean },
  ) => request<Histogram>(`/api/datasets/${id}/histogram`, json(body)),
  hourlyProfile: (id: string, body: { variable: string; group?: string | null; theta_deg: number }) =>
    request<HourlyProfile>(`/api/datasets/${id}/hourly-profile`, json(body)),
};
