import type { DatasetSummary } from "../api";

// Colour-blind-safe series palette (Okabe-Ito order). Used by every chart so groups keep their colour.
export const SERIES_COLORS = ["#0072B2", "#D55E00", "#009E73", "#CC79A7", "#E69F00", "#56B4E9"];

export function valueColumns(ds: DatasetSummary): string[] {
  const time = new Set(Object.values(ds.time_fields));
  return ds.columns.filter((c) => !time.has(c));
}

// Sequential colour ramp (light to dark blue) for heatmaps. t in [0, 1].
export function rampColor(t: number): string {
  const c = Math.min(1, Math.max(0, t));
  const from = [239, 246, 252];
  const to = [8, 48, 107];
  const mix = from.map((f, i) => Math.round(f + (to[i] - f) * c));
  return `rgb(${mix[0]}, ${mix[1]}, ${mix[2]})`;
}
