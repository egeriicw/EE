import { useMemo, useState } from "react";
import { Bar, BarChart, CartesianGrid, Legend, ResponsiveContainer, Tooltip, XAxis, YAxis } from "recharts";
import { analysisApi, type DatasetSummary, type Histogram } from "../api";
import { SERIES_COLORS, valueColumns } from "./chartTheme";

interface Props {
  dataset: DatasetSummary;
}

export default function HistogramPanel({ dataset }: Props) {
  const values = valueColumns(dataset);
  const [variable, setVariable] = useState(values[values.length - 1] ?? "");
  const [bins, setBins] = useState(12);
  const [byGroup, setByGroup] = useState(dataset.grouped);
  const [xmin, setXmin] = useState("");
  const [xmax, setXmax] = useState("");
  const [result, setResult] = useState<Histogram | null>(null);
  const [error, setError] = useState<string | null>(null);

  const run = async () => {
    setError(null);
    try {
      setResult(
        await analysisApi.histogram(dataset.id, {
          variable,
          n_bins: bins,
          xmin: xmin === "" ? null : Number(xmin),
          xmax: xmax === "" ? null : Number(xmax),
          by_group: byGroup && dataset.grouped,
        }),
      );
    } catch (e) {
      setError((e as Error).message);
    }
  };

  // one row per bin; each series contributes its percentage of records
  const rows = useMemo(
    () =>
      result
        ? result.midpoints.map((mid, i) => {
            const row: Record<string, number | string> = {
              bin: `${result.edges[i].toPrecision(3)}–${result.edges[i + 1].toPrecision(3)}`,
              mid,
            };
            result.series.forEach((s) => {
              row[s.group] = Number(s.percent[i].toFixed(2));
              row[`${s.group} count`] = s.counts[i];
            });
            return row;
          })
        : [],
    [result],
  );

  return (
    <section className="panel">
      <h2>Histogram</h2>
      <div className="row">
        <label>
          Variable
          <select value={variable} onChange={(e) => setVariable(e.target.value)}>
            {values.map((c) => (
              <option key={c}>{c}</option>
            ))}
          </select>
        </label>
        <label>
          Intervals
          <input type="number" min={1} max={200} value={bins} onChange={(e) => setBins(Number(e.target.value))} />
        </label>
        <label>
          Min (optional)
          <input value={xmin} placeholder="data min" onChange={(e) => setXmin(e.target.value)} />
        </label>
        <label>
          Max (optional)
          <input value={xmax} placeholder="data max" onChange={(e) => setXmax(e.target.value)} />
        </label>
        {dataset.grouped && (
          <label className="inline">
            <input type="checkbox" checked={byGroup} onChange={(e) => setByGroup(e.target.checked)} />
            Separate groups
          </label>
        )}
        <button onClick={run}>Plot histogram</button>
      </div>
      {error && <p className="error">{error}</p>}
      {result && (
        <>
          <ResponsiveContainer width="100%" height={300}>
            <BarChart data={rows} margin={{ top: 8, right: 16, bottom: 24, left: 8 }}>
              <CartesianGrid strokeDasharray="3 3" vertical={false} />
              <XAxis dataKey="bin" interval={0} angle={-30} textAnchor="end" height={60} tick={{ fontSize: 11 }} />
              <YAxis unit="%" width={56} />
              <Tooltip
                formatter={(v: number, name: string, item) => [`${v}%  (${item.payload[`${name} count`]} records)`, name]}
              />
              <Legend verticalAlign="top" />
              {result.series.map((s, i) => (
                <Bar key={s.group} dataKey={s.group} fill={SERIES_COLORS[i % SERIES_COLORS.length]} isAnimationActive={false} />
              ))}
            </BarChart>
          </ResponsiveContainer>
          <p className="muted">
            {result.series.map((s) => `${s.group}: N = ${s.n.toLocaleString()}`).join(" · ")} · range{" "}
            {result.xmin.toPrecision(4)} – {result.xmax.toPrecision(4)}
          </p>
        </>
      )}
    </section>
  );
}
