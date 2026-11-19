import { useMemo, useState } from "react";
import { CartesianGrid, Legend, Line, LineChart, ResponsiveContainer, Tooltip, XAxis, YAxis } from "recharts";
import { api, type DatasetSummary, type FitResponse, type SavingsResponse } from "../api";

interface Props {
  fit: FitResponse;
  dataset: DatasetSummary;
}

export default function SavingsPanel({ fit, dataset }: Props) {
  const [measErr, setMeasErr] = useState(5);
  const [ar1, setAr1] = useState(false);
  const [result, setResult] = useState<SavingsResponse | null>(null);
  const [error, setError] = useState<string | null>(null);
  const linear = ["Mean", "2P", "MLR"].includes(fit.model.model_type);

  const run = async () => {
    setError(null);
    try {
      setResult(
        await api.savings(fit.id, {
          measurement_error_pct: measErr,
          post_group: dataset.group_names[1],
          autocorrelation_corrected: ar1 && linear,
        }),
      );
    } catch (e) {
      setError((e as Error).message);
    }
  };

  const series = useMemo(
    () =>
      result
        ? result.series.timestamp.map((t, i) => ({
            t: t.slice(0, 10),
            actual: result.series.actual[i],
            baseline: result.series.adjusted_baseline[i],
          }))
        : [],
    [result],
  );

  return (
    <section className="panel">
      <h2>4 · Savings</h2>
      <div className="row">
        <label>
          Measurement error (±%)
          <input type="number" value={measErr} min={0} step={0.5} onChange={(e) => setMeasErr(Number(e.target.value))} />
        </label>
        <label className="inline">
          <input type="checkbox" checked={ar1} disabled={!linear} onChange={(e) => setAr1(e.target.checked)} />
          AR(1)-corrected uncertainty (Mean/2P/MLR)
        </label>
        <button onClick={run}>Calculate savings</button>
      </div>
      {error && <p className="error">{error}</p>}
      {result && (
        <>
          <pre className="summary">{result.report}</pre>
          {result.autocorrelated && (
            <p className="muted">
              AR(1) ρ = {result.autocorrelated.rho?.toFixed(3)} · savings uncertainty ±
              {result.autocorrelated.savings_error?.toLocaleString(undefined, { maximumFractionDigits: 1 })}
            </p>
          )}
          <ResponsiveContainer width="100%" height={300}>
            <LineChart data={series} margin={{ top: 8, right: 16, bottom: 16, left: 8 }}>
              <CartesianGrid strokeDasharray="3 3" />
              <XAxis dataKey="t" minTickGap={40} />
              <YAxis />
              <Tooltip />
              <Legend />
              <Line dataKey="actual" stroke="#4f7cac" dot={false} isAnimationActive={false} connectNulls={false} />
              <Line dataKey="baseline" name="adjusted baseline" stroke="#d1495b" strokeDasharray="5 3" dot={false} isAnimationActive={false} />
            </LineChart>
          </ResponsiveContainer>
        </>
      )}
    </section>
  );
}
