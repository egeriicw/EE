import { useEffect, useMemo, useState } from "react";
import { CartesianGrid, ComposedChart, Legend, Line, ResponsiveContainer, Scatter, Tooltip, XAxis, YAxis } from "recharts";
import { api, type DatasetSummary, type FitResponse, type ModelTypeSpec } from "../api";

interface Props {
  dataset: DatasetSummary;
  fit: FitResponse | null;
  onFit: (f: FitResponse) => void;
}

export default function ModelPanel({ dataset, fit, onFit }: Props) {
  const [types, setTypes] = useState<Record<string, ModelTypeSpec>>({});
  const [modelType, setModelType] = useState("3PC");
  const values = dataset.columns.filter((c) => !Object.values(dataset.time_fields).includes(c));
  const [y, setY] = useState(values[0] ?? "");
  const [xs, setXs] = useState<string[]>(values.length > 1 ? [values[values.length - 1]] : []);
  const [group, setGroup] = useState<string>("");
  const [error, setError] = useState<string | null>(null);

  useEffect(() => {
    api.modelTypes().then(setTypes).catch((e) => setError(String(e.message ?? e)));
  }, []);
  useEffect(() => setGroup(dataset.grouped ? dataset.group_names[0] : ""), [dataset.grouped, dataset.group_names]);

  const spec = types[modelType];
  const maxX = spec?.max_x ?? 1;

  const run = async () => {
    setError(null);
    try {
      onFit(await api.fit(dataset.id, { model_type: modelType, y, x: xs.slice(0, maxX), group: group || null }));
    } catch (e) {
      setError((e as Error).message);
    }
  };

  const scatter = useMemo(
    () => (fit?.points.x ? fit.points.x.map((x, i) => ({ x, y: fit.points.y[i] })) : []),
    [fit],
  );
  const curve = useMemo(() => (fit?.curve ? fit.curve.x.map((x, i) => ({ x, fit: fit.curve!.y[i] })) : []), [fit]);

  return (
    <section className="panel">
      <h2>3 · Baseline model</h2>
      <div className="row">
        <label>
          Model
          <select value={modelType} onChange={(e) => setModelType(e.target.value)}>
            {Object.entries(types).map(([k, v]) => (
              <option key={k} value={k} title={v.label}>
                {k}
              </option>
            ))}
          </select>
        </label>
        <label>
          Y
          <select value={y} onChange={(e) => setY(e.target.value)}>
            {values.map((c) => (
              <option key={c}>{c}</option>
            ))}
          </select>
        </label>
        {maxX > 0 && (
          <label>
            X{maxX > 1 ? ` (X1 = change-point var, up to ${maxX})` : ""}
            <select
              multiple={maxX > 1}
              value={maxX > 1 ? xs : xs[0] ?? ""}
              onChange={(e) => setXs(Array.from(e.target.selectedOptions, (o) => o.value))}
            >
              {values.map((c) => (
                <option key={c}>{c}</option>
              ))}
            </select>
          </label>
        )}
        {dataset.grouped && (
          <label>
            Fit on
            <select value={group} onChange={(e) => setGroup(e.target.value)}>
              <option value="">all records</option>
              {dataset.group_names.map((g) => (
                <option key={g}>{g}</option>
              ))}
            </select>
          </label>
        )}
        <button onClick={run}>Fit</button>
      </div>
      {error && <p className="error">{error}</p>}
      {fit && (
        <>
          <pre className="summary">{fit.summary}</pre>
          {scatter.length > 0 && (
            <ResponsiveContainer width="100%" height={320}>
              <ComposedChart margin={{ top: 8, right: 16, bottom: 16, left: 8 }}>
                <CartesianGrid strokeDasharray="3 3" />
                <XAxis dataKey="x" type="number" domain={["auto", "auto"]} name={xs[0]} />
                <YAxis type="number" domain={["auto", "auto"]} name={y} />
                <Tooltip />
                <Legend />
                <Scatter name="observed" data={scatter} dataKey="y" fill="#4f7cac" fillOpacity={0.55} />
                <Line name="model" data={curve} dataKey="fit" dot={false} stroke="#d1495b" strokeWidth={2} isAnimationActive={false} />
              </ComposedChart>
            </ResponsiveContainer>
          )}
        </>
      )}
    </section>
  );
}
