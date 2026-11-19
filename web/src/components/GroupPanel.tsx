import { useState } from "react";
import { api, type DatasetSummary } from "../api";

interface Props {
  dataset: DatasetSummary;
  onChange: (ds: DatasetSummary) => void;
}

export default function GroupPanel({ dataset, onChange }: Props) {
  const [start, setStart] = useState("");
  const [end, setEnd] = useState("");
  const [error, setError] = useState<string | null>(null);

  const apply = async (mode: "date" | "dow") => {
    setError(null);
    try {
      onChange(await api.group(dataset.id, { mode, start: start || undefined, end: end || undefined, names: ["Pre", "Post"] }));
    } catch (e) {
      setError((e as Error).message);
    }
  };

  const hint = { daily: "m/d/yyyy", monthly: "m/yyyy", hourly: "m/d/yyyy/h", yearly: "yyyy" }[dataset.interval] ?? "";

  return (
    <section className="panel">
      <h2>2 · Groups</h2>
      <div className="row">
        <label>
          Pre start
          <input value={start} placeholder={`start of file (${hint})`} onChange={(e) => setStart(e.target.value)} />
        </label>
        <label>
          Pre end
          <input value={end} placeholder={hint} onChange={(e) => setEnd(e.target.value)} />
        </label>
        <button onClick={() => apply("date")}>Group by date (Pre / Post)</button>
        <button className="secondary" onClick={() => apply("dow")} disabled={dataset.interval === "monthly" || dataset.interval === "yearly"}>
          Weekday / weekend
        </button>
      </div>
      {error && <p className="error">{error}</p>}
      {dataset.group_counts && (
        <p className="muted">
          {Object.entries(dataset.group_counts)
            .map(([k, v]) => `${k}: ${v}`)
            .join(" · ")}
        </p>
      )}
    </section>
  );
}
