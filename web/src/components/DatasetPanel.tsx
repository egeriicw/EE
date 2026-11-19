import { useEffect, useState } from "react";
import { api, type DatasetSummary } from "../api";

interface Props {
  dataset: DatasetSummary | null;
  onLoaded: (ds: DatasetSummary) => void;
}

export default function DatasetPanel({ dataset, onLoaded }: Props) {
  const [samples, setSamples] = useState<string[]>([]);
  const [dataFile, setDataFile] = useState<File | null>(null);
  const [namesFile, setNamesFile] = useState<File | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [busy, setBusy] = useState(false);

  useEffect(() => {
    api.samples().then(setSamples).catch((e) => setError(String(e.message ?? e)));
  }, []);

  const run = async (p: Promise<DatasetSummary>) => {
    setBusy(true);
    setError(null);
    try {
      onLoaded(await p);
    } catch (e) {
      setError((e as Error).message);
    } finally {
      setBusy(false);
    }
  };

  return (
    <section className="panel">
      <h2>1 · Data</h2>
      <div className="row">
        <label>
          Sample
          <select defaultValue="" onChange={(e) => e.target.value && run(api.loadSample(e.target.value))}>
            <option value="" disabled>
              choose…
            </option>
            {samples.map((s) => (
              <option key={s}>{s}</option>
            ))}
          </select>
        </label>
        <span className="muted">or upload</span>
        <label>
          Data file
          <input type="file" onChange={(e) => setDataFile(e.target.files?.[0] ?? null)} />
        </label>
        <label>
          Names (.var)
          <input type="file" accept=".var,.dvn,.txt" onChange={(e) => setNamesFile(e.target.files?.[0] ?? null)} />
        </label>
        <button disabled={!dataFile || busy} onClick={() => dataFile && run(api.upload(dataFile, namesFile))}>
          Upload
        </button>
      </div>
      {error && <p className="error">{error}</p>}
      {dataset && (
        <p className="muted">
          <strong>{dataset.name}</strong> · {dataset.n_records.toLocaleString()} records · {dataset.interval} ·{" "}
          {dataset.columns.join(", ")}
        </p>
      )}
    </section>
  );
}
