import { useState } from "react";
import type { DatasetSummary, FitResponse } from "./api";
import DatasetPanel from "./components/DatasetPanel";
import GroupPanel from "./components/GroupPanel";
import ModelPanel from "./components/ModelPanel";
import SavingsPanel from "./components/SavingsPanel";

// Workflow mirrors the original program: open data -> group pre/post -> fit baseline -> savings.
export default function App() {
  const [dataset, setDataset] = useState<DatasetSummary | null>(null);
  const [fit, setFit] = useState<FitResponse | null>(null);

  return (
    <div className="app">
      <header>
        <h1>Energy Explorer</h1>
        <p className="muted">Change-point models and M&amp;V savings for measured energy data</p>
      </header>
      <main>
        <DatasetPanel
          dataset={dataset}
          onLoaded={(ds) => {
            setDataset(ds);
            setFit(null);
          }}
        />
        {dataset && <GroupPanel dataset={dataset} onChange={setDataset} />}
        {dataset && <ModelPanel key={dataset.id} dataset={dataset} fit={fit} onFit={setFit} />}
        {dataset?.grouped && fit && <SavingsPanel fit={fit} dataset={dataset} />}
      </main>
    </div>
  );
}
