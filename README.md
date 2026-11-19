# Energy Explorer — Python port

**ENERGY EXPLORER** is a Visual Basic 6 desktop program for analysing measured building-energy data. It was written
by Kelly Kissock and colleagues, Dept. of Mechanical Engineering, University of Dayton. Its features:

- change-point regression baselines (Mean, Median, 2P, MLR, 3PC/3PH, 4P, 5P, each with optional extra regressors, and Bin Fit)
- measurement-and-verification savings with uncertainty
- Lean Energy Analysis
- a simplified two-zone HVAC simulation

This repository holds the original VB source plus a Python port, laid out so it can grow into an API and a web app.

```
src/               original VB6 source, help files, user guide (untouched)
energy_explorer/   Python library (pure numpy/pandas, no UI)
notebooks/         Jupyter walkthrough of the whole workflow
api/               FastAPI service wrapping the library
web/               React + Vite + TypeScript client for the API
tests/             pytest suite (library + API)
data/              SYNTHETIC sample data (see scripts/make_sample_data.py)
docs/PORTING.md    VB procedure → Python mapping, deviations, and known VB quirks
docs/user_guide_digest.md  technical digest of the 2005 user guide
```

## Quick start

```bash
uv venv --python 3.12 .venv
```

```bash
uv pip install -p .venv/bin/python -e ".[api,notebook,dev]"
```

```bash
.venv/bin/python -m pytest
```

### Library

```python
import energy_explorer as ee

ds = ee.read_data("data/sample_daily.dat")               # auto-loads sample_daily.var; -99 -> NaN
ds = ee.group_by_date(ds, end="12/31/2019", names=("Pre", "Post"))
model = ee.fit(ds, "3PC", y="CW", x=["Temp"], group="Pre")
print(model.summary())
print(ee.calculate_savings(model, ds, measurement_error_pct=5).report())
```

### Notebook

```bash
.venv/bin/jupyter lab notebooks/energy_explorer_walkthrough.ipynb
```

### API + web app

```bash
.venv/bin/uvicorn api.app.main:app --reload
```

```bash
cd web && npm install && npm run dev
```

The API's interactive docs are at http://localhost:8000/docs and the web app runs at http://localhost:5173. Vite
forwards `/api` requests to port 8000.

## Architecture

- **`energy_explorer`** holds all the numerical behaviour of the VB program. Its functions operate on an immutable
  `Dataset` (a pandas DataFrame plus groups and metadata) and return new objects. Nothing depends on global state,
  so the same calls work from notebooks, scripts and the web API.
- Fitted models are dataclasses. Each one has `predict()`, `equation()`, `summary()`, and
  `to_dict()` / `model_from_dict()`. A baseline can therefore be stored as JSON and reused later to compute savings.
- **`api/`** is a thin FastAPI layer. It contains no maths: it validates requests, calls the library, and serialises
  the results. Storage sits behind a `Store` protocol (`api/app/store.py`). The in-memory implementation can be
  swapped for a database or object store without changing the routers.
- **`web/`** is a React client for the standard M&V workflow: load data → group Pre/Post → fit baseline →
  calculate savings, with charts drawn by Recharts.

## Status and caveats

- **Full parity with `EnergyExplorer.exe` is not verified.** The only reference outputs available are two
  screenshots in the user guide. The tests reproduce both: the savings uncertainty ±213.9 and the LEA facility
  total 715.0. The algorithms are ported structurally from the VB code: the same grid searches, indicator conventions,
  statistics and uncertainty formulas. Tests check the formulas directly and confirm that the models recover known
  parameters from synthetic data. To verify parity, run the exe and the port on the same files and compare.
- The data in `data/` is **synthetic** (`scripts/make_sample_data.py`).
- Graphics, zoom, the text editor, printing, help and the Enertel database import were not ported. Matplotlib (in
  the notebook) and the web client replace the graphics. See `docs/PORTING.md`.
