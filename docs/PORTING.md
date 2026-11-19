# Porting notes: ENERGY EXPLORER (VB6) → `energy_explorer` (Python)

Source reviewed: `src/EnergyExplorer.vbp` (version 5.0, Dept. of Mech. Eng., University of Dayton). The project
contains about 17k lines of VB:

- `MAIN.FRM` (8.5k lines): every menu and button handler
- `SUBS.BAS` (2.5k lines): the regression engine, change-point searches and date helpers
- `GLOBAL.BAS`: about 150 global variables
- 25 dialog forms

The 2005 user guide (`src/EnergyExplorerUserGuide.pdf`) was also read. **Where the guide and the code disagree, the
port follows the code**, and each case is noted below.

## How the VB program is organised (and how the port differs)

The VB program keeps all its state in globals:

- data matrix `d(rec, fld)`, names `vn$()`, groups `grpd()`
- the current model in `modtype$`
- the baseline model's parameters in the `*pre` globals: `betapre()`, `xcppre`, `ycppre`, `ivcoefspre()`, …

Button handlers mix calculation, drawing on the `GraphBox` picture control, and printing to the `ValueBar` text box.
Missing data uses a sentinel, `ndflag` (default -99), which is checked by hand everywhere.

The port separates these concerns:

| Concern | VB | Python |
|---|---|---|
| Data + metadata | `d()`, `vn$()`, `numrecs`, `timeint`, `grpd()`, `g1name$` … | `Dataset` (immutable; a DataFrame plus groups). NaN replaces the sentinel and is converted back only at I/O |
| Regression engine | `Reg`, `Inf`, `Invert`, `Mult`, `Trans` | `regression.ols` (numpy `lstsq` + `pinv`) |
| Fitted baseline | `premodtype$`, `betapre()`, `xcppre`, … | Model dataclasses with `predict` / `to_dict` |
| Output text | `ValueBar.Print` | `model.summary()`, `SavingsResult.report()` |
| Graphics | `GraphBox` drawing code | not ported; matplotlib in the notebook and Recharts in `web/` |

## Procedure map

| VB procedure (file) | Python | Notes |
|---|---|---|
| `OpenDataCmd_Click` (MAIN) | `io.read_data`, `io.read_data_text` | Delimiter is detected from line 1: comma, else tab, else whitespace |
| `OpenVar` (SUBS) | `io.read_var_names`, `Dataset.time_fields` | `.var`/`.dvn` sidecar files. `yr/mo/dy/hr/mn` are recognised by name. Two-digit years are expanded |
| `OpenDataOldCmd_Click`, `OpenDailyCmd_Click` | — | Older open paths; superseded by `read_data` |
| `SaveDataCmd_Click`, `SaveNamesCmd_Click` | `io.write_data`, `io.write_var_names` | |
| `CalendarDate`, `SerDate`, `SerialDate`, `TxtToSer`, `SerToCal` | `Dataset.serial_dates`, `grouping.to_serial` | Serial dates use the yyyymmddhh form |
| `GroupByDateCmd_Click` + `AddGrpDB` | `grouping.group_by_date` | Group 1 = start ≤ date ≤ end (inclusive) |
| `GroupByDOWCmd_Click` | `grouping.group_by_day_of_week` | Sat/Sun → group 2 ("WE") |
| `TimeScaleChange`, `ToDaily/ToMonthly/ToYearlyCmd_Click`, `TSOptsDB` | `timescale.change_timescale` | See *Time scales* |
| `ToHourlyCmd_Click` (restore original) | keep the original `Dataset` | The API returns a new dataset id instead |
| `AvgIntDataCmd_Click` | `timescale.average_intermittent` | |
| `NewVarDB` (Create Variable) | `variables.create_variable` | Copy, or `A op B` |
| `ModValDB` (Modify Values) | `variables.modify_values`, `MODIFIERS` | |
| `DeleteDB` | `variables.delete_variable` | |
| `AddVar` | `Dataset.with_data` | |
| `SetAll`, `SetGrp`, `FillXYGrp` | `modeling.select_xy` | Selects rows where y and every x are present and the row is in the chosen group |
| `MeanBut_Click` | `changepoint.fit_mean` | |
| `MedianBut_Click`, `Median`, `Percentile` | `changepoint.fit_median`, `stats.vb_percentile`, `stats.median_with_se` | |
| `TwoPBut_Click`, `TwoP` | `changepoint.fit_linear` | |
| `MLRBut_Click` + `MultRgDB` | `changepoint.fit_linear` (k > 1) | |
| `ThreePBut_Click`, `ThreePMVRBut_Click`, `ThreePMVR` | `changepoint.fit_3p` | |
| `FourPBut_Click`, `FourPMVRBut_Click`, `FourPMVR` | `changepoint.fit_4p` | The plain button and the MVR button use different grid sizes (see below) |
| `FivePBut_Click`, `FivePMVRBut_Click`, `FivePMVR` | `changepoint.fit_5p` | |
| `BinFitBut_Click` + `BinFitDB` | `changepoint.fit_binfit`, `BinFitModel` | |
| `DWTest` | `stats.dw_test` | |
| `CalcSavingsCmd_Click` | `savings.calculate_savings` | The savings shown in the UI |
| `ProjectSavingsCmd_Click`, `CalcSavingsPECmd_Click`, `WYZCalc` | `savings.autocorrelated_savings_uncertainty` | Mean/2P/MLR only, as in VB |
| `DefineMeasurementErrorCmd_Click` + `MeasErr` | `measurement_error_pct=` argument | |
| `CalcLEACmd_Click` | `savings.lean_energy_analysis` | |
| `CreateModelProjectionCmd_Click` | `savings.create_model_projection` | |
| `RunSimulationCmd_Click` + `SystemDB` + `LoadDB` | `simulation.simulate`, `SystemParams`, `add_simulation` | `SIMULATE.BAS` is a binary-saved copy of ESIM.BAS and is not in the `.vbp` |
| `MergeCmd_Click`, `SaveFileParams`, `CalendarDate2` | **not yet ported** | Merging a fine file into a coarse one: average the fine file over each coarse period. Planned next |
| `SelDatesCmd_Click` + `SelDatDB`, `ViewDatesCmd_Click` | **not yet ported** | Trim to a date window. Easy to add with pandas filtering |
| `CalcTotalCmd_Click` | not wrapped | Prints N and Σy per group: `ds.data.loc[ds.group_mask(g), y].sum()` |
| `SetNodataFlagCmd_Click` | `read_data(no_data=…)` | |
| `TimeSeries*`, `XY*`, `Hist*`, `HourlyProfiles*`, `D3Line`, `DrawModelLine`, `Cursor`, `Zoom*`, `GraphBox_*`, `SetScalesDB`, `RotateDB`, `AnCntlDB` | not ported (UI) | Replaced by the notebook and the web charts |
| `ESEditor`, `FileOpen/FileSave`, `FindDB`, `NotSave`, `PerDone`, `UpdatePerDone`, `Print*`, `Copy*`, `ShowHelpCmd`, `CONTXTID.BAS`, `EEHELP.*` | not ported (UI) | |
| Enertel/Deskactive import (`Config.cfg`) | not ported | The guide describes it, but no database-access code was found in the VB source (grep for `mdb`, `OpenDatabase`, `Recordset`) |

## Algorithms as implemented (matching the VB code)

**Variable limits:** the Multiple Regression dialog offers X1..X7, and `Invert` handles at most 8 regression columns. That gives up to 7 X variables for MLR and 3P-MVR, and 6 for 4P/5P-MVR. The guide quotes lower limits (4/3/3); the port follows the code.

**Regression statistics (`Inf`)**

- p is the number of regression columns. Change points are *not* counted in p.
- RMSE = √(SSE/(n − p)), and 0 when n = p.
- CV-RMSE = RMSE/ȳ·100.
- R² = 1 − SSE/SST.
- adj R² = 1 − (n−1)/(n−p)·SSE/SST.
- Lag-1 autocorrelation = Σ rᵢ₋₁rᵢ / Σ rᵢ₋₁².
- DW = Σ(Δr)² / Σr².
- Coefficient SE = RMSE·√diag((XᵀX)⁻¹).

**Change-point search** (the guide says "10 intervals"; the code differs, and the port follows the code):

- 3PC/3PH (± MVR):
  - Coarse pass: 100 intervals over [xmin, xmax], trying cp = xmin + i·inc for i = 1..99.
  - Fine pass: start at best − inc, use step 2·inc/100, and try i = 1..99 again.
  - The incumbent is reset at the start of each pass.
- 4P:
  - Same scheme as 3P.
  - `FourPBut_Click` uses **10** intervals; `FourPMVR` uses **100**. `fit_4p` reproduces both through its
    `n_intervals` default.
- 5P:
  - Coarse pass: 9 intervals, all pairs from the 8 interior points with cp1 < cp2.
  - Fine pass: each change point ± 0.333·inc and ± 0.666·inc.
- Selection is by minimum RMSE. The first candidate wins ties.
- A point with x == cp counts as left of the change point (`x <= cp`).
- The "standard error" of Xcp is the final grid step, not a statistical SE. The port names it `xcp_step`.
- 4P:
  - ycp = b0 + b1·xcp
  - ls = b1
  - rs = b1 + b2
  - SE(rs) = √(SE(b1)² + SE(b2)²)
  - SE(ycp) uses the VB propagation formula, which ignores the b0–b1 covariance.
- Prediction interval: (1.96 + 2.7/(n − p_total)) · RMSE · √(1 + 1/n), where p_total counts change points.

**Percentiles** use position (n+1)·p, interpolated linearly and clamped to the data range. This is not numpy's
default method. SE(median) = (P75 − P25)/√n.

**Bin Fit**

- Bins are equal-width over [xmin, xmax].
- Bin 1 is closed, [lo, hi]. Every other bin is (lo, hi].
- An empty bin predicts:
  - for bin 1, the next non-empty bin;
  - otherwise, the average of the nearest non-empty neighbours on each side (or the one side that exists).
- RMSE divides by n − 1. The VB code divides by `n - j`, but `j` is a leaked loop index that is 1 in practice.

**Savings (`CalcSavingsCmd_Click`)**

- Savings = Σ prediction − Σ measured, over Post records where y and the model inputs are present.
- ε_baseline = 1.96·RMSE·√(1+2/n_pre)·√n_post.
- ε_meas = Σmeasured·pct/100.
- ε_savings is the two combined in quadrature.
- The measurement error defaults to 0, which is the VB global's value until "Define Measurement Error" is run. That
  dialog proposes 5%; the guide says 2%.

**AR(1)-corrected uncertainty (`ProjectSavingsCmd_Click`)**

- Ports the c11/c12/c22/c23 variance terms and the effective-degrees-of-freedom correction (n − trace).
- Regressors are centred on the Pre means.
- `WYZCalc` is evaluated in O(n) with forward and backward AR recursions, so hourly data does not need an n×n matrix.

**LEA**

- 3PC/3PH(-MVR) models only.
- Uses every record with complete inputs (not just one group). This matches VB.
- Results are per-observation averages of three parts:
  - facility = ycp
  - weather = slope·(x − xcp)±
  - production = Σ bₖ·xₖ

**Time scales**

- A new output record starts whenever the (yr[, mo[, dy]]) key changes between consecutive records.
- Every variable is totalled by default. Variables listed in `averages=` are averaged.
- With `strict_no_data=True`, any missing value makes the whole period missing. A period with no data at all is
  always missing.
- Groups are dropped, as in VB.

## Deliberate deviations (bugs or hazards in the VB code)

| # | VB behaviour | Port |
|---|---|---|
| 1 | `Invert` swaps columns wrongly when undoing pivoting: `B(k, irow)` is assigned twice and `B(k, icol)` never. The bug is **latent**: XᵀX is symmetric positive-definite, so full pivoting picks diagonal pivots and no swap happens, except on exact ties in degenerate (collinear) designs | numpy `pinv`; no difference expected in normal use |
| 2 | `Inf` tests `pdemnom` (typo) instead of `pdenom`, so the residual autocorrelation it returns (`roe`) is never set. 2P/MLR "project savings" therefore ran with ρ = 0 | ρ is estimated by default; pass `rho=0` to reproduce VB |
| 3 | `CreateModelProjection`, Median case, projects `ymeanpre` (the mean) | Uses the median |
| 4 | `ModValDB` Degs↔Rads uses π ≈ 3.1428 | `math.pi` |
| 5 | Simulation: the optimised hot-deck reset uses `foa` *before* the economizer computes it for the current hour, so it gets the previous hour's value (and 0 on the first hour) | Economizer is computed first |
| 6 | Simulation reads inputs from hard-coded columns 9/5/6/7 (`hourly.dat` layout) | Columns are passed by name |
| 7 | `ProjectSavingsCmd_Click` 3P/5P branches reference variables that are never set (`hcindex`, `cp1pre`, `cp2pre`, `betapre` for 3P) | Not ported; `calculate_savings` covers every model type |
| 8 | Time-scale periods are detected by a change in a *single* field (e.g. `dy`). Consecutive records with equal day numbers in different months would merge | The full (yr, mo, dy) key is used. Results are identical on normally ordered data |
| 9 | `FourPBut_Click` (non-MVR) uses an uninitialised `sexcp` (= 0) in the SE(ycp) formula; `FourPMVR` uses the grid step | The grid step is always used |
| 10 | Mean/Median use only rows where the currently selected X is also present (`SetAll`) | `x` is optional for Mean/Median; if given, it restricts rows the same way |
| 11 | `AvgIntDataCmd_Click` blanks *every* non-time variable in record 1 | Only the variables being averaged are blanked |
| 12 | `SerialDate` with hourly data and a date that has no hour (`m/d/yyyy`) misparses the year as the hour | Raises an error asking for `m/d/yyyy/h` |

## Unverified items / next steps

1. **Golden-file parity.** So far, only two numbers from the real program are tested: the savings and LEA screenshots in the user guide. `tests/test_core.py::test_reference_values_from_user_guide_screenshots` reproduces both. Run `EnergyExplorer.exe` (Windows/VM) on `data/*.dat` and on the 1995 file `dhr.tmp` in
   `EnergyExplorerSource.ZIP`. Capture the Value Window output for each model and for savings, then add the results
   as regression fixtures under `tests/`.
2. Port `Merge` and `Select Dates`, and the hourly-profile / histogram analyses as data functions with no plotting.
3. API: persistent store (e.g. SQLite/Postgres + Parquet), authentication, background jobs for large hourly fits,
   an OpenAPI-generated TypeScript client.
4. Web: dataset table and time-series explorer, a variable editor, time-scale and simulation forms, a model
   comparison view.
