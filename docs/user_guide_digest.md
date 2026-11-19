# ENERGY EXPLORER User Guide: Technical Digest

> This digest of `src/EnergyExplorerUserGuide.pdf` was written during the port as a quick reference. Where it disagrees with the VB code, the port follows the code; see `PORTING.md`.

Source: `EnergyExplorerUserGuide.pdf`, 47 pages, "Version 2005-05-24" (the body says program version 20050518). Author: Kelly Kissock, copyright 2000. I read all 47 pages: text came out through macOS PDFKit, and pages 33, 41, 42, 44, 45 and 46 were rendered as images so I could read the equations and screenshots.

**Scope warning.** This 2005 guide does not cover a large part of what you asked about. It has nothing on the building simulation (CAV/VAV, economizer, hot-deck reset, latent loads), Durbin-Watson or autocorrelation, net determination bias, fractional savings uncertainty, "project savings", model projection, hourly profiles or 3D graphs, a formula syntax for new variables, or a separate "Group by DOW" weekday-type selection beyond weekday/weekend. If your VB6 source has these features, they were added after this guide or never documented, and you will need to take their behaviour from the code. Anything marked **[inferred]** is my reading and is not stated in the guide.

---

## 1. Input data format

- **Sources:**
  - (a) ASCII text file.
  - (b) An Enertel "Deskactive" database. This upload needs EE installed in the same directory as Deskactive, and a dialog where you pick site, channels and dates, dragging channels into "Selected Channels".
- **ASCII rules:**
  - Numbers only.
  - Every row has the same number of fields.
  - Fields are separated by **one or more spaces**. Commas and tabs are not mentioned.
  - There is **no header row**. Names come from a separate file.
  - One row is one record (time step), and one column is one channel.
- **No-data flag:**
  - The default is **-99** for ASCII files and **-9.99** for Enertel databases.
  - Any number can be set with *Data > Set No-Data Flag* or in the Define Variable Names dialog.
  - Missing values must be written as the flag so that all rows keep the same column count.
- **Names file:**
  - Same basename as the data file with extension **`.VAR`** (for example `DAILY2.DAT` goes with `DAILY2.VAR`), in the same directory.
  - One variable name per line, in column order.
  - It loads automatically when found. Otherwise the "Define Variable Names" dialog opens with default names, which you can change, Load, or Save-and-Load to a .var file.
  - *File > Open Names from Text File* loads a different names file.
- **Date and time columns** are recognised by their exact names. They can be in any order or position, and you only include the ones needed for the time scale (monthly data needs only Yr and Mo):

| Field | Values | Required name |
|---|---|---|
| Year | 1993 (4-digit strongly encouraged; some 2-digit years may work) | `Yr` |
| Month | 1–12 | `Mo` |
| Day | 1–31 | `Dy` |
| Hour | 0–23 or 1–24 | `Hr` |
| Minute | 1–60 | `Mn` |
| Day of week | 1–7, 1=Sun, 2=Mon… | any name |

- **Time intervals:** yearly, monthly, daily, hourly and sub-hourly are all supported. The interval is implied by which timestamp fields are present. No explicit interval field is described.
- **Sample files:** `DAILY2.DAT/.VAR`, `HOURLY.DAT/.VAR`, `MONTHLYENGY.DAT/.VAR`, `DAILYTEMPS.DAT/.VAR`.
  - DAILY2 has 9 columns: `Site, Mo, Dy, Yr, DOW, CW (MBtu/day), HW (MBtu/day), WBE (kWh/day), Temp (F)`.
  - Example row: `114 10 16 1990 1 61.8 27.23 -99 76`.
  - Names can contain spaces and parentheses.
- **On open**, the Value Window shows the file name, the number of records (rows) and the number of fields (columns).
- **Data lives in RAM.** Edits never change the source file.

## 2. Menus and controls

Menu bar from the screenshots: **File, Edit, Data, Graphics, Analysis, Help**. The guide also refers to a "Savings" menu (Savings > Calculate Savings, Savings > Define Measurement Error), while the overview says *Analysis > Calculate Savings*. **[inferred]** The savings items most likely sit under Analysis in this build.

**File**
- Open Data from Deskactive Database: Enertel upload dialog (site, channels, dates).
- Open Data from Text File: opens an ASCII data file and auto-loads the matching .VAR.
- Open Names from Text File: loads a different names file.
- Print: prints the Value Window and Graph Window to the default Windows printer.
- Save Data: writes the active data set to a file (Save As). The output format is not specified. **[inferred]** It is the same space-delimited ASCII format.
- Exit.
- (Editor: the guide lists it under both File and Edit.)

**Edit**
- Copy Value Window: copies the text to the clipboard.
- Copy Graph: copies the graph to the clipboard.
- Editor: built-in ASCII text editor, used for example to write .VAR files.

**Data**
- Name Variables: Define Variable Names dialog (rename, load or save names, change the no-data flag).
- Set No-Data Flag.
- Select Dates: sets a new Start and End date, and/or removes one or more blocks with "Delete From/To" plus Set. Removal cannot be undone except by reloading.
- View Dates: prints the start and end of the data set, the grouping, and removed blocks to the Value Window.
- Modify Variable: applies a built-in function to a variable in place (see §5).
- Create Variable: duplicates a variable, or combines two variables with an operator (see §5).
- Delete Variable.
- Group by Date: Group 1 is defined by a Start and/or End date, the rest is Group 2, and groups can be named.
- Group by DOW: automatic weekday/weekend grouping.
- Timescales…: changes the time scale with a sum or average choice per variable (see §5).
- Average Intermittent Data…: time-averages unevenly spaced data. The guide gives no algorithm.
- Merge With Previous Data File (see §5).

**Graphics**
- Single Scale (toggle): overlaid time-series traces share one Y scale.
- Zoom In (mode): drag a box to zoom, repeatable.
- Zoom Out: full-range time series of the current Y variable.
- Point (mode, the default): click a point to identify it.
- Grid (toggle): grid lines at the axis labels.
- Prediction Intervals (toggle): draws 95% prediction-interval lines with models.

**Analysis**
- Calculate Savings.
- Define Measurement Error (referred to as Savings > …).
- Calculate Lean Energy Breakdown.

**Help**: contents not described.

**Work-window controls**
- Value Window: text output, with a clear box.
- Graph Window: with a clear box.
- Y Variable and X Variable list boxes.
- Graph buttons: TimeSeries, "2" (overlay), XY, XY Animate, Histogram.
- Model buttons:
  - Mean / Median
  - 2P [M]
  - 3P-C [M] / 3P-H [M]
  - 4P [M]
  - 5P [M]
  - Bin Fit
- Grouping radio buttons: Single Group / G1|G2.
- Arrow keys step the point cursor to the next or previous point in time.

## 3. Regression models

General points:
- "P" is the number of regression coefficients.
- Notation: ( )+ means the term is set to 0 when negative, and ( )− means the term is set to 0 when positive.
- Models need an XY plot to be shown first. Mean and Median need only Y and a time-series or XY plot.
- With zero or one X variable, predictions are drawn as a solid line. With several X variables, predictions are drawn as heavy boxes.

| Model | Equation (guide notation) | Max X vars |
|---|---|---|
| Mean | Ymean = ΣY / n | 0 |
| Median | value with equal numbers of observations above and below it | 0 |
| 2P | Y = β1 + β2·X1 | 1 |
| 2P+MVR (MLR) | Y = β1 + β2X1 + β3X2 + … + β7X6 | **6** in §7 (the overview and dialog say "up to seven", X1..X7, which is a guide inconsistency) |
| 3PC | Y = β1 + β2(X1 − β3)+ | 1 |
| 3PH | Y = β1 + β2(X1 − β3)− | 1 |
| 3P+MVR | Y = β1 + β2(X1 − β3)± + β4X2 + β5X3 + β6X4 | 4 |
| 4P | Y = β1 + β2(X1 − β4)− + β3(X1 − β4)+ | 1 |
| 4P+MVR | … + β5X2 + β6X3 | 3 |
| 5P | Y = β1 + β2(X1 − β4)− + β3(X1 − β5)+ | 1 |
| 5P+MVR | … + β6X2 + β7X3 | 3 (guide text says "up to two independent variables" but the equation has X1..X3) |

What the parameters mean:
- 3P: β1 = Ycp (y of the change point), β2 = slope, β3 = Xcp.
- 4P: β1 = Ycp, β2 = left slope, β3 = right slope, β4 = Xcp.
- 5P: β1 = Ycp (the flat segment), β2 = left slope, β3 = right slope, β4 = Xcp1, β5 = Xcp2.
- In the MVR variants, X1 always carries the change-point structure and X2 onward enter linearly. Variables are assigned in the "Multiple Regression" dialog (Regressor Y, X1…Xk, then Variable Name, then Set).

**3P algorithm (indicator-variable regression):**
- y = a + b·I·(x − xcp).
- 3PC: I = 0 when x ≤ xcp, and I = 1 when x > xcp.
- 3PH: I = 1 when x ≤ xcp, and I = 0 when x > xcp.
- a and b are found by OLS at each candidate xcp.

**4P algorithm:**
- y = a + b·x + c·I·(x − xcp), with I = 0 when x ≤ xcp and I = 1 when x > xcp.
- This reduces to y = a + b·x for x ≤ xcp, and y = (a − c·xcp) + (b + c)·x for x > xcp.
- **[note]** As written, `a` is the intercept at x = 0 and not literally the y at the change point. Yet the guide says the SE of the y change point is reported as SE(a). Check the VB code to see whether x is centred on xcp, i.e. y = a + b(x − xcp) + c·I(x − xcp).

**5P algorithm:**
- Y = β1 + β2(X − Xcp1)·I1 + β3(X − Xcp2)·I2.
- I1 = 1 when X ≤ Xcp1, otherwise 0.
- I2 = 1 when X > Xcp2, otherwise 0.
- This reduces to:
  - β1 + β2(X − Xcp1) for X ≤ Xcp1
  - β1 for Xcp1 < X ≤ Xcp2
  - β1 + β3(X − Xcp2) for X > Xcp2

**Change-point search (two-pass grid, minimum SSE):**
1. Pass 1: split [xmin, xmax] into 10 equal sub-intervals. Each sub-interval boundary is a candidate xcp. **[inferred]** That gives 11 boundaries; whether the endpoints are included is unspecified. Fit at each candidate and keep the xcp with the lowest SSE.
2. Pass 2: take the span from one sub-interval left to one sub-interval right of the best xcp and split it into 10 new sub-intervals. If the best xcp was an endpoint of the original range, only the one adjacent sub-interval is subdivided. Repeat the fits and pick the lowest SSE.
3. 5P: pass 1 tries all pairs of boundaries with Xcp2 > Xcp1. Pass 2 refines a ±1 sub-interval window around each change point into 10 sub-intervals each and tries all pairs with Xcp2 > Xcp1.
4. MVR variants: same search on X1. **[inferred]** The extra X columns are included in each OLS fit.

**Standard errors reported:**
- 3P: SE(Ycp) = SE(a) and SE(slope) = SE(b).
- 4P: SE(Ycp) = SE(a), SE(LS) = SE(b), and SE(RS) = √(SE(b)² + SE(c)²).
- 5P: SE(Ycp) = SE(β1), SE(LS) = SE(β2), SE(RS) = SE(β3). The text has a typo, "β23".
- Xcp in every model: SE = width of the smallest search sub-interval, i.e. the pass-2 step.
- Each SE is printed in parentheses after its parameter.
- Interpretation given: P_true = P_est ± t(1−α/2, n−p)·SE(P).

**Bin Fit:**
- The user gives the number of X bins and Y is averaged within each bin.
- The graph joins the bin means with a line.
- The Value Window lists one pair per bin: (bin midpoint, mean Y).
- **[inferred]** Bins are equal width over [xmin, xmax].
- The idea is credited to J. Haberl.

**Statistics (p = number of regression coefficients, ŷ = prediction, ȳ = mean):**
- N = number of observations used.
- Ymean = mean of Y.
- StdDev = [Σ(y − ȳ)² / (n − 1)]^½
- CV-StdDev = StdDev / ȳ × 100 (%)
- R2 = 1 − Σ(y − ŷ)² / Σ(y − ȳ)²
- RMSE = [Σ(y − ŷ)² / (n − p)]^½
- CV-RMSE = RMSE / ȳ × 100 (%)
- 95% prediction interval: PI = t(1−α/2, n−p) · RMSE · (1 + 1/n + X′)^½ with α = 0.05. X′ (the leverage / distance-from-mean term) is **assumed to be 0**, so the band is Ψ ± PI, a constant width.
- Parameter labels printed: `N, Ymean, Yint, Xn, LS, RS, Xcp, Xcp1, Xcp2, Ycp`.

**Output line format** (from the screenshots):
- `Baseline Model Stats: Type = 4P  N = 250  R2 = 0.90  RMSE = 5.33  CV-RMSE = 10.3%`
- `Baseline Model Eqn: CW (MBtu/day) = 47.64 - 0.74 (66.50 - Temp (F))+ + 1.87 (Temp (F) - 66.50)+`
  - The left-slope term is shown as −LS·(Xcp − X)+.

**Residuals:**
- residual = y_obs − y_pred.
- Stored automatically as new variables `G1_Resids`, `G2_Resids`, `SG_Resids`, overwritten by each model run.
- They can be graphed and modelled like any other variable.

**Groups:**
- In G1|G2 mode, the first model click fits Group 1 and the second click fits Group 2.
- **[inferred]** The clicks alternate after that.

**Not in guide:** Durbin-Watson, autocorrelation, net determination bias, t-stats or p-values, F-statistic, adjusted R².

## 4. Savings and Lean Energy Analysis (LEA)

**Workflow:**
1. Group the data by date into pre and post periods (G1 = pre).
2. Select G1|G2.
3. Fit the baseline model on the pre-retrofit group (any model: mean, regression or change-point).
4. Run Calculate Savings.

**Savings:**
- ΣE_save,j = ΣÊ_pre,j − ΣE_meas,j, summed over j = 1..m.
- m = number of post-retrofit observations.
- Ê_pre = the pre-model's prediction at post-period conditions.

**Uncertainty (95%):**
- ε_pre = 1.96 · RMSE · [m·(1 + 2/n)]^½
  - RMSE comes from the pre model, and n = number of pre-retrofit observations.
- ε_meas = (percent error) · ΣE_meas,j
  - The default percent error is **2%**, changeable with Define Measurement Error.
- ε_save = (ε_pre² + ε_meas²)^½
- The guide has **no autocorrelation correction** and no fractional-savings-uncertainty formula. Percent uncertainty is only shown in the output (see below).

**Output** (screenshot, 4P baseline; each line is value ± absolute uncertainty ± percent):

```
Observations During Post-retrofit Period: 416
Adjusted Baseline Use During Post-retrofit Period = 25,244.3 ±213.9 ±0.8%
Actual Use During Post-retrofit Period = 13,500.3 ±0.0 ±0.0%
Total Savings During Post-retrofit Period = 11,744.0 ±213.9 ±1.8%
Average Savings Per Observation During Post-retrofit Period = 28.2 ±0.5 ±1.8%
```

- **[inferred]** Average = Total / m, and its ± is ε_save / m.
- **[inferred]** The percent is ε / value.
- In this screenshot ε_meas is shown as 0, so the measurement error appears to have been set to 0.
- The graph overlays the measured series with a dashed "Adjusted Baseline" (the pre-model's prediction) over the post period.

**LEA ("Lean Energy Analysis"):**
- Splits total energy into three parts:
  - **Facility**: independent of weather and production.
  - **Weather-dependent.**
  - **Production-dependent.**
- Procedure: fit a **3PH-MVR or 3PC-MVR** with Y = energy, X1 = temperature, X2 = production. Then run *Analysis > Calculate Lean Energy Breakdown*.
- Results are shown as numbers and as a time-series graph with three dashed component lines ("Fac Use", "Wea-dep Use", "Prod-dep Use") plus the total.
- Screenshot, monthly gas (N = 12):
  - Equation: `Nat Gas = 59.58 - -9.37 (62.06 - Toa)+ + 0.00 (Toa - 62.06)+ + 0.02 Prod`
  - Breakdown: `Facility = 715.0 (13.6 %)  Wea-dep = 1,441.4 (27.5 %)  Prod-dep = 3,083.6 (58.8 %)  Total = 5,240.0 (100%)`
  - Note that the 3P model is printed in 4P-style two-slope form.
- **[inferred, consistent with the numbers: 59.58 × 12 ≈ 715]** The components are:
  - Facility = Σ β1, i.e. β1 · N.
  - Weather = Σ of the change-point slope term.
  - Production = Σ β4 · X2.
  - Total = the sum of the three, i.e. the sum of predictions, not the measured total.
  - Percentages are each component / Total.

**Not in guide:** "project savings" and model projection. The closest thing is that savings apply the pre model to post-period conditions.

## 5. Data transforms

**Change time scale (Data > Timescales…):**
- Aggregates up to a coarser scale, for example hourly to daily, monthly or yearly.
- You can return to the original scale at any time. **[inferred]** The original data is kept in memory.
- Each variable is **summed** by default. Variables moved to the "Average" list are averaged. The example given: sum energy, average temperature.
- No-data handling is a global option:
  - "If NoData then set Avg/Total to NoData", or
  - "If NoData then calculate Avg/Total using available data".
- **[inferred]** Periods follow the calendar boundaries defined by the timestamp fields.

**Average Intermittent Data:** time-average of unevenly spaced records. No details are given.

**Group by Date:**
- Group 1 is defined by a Start and/or End date, and everything else is Group 2.
- Groups can be given custom names.
- All graphs colour the groups differently (blue and red in the screenshots). Models fit one group at a time.
- The date entry format is shown in the dialog but not in the guide. The screenshot axis shows m/d/yyyy.

**Group by DOW:** automatic weekday vs weekend groups. **[inferred]** This uses the timestamp, or possibly a 1=Sun day-of-week column.

**Select Dates:** trims the data to a start/end window and deletes listed blocks permanently.

**Modify Variable:**
- Applies a function to a variable in place.
- No-data values are left unchanged.
- Built-in functions:
  - F to C, C to F
  - ( )^c, ( )+c, ( )−c, ( )*c, ( )\c (divide)
  - Abs, ln, Exp, Sin, Cos (radians)
  - Degs↔Rads
  - MBtu→kWh, kWh→MBtu (the guide mislabels these as "Kilowatts")
- **[inferred]** Conversion constants: F→C = (F − 32)/1.8, and 1 MBtu = 293.07 kWh, taking MBtu as 10⁶ Btu.

**Create Variable:**
- Option 1: duplicate an existing variable.
- Option 2: `VarA <op> VarB` with op ∈ {×, ÷, +, −}, built by clicking.
- There is **no free-text formula syntax**.
- The result is no-data if either input value is no-data.
- More complex expressions come from repeating Modify and Create.

**Delete Variable:** removes a variable from the active set.

**Merge With Previous Data File:**
1. Open file A, then open file B, then run the command.
2. The file with the finer time resolution (for example daily temperatures) is merged into the coarser one (for example monthly bills).
3. For each coarse period, the **mean** of every non-timestamp field in the fine file is computed and added as a new variable.
4. Both files must use the standard timestamp names. Save the result with Save Data and the names with Name Variables.
5. **[inferred]** Bill periods are calendar periods from the Yr/Mo/Dy fields. Meter-read date ranges are not described.

## 6. Building simulation

**Not present in this guide.** It has no mention of CAV/VAV, economizer, hot-deck reset, latent loads, or any simulation inputs or equations. Take these from the VB6 source or another document. The only HVAC remark is that a 4P change point can match a cooling-coil dew point or a VAV set point (Kissock et al. 1999), and that VAV fan power can look U-shaped (suits 5P).

## 7. Graphs

- **Time series:**
  - Y against date-time.
  - The "2" button overlays more traces, with the right axis showing the most recently added variable.
  - Single Scale forces one shared Y axis.
  - Zoom In/Out.
  - Groups are colour-coded.
- **XY:** scatter of Y vs X. This is the base plot for model lines.
- **XY Animate:**
  - Plots points in time order.
  - Animation Control dialog settings:
    - Speed.
    - Number of Points To Display: a trailing window.
    - Clear Screen: off keeps existing items such as a model line.
    - Begin and Quit buttons.
  - An animation cannot be stopped once started.
- **Histogram:**
  - The user enters Ymin, Ymax and the number of intervals. Width = (Ymax − Ymin) / N.
  - Left axis = count, right axis = probability.
  - In G1|G2 mode the two groups are shown side by side, with probabilities labelled "G1/G2".
  - The Value Window lists (interval #, count, frequency) for each interval.
- **Point identification:**
  - Click a point to show its date and value in the Value Window.
  - Overlapping points are all listed, and the cursor goes to the last one.
  - The ← and → keys step the cursor through time.
- **Grid, prediction-interval bands, model lines/boxes:** see §2 and §3.
- **Not in guide:** hourly profiles, 3D or contour plots.

## References cited in the guide

- PRISM, the basis of the 3P and 5P models: Fels 1986, *Energy and Buildings* 9(1–2); Fels, Kissock et al. 1995.
- 4P model: Ruch and Claridge 1991, ASME Solar Engineering.
- Animate: Sparks and Haberl 1993.
- Statistics: Miller 1981; Neter, Wasserman and Kutner 1989.
