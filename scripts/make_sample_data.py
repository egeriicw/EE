"""Generate the SYNTHETIC sample data sets in ``data/`` (deterministic, seed 42).

These are NOT measured data. They mimic the layout of the original ENERGY EXPLORER sample files
(DAILY2.DAT / HOURLY.DAT: space-delimited numbers, no header, names in a ``.var`` sidecar,
-99 = no data) so the notebook, tests and API have something realistic to load.

* ``sample_daily.dat`` — two years of daily data for one building. 2019 is the "pre" period,
  2020 the "post" period after a retrofit that lowered the cooling change point slope and base.
  CW (MBtu/day) follows a 3PC model, HW (MBtu/day) a 3PH model, WBE (kWh/day) depends on weekday.
* ``sample_monthly_lea.dat`` — 24 months of gas use vs temperature and production, for Lean Energy
  Analysis (3PH-MVR).
* ``sample_hourly.dat`` — one year of hourly Toa / Woa / WBE for the HVAC simulation.

Run:  python scripts/make_sample_data.py
"""

from __future__ import annotations

from pathlib import Path

import numpy as np
import pandas as pd

OUT = Path(__file__).resolve().parents[1] / "data"
ND = -99


def _write(df: pd.DataFrame, stem: str) -> None:
    OUT.mkdir(exist_ok=True)
    df.fillna(ND).to_csv(OUT / f"{stem}.dat", sep=" ", header=False, index=False,
                         float_format="%.3f")
    (OUT / f"{stem}.var").write_text("\n".join(df.columns) + "\n")


def daily(rng: np.random.Generator) -> None:
    dates = pd.date_range("2019-01-01", "2020-12-31", freq="D")
    doy = dates.day_of_year.to_numpy()
    temp = 55 + 22 * np.sin(2 * np.pi * (doy - 105) / 365) + rng.normal(0, 6, len(dates))
    post = dates.year == 2020
    # cooling: 3PC, change point 62 F
    base_cw = np.where(post, 22.0, 30.0)
    slope_cw = np.where(post, 1.3, 1.8)
    cw = base_cw + slope_cw * np.clip(temp - 62, 0, None) + rng.normal(0, 3.0, len(dates))
    # heating: 3PH, change point 58 F
    hw = 15 + 1.1 * np.clip(58 - temp, 0, None) + rng.normal(0, 2.0, len(dates))
    weekday = dates.day_of_week < 5
    wbe = np.where(weekday, 42_000, 28_000) * np.where(post, 0.93, 1.0) + rng.normal(0, 900, len(dates))
    df = pd.DataFrame({
        "Site": 114,
        "Mo": dates.month,
        "Dy": dates.day,
        "Yr": dates.year,
        "DOW": (dates.day_of_week + 1) % 7 + 1,  # 1 = Sunday
        "CW": cw.round(2),
        "HW": hw.round(2),
        "WBE": wbe.round(0),
        "Temp": temp.round(1),
    })
    # sprinkle missing values like a real logger
    for col in ("CW", "WBE", "Temp"):
        idx = rng.choice(len(df), 8, replace=False)
        df.loc[idx, col] = np.nan
    _write(df, "sample_daily")


def monthly_lea(rng: np.random.Generator) -> None:
    months = pd.date_range("2018-01-01", periods=24, freq="MS")
    toa = 52 + 24 * np.sin(2 * np.pi * (months.month.to_numpy() - 4) / 12) + rng.normal(0, 2, 24)
    prod = rng.uniform(2_000, 6_000, 24).round(0)
    gas = 60 + 9.0 * np.clip(62 - toa, 0, None) + 0.02 * prod + rng.normal(0, 8, 24)
    df = pd.DataFrame({"Mo": months.month, "Yr": months.year, "Toa": toa.round(1),
                       "Prod": prod, "NatGas": gas.round(1)})
    _write(df, "sample_monthly_lea")


def hourly(rng: np.random.Generator) -> None:
    ts = pd.date_range("2019-01-01 00:00", "2019-12-31 23:00", freq="h")
    doy = ts.day_of_year.to_numpy()
    hr = ts.hour.to_numpy()
    toa = (55 + 22 * np.sin(2 * np.pi * (doy - 105) / 365) + 8 * np.sin(2 * np.pi * (hr - 9) / 24)
           + rng.normal(0, 3, len(ts)))
    woa = np.clip(0.004 + 0.012 * (toa - 20) / 80 + rng.normal(0, 0.001, len(ts)), 0.001, 0.02)
    occupied = (ts.day_of_week < 5) & (hr >= 7) & (hr < 19)
    wbe = np.where(occupied, 1_600, 700) + rng.normal(0, 40, len(ts))
    df = pd.DataFrame({"Mo": ts.month, "Dy": ts.day, "Yr": ts.year, "Hr": hr + 1,
                       "Toa": toa.round(1), "Woa": woa.round(5), "WBE": wbe.round(0)})
    _write(df, "sample_hourly")


if __name__ == "__main__":
    rng = np.random.default_rng(42)
    daily(rng)
    monthly_lea(rng)
    hourly(rng)
    print(f"wrote synthetic sample data to {OUT}")
