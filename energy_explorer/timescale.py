"""Changing time scales (port of ``TimeScaleChange`` in SUBS.BAS and ``AvgIntDataCmd_Click``)."""

from __future__ import annotations

from collections.abc import Iterable

import numpy as np
import pandas as pd

from energy_explorer.dataset import Dataset, TimeInterval

_TARGETS = {
    "yearly": (TimeInterval.YEARLY, ("yr",)),
    "monthly": (TimeInterval.MONTHLY, ("yr", "mo")),
    "daily": (TimeInterval.DAILY, ("yr", "mo", "dy")),
}


def change_timescale(
    ds: Dataset,
    to: str,
    *,
    totals: Iterable[str] | None = None,
    averages: Iterable[str] | None = None,
    strict_no_data: bool = False,
) -> Dataset:
    """Aggregate to a coarser time scale ("daily", "monthly" or "yearly").

    Each non-time variable is either **totalled** or **averaged** over the period. VB's options
    dialog defaults every variable to "total"; pass ``averages=[...]`` (e.g. temperatures) to average
    those instead, or ``totals=[...]`` to total only those and average the rest.

    Missing values are skipped. With ``strict_no_data=True`` ("If NoData then set Avg/Total to
    NoData") any missing value in a period makes that period's result missing. A period with no
    data at all is always missing.

    Periods are consecutive runs of records with the same (yr[, mo[, dy]]) key, as in VB, which
    starts a new output record whenever the period field changes. Grouping is not preserved: VB
    discards groups when the time scale changes.
    """
    if to not in _TARGETS:
        raise ValueError(f"'to' must be one of {list(_TARGETS)}")
    target, keys = _TARGETS[to]
    if ds.interval <= target:
        raise ValueError(f"Data are {ds.interval.name.lower()}; can only aggregate to coarser scales.")
    tf = ds.time_fields
    key_cols = [tf[k] for k in keys]
    values = ds.value_columns
    if totals is not None and averages is not None:
        raise ValueError("Give either 'totals' or 'averages', not both.")
    if averages is not None:
        avg_set = set(averages)
    elif totals is not None:
        avg_set = set(values) - set(totals)
    else:
        avg_set = set()
    unknown = avg_set - set(values)
    if unknown:
        raise KeyError(f"Unknown variables: {sorted(unknown)}")

    df = ds.data
    key = df[key_cols].to_numpy()
    run_id = np.concatenate([[0], np.cumsum(np.any(key[1:] != key[:-1], axis=1))])
    g = df.groupby(run_id, sort=False)

    out = g[key_cols].first()
    for col in values:
        series = g[col]
        agg = series.mean() if col in avg_set else series.sum(min_count=1)
        if strict_no_data:
            agg = agg.where(series.count() == series.size())
        out[col] = agg
    # drop the finer time fields; keep column order
    ordered = [c for c in df.columns if c in out.columns]
    return ds.with_data(out[ordered].reset_index(drop=True), groups=None)


def average_intermittent(ds: Dataset, variables: Iterable[str]) -> Dataset:
    """Divide each reading by the elapsed time since the previous record (``AvgIntDataCmd``).

    For daily data the divisor is days, for hourly data hours. Intended for meter readings taken
    at irregular intervals (e.g. cumulative consumption between reads). The first record has no
    predecessor, so the selected variables are set to missing there. (VB blanked *every* non-time
    variable in record 1; the port only touches the variables being averaged.)
    """
    if ds.interval not in (TimeInterval.DAILY, TimeInterval.HOURLY):
        raise ValueError("Average Intermittent Data is only for daily or hourly data.")
    ts = ds.timestamps()
    unit = pd.Timedelta(days=1) if ds.interval == TimeInterval.DAILY else pd.Timedelta(hours=1)
    elapsed = (ts.diff() / unit).to_numpy()
    df = ds.data.copy()
    for col in variables:
        v = df[col].to_numpy(dtype=float).copy()
        with np.errstate(divide="ignore", invalid="ignore"):
            v = np.where(elapsed != 0, v / elapsed, v)
        v[0] = np.nan
        df[col] = v
    return ds.with_data(df)
