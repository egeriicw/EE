"""Grouping records into two groups (port of ``GroupByDateCmd_Click`` / ``GroupByDOWCmd_Click``)."""

from __future__ import annotations

from datetime import date, datetime

import numpy as np
import pandas as pd

from energy_explorer.dataset import Dataset, TimeInterval


def to_serial(value, interval: TimeInterval) -> float:
    """Convert a user date to the VB serial form yyyymmddhh (``SerialDate`` in SUBS.BAS).

    Accepts a number already in serial form, a ``date``/``datetime``/``Timestamp``, or text:
    ``"yyyy"`` (yearly), ``"m/yyyy"`` (monthly), ``"m/d/yyyy"`` (daily), ``"m/d/yyyy/h"`` (hourly).
    Fields finer than the data's interval are dropped, as in VB.
    """
    if isinstance(value, (int, float, np.integer, np.floating)) and value > 9999:
        return float(value)
    hr = 0
    if isinstance(value, str):
        txt = value.strip()
        if interval == TimeInterval.HOURLY:
            if txt.count("/") != 3:
                raise ValueError(f"Hourly data needs dates as m/d/yyyy/h (got {value!r}).")
            txt, hr_txt = txt.rsplit("/", 1)
            hr = int(float(hr_txt))
        if interval == TimeInterval.YEARLY or txt.isdigit():
            yr = int(txt)
            ts = pd.Timestamp(year=yr + (1900 if 50 <= yr <= 99 else 2000 if yr < 50 else 0), month=1, day=1)
        elif txt.count("/") == 1:
            mo, yr = txt.split("/")
            ts = pd.Timestamp(year=int(yr), month=int(mo), day=1)
        else:
            ts = pd.Timestamp(txt)
    elif isinstance(value, (date, datetime, pd.Timestamp)):
        ts = pd.Timestamp(value)
        hr = ts.hour
    else:
        raise TypeError(f"Cannot interpret {value!r} as a date")
    yr = ts.year
    if yr < 100:
        yr += 1900 if yr >= 50 else 2000
    mo = ts.month if interval >= TimeInterval.MONTHLY else 0
    dy = ts.day if interval >= TimeInterval.DAILY else 0
    hr = hr if interval >= TimeInterval.HOURLY else 0
    return yr * 1e6 + mo * 1e4 + dy * 1e2 + hr


def group_by_date(
    ds: Dataset,
    start=None,
    end=None,
    names: tuple[str, str] = ("G1", "G2"),
) -> Dataset:
    """Group 1 = records with start <= date <= end (inclusive); everything else is group 2.

    ``start=None`` means "start of data file", ``end=None`` means "end of data file". Name the
    groups ``("Pre", "Post")`` for a savings analysis.
    """
    interval = ds.interval
    serial = ds.serial_dates()
    lo = 0.0 if start in (None, "") else to_serial(start, interval)
    hi = 9_999_999_999.0 if end in (None, "") else to_serial(end, interval)
    groups = np.where((serial >= lo) & (serial <= hi), 1, 2)
    return ds.with_groups(groups, names)


def group_by_day_of_week(ds: Dataset, names: tuple[str, str] = ("WD", "WE")) -> Dataset:
    """Weekdays -> group 1, Saturdays/Sundays -> group 2. Needs Yr, Mo and Dy columns."""
    tf = ds.time_fields
    if not {"yr", "mo", "dy"} <= tf.keys():
        raise ValueError("Group by day-of-week requires Yr, Mo and Dy columns.")
    ts = pd.to_datetime(
        pd.DataFrame(
            {"year": ds.data[tf["yr"]], "month": ds.data[tf["mo"]], "day": ds.data[tf["dy"]]}
        ),
        errors="coerce",
    )
    weekend = ts.dt.day_of_week.isin([5, 6]).to_numpy()
    return ds.with_groups(np.where(weekend, 2, 1), names)
