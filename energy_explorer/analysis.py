"""Histogram and hourly-profile analyses (ports ``HistBut_Click`` and ``HourlyProfilesBut_Click``).

Both are pure data functions. The VB versions drew straight onto a picture box; here they return
numbers, and the API / web client / notebook decide how to draw them.

Histogram
---------
Bins are equal width over [xmin, xmax] (defaults: the range of the selected data). Bin 1 is closed,
[lo, hi]; every later bin is (lo, hi], matching ``HistBut_Click``. Counts are per group, and the
probability is count / n for that group.

Hourly profile
--------------
Records are arranged as one row per day and one column per hour, as the VB code did when it looked
for ``hr == 0`` to start each day. Values are scaled by a rounded maximum (``vb_round_up``), as VB
did, and a 3D oblique projection (``project_3d``, the same formula as ``HourlyProfilesBut_Click``)
turns each day into a polyline.
"""

from __future__ import annotations

import math
from dataclasses import asdict, dataclass

import numpy as np
import pandas as pd

from energy_explorer.dataset import Dataset
from energy_explorer.models import _clean


# ----------------------------------------------------------------------------- binning
def bin_index(x, xmin: float, xmax: float, n_bins: int) -> np.ndarray:
    """0-based bin for each value, -1 if outside [xmin, xmax] or missing.

    Bin 0 is closed [xmin, edge1]; the rest are (lo, hi] (as VB's ``Case 1`` / ``Case Else``).
    """
    if n_bins < 1:
        raise ValueError("n_bins must be at least 1.")
    if not xmax > xmin:
        raise ValueError(f"Bin range needs xmax > xmin (got {xmin} .. {xmax}).")
    x = np.asarray(x, dtype=float)
    edges = xmin + np.arange(n_bins + 1) * (xmax - xmin) / n_bins
    # side="left": a value equal to edge k lands in bin k-1, i.e. (lo, hi]
    idx = np.searchsorted(edges, np.nan_to_num(x, nan=np.inf), side="left") - 1
    idx = np.where(x == xmin, 0, idx)
    ok = (x >= xmin) & (x <= edges[-1]) & ~np.isnan(x)
    return np.where(ok, np.clip(idx, 0, n_bins - 1), -1).astype(int)


def bin_edges(xmin: float, xmax: float, n_bins: int) -> np.ndarray:
    return xmin + np.arange(n_bins + 1) * (xmax - xmin) / n_bins


# ----------------------------------------------------------------------------- histogram
@dataclass
class HistogramSeries:
    group: str  # "All" or a group name
    n: int
    counts: list[int]
    probability: list[float]  # 0..1
    percent: list[float]  # 0..100 (the VB value window prints this)


@dataclass
class Histogram:
    variable: str
    n_bins: int
    xmin: float
    xmax: float
    edges: list[float]
    midpoints: list[float]
    series: list[HistogramSeries]

    def to_dict(self) -> dict:
        return _clean(asdict(self))

    def text(self) -> str:
        """Plain-text table in the style of the VB value window."""
        lines = [f"Histogram of {self.variable}: N = {[s.n for s in self.series]}, bins = {self.n_bins}"]
        for s in self.series:
            lines.append(f"  [{s.group}]")
            for i, (lo, hi) in enumerate(zip(self.edges[:-1], self.edges[1:]), start=1):
                lines.append(
                    f"    ({i}) {lo:>12,.2f} < x <= {hi:<12,.2f}  count {s.counts[i - 1]:>6}  "
                    f"{s.percent[i - 1]:6.2f}%"
                )
        return "\n".join(lines)


def histogram(
    ds: Dataset,
    variable: str,
    *,
    n_bins: int = 10,
    xmin: float | None = None,
    xmax: float | None = None,
    group: int | str | None = None,
    by_group: bool = False,
) -> Histogram:
    """Histogram of ``variable``.

    ``group=None`` uses every record. If ``by_group`` is true and the data are grouped, each group
    gets its own series (the G1|G2 display in VB). The bin range is shared so the series line up.
    """
    values = ds.column(variable).to_numpy(dtype=float)
    present = ~np.isnan(values)
    if by_group and ds.is_grouped:
        selections = [(ds.group_names[0], 1), (ds.group_names[1], 2)]
    else:
        g = ds.group_number(group)
        selections = [("All" if g is None else ds.group_names[g - 1], g)]
    # group_mask(None) selects every record
    masks = [present & ds.group_mask(g) for _, g in selections]
    pooled = np.concatenate([values[m] for m in masks])
    if pooled.size == 0:
        raise ValueError("No data is available for the histogram.")
    lo = float(pooled.min()) if xmin is None else float(xmin)
    hi = float(pooled.max()) if xmax is None else float(xmax)
    series = []
    for (label, _), m in zip(selections, masks):
        v = values[m]
        idx = bin_index(v, lo, hi, n_bins)
        counts = np.bincount(idx[idx >= 0], minlength=n_bins)
        n = int(v.size)
        prob = counts / n if n else np.zeros(n_bins)
        series.append(
            HistogramSeries(
                group=label,
                n=n,
                counts=counts.astype(int).tolist(),
                probability=prob.tolist(),
                percent=(prob * 100).tolist(),
            )
        )
    edges = bin_edges(lo, hi, n_bins)
    return Histogram(
        variable=variable,
        n_bins=n_bins,
        xmin=lo,
        xmax=hi,
        edges=edges.tolist(),
        midpoints=((edges[:-1] + edges[1:]) / 2).tolist(),
        series=series,
    )


# ----------------------------------------------------------------------------- hourly profile
def vb_round_up(value: float) -> float:
    """VB's rounded scale maximum: leading digit + 1, times the power of ten (``HourlyProfiles``).

    e.g. 1234 -> 2000, 9.5 -> 10. Non-positive values return 1 so the scale stays usable.
    """
    if not value or value <= 0 or not math.isfinite(value):
        return 1.0
    power = math.floor(math.log10(value))
    lead = int(value / 10**power)
    return float((lead + 1) * 10**power)


def hour_of_day(ds: Dataset) -> pd.Series:
    """Hour 0..23 for each record.

    The VB code assumed 0..23. The guide also allows 1..24 (hour-ending); when the column has 24 and
    no 0, it is shifted to 0..23 so the day lines up.
    """
    tf = ds.time_fields
    if "hr" not in tf:
        raise ValueError("Hourly profiles need an 'Hr' column.")
    hr = ds.data[tf["hr"]].astype(float)
    if (hr == 24).any() and not (hr == 0).any():
        hr = hr - 1
    if ((hr < 0) | (hr > 23)).any():
        raise ValueError("Hr values must be 0-23 (or 1-24).")
    return hr.astype(int)


@dataclass
class HourlyProfile:
    variable: str
    group: str
    dates: list[str]  # one per day, oldest first
    hours: list[int]  # 0..23
    values: list[list[float | None]]  # days x hours, original units
    scale: float  # VB rounded maximum used for drawing
    n: int  # number of valid observations

    def to_dict(self) -> dict:
        return _clean(asdict(self))


def hourly_profile(ds: Dataset, variable: str, group: int | str | None = None) -> HourlyProfile:
    tf = ds.time_fields
    for need in ("yr", "mo", "dy"):
        if need not in tf:
            raise ValueError(f"Hourly profiles need Yr, Mo, Dy and Hr columns (missing {need}).")
    gnum = ds.group_number(group)
    mask = ds.group_mask(gnum) & ds.column(variable).notna().to_numpy()
    if not mask.any():
        raise ValueError("No data is available for the hourly profile.")
    days = pd.to_datetime(
        pd.DataFrame(
            {
                "year": ds.data[tf["yr"]],
                "month": ds.data[tf["mo"]],
                "day": ds.data[tf["dy"]],
            }
        ),
        errors="coerce",
    )
    frame = pd.DataFrame(
        {"day": days.to_numpy(), "hour": hour_of_day(ds).to_numpy(), "value": ds.column(variable).to_numpy(dtype=float)}
    )[mask]
    table = (
        frame.pivot_table(index="day", columns="hour", values="value", aggfunc="mean")
        .reindex(columns=range(24))
        .sort_index()
    )
    values = [[None if pd.isna(v) else float(v) for v in row] for row in table.to_numpy()]
    peak = float(np.nanmax(table.to_numpy())) if table.notna().any().any() else 0.0
    return HourlyProfile(
        variable=variable,
        group=("All" if gnum is None else ds.group_names[gnum - 1]),
        dates=[pd.Timestamp(d).strftime("%Y-%m-%d") for d in table.index],
        hours=list(range(24)),
        values=values,
        scale=vb_round_up(peak),
        n=int(mask.sum()),
    )


def project_3d(x, y, z, theta_deg: float, res: float):
    """Oblique projection from ``HourlyProfilesBut_Click`` / ``D3Line`` (VB's 3.1416 pi kept as math.pi).

    Maps (x = day, y = value, z = hour) to 2-D screen coordinates for a view rotated by ``theta_deg``.
    """
    t = math.radians(theta_deg)
    s, c = math.sin(t), math.cos(t)
    xs = (x - z * c + res * c) / (c + 1)
    ys = (y - z * s + res * s) / (s + 1)
    return xs, ys


def profile_polylines(
    profile: HourlyProfile, *, theta_deg: float = 45.0, res: float = 400.0
) -> list[dict]:
    """One projected polyline per day, as in the VB 3D display. Days are laid out front to back in
    the same order VB used (the first day is drawn at the largest x)."""
    n_days = len(profile.dates)
    lines = []
    for k, date in enumerate(profile.dates):
        x = (n_days - 1 - k) / n_days * res
        pts = []
        for h, v in zip(profile.hours, profile.values[k]):
            if v is None:
                pts.append(None)
                continue
            y = v / profile.scale * res
            z = res * (1 - h / 23)
            pts.append(list(project_3d(x, y, z, theta_deg, res)))
        lines.append({"date": date, "points": pts})
    return lines


__all__ = [
    "Histogram",
    "HistogramSeries",
    "HourlyProfile",
    "bin_edges",
    "bin_index",
    "histogram",
    "hour_of_day",
    "hourly_profile",
    "profile_polylines",
    "project_3d",
    "vb_round_up",
]
