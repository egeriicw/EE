"""Small statistics helpers ported from SUBS.BAS (``Percentile``, ``Median``, ``DWTest``)."""

from __future__ import annotations

import math

import numpy as np


def vb_percentile(values, p: float) -> float:
    """Percentile exactly as ENERGY EXPLORER computes it.

    Rank position k = (n + 1)·p on the sorted data, linear interpolation between neighbours,
    clamped to the first/last value. (This is NOT numpy's default "linear" method.)
    """
    v = np.sort(np.asarray(values, dtype=float)[~np.isnan(np.asarray(values, dtype=float))])
    n = len(v)
    if n == 0:
        return float("nan")
    k = (n + 1) * p
    ip = math.floor(k)
    frac = k - ip
    if ip <= 0:
        return float(v[0])
    if ip >= n:
        return float(v[n - 1])
    return float((1 - frac) * v[ip - 1] + frac * v[ip])


def median_with_se(values) -> tuple[float, float]:
    """Median and its standard error, SE = (P75 − P25) / √n (SUBS.BAS ``Median``)."""
    v = np.asarray(values, dtype=float)
    v = v[~np.isnan(v)]
    n = len(v)
    med = vb_percentile(v, 0.5)
    se = (vb_percentile(v, 0.75) - vb_percentile(v, 0.25)) / math.sqrt(n) if n else float("nan")
    return med, se


def box_percentiles(values) -> dict[str, float]:
    """P10/P25/P50/P75/P90 used by the box-and-whisker Bin Fit display."""
    return {f"p{int(p * 100)}": vb_percentile(values, p) for p in (0.1, 0.25, 0.5, 0.75, 0.9)}


def lag1_autocorrelation(resid: np.ndarray) -> float:
    """ρ = Σ r[i-1]·r[i] / Σ r[i-1]²  (``autocor`` in ``Inf``)."""
    r = np.asarray(resid, dtype=float)
    if len(r) < 2:
        return float("nan")
    den = float(np.sum(r[:-1] ** 2))
    return float(np.sum(r[:-1] * r[1:]) / den) if den else float("nan")


def durbin_watson(resid: np.ndarray) -> float:
    r = np.asarray(resid, dtype=float)
    den = float(np.sum(r**2))
    return float(np.sum(np.diff(r) ** 2) / den) if den else float("nan")


def dw_test(n_params: int, n: int, dw: float) -> str | None:
    """Approximate 5% Durbin–Watson bounds test from ``DWTest`` (fits for p = 2 or 3 only).

    Returns ``"p>0"`` (positive autocorrelation), ``"p=0"`` (none) or ``"I"`` (inconclusive).
    """
    if n <= 1 or not np.isfinite(dw):
        return None
    if n_params == 2:
        dl, du = 0.31624 + 0.29655 * math.log(n), 0.88254 + 0.17875 * math.log(n)
    elif n_params == 3:
        dl, du = 0.046637 + 0.35361 * math.log(n), 1.24391 + 1.00448 * math.log(n)
    else:
        return None
    if dw < dl:
        return "p>0"
    if dw > du:
        return "p=0"
    return "I"


def t_approx(dof: float) -> float:
    """ENERGY EXPLORER's approximation to the two-sided 95% t value: 1.96 + 2.7 / dof."""
    return 1.96 + 2.7 / dof if dof > 0 else float("inf")
