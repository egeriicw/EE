"""Model fitting: mean, median, 2P/MLR, change-point (3P/4P/5P, ± MVR) and Bin Fit.

Ports ``MeanBut_Click``, ``MedianBut_Click``, ``TwoP``, ``MLRBut_Click``, ``ThreePMVR``,
``FourPBut_Click`` / ``FourPMVR``, ``FivePMVR`` and ``BinFitBut_Click``.

The change-point search is a faithful port of the VB grid search (deliberately *not* replaced by
a continuous optimiser, so results match the original program's choice of change point):

* 3P (C/H, with or without MVR): 100 intervals over [xmin, xmax]; candidates xmin + i·inc,
  i = 1..99. Then a fine pass from (best − inc) with step 2·inc/100, again i = 1..99.
* 4P: same two-pass scheme. The plain 4P button used **10** intervals, the 4P-MVR button 100
  (``n_intervals`` defaults reproduce that).
* 5P: 9 intervals; all (cp1, cp2) pairs from the 8 interior grid points with cp1 < cp2; then a
  fine pass trying cp ± 0.333·inc and ± 0.666·inc around each best change point.
* Selection is by minimum RMSE (equivalently SSE); ties keep the first candidate.
* Indicator convention: points with x == cp belong to the *left* segment (x <= cp).
"""

from __future__ import annotations

from collections.abc import Callable

import numpy as np

from energy_explorer.models import (
    BinFitModel,
    FivePModel,
    FourPModel,
    LinearModel,
    MeanModel,
    MedianModel,
    ModelStats,
    ThreePModel,
)
from energy_explorer.regression import OLSResult, ols
from energy_explorer.stats import dw_test, median_with_se, t_approx


def _stats(fit: OLSResult, n_total_params: int) -> ModelStats:
    n = fit.n
    return ModelStats(
        n=n,
        n_params=fit.n_params,
        rmse=fit.rmse,
        cv_rmse=fit.cv_rmse,
        r2=fit.r2,
        adj_r2=fit.adj_r2,
        ymean=fit.ymean,
        autocorrelation=fit.autocorrelation,
        durbin_watson=fit.durbin_watson,
        dw_test=dw_test(fit.n_params, n, fit.durbin_watson),
        prediction_interval=t_approx(n - n_total_params) * fit.rmse * (1 + 1 / n) ** 0.5,
    )


def _columns(a) -> list[np.ndarray]:
    """(n,) or (n, k) array -> list of k column vectors."""
    a = np.asarray(a, dtype=float)
    if a.ndim == 1:
        a = a[:, None]
    return [a[:, j] for j in range(a.shape[1])]


def _design(cols: list[np.ndarray]) -> np.ndarray:
    return np.column_stack([np.ones_like(cols[0])] + cols)


# --------------------------------------------------------------------------- simple models
def fit_mean(y: np.ndarray, *, y_name: str = "y") -> MeanModel:
    y = np.asarray(y, dtype=float)
    n = len(y)
    if n == 0:
        raise ValueError("No data is available to model.")
    mean = float(y.mean())
    std = float(y.std(ddof=1)) if n > 1 else 0.0
    return MeanModel(
        model_type="Mean",
        y=y_name,
        mean=mean,
        std=std,
        stats=ModelStats(
            n=n,
            n_params=1,
            rmse=std,
            cv_rmse=std / mean * 100 if mean else float("nan"),
            ymean=mean,
            prediction_interval=t_approx(n - 1) * std,
        ),
    )


def fit_median(y: np.ndarray, *, y_name: str = "y") -> MedianModel:
    y = np.asarray(y, dtype=float)
    n = len(y)
    if n == 0:
        raise ValueError("No data is available to model.")
    med, se = median_with_se(y)
    return MedianModel(
        model_type="Median",
        y=y_name,
        median=med,
        se_median=se,
        stats=ModelStats(
            n=n,
            n_params=1,
            rmse=se,
            cv_rmse=se / med * 100 if med else float("nan"),
            ymean=float(y.mean()),
            prediction_interval=t_approx(n - 1) * se,
        ),
    )


def fit_linear(
    X: np.ndarray, y: np.ndarray, *, y_name: str = "y", x_names: list[str] | None = None
) -> tuple[LinearModel, OLSResult]:
    cols = _columns(X)
    k = len(cols)
    fit = ols(_design(cols), y)
    model = LinearModel(
        model_type="2P" if k == 1 else "MLR",
        y=y_name,
        x=list(x_names or [f"x{j + 1}" for j in range(k)]),
        coefs=fit.beta.tolist(),
        se=fit.se.tolist(),
        stats=_stats(fit, k + 1),
    )
    return model, fit


# --------------------------------------------------------------------------- grid search
def _two_pass_search(
    x: np.ndarray,
    build: Callable[[float], np.ndarray],
    y: np.ndarray,
    n_intervals: int,
) -> tuple[float, float]:
    """VB coarse→fine grid search used by 3P and 4P. Returns (best cp, final step)."""
    xmin, xmax = float(np.min(x)), float(np.max(x))
    inc = (xmax - xmin) / n_intervals
    start = xmin
    best_cp, best_rmse = xmin, np.inf
    for iteration in (1, 2):
        if iteration == 2:
            start = best_cp - inc
            inc = 2 * inc / n_intervals
        for i in range(1, n_intervals):
            cp = start + i * inc
            rmse = ols(build(cp), y).rmse
            # VB resets the incumbent at i == 1 of *each* pass, so the fine pass result is the
            # best point of the fine grid (which brackets the coarse optimum).
            if i == 1 or rmse < best_rmse:
                best_cp, best_rmse = cp, rmse
    return float(best_cp), float(inc)


def fit_3p(
    x: np.ndarray,
    y: np.ndarray,
    *,
    heating: bool,
    extra: np.ndarray | None = None,
    n_intervals: int = 100,
    y_name: str = "y",
    x_names: list[str] | None = None,
) -> tuple[ThreePModel, OLSResult]:
    x = np.asarray(x, dtype=float)
    y = np.asarray(y, dtype=float)
    if len(x) == 0:
        raise ValueError("No data is available to model.")
    extra_cols = [] if extra is None else _columns(extra)

    def build(cp: float) -> np.ndarray:
        ind = (x <= cp) if heating else (x > cp)
        return _design([ind * (x - cp)] + extra_cols)

    cp, step = _two_pass_search(x, build, y, n_intervals)
    fit = ols(build(cp), y)
    k = 1 + len(extra_cols)
    model = ThreePModel(
        model_type="3PH" if heating else "3PC",
        y=y_name,
        x=list(x_names or [f"x{j + 1}" for j in range(k)]),
        heating=heating,
        xcp=cp,
        xcp_step=step,
        ycp=float(fit.beta[0]),
        se_ycp=float(fit.se[0]),
        slope=float(fit.beta[1]),
        se_slope=float(fit.se[1]),
        extra_coefs=fit.beta[2:].tolist(),
        extra_se=fit.se[2:].tolist(),
        stats=_stats(fit, fit.n_params + 1),
    )
    return model, fit


def fit_4p(
    x: np.ndarray,
    y: np.ndarray,
    *,
    extra: np.ndarray | None = None,
    n_intervals: int | None = None,
    y_name: str = "y",
    x_names: list[str] | None = None,
) -> tuple[FourPModel, OLSResult]:
    x = np.asarray(x, dtype=float)
    y = np.asarray(y, dtype=float)
    if len(x) == 0:
        raise ValueError("No data is available to model.")
    extra_cols = [] if extra is None else _columns(extra)
    if n_intervals is None:
        n_intervals = 100 if extra_cols else 10  # FourPMVR vs FourPBut_Click

    def build(cp: float) -> np.ndarray:
        return _design([x, (x > cp) * (x - cp)] + extra_cols)

    cp, step = _two_pass_search(x, build, y, n_intervals)
    fit = ols(build(cp), y)
    b0, b1, b2 = (float(v) for v in fit.beta[:3])
    s0, s1, s2 = (float(v) for v in fit.se[:3])
    ycp = b0 + b1 * cp
    # VB error propagation (ignores the b0/b1 covariance); kept for fidelity.
    if abs(b1) > 0 and abs(cp) > 0:
        se_ycp = (s0**2 + (b1 * cp) ** 2 * ((s1 / b1) ** 2 + (step / cp) ** 2)) ** 0.5
    else:
        se_ycp = step
    k = 1 + len(extra_cols)
    model = FourPModel(
        model_type="4P",
        y=y_name,
        x=list(x_names or [f"x{j + 1}" for j in range(k)]),
        xcp=cp,
        xcp_step=step,
        ycp=ycp,
        se_ycp=se_ycp,
        left_slope=b1,
        se_left_slope=s1,
        right_slope=b1 + b2,
        se_right_slope=(s1**2 + s2**2) ** 0.5,
        beta=[b0, b1, b2],
        extra_coefs=fit.beta[3:].tolist(),
        extra_se=fit.se[3:].tolist(),
        stats=_stats(fit, fit.n_params + 1),
    )
    return model, fit


def fit_5p(
    x: np.ndarray,
    y: np.ndarray,
    *,
    extra: np.ndarray | None = None,
    n_intervals: int = 9,
    y_name: str = "y",
    x_names: list[str] | None = None,
) -> tuple[FivePModel, OLSResult]:
    x = np.asarray(x, dtype=float)
    y = np.asarray(y, dtype=float)
    if len(x) == 0:
        raise ValueError("No data is available to model.")
    extra_cols = [] if extra is None else _columns(extra)

    def build(cp1: float, cp2: float) -> np.ndarray:
        return _design([(x <= cp1) * (x - cp1), (x > cp2) * (x - cp2)] + extra_cols)

    xmin, xmax = float(np.min(x)), float(np.max(x))
    inc = (xmax - xmin) / n_intervals
    grid = [xmin + i * inc for i in range(1, n_intervals)]
    best = None
    best_rmse = np.inf
    for gridpass in (1, 2):
        if gridpass == 1:
            left = right = grid
        else:
            offs = (-0.666, -0.333, 0.333, 0.666)
            left = [best[0] + o * inc for o in offs]
            right = [best[1] + o * inc for o in offs]
        for cp1 in left:
            for cp2 in right:
                if cp1 >= cp2:
                    continue
                rmse = ols(build(cp1, cp2), y).rmse
                if best is None or rmse < best_rmse:
                    best, best_rmse = (cp1, cp2), rmse
    if best is None:
        raise ValueError("5P search found no valid change-point pair.")
    fit = ols(build(*best), y)
    k = 1 + len(extra_cols)
    model = FivePModel(
        model_type="5P",
        y=y_name,
        x=list(x_names or [f"x{j + 1}" for j in range(k)]),
        xcp1=float(best[0]),
        xcp2=float(best[1]),
        xcp_step=inc * 0.333,
        ycp=float(fit.beta[0]),
        se_ycp=float(fit.se[0]),
        left_slope=float(fit.beta[1]),
        se_left_slope=float(fit.se[1]),
        right_slope=float(fit.beta[2]),
        se_right_slope=float(fit.se[2]),
        extra_coefs=fit.beta[3:].tolist(),
        extra_se=fit.se[3:].tolist(),
        stats=_stats(fit, fit.n_params + 2),
    )
    return model, fit


def fit_binfit(
    x: np.ndarray,
    y: np.ndarray,
    *,
    n_bins: int = 10,
    xmin: float | None = None,
    xmax: float | None = None,
    y_name: str = "y",
    x_name: str = "x",
) -> tuple[BinFitModel, np.ndarray]:
    x = np.asarray(x, dtype=float)
    y = np.asarray(y, dtype=float)
    if len(x) == 0:
        raise ValueError("No data is available to model.")
    xmin = float(np.min(x)) if xmin is None else float(xmin)
    xmax = float(np.max(x)) if xmax is None else float(xmax)
    shell = BinFitModel(
        model_type="BinFit",
        y=y_name,
        x=[x_name],
        xmin=xmin,
        xmax=xmax,
        n_bins=n_bins,
        bin_means=[],
        stats=ModelStats(n=0, n_params=n_bins, rmse=0, cv_rmse=0),
    )
    idx = shell.bin_index(x)
    counts = np.bincount(idx[idx >= 0], minlength=n_bins)
    sums = np.bincount(idx[idx >= 0], weights=y[idx >= 0], minlength=n_bins)
    with np.errstate(invalid="ignore", divide="ignore"):
        means = np.where(counts > 0, sums / np.maximum(counts, 1), np.nan)
    fitted = np.where(idx >= 0, means[np.clip(idx, 0, None)], np.nan)
    resid = y - fitted
    ssm = float(np.nansum(resid**2))
    n = len(y)
    # VB divides by (n − j) where j is a leaked loop index that is 1 after the last binned point,
    # i.e. effectively n − 1. Reproduced here.
    rmse = (ssm / (n - 1)) ** 0.5 if n > 1 else 0.0
    ymean = float(y.mean())
    shell.bin_means = means.tolist()
    shell.bin_counts = counts.astype(int).tolist()
    shell.stats = ModelStats(
        n=n,
        n_params=n_bins,
        rmse=rmse,
        cv_rmse=rmse / ymean * 100 if ymean else float("nan"),
        ymean=ymean,
    )
    return shell, resid
