"""OLS engine (replaces ``Reg``, ``Inf``, ``Invert``, ``Mult``, ``Trans`` in SUBS.BAS).

The VB code normalised each regressor by its mean, formed XᵀX, inverted it with a hand-written
Gauss–Jordan routine (max 8×8), and back-scaled. Here we use numpy: coefficients from
``lstsq`` and coefficient variances from the pseudo-inverse of XᵀX, which also behaves sensibly
for the rank-deficient design matrices that occur at the edges of a change-point grid search.
"""

from __future__ import annotations

from dataclasses import asdict, dataclass

import numpy as np

from energy_explorer.stats import durbin_watson, lag1_autocorrelation


@dataclass(frozen=True)
class OLSResult:
    beta: np.ndarray
    se: np.ndarray
    fitted: np.ndarray
    resid: np.ndarray
    n: int
    n_params: int
    sse: float
    rmse: float
    cv_rmse: float
    r2: float
    adj_r2: float
    ymean: float
    autocorrelation: float
    durbin_watson: float

    def stats_dict(self) -> dict:
        d = asdict(self)
        for k in ("beta", "se", "fitted", "resid"):
            d.pop(k)
        return d


def ols(X: np.ndarray, y: np.ndarray) -> OLSResult:
    """Fit y = X·β. ``X`` must already contain the intercept column.

    Statistics follow ``Inf``: RMSE uses n − p degrees of freedom (0 when n == p),
    CV-RMSE = RMSE / ȳ × 100, R² = 1 − SSE/SST, adjusted R² = 1 − (n−1)/(n−p)·SSE/SST.
    """
    X = np.asarray(X, dtype=float)
    y = np.asarray(y, dtype=float).ravel()
    n, p = X.shape
    if n == 0:
        raise ValueError("No data is available to model.")
    beta, *_ = np.linalg.lstsq(X, y, rcond=None)
    fitted = X @ beta
    resid = y - fitted
    ymean = float(y.mean())
    sse = float(resid @ resid)
    sst = float(((y - ymean) ** 2).sum())
    dof = n - p
    mse = sse / dof if dof > 0 else 0.0
    rmse = mse**0.5
    xtx_inv = np.linalg.pinv(X.T @ X)
    se = rmse * np.sqrt(np.clip(np.diag(xtx_inv), 0, None))
    return OLSResult(
        beta=beta,
        se=se,
        fitted=fitted,
        resid=resid,
        n=n,
        n_params=p,
        sse=sse,
        rmse=rmse,
        cv_rmse=rmse / ymean * 100 if ymean else float("nan"),
        r2=1 - sse / sst if sst else float("nan"),
        adj_r2=(1 - (n - 1) / dof * sse / sst) if (dof > 0 and sst) else 1.0,
        ymean=ymean,
        autocorrelation=lag1_autocorrelation(resid),
        durbin_watson=durbin_watson(resid),
    )
