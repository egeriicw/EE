"""Savings, uncertainty, Lean Energy Analysis and model projection.

Ports ``CalcSavingsCmd_Click`` (the savings shown in the UI), the autocorrelation-corrected
uncertainty in ``ProjectSavingsCmd_Click`` / ``CalcSavingsPECmd_Click``, ``CalcLEACmd_Click`` and
``CreateModelProjectionCmd_Click``.
"""

from __future__ import annotations

import math
from dataclasses import asdict, dataclass, field

import numpy as np
import pandas as pd

from energy_explorer.dataset import Dataset
from energy_explorer.models import BaselineModel, LinearModel, MeanModel, ThreePModel, _clean
from energy_explorer.stats import t_approx


def _pct(err: float, value: float) -> float:
    return err / value * 100 if value else float("nan")


@dataclass
class SavingsResult:
    model_type: str
    n_pre: int
    n_post: int
    adjusted_baseline: float  # Σ predicted (pre model at post conditions)
    actual: float  # Σ measured in post period
    savings: float  # adjusted_baseline − actual
    baseline_error: float  # ±, 95 %
    measurement_error: float
    savings_error: float
    measurement_error_pct: float
    #: per-record predicted baseline over the post period (NaN elsewhere); not serialised
    predicted: pd.Series | None = field(default=None, repr=False)

    @property
    def savings_per_obs(self) -> float:
        return self.savings / self.n_post if self.n_post else float("nan")

    @property
    def savings_error_per_obs(self) -> float:
        return self.savings_error / self.n_post if self.n_post else float("nan")

    @property
    def fractional_savings_uncertainty(self) -> float:
        return _pct(self.savings_error, self.savings)

    def report(self) -> str:
        """Text in the layout of the VB Value Window."""

        def line(label, v, e):
            return f" {label} = {v:,.1f}  ±{e:,.1f}  ±{_pct(e, v):.1f}%"

        return "\n".join(
            [
                f" Observations During Post-retrofit Period: {self.n_post:,}",
                line("Adjusted Baseline Use During Post-retrofit Period", self.adjusted_baseline,
                     self.baseline_error),
                line("Actual Use During Post-retrofit Period", self.actual, self.measurement_error),
                line("Total Savings During Post-retrofit Period", self.savings, self.savings_error),
                line("Average Savings Per Observation During Post-retrofit Period",
                     self.savings_per_obs, self.savings_error_per_obs),
            ]
        )

    def to_dict(self) -> dict:
        d = asdict(self)
        d.pop("predicted")
        d["savings_per_obs"] = self.savings_per_obs
        d["savings_error_per_obs"] = self.savings_error_per_obs
        d["fractional_savings_uncertainty_pct"] = self.fractional_savings_uncertainty
        return _clean(d)


def calculate_savings(
    model: BaselineModel,
    ds: Dataset,
    *,
    post_group: int | str = 2,
    measurement_error_pct: float = 0.0,
) -> SavingsResult:
    """Savings = Σ(baseline prediction) − Σ(measured) over the post-retrofit group.

    Uncertainty (95 %), as implemented in ``CalcSavingsCmd_Click``:

    * baseline: 1.96 · RMSE_pre · √(1 + 2/n_pre) · √n_post
    * measurement: Σ measured · (measurement_error_pct / 100)
    * savings: √(baseline² + measurement²)

    Records count toward the post period only where y and all model inputs are present (and, for
    Bin Fit, where x falls inside the fitted bins). ``measurement_error_pct`` defaults to 0, which
    is the value of the VB global until "Define Measurement Error" is used (its dialog proposes 5%).
    """
    gnum = ds.group_number(post_group)
    mask = ds.group_mask(gnum)
    y = ds.column(model.y).to_numpy(dtype=float)
    pred = model.predict(ds.data)
    use = mask & ~np.isnan(y) & ~np.isnan(pred)
    n_post = int(use.sum())
    adj = float(pred[use].sum())
    actual = float(y[use].sum())
    n_pre = model.stats.n
    rmse = model.stats.rmse
    base_err = 1.96 * rmse * math.sqrt(1 + 2 / n_pre) * math.sqrt(n_post) if n_post else 0.0
    meas_err = actual * measurement_error_pct / 100
    series = pd.Series(np.where(use, pred, np.nan), index=ds.data.index, name="Adjusted_Baseline")
    return SavingsResult(
        model_type=model.model_type,
        n_pre=n_pre,
        n_post=n_post,
        adjusted_baseline=adj,
        actual=actual,
        savings=adj - actual,
        baseline_error=base_err,
        measurement_error=meas_err,
        savings_error=math.hypot(base_err, meas_err),
        measurement_error_pct=measurement_error_pct,
        predicted=series,
    )


# ---------------------------------------------------------------- AR(1)-corrected uncertainty
def _toeplitz_apply(z: np.ndarray, rho: float) -> np.ndarray:
    """(R·z) with R[i, j] = rho^|i−j|, in O(n) (forward + backward AR recursions)."""
    n = len(z)
    f = np.empty(n)
    b = np.empty(n)
    acc = 0.0
    for i in range(n):
        acc = z[i] + rho * acc
        f[i] = acc
    acc = 0.0
    for i in range(n - 1, -1, -1):
        acc = z[i] + rho * acc
        b[i] = acc
    return f + b - z


def _wyz(w: np.ndarray, z: np.ndarray, rho: float) -> float:
    """VB ``WYZCalc``: Σⱼ wⱼ Σᵢ zᵢ ρ^|i−j|."""
    return float(np.dot(w, _toeplitz_apply(np.asarray(z, float), rho)))


@dataclass
class AutocorrelatedUncertainty:
    rho: float
    n_pre: int
    n_post: int
    savings: float
    post_use: float
    adjusted_baseline: float
    model_error: float
    measurement_error: float
    savings_error: float

    def to_dict(self) -> dict:
        return _clean(asdict(self))


def autocorrelated_savings_uncertainty(
    X_pre: np.ndarray,
    y_pre: np.ndarray,
    X_post: np.ndarray,
    y_post: np.ndarray,
    *,
    rho: float | None = None,
    measurement_error_pct: float = 0.0,
) -> AutocorrelatedUncertainty:
    """Savings uncertainty for linear baselines (Mean / 2P / MLR) with AR(1) residuals.

    Port of the "project savings" math in ``ProjectSavingsCmd_Click``. ``X_pre``/``X_post`` are
    the regressors *without* the intercept column (shape (n, 0) for a Mean model). The lag-1
    autocorrelation ρ is estimated from the baseline residuals as Σ rᵢ·rᵢ₋₁ / Σᵢ≥₂ rᵢ² unless given.

    Note: in VB only the Mean branch actually estimated ρ. For 2P/MLR the value came back from
    ``Inf`` through a misspelt variable (``pdemnom``) and was therefore always 0, which reduces
    this formula to the ordinary (uncorrelated) prediction variance. Pass ``rho=0`` to reproduce.
    """
    Xp = np.asarray(X_pre, float).reshape(len(y_pre), -1)
    Xq = np.asarray(X_post, float).reshape(len(y_post), -1)
    y_pre = np.asarray(y_pre, float)
    y_post = np.asarray(y_post, float)
    n, m = len(y_pre), len(y_post)
    # centre regressors on the pre means (as VB does) and add intercept
    mu = Xp.mean(axis=0) if Xp.shape[1] else np.empty(0)
    X = np.column_stack([np.ones(n), Xp - mu])
    Xpost = np.column_stack([np.ones(m), Xq - mu])
    p = X.shape[1]
    Q = np.linalg.pinv(X.T @ X) @ X.T  # (p, n)
    beta = Q @ y_pre
    resid = y_pre - X @ beta
    rmse = math.sqrt(float(resid @ resid) / (n - p)) if n > p else 0.0
    if rho is None:
        den = float(np.sum(resid[1:] ** 2))
        rho = float(np.sum(resid[1:] * resid[:-1]) / den) if den else 0.0
    if abs(1 - rho) < 1e-12:
        raise ValueError("Autocorrelation of 1 makes the uncertainty undefined.")

    adjusted = float((Xpost @ beta).sum())
    post_use = float(y_post.sum())
    savings = adjusted - post_use

    xb = Xpost.mean(axis=0)  # [1, mean centred post x]
    k = rho * (1 - rho**m) / (1 - rho)
    v = k * rho ** (n - 1 - np.arange(n))
    a = X[-1] @ Q
    a[-1] -= 1
    mvec = -k * a
    xbq = xb @ Q
    c11 = _wyz(xbq, xbq, rho) * m**2
    c23 = -2 * float(mvec @ v) + _wyz(mvec, mvec, rho)
    c12 = 2 * m * (_wyz(xbq, mvec, rho) - float(xbq @ v))
    c22 = (2 * rho ** (m + 1) - m * (rho**2 - 1) - 2 * rho) / (1 - rho) ** 2
    trace = sum(_wyz(X[:, i], Q.T[:, i], rho) for i in range(p))
    sig2 = rmse**2 * (n - p) / (n - trace)
    var = sig2 * (c11 + c12 + c22 + c23)
    model_err = t_approx(n - 2) * math.sqrt(max(var, 0.0))
    meas_err = post_use * measurement_error_pct / 100
    return AutocorrelatedUncertainty(
        rho=rho,
        n_pre=n,
        n_post=m,
        savings=savings,
        post_use=post_use,
        adjusted_baseline=adjusted,
        model_error=model_err,
        measurement_error=meas_err,
        savings_error=math.hypot(model_err, meas_err),
    )


def autocorrelated_savings_for_model(
    model: BaselineModel, ds: Dataset, *, pre_group=1, post_group=2, **kw
) -> AutocorrelatedUncertainty:
    """Convenience wrapper: pull pre/post rows for a fitted Mean/2P/MLR model."""
    if not isinstance(model, (MeanModel, LinearModel)):
        raise TypeError("Autocorrelation-corrected uncertainty is only available for Mean/2P/MLR.")
    from energy_explorer.modeling import select_xy

    Xa, ya, _ = select_xy(ds, model.y, model.x, pre_group)
    Xb, yb, _ = select_xy(ds, model.y, model.x, post_group)
    return autocorrelated_savings_uncertainty(Xa, ya, Xb, yb, **kw)


# ---------------------------------------------------------------- Lean Energy Analysis
@dataclass
class LeanEnergyResult:
    n: int
    facility: float  # per-observation average of the independent (ycp) part
    weather: float
    production: float
    components: pd.DataFrame = field(repr=False)

    @property
    def total(self) -> float:
        return self.facility + self.weather + self.production

    def shares(self) -> dict[str, float]:
        t = self.total
        return {k: _pct(getattr(self, k), t) for k in ("facility", "weather", "production")}

    def to_dict(self) -> dict:
        return _clean(
            {
                "n": self.n,
                "facility": self.facility,
                "weather": self.weather,
                "production": self.production,
                "total": self.total,
                "shares_pct": self.shares(),
                "totals": {
                    "facility": self.facility * self.n,
                    "weather": self.weather * self.n,
                    "production": self.production * self.n,
                    "total": self.total * self.n,
                },
            }
        )


def lean_energy_analysis(model: BaselineModel, ds: Dataset) -> LeanEnergyResult:
    """Split predicted energy into facility (independent), weather- and production-dependent parts.

    Requires a 3PC/3PH(-MVR) model with X1 = temperature and any further X = production
    variables. As in VB, the breakdown uses **all** records with complete inputs (not just one
    group); returned values are per-observation averages (multiply by ``n`` for totals).
    """
    if not isinstance(model, ThreePModel):
        raise TypeError("Lean Energy breakdown is available only with 3PH-MVR or 3PC-MVR models.")
    cols = [ds.column(c).to_numpy(dtype=float) for c in model.x]
    ok = np.ones(ds.n_records, dtype=bool)
    for c in cols:
        ok &= ~np.isnan(c)
    fac = np.where(ok, model.ycp, np.nan)
    wea = np.where(ok, model.weather_term(cols[0]), np.nan)
    prod = np.where(ok, model._extra_terms(cols, model.extra_coefs) if len(cols) > 1 else 0.0, np.nan)
    comp = pd.DataFrame(
        {"Fac_Use": fac, "Wea_dep_Use": wea, "Prod_dep_Use": prod, "Total_Pred": fac + wea + prod},
        index=ds.data.index,
    )
    n = int(ok.sum())
    mean = (lambda a: float(np.nanmean(a)) if n else 0.0)
    return LeanEnergyResult(n=n, facility=mean(fac), weather=mean(wea), production=mean(prod),
                            components=comp)


def create_model_projection(ds: Dataset, model: BaselineModel, name: str | None = None) -> Dataset:
    """Add the model's prediction for every record as a new variable (``CreateModelProjection``).

    (VB's Median branch mistakenly projected the *mean*; this port uses the median.)
    """
    name = name or f"{model.model_type}_Proj"
    df = ds.data.copy()
    df[name] = model.predict(ds.data)
    return ds.with_data(df)
