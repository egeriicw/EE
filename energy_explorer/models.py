"""Fitted baseline models.

Each class holds the parameters the VB program kept in its ``*pre`` globals (``ymeanpre``,
``betapre()``, ``xcppre``, ``ycppre``, ``slopepre``, ``ivcoefspre()``, ``ybarbf()``...) plus the fit
statistics, and knows how to

* ``predict`` from a DataFrame / mapping of columns (NaN where an input is missing),
* render the model ``equation`` and a text ``summary`` like the VB Value Window,
* serialise to / from plain JSON-able dicts (``to_dict`` / ``model_from_dict``) so a fitted
  baseline can be stored by the API and reused for savings calculations.

Fitting lives in ``energy_explorer.changepoint`` and ``energy_explorer.modeling``.
"""

from __future__ import annotations

import math
from collections.abc import Mapping
from dataclasses import asdict, dataclass, field, fields
from typing import Any, ClassVar

import numpy as np
import pandas as pd


@dataclass(kw_only=True)
class ModelStats:
    n: int
    n_params: int  # regression columns, as counted by VB ``Inf`` (change points not counted)
    rmse: float  # for Mean: standard deviation; for Median: SE of the median; BinFit: RMSE
    cv_rmse: float
    r2: float | None = None
    adj_r2: float | None = None
    ymean: float | None = None
    autocorrelation: float | None = None
    durbin_watson: float | None = None
    dw_test: str | None = None
    prediction_interval: float | None = None  # ± half-width (approx 95%)


def _clean(v: Any) -> Any:
    """Make values JSON-friendly (numpy -> python, NaN -> None)."""
    if isinstance(v, dict):
        return {k: _clean(x) for k, x in v.items()}
    if isinstance(v, (list, tuple, np.ndarray)):
        return [_clean(x) for x in v]
    if isinstance(v, (np.floating, float)):
        f = float(v)
        return None if math.isnan(f) or math.isinf(f) else f
    if isinstance(v, np.integer):
        return int(v)
    return v


def _fmt(v: float, spec: str = ",.2f") -> str:
    return "nan" if v is None or (isinstance(v, float) and math.isnan(v)) else format(v, spec)


_REGISTRY: dict[str, type[BaselineModel]] = {}


@dataclass(kw_only=True)
class BaselineModel:
    kind: ClassVar[str] = "base"

    model_type: str
    y: str
    x: list[str] = field(default_factory=list)
    stats: ModelStats
    #: residuals aligned to the source dataset rows (NaN where not used). Not serialised.
    residuals: pd.Series | None = field(default=None, repr=False, compare=False)

    def __init_subclass__(cls, **kw):
        super().__init_subclass__(**kw)
        _REGISTRY[cls.kind] = cls

    # -------------------------------------------------------------- prediction
    def _inputs(self, data: pd.DataFrame | Mapping[str, Any]) -> list[np.ndarray]:
        cols = []
        for name in self.x:
            if name not in data:
                raise KeyError(f"Model needs variable {name!r}")
            cols.append(np.asarray(data[name], dtype=float))
        return cols

    def predict(self, data: pd.DataFrame | Mapping[str, Any]) -> np.ndarray:
        cols = self._inputs(data)
        if cols:
            out = self._predict(cols)
            bad = np.zeros_like(out, dtype=bool)
            for c in cols:
                bad |= np.isnan(c)
            return np.where(bad, np.nan, out)
        n = len(next(iter(data.values()))) if isinstance(data, Mapping) else len(data)
        return self._predict_const(n)

    def _predict(self, cols: list[np.ndarray]) -> np.ndarray:  # pragma: no cover - abstract
        raise NotImplementedError

    def _predict_const(self, n: int) -> np.ndarray:  # pragma: no cover - abstract
        raise NotImplementedError

    @property
    def n_extra(self) -> int:
        return max(len(self.x) - 1, 0)

    def _extra_terms(self, cols: list[np.ndarray], coefs: list[float]) -> np.ndarray:
        total = np.zeros_like(cols[0])
        for c, b in zip(cols[1:], coefs):
            total = total + b * c
        return total

    # -------------------------------------------------------------- presentation
    def params(self) -> dict[str, Any]:
        skip = {"model_type", "y", "x", "stats", "residuals"}
        return {f.name: getattr(self, f.name) for f in fields(self) if f.name not in skip}

    def equation(self) -> str:  # pragma: no cover - abstract
        raise NotImplementedError

    def summary(self) -> str:
        s = self.stats
        parts = [f"Type = {self.model_type}", f"N = {s.n}"]
        if s.r2 is not None:
            parts.append(f"R2 = {_fmt(s.r2)}")
        label = {"Mean": "Std Dev", "Median": "Std Err"}.get(self.model_type, "RMSE")
        parts.append(f"{label} = {_fmt(s.rmse, ',.4f')}")
        parts.append(f"CV-{label} = {_fmt(s.cv_rmse)}%")
        if s.durbin_watson is not None:
            parts.append(f"DW = {_fmt(s.durbin_watson)}")
        if s.autocorrelation is not None:
            parts.append(f"AutoCorr = {_fmt(s.autocorrelation)}")
        return "  ".join(parts) + "\nModel: " + self.equation()

    def to_dict(self) -> dict[str, Any]:
        d = {f.name: getattr(self, f.name) for f in fields(self) if f.name != "residuals"}
        d["stats"] = asdict(self.stats)
        d["kind"] = self.kind
        d["equation"] = self.equation()
        return _clean(d)

    def __repr__(self) -> str:  # pragma: no cover - cosmetic
        return f"<{type(self).__name__} {self.model_type}: {self.equation()}>"


def model_from_dict(d: Mapping[str, Any]) -> BaselineModel:
    d = dict(d)
    cls = _REGISTRY[d.pop("kind")]
    d.pop("equation", None)
    d["stats"] = ModelStats(**d["stats"])
    names = {f.name for f in fields(cls)}
    kwargs = {k: v for k, v in d.items() if k in names}
    for k, v in list(kwargs.items()):
        if v is None and k in {"se", "extra_coefs", "extra_se", "bin_means"}:
            kwargs[k] = []
    if "bin_means" in kwargs:
        kwargs["bin_means"] = [np.nan if v is None else v for v in kwargs["bin_means"]]
    return cls(**kwargs)


# ======================================================================== concrete models
@dataclass(kw_only=True, repr=False)
class MeanModel(BaselineModel):
    kind: ClassVar[str] = "mean"
    mean: float
    std: float

    def _predict_const(self, n: int) -> np.ndarray:
        return np.full(n, self.mean)

    def _predict(self, cols):
        return np.full_like(cols[0], self.mean)

    def equation(self) -> str:
        return f"{self.y} = {_fmt(self.mean)}"


@dataclass(kw_only=True, repr=False)
class MedianModel(BaselineModel):
    kind: ClassVar[str] = "median"
    median: float
    se_median: float

    def _predict_const(self, n: int) -> np.ndarray:
        return np.full(n, self.median)

    def _predict(self, cols):
        return np.full_like(cols[0], self.median)

    def equation(self) -> str:
        return f"{self.y} = {_fmt(self.median)}"


@dataclass(kw_only=True, repr=False)
class LinearModel(BaselineModel):
    """2P (one X) and MLR (several X): y = b0 + b1·x1 + b2·x2 + ..."""

    kind: ClassVar[str] = "linear"
    coefs: list[float]
    se: list[float] = field(default_factory=list)

    def _predict(self, cols):
        out = np.full_like(cols[0], self.coefs[0])
        for c, b in zip(cols, self.coefs[1:]):
            out = out + b * c
        return out

    def _predict_const(self, n: int) -> np.ndarray:
        return np.full(n, self.coefs[0])

    def equation(self) -> str:
        terms = [_fmt(self.coefs[0])] + [f"{_fmt(b)} {n}" for b, n in zip(self.coefs[1:], self.x)]
        return f"{self.y} = " + "  +  ".join(terms)


@dataclass(kw_only=True, repr=False)
class ThreePModel(BaselineModel):
    """3PC: y = ycp + slope·(x − xcp)+     3PH: y = ycp + slope·(x − xcp)−   (+ Σ bₖ·xₖ for MVR).

    For heating (3PH) the fitted ``slope`` is negative (energy rises as x falls).
    """

    kind: ClassVar[str] = "3p"
    heating: bool
    xcp: float
    xcp_step: float  # final grid step; VB reports it as "SE" of Xcp
    ycp: float
    se_ycp: float
    slope: float
    se_slope: float
    extra_coefs: list[float] = field(default_factory=list)
    extra_se: list[float] = field(default_factory=list)

    def indicator(self, x: np.ndarray) -> np.ndarray:
        return (x <= self.xcp).astype(float) if self.heating else (x > self.xcp).astype(float)

    def weather_term(self, x: np.ndarray) -> np.ndarray:
        return self.slope * (x - self.xcp) * self.indicator(x)

    def _predict(self, cols):
        return self.ycp + self.weather_term(cols[0]) + self._extra_terms(cols, self.extra_coefs)

    @property
    def left_slope(self) -> float:
        return self.slope if self.heating else 0.0

    @property
    def right_slope(self) -> float:
        return 0.0 if self.heating else self.slope

    def equation(self) -> str:
        x1 = self.x[0]
        eq = (
            f"{self.y} = {_fmt(self.ycp)}  -  {_fmt(self.left_slope)} ({_fmt(self.xcp)} - {x1})+"
            f"  +  {_fmt(self.right_slope)} ({x1} - {_fmt(self.xcp)})+"
        )
        for b, n in zip(self.extra_coefs, self.x[1:]):
            eq += f"  +  {_fmt(b, ',.4g')} {n}"
        return eq


@dataclass(kw_only=True, repr=False)
class FourPModel(BaselineModel):
    """y = b0 + b1·x + b2·(x − xcp)+  ≡  ycp + ls·(x − xcp)− + rs·(x − xcp)+  (+ MVR terms)."""

    kind: ClassVar[str] = "4p"
    xcp: float
    xcp_step: float
    ycp: float
    se_ycp: float
    left_slope: float
    se_left_slope: float
    right_slope: float
    se_right_slope: float
    beta: list[float]  # [b0, b1, b2] of the indicator regression
    extra_coefs: list[float] = field(default_factory=list)
    extra_se: list[float] = field(default_factory=list)

    def _predict(self, cols):
        x = cols[0]
        b0, b1, b2 = self.beta
        base = b0 + b1 * x + b2 * (x - self.xcp) * (x > self.xcp)
        return base + self._extra_terms(cols, self.extra_coefs)

    def equation(self) -> str:
        x1 = self.x[0]
        eq = (
            f"{self.y} = {_fmt(self.ycp)}  -  {_fmt(self.left_slope)} ({_fmt(self.xcp)} - {x1})+"
            f"  +  {_fmt(self.right_slope)} ({x1} - {_fmt(self.xcp)})+"
        )
        for b, n in zip(self.extra_coefs, self.x[1:]):
            eq += f"  +  {_fmt(b, ',.4g')} {n}"
        return eq


@dataclass(kw_only=True, repr=False)
class FivePModel(BaselineModel):
    """y = ycp + ls·(x − xcp1)− + rs·(x − xcp2)+  (+ MVR terms)."""

    kind: ClassVar[str] = "5p"
    xcp1: float
    xcp2: float
    xcp_step: float
    ycp: float
    se_ycp: float
    left_slope: float
    se_left_slope: float
    right_slope: float
    se_right_slope: float
    extra_coefs: list[float] = field(default_factory=list)
    extra_se: list[float] = field(default_factory=list)

    def _predict(self, cols):
        x = cols[0]
        out = (
            self.ycp
            + self.left_slope * (x - self.xcp1) * (x <= self.xcp1)
            + self.right_slope * (x - self.xcp2) * (x > self.xcp2)
        )
        return out + self._extra_terms(cols, self.extra_coefs)

    def equation(self) -> str:
        x1 = self.x[0]
        eq = (
            f"{self.y} = {_fmt(self.ycp)}  -  {_fmt(self.left_slope)} ({_fmt(self.xcp1)} - {x1})+"
            f"  +  {_fmt(self.right_slope)} ({x1} - {_fmt(self.xcp2)})+"
        )
        for b, n in zip(self.extra_coefs, self.x[1:]):
            eq += f"  +  {_fmt(b, ',.4g')} {n}"
        return eq


@dataclass(kw_only=True, repr=False)
class BinFitModel(BaselineModel):
    """Mean of y in equal-width x bins over [xmin, xmax].

    Bin 1 is closed [lo, hi]; later bins are half-open (lo, hi]. Values outside [xmin, xmax]
    cannot be predicted (NaN). Empty bins borrow from neighbours exactly like the VB savings code:
    bin 1 uses the next non-empty bin; other bins average the nearest non-empty bins on each side
    (or use whichever side exists).
    """

    kind: ClassVar[str] = "binfit"
    xmin: float
    xmax: float
    n_bins: int
    bin_means: list[float]
    bin_counts: list[int] = field(default_factory=list)

    @property
    def edges(self) -> np.ndarray:
        return self.xmin + np.arange(self.n_bins + 1) * (self.xmax - self.xmin) / self.n_bins

    @property
    def midpoints(self) -> np.ndarray:
        e = self.edges
        return (e[:-1] + e[1:]) / 2

    def bin_index(self, x: np.ndarray) -> np.ndarray:
        """0-based bin per value, -1 if outside."""
        x = np.asarray(x, dtype=float)
        edges = self.edges
        # searchsorted(side="left") puts a value equal to an upper edge in the lower bin: (lo, hi]
        idx = np.searchsorted(edges, np.nan_to_num(x, nan=np.inf), side="left") - 1
        idx = np.where(x == self.xmin, 0, idx)
        ok = (x >= self.xmin) & (x <= edges[-1]) & ~np.isnan(x)
        return np.where(ok, np.clip(idx, 0, self.n_bins - 1), -1).astype(int)

    def effective_means(self) -> np.ndarray:
        m = np.array(self.bin_means, dtype=float)
        eff = m.copy()
        for j in range(len(m)):
            if not np.isnan(m[j]):
                continue
            if j == 0:
                nxt = next((m[k] for k in range(1, len(m)) if not np.isnan(m[k])), np.nan)
                eff[j] = nxt
            else:
                prv = next((m[k] for k in range(j - 1, -1, -1) if not np.isnan(m[k])), np.nan)
                nxt = next((m[k] for k in range(j + 1, len(m)) if not np.isnan(m[k])), np.nan)
                if not np.isnan(prv) and not np.isnan(nxt):
                    eff[j] = (prv + nxt) / 2
                else:
                    eff[j] = nxt if np.isnan(prv) else prv
        return eff

    def _predict(self, cols):
        idx = self.bin_index(cols[0])
        eff = self.effective_means()
        return np.where(idx >= 0, eff[np.clip(idx, 0, None)], np.nan)

    def equation(self) -> str:
        return f"{self.y} = mean of each of {self.n_bins} intervals of {self.x[0]}"
