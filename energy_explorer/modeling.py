"""High-level model fitting on a ``Dataset`` (replaces ``SetAll`` / ``SetGrp`` / ``FillXYGrp``
plus the model button handlers in MAIN.FRM)."""

from __future__ import annotations

from collections.abc import Sequence

import numpy as np
import pandas as pd

from energy_explorer import changepoint as cp
from energy_explorer.dataset import Dataset
from energy_explorer.models import BaselineModel

#: Model catalogue: type -> (description, min X vars, max X vars).  X1 always carries the change
#: point; additional X variables enter linearly ("MVR"). Limits are the VB program's effective ones:
#: the Multiple Regression dialog offers X1..X7 and ``Invert`` handles at most 8 regression columns
#: (3P: 1 + k, 4P/5P: 2 + k). The 2005 user guide quotes lower limits (4/3/3); the code allows more.
MODEL_TYPES: dict[str, dict] = {
    "Mean": {"label": "Mean", "min_x": 0, "max_x": 0},
    "Median": {"label": "Median", "min_x": 0, "max_x": 0},
    "2P": {"label": "Two-parameter linear", "min_x": 1, "max_x": 1},
    "MLR": {"label": "Multiple linear regression (2P-MVR)", "min_x": 1, "max_x": 7},
    "3PC": {"label": "Three-parameter cooling (± MVR)", "min_x": 1, "max_x": 7},
    "3PH": {"label": "Three-parameter heating (± MVR)", "min_x": 1, "max_x": 7},
    "4P": {"label": "Four-parameter (± MVR)", "min_x": 1, "max_x": 6},
    "5P": {"label": "Five-parameter (± MVR)", "min_x": 1, "max_x": 6},
    "BinFit": {"label": "Bin fit (mean y per x interval)", "min_x": 1, "max_x": 1},
}


def select_xy(
    ds: Dataset, y: str, x: Sequence[str] = (), group: int | str | None = None
) -> tuple[np.ndarray, np.ndarray, np.ndarray]:
    """Rows where y and every x are present (and in ``group``). Returns (X (n,k), y (n,), row index)."""
    gnum = ds.group_number(group)
    mask = ds.group_mask(gnum) & ds.column(y).notna().to_numpy()
    for name in x:
        mask &= ds.column(name).notna().to_numpy()
    rows = np.flatnonzero(mask)
    cols = [ds.data.columns.get_loc(c) for c in x]
    X = ds.data.iloc[rows, cols].to_numpy(dtype=float) if x else np.empty((len(rows), 0))
    return X, ds.column(y).iloc[rows].to_numpy(dtype=float), rows


def fit(
    ds: Dataset,
    model_type: str,
    y: str,
    x: Sequence[str] | str | None = None,
    group: int | str | None = None,
    **options,
) -> BaselineModel:
    """Fit a model of ``y`` on ``x`` using the rows of ``group`` (1/2 or a group name).

    ``model_type``: one of ``MODEL_TYPES``. Options:

    * 3PC/3PH/4P: ``n_intervals`` (grid size; defaults 100, 4P: 10 without MVR / 100 with MVR)
    * 5P: ``n_intervals`` (default 9)
    * BinFit: ``n_bins`` (default 10), ``xmin``, ``xmax``

    For Mean/Median an ``x`` may be given; as in VB it only restricts rows to those where x is
    present (``SetAll`` requires both selected variables).

    The returned model's ``residuals`` is a Series aligned to ``ds.data`` (NaN for unused rows),
    the equivalent of the ``SG_Resids`` / ``G1_Resids`` / ``G2_Resids`` variables.
    """
    if model_type not in MODEL_TYPES:
        raise ValueError(f"Unknown model type {model_type!r}; choose from {list(MODEL_TYPES)}")
    xs = [x] if isinstance(x, str) else list(x or [])
    spec = MODEL_TYPES[model_type]
    if model_type not in ("Mean", "Median") and not (spec["min_x"] <= len(xs) <= spec["max_x"]):
        raise ValueError(
            f"{model_type} needs between {spec['min_x']} and {spec['max_x']} X variables "
            f"(got {len(xs)})."
        )
    X, yv, rows = select_xy(ds, y, xs, group)
    if len(yv) == 0:
        raise ValueError("No data is available to model.")

    extra = X[:, 1:] if X.shape[1] > 1 else None
    resid: np.ndarray
    if model_type == "Mean":
        model = cp.fit_mean(yv, y_name=y)
        resid = yv - model.mean
    elif model_type == "Median":
        model = cp.fit_median(yv, y_name=y)
        resid = yv - model.median
    elif model_type in ("2P", "MLR"):
        model, res = cp.fit_linear(X, yv, y_name=y, x_names=xs)
        model.model_type = "2P" if len(xs) == 1 and model_type == "2P" else "MLR"
        resid = res.resid
    elif model_type in ("3PC", "3PH"):
        model, res = cp.fit_3p(
            X[:, 0], yv, heating=model_type == "3PH", extra=extra, y_name=y, x_names=xs, **options
        )
        resid = res.resid
    elif model_type == "4P":
        model, res = cp.fit_4p(X[:, 0], yv, extra=extra, y_name=y, x_names=xs, **options)
        resid = res.resid
    elif model_type == "5P":
        model, res = cp.fit_5p(X[:, 0], yv, extra=extra, y_name=y, x_names=xs, **options)
        resid = res.resid
    else:  # BinFit
        model, resid = cp.fit_binfit(X[:, 0], yv, y_name=y, x_name=xs[0], **options)

    full = np.full(ds.n_records, np.nan)
    full[rows] = resid
    model.residuals = pd.Series(full, index=ds.data.index, name=f"{model.model_type}_Resids")
    return model
