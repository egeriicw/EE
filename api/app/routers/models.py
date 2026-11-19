"""Model endpoints: catalogue, fit, savings, Lean Energy Analysis, projection."""

from __future__ import annotations

import math

import numpy as np
from fastapi import APIRouter, Depends, HTTPException

import energy_explorer as ee
from energy_explorer.modeling import select_xy
from energy_explorer.savings import autocorrelated_savings_for_model

from ..schemas import FitRequest, FitResponse, SavingsRequest
from ..store import NotFound, Store, get_store
from .datasets import _get, summary

router = APIRouter(prefix="/api", tags=["models"])


def _f(v) -> float | None:
    v = float(v)
    return None if math.isnan(v) or math.isinf(v) else v


@router.get("/model-types")
def model_types():
    return ee.MODEL_TYPES


@router.post("/datasets/{ds_id}/models", response_model=FitResponse)
def fit_model(ds_id: str, req: FitRequest, store: Store = Depends(get_store)):
    ds = _get(store, ds_id)
    try:
        model = ee.fit(ds, req.model_type, req.y, req.x, group=req.group, **req.options)
    except (ValueError, KeyError, TypeError) as exc:
        raise HTTPException(422, str(exc)) from exc
    model_id = store.put_model(model, ds_id)

    # chart payload: observed points used in the fit and (for one-X models) the model curve
    X, y, rows = select_xy(ds, req.y, req.x, ds.group_number(req.group))
    points = {"y": [_f(v) for v in y], "row": rows.tolist()}
    curve = None
    if X.shape[1] >= 1:
        points["x"] = [_f(v) for v in X[:, 0]]
    if X.shape[1] == 1:
        grid = np.linspace(X[:, 0].min(), X[:, 0].max(), 200)
        curve = {"x": grid.tolist(), "y": [_f(v) for v in model.predict({req.x[0]: grid})]}
    return FitResponse(
        id=model_id, dataset_id=ds_id, model=model.to_dict(), summary=model.summary(),
        points=points, curve=curve,
    )


def _model(store: Store, model_id: str):
    try:
        return store.get_model(model_id)
    except NotFound:
        raise HTTPException(404, f"Model {model_id} not found") from None


@router.get("/models/{model_id}")
def get_model(model_id: str, store: Store = Depends(get_store)):
    model, ds_id = _model(store, model_id)
    return {"id": model_id, "dataset_id": ds_id, "model": model.to_dict(), "summary": model.summary()}


@router.post("/models/{model_id}/savings")
def savings(model_id: str, req: SavingsRequest, store: Store = Depends(get_store)):
    model, ds_id = _model(store, model_id)
    ds = _get(store, req.dataset_id or ds_id)
    try:
        res = ee.calculate_savings(model, ds, post_group=req.post_group,
                                   measurement_error_pct=req.measurement_error_pct)
        out = {"savings": res.to_dict(), "report": res.report()}
        if req.autocorrelation_corrected:
            out["autocorrelated"] = autocorrelated_savings_for_model(
                model, ds, post_group=req.post_group,
                measurement_error_pct=req.measurement_error_pct,
            ).to_dict()
    except (ValueError, TypeError, KeyError) as exc:
        raise HTTPException(422, str(exc)) from exc
    ts = ds.timestamps()
    pred = res.predicted
    out["series"] = {
        "timestamp": [str(t) for t in ts.tolist()],
        "actual": [_f(v) for v in ds.column(model.y).tolist()],
        "adjusted_baseline": [_f(v) for v in pred.tolist()],
    }
    return out


@router.post("/models/{model_id}/lea")
def lean_energy(model_id: str, dataset_id: str | None = None, store: Store = Depends(get_store)):
    model, ds_id = _model(store, model_id)
    try:
        return ee.lean_energy_analysis(model, _get(store, dataset_id or ds_id)).to_dict()
    except TypeError as exc:
        raise HTTPException(422, str(exc)) from exc


@router.post("/models/{model_id}/projection")
def projection(model_id: str, name: str | None = None, store: Store = Depends(get_store)):
    """Add the model prediction as a new variable on the model's dataset."""
    model, ds_id = _model(store, model_id)
    ds = ee.create_model_projection(_get(store, ds_id), model, name)
    store.replace_dataset(ds_id, ds)
    return summary(ds_id, ds)
