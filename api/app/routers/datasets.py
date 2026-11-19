"""Dataset endpoints: upload, inspect, group, re-scale, derive variables, simulate."""

from __future__ import annotations

import math
from pathlib import Path

import pandas as pd
from fastapi import APIRouter, Depends, File, Form, HTTPException, UploadFile

import energy_explorer as ee
from energy_explorer.simulation import SystemParams, add_simulation

from ..schemas import (
    CreateVariableRequest,
    DatasetSummary,
    GroupRequest,
    ModifyRequest,
    SimulationRequest,
    TimescaleRequest,
)
from ..store import NotFound, Store, get_store

router = APIRouter(prefix="/api/datasets", tags=["datasets"])
SAMPLES = Path(__file__).resolve().parents[3] / "data"


def summary(ds_id: str, ds: ee.Dataset) -> DatasetSummary:
    return DatasetSummary(id=ds_id, **ds.summary())


def _get(store: Store, ds_id: str) -> ee.Dataset:
    try:
        return store.get_dataset(ds_id)
    except NotFound:
        raise HTTPException(404, f"Dataset {ds_id} not found") from None


def _bad(exc: Exception) -> HTTPException:
    return HTTPException(422, str(exc))


@router.get("", response_model=list[DatasetSummary])
def list_datasets(store: Store = Depends(get_store)):
    return [summary(k, v) for k, v in store.list_datasets().items()]


@router.get("/samples", response_model=list[str])
def list_samples():
    return sorted(p.stem for p in SAMPLES.glob("*.dat"))


@router.post("/samples/{name}", response_model=DatasetSummary)
def load_sample(name: str, store: Store = Depends(get_store)):
    path = SAMPLES / f"{name}.dat"
    if not path.is_file() or path.parent != SAMPLES:
        raise HTTPException(404, f"Unknown sample {name!r}")
    ds = ee.read_data(path)
    return summary(store.put_dataset(ds), ds)


@router.post("", response_model=DatasetSummary)
async def upload(
    data: UploadFile = File(..., description="ASCII data file (space/comma/tab delimited)"),
    names: UploadFile | None = File(None, description="optional .var names file"),
    no_data: float = Form(-99.0),
    store: Store = Depends(get_store),
):
    try:
        text = (await data.read()).decode("utf-8", errors="replace")
        names_text = (await names.read()).decode("utf-8", errors="replace") if names else None
        ds = ee.read_data_text(text, names_text, no_data=no_data, name=data.filename or "upload")
    except ValueError as exc:
        raise _bad(exc) from exc
    return summary(store.put_dataset(ds), ds)


@router.get("/{ds_id}", response_model=DatasetSummary)
def get_dataset(ds_id: str, store: Store = Depends(get_store)):
    return summary(ds_id, _get(store, ds_id))


def _jsonable(values) -> list:
    return [None if (v is None or (isinstance(v, float) and math.isnan(v))) else v for v in values]


@router.get("/{ds_id}/records")
def records(
    ds_id: str,
    columns: str | None = None,
    offset: int = 0,
    limit: int = 5000,
    store: Store = Depends(get_store),
):
    """Column-oriented slice of the data plus ISO timestamps and group numbers (for charts)."""
    ds = _get(store, ds_id)
    cols = columns.split(",") if columns else ds.columns
    sl = slice(offset, offset + limit)
    try:
        out = {c: _jsonable(ds.column(c).iloc[sl].astype(float).tolist()) for c in cols}
    except KeyError as exc:
        raise _bad(exc) from exc
    ts = ds.timestamps().iloc[sl]
    return {
        "offset": offset,
        "total": ds.n_records,
        "timestamp": [None if pd.isna(t) else str(t) for t in ts.tolist()],
        "group": ds.groups[sl].tolist() if ds.groups is not None else None,
        "columns": out,
    }


@router.post("/{ds_id}/groups", response_model=DatasetSummary)
def set_groups(ds_id: str, req: GroupRequest, store: Store = Depends(get_store)):
    ds = _get(store, ds_id)
    try:
        if req.mode == "dow":
            ds = ee.group_by_day_of_week(ds)
        else:
            ds = ee.group_by_date(ds, req.start, req.end, names=tuple(req.names))
    except (ValueError, TypeError) as exc:
        raise _bad(exc) from exc
    store.replace_dataset(ds_id, ds)
    return summary(ds_id, ds)


@router.post("/{ds_id}/timescale", response_model=DatasetSummary)
def timescale(ds_id: str, req: TimescaleRequest, store: Store = Depends(get_store)):
    """Returns a NEW dataset (the original is kept, like VB's ability to return to hourly)."""
    try:
        ds = ee.change_timescale(_get(store, ds_id), req.to, averages=req.averages,
                                 strict_no_data=req.strict_no_data)
    except (ValueError, KeyError) as exc:
        raise _bad(exc) from exc
    return summary(store.put_dataset(ds, parent=ds_id), ds)


@router.post("/{ds_id}/variables", response_model=DatasetSummary)
def create_variable(ds_id: str, req: CreateVariableRequest, store: Store = Depends(get_store)):
    try:
        ds = ee.create_variable(_get(store, ds_id), req.name, req.a, req.op, req.b)
    except (ValueError, KeyError) as exc:
        raise _bad(exc) from exc
    store.replace_dataset(ds_id, ds)
    return summary(ds_id, ds)


@router.get("/meta/modifiers")
def modifiers():
    return {k: label for k, (label, _) in ee.MODIFIERS.items()}


@router.post("/{ds_id}/modify", response_model=DatasetSummary)
def modify(ds_id: str, req: ModifyRequest, store: Store = Depends(get_store)):
    try:
        ds = ee.modify_values(_get(store, ds_id), req.column, req.modifier, req.constant)
    except (ValueError, KeyError) as exc:
        raise _bad(exc) from exc
    store.replace_dataset(ds_id, ds)
    return summary(ds_id, ds)


@router.post("/{ds_id}/simulate", response_model=DatasetSummary)
def simulate(ds_id: str, req: SimulationRequest, store: Store = Depends(get_store)):
    ds = _get(store, ds_id)
    try:
        params = SystemParams(**req.params)
        run = 1 + sum(c.startswith("CC") and c.endswith("_(MBtu/hr)") for c in ds.columns)
        ds = add_simulation(ds, params, run=run, wbe=req.wbe, toa=req.toa, woa=req.woa, hsol=req.hsol)
    except (ValueError, KeyError, TypeError) as exc:
        raise _bad(exc) from exc
    store.replace_dataset(ds_id, ds)
    return summary(ds_id, ds)

