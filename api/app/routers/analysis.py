"""Histogram and hourly-profile endpoints (thin wrappers over ``energy_explorer.analysis``)."""

from __future__ import annotations

from fastapi import APIRouter, Depends, HTTPException

from energy_explorer.analysis import histogram, hourly_profile, profile_polylines

from ..schemas import HistogramRequest, HourlyProfileRequest
from ..store import Store, get_store
from .datasets import _get

router = APIRouter(prefix="/api/datasets", tags=["analysis"])


@router.post("/{ds_id}/histogram")
def get_histogram(ds_id: str, req: HistogramRequest, store: Store = Depends(get_store)):
    ds = _get(store, ds_id)
    try:
        return histogram(
            ds, req.variable, n_bins=req.n_bins, xmin=req.xmin, xmax=req.xmax,
            group=req.group, by_group=req.by_group,
        ).to_dict()
    except (ValueError, KeyError, IndexError) as exc:
        raise HTTPException(422, str(exc)) from exc


@router.post("/{ds_id}/hourly-profile")
def get_hourly_profile(ds_id: str, req: HourlyProfileRequest, store: Store = Depends(get_store)):
    ds = _get(store, ds_id)
    try:
        profile = hourly_profile(ds, req.variable, group=req.group)
    except (ValueError, KeyError, IndexError) as exc:
        raise HTTPException(422, str(exc)) from exc
    out = profile.to_dict()
    out["polylines"] = profile_polylines(profile, theta_deg=req.theta_deg, res=req.res)
    return out

