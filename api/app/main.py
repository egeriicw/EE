"""FastAPI application exposing the energy_explorer library.

Run from the repository root::

    uvicorn api.app.main:app --reload

Interactive docs at http://localhost:8000/docs. The React client in ``web/`` proxies ``/api`` here.
"""

from __future__ import annotations

import os

from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware

import energy_explorer

from .routers import analysis, datasets, models

app = FastAPI(
    title="Energy Explorer API",
    version=energy_explorer.__version__,
    description="Change-point regression, M&V savings and HVAC simulation for measured energy data "
    "(Python port of ENERGY EXPLORER, University of Dayton).",
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=os.environ.get("EE_CORS_ORIGINS", "http://localhost:5173").split(","),
    allow_methods=["*"],
    allow_headers=["*"],
)

app.include_router(datasets.router)
app.include_router(models.router)
app.include_router(analysis.router)


@app.get("/api/health", tags=["meta"])
def health():
    return {"status": "ok", "version": energy_explorer.__version__}
