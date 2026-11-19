from __future__ import annotations

from pathlib import Path

import pytest
from fastapi.testclient import TestClient

from api.app.main import app
from api.app.store import InMemoryStore, get_store

DATA = Path(__file__).resolve().parents[1] / "data"


@pytest.fixture()
def client():
    store = InMemoryStore()
    app.dependency_overrides[get_store] = lambda: store
    yield TestClient(app)
    app.dependency_overrides.clear()


def test_full_workflow(client):
    assert client.get("/api/health").json()["status"] == "ok"
    with open(DATA / "sample_daily.dat", "rb") as f, open(DATA / "sample_daily.var", "rb") as v:
        r = client.post("/api/datasets", files={"data": ("d.dat", f), "names": ("d.var", v)})
    assert r.status_code == 200, r.text
    ds = r.json()
    assert ds["interval"] == "daily" and "CW" in ds["columns"]

    r = client.post(f"/api/datasets/{ds['id']}/groups", json={"mode": "date", "end": "12/31/2019"})
    assert r.json()["group_counts"] == {"Pre": 365, "Post": 366}

    r = client.post(f"/api/datasets/{ds['id']}/models",
                    json={"model_type": "3PC", "y": "CW", "x": ["Temp"], "group": "Pre"})
    assert r.status_code == 200, r.text
    fit = r.json()
    assert fit["model"]["kind"] == "3p" and len(fit["curve"]["x"]) == 200

    r = client.post(f"/api/models/{fit['id']}/savings", json={"measurement_error_pct": 5})
    body = r.json()
    assert body["savings"]["n_post"] > 300 and body["savings"]["savings"] > 0
    assert len(body["series"]["actual"]) == ds["n_records"]

    r = client.get(f"/api/datasets/{ds['id']}/records", params={"columns": "CW,Temp", "limit": 5})
    assert len(r.json()["columns"]["CW"]) == 5


def test_sample_lea_and_errors(client):
    ds = client.post("/api/datasets/samples/sample_monthly_lea").json()
    fit = client.post(f"/api/datasets/{ds['id']}/models",
                      json={"model_type": "3PH", "y": "NatGas", "x": ["Toa", "Prod"]}).json()
    lea = client.post(f"/api/models/{fit['id']}/lea").json()
    assert set(lea["shares_pct"]) == {"facility", "weather", "production"}
    bad = client.post(f"/api/datasets/{ds['id']}/models", json={"model_type": "9P", "y": "NatGas"})
    assert bad.status_code == 422
    assert client.get("/api/datasets/nope").status_code == 404


def test_timescale_and_simulation(client):
    ds = client.post("/api/datasets/samples/sample_hourly").json()
    r = client.post(f"/api/datasets/{ds['id']}/simulate",
                    json={"params": {"latent": True}, "wbe": "WBE", "toa": "Toa", "woa": "Woa"})
    assert "CC1_(MBtu/hr)" in r.json()["columns"]
    daily = client.post(f"/api/datasets/{ds['id']}/timescale",
                        json={"to": "daily", "averages": ["Toa", "Woa"]}).json()
    assert daily["n_records"] == 365 and daily["id"] != ds["id"]


def test_histogram_and_profile_endpoints(client):
    ds = client.post("/api/datasets/samples/sample_daily").json()
    client.post(f"/api/datasets/{ds['id']}/groups", json={"mode": "date", "end": "12/31/2019"})
    h = client.post(f"/api/datasets/{ds['id']}/histogram",
                    json={"variable": "CW", "n_bins": 12, "by_group": True}).json()
    assert [s["group"] for s in h["series"]] == ["Pre", "Post"] and len(h["edges"]) == 13
    bad = client.post(f"/api/datasets/{ds['id']}/histogram", json={"variable": "nope"})
    assert bad.status_code == 422

    hr = client.post("/api/datasets/samples/sample_hourly").json()
    p = client.post(f"/api/datasets/{hr['id']}/hourly-profile",
                    json={"variable": "Toa", "theta_deg": 45}).json()
    assert len(p["polylines"]) == len(p["dates"]) and len(p["hours"]) == 24
