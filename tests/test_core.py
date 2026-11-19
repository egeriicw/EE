"""Behavioural tests for the Python port.

There are no golden outputs from the original EnergyExplorer.exe in the repository, so these
tests check (a) the documented formulas directly, (b) recovery of known parameters from
noise-free synthetic data, and (c) the VB conventions that are easy to get subtly wrong
(percentile method, change-point tie handling, bin edges, period grouping).
"""

from __future__ import annotations

import json
import math
from pathlib import Path

import numpy as np
import pandas as pd
import pytest

import energy_explorer as ee
from energy_explorer.changepoint import fit_3p, fit_4p, fit_5p, fit_binfit
from energy_explorer.dataset import Dataset
from energy_explorer.savings import autocorrelated_savings_uncertainty
from energy_explorer.stats import dw_test, median_with_se, vb_percentile

DATA = Path(__file__).resolve().parents[1] / "data"


# --------------------------------------------------------------------------- I/O
@pytest.mark.parametrize("sep", [" ", "  ", ",", "\t"])
def test_parse_delimiters_and_no_data(sep):
    text = sep.join(["1", "2", "95", "-99"]) + "\n" + sep.join(["1", "3", "5", "4.5"]) + "\n\n"
    ds = ee.read_data_text(text, ["Mo", "Dy", "Yr", "kWh"])
    assert ds.n_records == 2
    assert np.isnan(ds.data.loc[0, "kWh"])
    assert ds.data["Yr"].tolist() == [1995, 2005]  # two-digit year expansion
    assert ds.interval == ee.TimeInterval.DAILY


def test_default_names_and_mismatch():
    ds = ee.read_data_text("1 2 3\n4 5 6\n")
    assert ds.columns == ["V01", "V02", "V03"]
    with pytest.raises(ValueError, match="variable names"):
        ee.read_data_text("1 2 3\n", ["a", "b"])
    with pytest.raises(ValueError, match="number of columns"):
        ee.read_data_text("1 2 3\n4 5\n")


def test_sample_file_roundtrip(tmp_path):
    ds = ee.read_data(DATA / "sample_daily.dat")
    assert ds.columns[:4] == ["Site", "Mo", "Dy", "Yr"]
    out = ee.write_data(ds, tmp_path / "copy.dat")
    back = ee.read_data(out)
    pd.testing.assert_frame_equal(ds.data, back.data, check_exact=False, rtol=1e-5)


# --------------------------------------------------------------------------- stats
def test_vb_percentile_method():
    v = np.arange(1, 11, dtype=float)
    assert vb_percentile(v, 0.25) == pytest.approx(2.75)  # (n+1)p = 2.75
    assert vb_percentile(v, 0.5) == pytest.approx(5.5)
    assert vb_percentile(v, 0.9) == pytest.approx(9.9)
    assert vb_percentile([3.0], 0.9) == 3.0
    med, se = median_with_se(v)
    assert se == pytest.approx((8.25 - 2.75) / math.sqrt(10))


def test_dw_bounds():
    assert dw_test(2, 100, 0.5) == "p>0"
    assert dw_test(2, 100, 2.5) == "p=0"
    assert dw_test(5, 100, 2.0) is None


# --------------------------------------------------------------------------- models
def _xgrid(n=201, lo=20.0, hi=90.0):
    return np.linspace(lo, hi, n)


def test_linear_and_mean_stats():
    x = np.array([1.0, 2, 3, 4, 5])
    y = np.array([2.1, 3.9, 6.2, 7.8, 10.1])
    ds = Dataset(pd.DataFrame({"x": x, "y": y}))
    m = ee.fit(ds, "2P", y="y", x="x")
    b1, b0 = np.polyfit(x, y, 1)
    assert m.coefs == pytest.approx([b0, b1])
    resid = y - (b0 + b1 * x)
    assert m.stats.rmse == pytest.approx(math.sqrt((resid**2).sum() / 3))
    mean = ee.fit(ds, "Mean", y="y")
    assert mean.stats.rmse == pytest.approx(np.std(y, ddof=1))


@pytest.mark.parametrize("heating", [False, True])
def test_3p_recovers_change_point(heating):
    x = _xgrid()
    xcp, ycp, slope = 61.3, 40.0, (-1.4 if heating else 2.2)
    ind = (x <= xcp) if heating else (x > xcp)
    y = ycp + slope * (x - xcp) * ind
    m, fit = fit_3p(x, y, heating=heating)
    assert abs(m.xcp - xcp) <= m.xcp_step * 1.01
    assert m.slope == pytest.approx(slope, rel=0.02)
    assert m.ycp == pytest.approx(ycp, abs=0.2)
    assert m.model_type == ("3PH" if heating else "3PC")


def test_3p_mvr_extra_variable():
    rng = np.random.default_rng(0)
    x = _xgrid()
    prod = rng.uniform(0, 100, len(x))
    y = 10 + 1.5 * np.clip(x - 55, 0, None) + 0.3 * prod
    m, _ = fit_3p(x, y, heating=False, extra=prod)
    assert m.extra_coefs[0] == pytest.approx(0.3, rel=0.02)
    assert m.xcp == pytest.approx(55, abs=0.2)


def test_4p_recovers_slopes():
    x = _xgrid()
    xcp = 58.0
    y = 50 + np.where(x <= xcp, -0.8 * (x - xcp), 2.0 * (x - xcp))
    m, _ = fit_4p(x, y, n_intervals=100)
    assert m.xcp == pytest.approx(xcp, abs=m.xcp_step * 1.01)
    assert m.left_slope == pytest.approx(-0.8, rel=0.05)
    assert m.right_slope == pytest.approx(2.0, rel=0.05)
    assert m.ycp == pytest.approx(50, abs=0.5)


def test_4p_default_grid_matches_vb_buttons():
    x = _xgrid()
    y = 50 + 2.0 * np.clip(x - 58, 0, None)
    plain, _ = fit_4p(x, y)  # FourPBut_Click: 10 intervals
    mvr, _ = fit_4p(x, y, extra=np.sin(x))  # FourPMVR: 100 intervals
    assert plain.xcp_step == pytest.approx(2 * 7 / 10)
    assert mvr.xcp_step == pytest.approx(2 * 0.7 / 100)


def test_5p_recovers_both_change_points():
    x = _xgrid(301)
    y = 30 + np.where(x <= 45, -1.2 * (x - 45), 0) + np.where(x > 70, 1.8 * (x - 70), 0)
    m, _ = fit_5p(x, y)
    tol = 0.34 * (x.max() - x.min()) / 9
    assert m.xcp1 == pytest.approx(45, abs=tol)
    assert m.xcp2 == pytest.approx(70, abs=tol)
    assert m.xcp1 < m.xcp2


def test_indicator_tie_goes_left():
    m = ee.ThreePModel(
        model_type="3PC", y="y", x=["x"], heating=False, xcp=60.0, xcp_step=0.1, ycp=10.0,
        se_ycp=0, slope=2.0, se_slope=0, stats=ee.models.ModelStats(n=1, n_params=2, rmse=0, cv_rmse=0),
    )
    assert m.predict({"x": [60.0, 61.0]}).tolist() == [10.0, 12.0]


def test_binfit_edges_and_empty_bins():
    x = np.array([0.0, 1.0, 1.5, 2.0, 9.0, 10.0])
    y = np.array([1.0, 1.0, 5.0, 5.0, 9.0, 11.0])
    m, _ = fit_binfit(x, y, n_bins=5)  # edges 0,2,4,6,8,10
    assert m.bin_counts == [4, 0, 0, 0, 2]  # 2.0 belongs to bin 1 (lo, hi]
    assert m.bin_means[0] == pytest.approx(3.0)
    eff = m.effective_means()
    assert eff[1] == pytest.approx((3.0 + 10.0) / 2)  # interior empty -> average of neighbours
    assert np.isnan(m.predict({"x": [10.5]})[0])


def test_model_serialisation_roundtrip():
    ds = ee.read_data(DATA / "sample_daily.dat")
    for mt, x in [("Mean", []), ("Median", []), ("2P", ["Temp"]), ("MLR", ["Temp", "WBE"]),
                  ("3PC", ["Temp"]), ("3PH", ["Temp"]), ("4P", ["Temp"]), ("5P", ["Temp"]),
                  ("BinFit", ["Temp"])]:
        m = ee.fit(ds, mt, y="CW", x=x)
        d = json.loads(json.dumps(m.to_dict()))
        m2 = ee.model_from_dict(d)
        np.testing.assert_allclose(m2.predict(ds.data), m.predict(ds.data), equal_nan=True)
        assert m.residuals is not None and m.residuals.notna().sum() == m.stats.n


# --------------------------------------------------------------------------- data ops
def test_group_by_date_inclusive_and_dow():
    ds = ee.read_data(DATA / "sample_daily.dat")
    g = ee.group_by_date(ds, start="1/1/2019", end="12/31/2019", names=("Pre", "Post"))
    assert (g.groups == 1).sum() == 365 and (g.groups == 2).sum() == 366
    assert g.group_number("post") == 2
    w = ee.group_by_day_of_week(ds)
    # 2019-01-05 is a Saturday
    assert w.groups[4] == 2 and w.groups[0] == 1


def test_timescale_total_average_strict():
    df = pd.DataFrame({"Yr": [2020] * 4, "Mo": [1, 1, 2, 2], "Dy": [1, 2, 1, 2],
                       "E": [1.0, 2.0, 3.0, np.nan], "T": [10.0, 20.0, 30.0, 50.0]})
    ds = Dataset(df)
    m = ee.change_timescale(ds, "monthly", averages=["T"])
    assert m.columns == ["Yr", "Mo", "E", "T"]
    assert m.data["E"].tolist() == [3.0, 3.0]
    assert m.data["T"].tolist() == [15.0, 40.0]
    strict = ee.change_timescale(ds, "monthly", averages=["T"], strict_no_data=True)
    assert np.isnan(strict.data.loc[1, "E"])
    with pytest.raises(ValueError):
        ee.change_timescale(m, "daily")


def test_variables():
    ds = Dataset(pd.DataFrame({"a": [1.0, 2.0, np.nan], "b": [2.0, 0.0, 1.0]}))
    out = ee.create_variable(ds, "r", "a", "/", "b")
    assert out.data["r"].iloc[0] == 0.5 and np.isnan(out.data["r"].iloc[1])
    assert np.isnan(out.data["r"].iloc[2])
    f = ee.modify_values(Dataset(pd.DataFrame({"t": [32.0, 212.0]})), "t", "ftoc")
    assert f.data["t"].tolist() == pytest.approx([0.0, 100.0])
    ln = ee.modify_values(ds, "b", "ln")
    assert np.isnan(ln.data["b"].iloc[1])


# --------------------------------------------------------------------------- savings
def test_savings_formula():
    rng = np.random.default_rng(1)
    t = rng.uniform(40, 90, 200)
    y = 20 + 1.5 * np.clip(t - 60, 0, None) + rng.normal(0, 1, 200)
    y[100:] -= 5
    ds = Dataset(pd.DataFrame({"T": t, "E": y}), groups=np.r_[np.ones(100), 2 * np.ones(100)].astype(int),
                 group_names=("Pre", "Post"))
    m = ee.fit(ds, "3PC", y="E", x="T", group="Pre")
    r = ee.calculate_savings(m, ds, measurement_error_pct=2)
    pred = m.predict(ds.data)[100:]
    assert r.n_post == 100
    assert r.savings == pytest.approx(pred.sum() - y[100:].sum())
    assert r.baseline_error == pytest.approx(1.96 * m.stats.rmse * math.sqrt(1 + 2 / 100) * 10)
    assert r.measurement_error == pytest.approx(0.02 * y[100:].sum())
    assert r.savings == pytest.approx(500, rel=0.1)


def test_autocorrelated_uncertainty_reduces_to_ols_when_rho_zero():
    rng = np.random.default_rng(2)
    xa, xb = rng.uniform(0, 10, 50), rng.uniform(0, 10, 30)
    ya, yb = 3 + 2 * xa + rng.normal(0, 1, 50), 2 + 2 * xb + rng.normal(0, 1, 30)
    r = autocorrelated_savings_uncertainty(xa[:, None], ya, xb[:, None], yb, rho=0.0)
    X = np.column_stack([np.ones(50), xa])
    beta, *_ = np.linalg.lstsq(X, ya, rcond=None)
    rmse = math.sqrt(((ya - X @ beta) ** 2).sum() / 48)
    xbar = np.array([1, xb.mean()])
    var = rmse**2 * (30 + 30**2 * xbar @ np.linalg.inv(X.T @ X) @ xbar)
    assert r.model_error == pytest.approx((1.96 + 2.7 / 48) * math.sqrt(var), rel=1e-6)
    r_ar = autocorrelated_savings_uncertainty(xa[:, None], ya, xb[:, None], yb, rho=0.6)
    assert r_ar.model_error > r.model_error


def test_lean_energy_components_sum_to_prediction():
    ds = ee.read_data(DATA / "sample_monthly_lea.dat")
    m = ee.fit(ds, "3PH", y="NatGas", x=["Toa", "Prod"])
    lea = ee.lean_energy_analysis(m, ds)
    np.testing.assert_allclose(lea.components["Total_Pred"], m.predict(ds.data))
    assert sum(lea.shares().values()) == pytest.approx(100)
    with pytest.raises(TypeError):
        ee.lean_energy_analysis(ee.fit(ds, "2P", y="NatGas", x="Toa"), ds)


# --------------------------------------------------------------------------- simulation
def test_vav_cooling_energy_balance():
    # Toa = Tr (no envelope load), no outdoor air, large internal gains -> all gains go to the
    # cooling coil: Qcc [MBtu/h] = WBE * 3412 / 1e6
    df = pd.DataFrame({"WBE": [2000.0], "Toa": [72.0]})
    p = ee.SystemParams(system="vav", foa_fixed=0.0, n_occupants=0, floor_area=10_000)
    out = ee.simulate(df, p, wbe="WBE", toa="Toa")
    assert out.loc[0, "CC"] == pytest.approx(2000 * 3412 / 1e6)
    assert out.loc[0, "HC"] == pytest.approx(0.0)
    assert not out.loc[0, "infeasible"]


def test_simulation_skips_missing_inputs():
    df = pd.DataFrame({"WBE": [np.nan, 500.0], "Toa": [50.0, 50.0]})
    out = ee.simulate(df, ee.SystemParams(system="cav"), wbe="WBE", toa="Toa")
    assert np.isnan(out.loc[0, "CC"]) and not np.isnan(out.loc[1, "CC"])


# --------------------------------------------------------------------------- review follow-ups
def test_select_xy_with_non_default_index():
    idx = pd.date_range("2020-01-01", periods=6, freq="D")
    df = pd.DataFrame({"x": [1.0, np.nan, 3, 4, 5, 6], "y": [2.0, 4, np.nan, 8, 10, 12]}, index=idx)
    ds = Dataset(df)
    X, y, rows = ee.select_xy(ds, "y", ["x"])
    assert rows.tolist() == [0, 3, 4, 5]
    assert X[:, 0].tolist() == [1, 4, 5, 6] and y.tolist() == [2, 8, 10, 12]
    m = ee.fit(ds, "2P", y="y", x="x")
    assert m.coefs == pytest.approx([0.0, 2.0], abs=1e-9)


def test_hourly_group_dates_need_hour():
    df = pd.DataFrame({"Mo": [1, 1], "Dy": [1, 1], "Yr": [2020, 2020], "Hr": [1, 2], "E": [1.0, 2.0]})
    ds = Dataset(df)
    g = ee.group_by_date(ds, end="1/1/2020/1")
    assert g.groups.tolist() == [1, 2]
    with pytest.raises(ValueError, match="m/d/yyyy/h"):
        ee.group_by_date(ds, end="1/1/2020")


def test_reference_values_from_user_guide_screenshots():
    """Two outputs of the real program quoted in the 2005 user guide.

    Savings screenshot (4P baseline): N_pre = 250, RMSE = 5.33, 416 post observations,
    adjusted-baseline uncertainty ±213.9. LEA screenshot: Ycp = 59.58 on 12 months -> Facility 715.0.
    """
    rng = np.random.default_rng(3)
    stats = ee.models.ModelStats(n=250, n_params=3, rmse=5.33, cv_rmse=0)
    m = ee.MeanModel(model_type="Mean", y="E", mean=60.0, std=5.33, stats=stats)
    y = rng.normal(30, 1, 416)
    ds = Dataset(pd.DataFrame({"E": y}), groups=np.full(416, 2), group_names=("Pre", "Post"))
    r = ee.calculate_savings(m, ds)
    assert round(r.baseline_error, 1) == 213.9

    lea_stats = ee.models.ModelStats(n=12, n_params=3, rmse=1, cv_rmse=0)
    lm = ee.ThreePModel(model_type="3PH", y="G", x=["Toa", "Prod"], heating=True, xcp=62.06,
                        xcp_step=0.1, ycp=59.58, se_ycp=0, slope=-9.37, se_slope=0,
                        extra_coefs=[0.02], extra_se=[0], stats=lea_stats)
    lds = Dataset(pd.DataFrame({"Toa": rng.uniform(20, 80, 12), "Prod": rng.uniform(1e4, 2e4, 12)}))
    lea = ee.lean_energy_analysis(lm, lds)
    assert round(lea.to_dict()["totals"]["facility"], 1) == 715.0


def test_autocorrelated_uncertainty_mean_model_path():
    ds = ee.group_by_date(ee.read_data(DATA / "sample_daily.dat"), end="12/31/2019",
                          names=("Pre", "Post"))
    from energy_explorer.savings import autocorrelated_savings_for_model

    m = ee.fit(ds, "Mean", y="CW", group="Pre")
    r = autocorrelated_savings_for_model(m, ds)
    assert r.n_post == int(((ds.groups == 2) & ds.data["CW"].notna()).sum())
    assert r.savings == pytest.approx(m.mean * r.n_post - r.post_use)
    assert r.model_error > 0 and -1 < r.rho < 1


def test_average_intermittent():
    df = pd.DataFrame({"Yr": [2020] * 3, "Mo": [1, 1, 1], "Dy": [1, 3, 4], "Read": [5.0, 10.0, 4.0],
                       "Other": [1.0, 1.0, 1.0]})
    out = ee.average_intermittent(Dataset(df), ["Read"])
    assert np.isnan(out.data["Read"][0])
    assert out.data["Read"].tolist()[1:] == [5.0, 4.0]
    assert out.data["Other"].tolist() == [1.0, 1.0, 1.0]
