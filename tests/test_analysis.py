from __future__ import annotations

from pathlib import Path

import numpy as np
import pandas as pd
import pytest

import energy_explorer as ee
from energy_explorer.analysis import (
    bin_index,
    histogram,
    hourly_profile,
    profile_polylines,
    project_3d,
    vb_round_up,
)
from energy_explorer.dataset import Dataset

DATA = Path(__file__).resolve().parents[1] / "data"


def test_bin_convention_first_bin_closed_others_half_open():
    idx = bin_index([0.0, 2.0, 2.0001, 10.0, 11.0, -1.0, np.nan], 0.0, 10.0, 5)
    assert idx.tolist() == [0, 0, 1, 4, -1, -1, -1]


def test_histogram_counts_probability_and_groups():
    df = pd.DataFrame({"E": [1.0, 2, 3, 4, 5, 6, 7, 8]})
    ds = Dataset(df, groups=np.array([1, 1, 1, 1, 2, 2, 2, 2]), group_names=("Pre", "Post"))
    h = histogram(ds, "E", n_bins=4, xmin=0, xmax=8, by_group=True)
    pre, post = h.series
    assert (pre.group, post.group) == ("Pre", "Post")
    assert pre.counts == [2, 2, 0, 0] and post.counts == [0, 0, 2, 2]
    assert sum(pre.probability) == pytest.approx(1.0)
    assert pre.percent[0] == pytest.approx(50.0)
    assert h.midpoints == pytest.approx([1, 3, 5, 7])


def test_histogram_defaults_to_data_range_and_all_records():
    ds = Dataset(pd.DataFrame({"E": [2.0, 4.0, np.nan, 6.0]}))
    h = histogram(ds, "E", n_bins=2)
    assert (h.xmin, h.xmax) == (2.0, 6.0)
    assert h.series[0].n == 3 and sum(h.series[0].counts) == 3
    assert "Histogram of E" in h.text()


def test_histogram_rejects_zero_range():
    with pytest.raises(ValueError, match="xmax > xmin"):
        histogram(Dataset(pd.DataFrame({"E": [3.0, 3.0]})), "E", n_bins=3)


def test_vb_round_up():
    assert vb_round_up(1234) == 2000
    assert vb_round_up(9.5) == 10
    assert vb_round_up(0) == 1.0


def test_project_3d_matches_vb_formula_at_zero_angle_limit():
    # theta = 0: xs = (x - z + res)/2 ... check against direct evaluation
    x, y, z, res = 100.0, 50.0, 30.0, 400.0
    xs, ys = project_3d(x, y, z, 45.0, res)
    c = s = np.sqrt(0.5)
    assert xs == pytest.approx((x - z * c + res * c) / (c + 1))
    assert ys == pytest.approx((y - z * s + res * s) / (s + 1))


def test_hourly_profile_pivot_and_hour_convention():
    ds = ee.read_data(DATA / "sample_hourly.dat")  # Hr is 1..24 (hour-ending)
    # group_by_date compares the raw Hr column (as VB does), so end at 24 to keep all of day 1
    grouped = ee.group_by_date(ds, end="1/1/2019/24", names=("Pre", "Post"))
    p = hourly_profile(grouped, "WBE", group="Pre")
    assert p.hours == list(range(24))
    assert p.dates[0] == "2019-01-01"
    # the 1..24 convention: Hr 1 becomes hour 0, so hour 0 of day 1 is present
    assert p.values[0][0] is not None and p.values[0][23] is not None
    assert p.scale >= max(v for row in p.values for v in row if v is not None)


def test_profile_polylines_one_per_day():
    ds = ee.read_data(DATA / "sample_hourly.dat")
    p = hourly_profile(ds, "Toa")
    lines = profile_polylines(p, theta_deg=45, res=400)
    assert len(lines) == len(p.dates)
    assert len(lines[0]["points"]) == 24
    assert lines[0]["points"][0] is not None


def test_profile_needs_hour_column():
    ds = Dataset(pd.DataFrame({"Yr": [2020], "Mo": [1], "Dy": [1], "E": [1.0]}))
    with pytest.raises(ValueError, match="Hr"):
        hourly_profile(ds, "E")
