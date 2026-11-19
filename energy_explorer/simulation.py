"""Simplified two-zone air-side HVAC simulation (port of ``RunSimulationCmd_Click`` / ESIM.BAS by
Kelly Kissock).

A building is split into an interior zone (fraction ``fi`` of internal gains) and an exterior zone
(remaining gains + envelope load UA·(Toa − Tr) [+ optional solar]). A constant-air-volume (CAV)
dual-duct / reheat system or a variable-air-volume (VAV) system meets each zone load by mixing
cold-deck (Tcc) and hot-deck (Thc) air. Options: air-side economizer, hot-deck reset schedules
and latent cooling loads. Output per record: cooling-coil and heating-coil loads (MBtu/hr) and
EDE = (Qcc − Qhc)/(Qcc + Qhc)·100 (%).

Inputs (by column name, not by the hard-coded column numbers 9/5/6/7 of the VB code):
whole-building electricity ``wbe`` (kWh/h), outdoor temperature ``toa`` (°F), outdoor humidity
ratio ``woa`` (lbw/lba, latent only) and ``hsol`` (Btu/ft²·h, solar only).
"""

from __future__ import annotations

from dataclasses import asdict, dataclass
from typing import Literal

import numpy as np
import pandas as pd

from energy_explorer.dataset import Dataset

PCP = 0.018  # Btu/ft³·°F  (air density × specific heat)
PHFG = 80.0  # Btu/ft³     (air density × latent heat of vaporisation)
S1 = 10_000.0  # Btu/hr solar load constant
S2 = 400.0  # Btu/hr·°F solar load slope


@dataclass
class SystemParams:
    """Defaults are taken from LOADDB/SYSTEMDB dialog defaults and ESIM.BAS comments."""

    floor_area: float = 300_000.0  # ft²
    fi: float = 0.65  # interior-zone fraction
    ua: float = 20_000.0  # Btu/hr·°F
    tr: float = 72.0  # room set point °F
    n_occupants: float = 0.0  # [assumption] no default in the VB dialog
    latent: bool = False
    system: Literal["cav", "vav"] = "vav"
    cav_flow: float = 1.67  # cfm/ft²
    vav_min_flow: float = 0.33  # cfm/ft²
    tcc: float = 55.0  # cooling-coil set point °F
    economizer: bool = False
    foa_fixed: float = 0.1  # fixed outside-air fraction
    hot_deck: Literal["fixed", "standard", "optimized"] = "fixed"
    thc_fixed: float = 120.0  # °F
    solar: bool = False  # VB hard-coded solaron = 0

    def to_dict(self) -> dict:
        return asdict(self)


def simulate(
    ds: Dataset | pd.DataFrame,
    params: SystemParams | None = None,
    *,
    wbe: str = "WBE",
    toa: str = "Toa",
    woa: str | None = None,
    hsol: str | None = None,
) -> pd.DataFrame:
    """Run the simulation; returns a DataFrame with ``CC``, ``HC`` (MBtu/hr), ``EDE`` (%) and a
    boolean ``infeasible`` flag (the VB warning "HVAC system could not meet the zone loads")."""
    p = params or SystemParams()
    df = ds.data if isinstance(ds, Dataset) else ds
    WBE = df[wbe].to_numpy(dtype=float)
    TOA = df[toa].to_numpy(dtype=float)
    WOA = df[woa].to_numpy(dtype=float) if woa else np.full(len(df), np.nan)
    HSOL = df[hsol].to_numpy(dtype=float) if hsol else np.full(len(df), np.nan)
    if p.latent and woa is None:
        raise ValueError("Latent loads need an outdoor humidity-ratio column (woa).")
    if p.solar and hsol is None:
        raise ValueError("Solar loads need a solar column (hsol).")

    vr = p.cav_flow * p.floor_area * 60  # ft³/hr
    vmini = p.vav_min_flow * p.floor_area * 60 * p.fi
    vmine = p.vav_min_flow * p.floor_area * 60 * (1 - p.fi)
    tr, tcc, fi = p.tr, p.tcc, p.fi

    n = len(df)
    cc = np.full(n, np.nan)
    hc = np.full(n, np.nan)
    ede = np.full(n, np.nan)
    bad = np.zeros(n, dtype=bool)

    for i in range(n):
        w, t = WBE[i], TOA[i]
        if np.isnan(w) or np.isnan(t) or (p.latent and np.isnan(WOA[i])) or (
            p.solar and np.isnan(HSOL[i])
        ):
            continue
        qi = w * 3412 + p.n_occupants * 255
        qli = fi * qi
        qle = (1 - fi) * qi + p.ua * (t - tr)
        if p.solar:
            qle += S1 + S2 * t  # as in VB (solar column itself is not used by the formula)

        # economizer (computed before the optimized hot-deck reset, which needs foa; VB used the
        # previous record's foa here — documented fix)
        if p.economizer:
            if t <= tcc:
                foa = (tcc - tr) / (t - tr)
            elif t <= tr:
                foa = 1.0
            else:
                foa = 0.1
        else:
            foa = p.foa_fixed

        if p.hot_deck == "standard":
            thc = 120.0 if t <= 50 else (80.0 if t >= 80 else 120 - (40 / 30) * (t - 50))
        elif p.hot_deck == "optimized":
            thc1 = foa * t + (1 - foa) * tr + 1
            thc2 = tr - qle / ((1 - fi) * vr * PCP) + 1
            thc = max(thc1, thc2)
        else:
            thc = p.thc_fixed

        if p.latent:
            qlv = vr * foa * PHFG * (WOA[i] - 0.0092) if WOA[i] > 0.0092 else 0.0
            qcl = (qlv + p.n_occupants * 255) / 1e6
        else:
            qcl = 0.0

        if p.system == "cav":
            vri, vre = fi * vr, (1 - fi) * vr
            vci = (vri * PCP * (tr - thc) - qli) / (PCP * (tcc - thc))
            vce = (vre * PCP * (tr - thc) - qle) / (PCP * (tcc - thc))
            vcc = vci + vce
            vhc = vr - vcc
        else:
            vci, vhi = _vav_zone(qli, vmini, tr, tcc, thc)
            vce, vhe = _vav_zone(qle, vmine, tr, tcc, thc)
            vcc = vci + vce
            vhc = vhi + vhe

        tma = foa * t + (1 - foa) * tr
        qcc = vcc * PCP * (tma - tcc) / 1e6 + qcl
        qhc = vhc * PCP * (thc - tma) / 1e6
        bad[i] = vci < 0 or vce < 0 or vhc < 0 or qcc < 0 or qhc < 0
        cc[i], hc[i] = qcc, qhc
        ede[i] = (qcc - qhc) / (qcc + qhc) * 100 if qcc + qhc > 0 else np.nan

    return pd.DataFrame({"CC": cc, "HC": hc, "EDE": ede, "infeasible": bad}, index=df.index)


def _vav_zone(q: float, vmin: float, tr: float, tcc: float, thc: float) -> tuple[float, float]:
    """VAV zone energy balance → (cold-deck flow, hot-deck flow), ft³/hr."""
    if q > 0:
        vc, vh = q / (PCP * (tr - tcc)), 0.0
        if vc < vmin:
            vc = (vmin * PCP * (tr - thc) - q) / (PCP * (tcc - thc))
            vh = vmin - vc
    else:
        vh, vc = q / (PCP * (tr - thc)), 0.0
        if vh < vmin:
            vc = (vmin * PCP * (tr - thc) - q) / (PCP * (tcc - thc))
            vh = vmin - vc
    return vc, vh


def add_simulation(ds: Dataset, params: SystemParams | None = None, *, run: int = 1, **cols) -> Dataset:
    """Run ``simulate`` and append ``CC{run}_(MBtu/hr)``, ``HC{run}_(MBtu/hr)``, ``EDE{run}_(%)``."""
    out = simulate(ds, params, **cols)
    df = ds.data.copy()
    df[f"CC{run}_(MBtu/hr)"] = out["CC"]
    df[f"HC{run}_(MBtu/hr)"] = out["HC"]
    df[f"EDE{run}_(%)"] = out["EDE"]
    return ds.with_data(df)
