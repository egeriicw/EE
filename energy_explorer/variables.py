"""Create / modify / delete variables (port of NEWVARDB.FRM, MODVALDB.FRM, DELETEDB.FRM)."""

from __future__ import annotations

import math
from collections.abc import Callable

import numpy as np

from energy_explorer.dataset import Dataset

_OPS: dict[str, Callable[[np.ndarray, np.ndarray], np.ndarray]] = {
    "+": np.add,
    "-": np.subtract,
    "*": np.multiply,
    "/": np.divide,
}


def create_variable(
    ds: Dataset, name: str, a: str, op: str | None = None, b: str | None = None
) -> Dataset:
    """New variable = copy of ``a``, or ``a <op> b`` with op in + - * /.

    Missing if either input is missing; division by zero gives missing (VB error 6 handler).
    More complex expressions are built by chaining calls, as in the original program. (A free-form
    expression API is intentionally left to the caller: ``ds.data.eval`` works on ``ds.data``.)
    """
    if not name:
        raise ValueError("Specify the new variable's name.")
    if name in ds.columns:
        raise ValueError(f"Variable {name!r} already exists.")
    va = ds.column(a).to_numpy(dtype=float)
    if op is None:
        new = va.copy()
    else:
        if op not in _OPS or b is None:
            raise ValueError("New variable must be a copy of one variable or 'A op B' (op: + - * /).")
        vb = ds.column(b).to_numpy(dtype=float)
        with np.errstate(divide="ignore", invalid="ignore"):
            new = _OPS[op](va, vb)
        new[~np.isfinite(new)] = np.nan
    df = ds.data.copy()
    df[name] = new
    return ds.with_data(df)


def delete_variable(ds: Dataset, name: str) -> Dataset:
    ds.column(name)
    return ds.with_data(ds.data.drop(columns=[name]))


def _safe(fn: Callable[[np.ndarray], np.ndarray]) -> Callable[[np.ndarray, float], np.ndarray]:
    def wrapped(v: np.ndarray, c: float) -> np.ndarray:
        return fn(v)

    return wrapped


def _ln(v: np.ndarray) -> np.ndarray:
    with np.errstate(divide="ignore", invalid="ignore"):
        return np.where(v > 0, np.log(np.where(v > 0, v, 1)), np.nan)


def _sqrt(v: np.ndarray) -> np.ndarray:
    return np.where(v > 0, np.sqrt(np.abs(v)), np.nan)


#: Modifiers from MODVALDB.FRM. Keys are the VB internal codes. Each takes (values, constant).
MODIFIERS: dict[str, tuple[str, Callable[[np.ndarray, float], np.ndarray]]] = {
    "ftoc": ("F to C", _safe(lambda v: (v - 32) / 1.8)),
    "ctof": ("C to F", _safe(lambda v: v * 1.8 + 32)),
    "+C": ("( ) + Constant", lambda v, c: v + c),
    "-C": ("( ) - Constant", lambda v, c: v - c),
    "*C": ("( ) * Constant", lambda v, c: v * c),
    "/C": ("( ) / Constant", lambda v, c: v / c if c != 0 else np.full_like(v, np.nan)),
    "^C": ("( ) ^ Constant", lambda v, c: np.power(v, c)),
    "exp": ("exp( )", _safe(np.exp)),
    "ln": ("ln( )", _safe(_ln)),
    "abs": ("abs( )", _safe(np.abs)),
    "sin": ("sin( )", _safe(np.sin)),
    "cos": ("cos( )", _safe(np.cos)),
    # VB used 3.1428 for pi; the port uses math.pi (documented deviation).
    "dtor": ("Degs to Rads", _safe(lambda v: v * math.pi / 180)),
    "rtod": ("Rads to Degs", _safe(lambda v: v * 180 / math.pi)),
    "mbtutokwh": ("MBtu to kWh", _safe(lambda v: v * 1e6 / 3412)),
    "kwhtombtu": ("kWh to MBtu", _safe(lambda v: v * 3412 / 1e6)),
    "()+": ("( )+  (negative values -> 0)", _safe(lambda v: np.where(v < 0, 0.0, v))),
    # present in VB code paths but not in the menu list:
    "sqr": ("( )^2", _safe(lambda v: v * v)),
    "sqrroot": ("sqrt( )", _safe(_sqrt)),
}


def modify_values(ds: Dataset, column: str, modifier: str, constant: float = 1.0) -> Dataset:
    """Apply a MODVALDB modifier in place (returns a new Dataset). Missing values stay missing;
    invalid results (overflow, log of non-positive) become missing."""
    if modifier not in MODIFIERS:
        raise ValueError(f"Unknown modifier {modifier!r}; choose from {list(MODIFIERS)}")
    v = ds.column(column).to_numpy(dtype=float)
    with np.errstate(all="ignore"):
        out = MODIFIERS[modifier][1](v, constant)
    out = np.where(np.isfinite(out), out, np.nan)
    df = ds.data.copy()
    df[column] = out
    return ds.with_data(df)
