"""Reading/writing ENERGY EXPLORER data files (port of ``OpenDataCmd_Click`` / ``OpenVar``).

File format (from the VB source and the 2005 user guide):

* ASCII, numbers only, no header row; one record per line.
* Delimiter auto-detected from the first non-blank line: comma, else tab, else whitespace
  (runs of spaces count as one separator).
* Variable names come from a sidecar file with the same base name and extension ``.var``
  (or ``.dvn``), one name per line. Without one, names default to V01, V02, ...
* Missing values are written as the no-data flag (default -99).
* Two-digit years are expanded: 50-99 -> 19xx, 0-49 -> 20xx (``OpenVar``).
"""

from __future__ import annotations

import io as _io
import re
from collections.abc import Iterable
from pathlib import Path

import numpy as np
import pandas as pd

from energy_explorer.dataset import Dataset, default_names

PathLike = str | Path


def _detect_delimiter(line: str) -> str | None:
    if "," in line:
        return ","
    if "\t" in line:
        return "\t"
    return None  # whitespace


def _split(line: str, delim: str | None) -> list[str]:
    if delim is None:
        return line.split()
    # VB collapses consecutive delimiters too (it checks the previous character)
    return [tok.strip() for tok in re.split(re.escape(delim) + "+", line.strip()) if tok.strip()]


def _vb_val(token: str) -> float:
    """VB ``Val()``: parse the leading numeric part, 0 if none."""
    m = re.match(r"\s*[-+]?(\d+\.?\d*|\.\d+)([eE][-+]?\d+)?", token)
    return float(m.group(0)) if m else 0.0


def parse_records(text: str) -> np.ndarray:
    lines = [ln.strip() for ln in text.splitlines()]
    lines = [ln for ln in lines if ln]
    if not lines:
        raise ValueError("Data file is empty.")
    delim = _detect_delimiter(lines[0])
    rows = [_split(ln, delim) for ln in lines]
    width = len(rows[0])
    for i, r in enumerate(rows, start=1):
        if len(r) != width:
            raise ValueError(f"Incorrect number of columns in line {i} ({len(r)} vs {width}).")
    return np.array([[_vb_val(t) for t in r] for r in rows], dtype=float)


def read_var_names(path_or_text: PathLike | str, *, is_text: bool = False) -> list[str]:
    text = path_or_text if is_text else Path(path_or_text).read_text(errors="replace")
    names = []
    for ln in text.splitlines():
        ln = ln.strip().strip('"').strip()
        if ln:
            names.append(ln)
    return names


def _expand_years(df: pd.DataFrame) -> pd.DataFrame:
    for col in df.columns:
        if str(col).strip().lower() == "yr":
            yr = df[col]
            df[col] = np.where((yr >= 50) & (yr <= 99), yr + 1900, yr)
            df[col] = np.where((df[col] >= 0) & (df[col] <= 49), df[col] + 2000, df[col])
    return df


def build_dataset(
    values: np.ndarray,
    names: Iterable[str] | None = None,
    *,
    no_data: float = -99.0,
    name: str = "",
) -> Dataset:
    n_cols = values.shape[1]
    names = list(names) if names is not None else default_names(n_cols)
    if len(names) < n_cols:
        raise ValueError(
            f"{n_cols} fields in the data file and {len(names)} variable names in the names file."
        )
    names = names[:n_cols]
    if len(set(names)) != len(names):
        raise ValueError("Variable names must be unique.")
    df = pd.DataFrame(values, columns=names)
    df = df.mask(np.isclose(df.to_numpy(), no_data))
    df = _expand_years(df)
    return Dataset(data=df, no_data=no_data, name=name)


def read_data_text(
    text: str, names: Iterable[str] | str | None = None, *, no_data: float = -99.0, name: str = ""
) -> Dataset:
    """Parse data from a string. ``names`` may be a list or the text of a .var file."""
    if isinstance(names, str):
        names = read_var_names(names, is_text=True)
    return build_dataset(parse_records(text), names, no_data=no_data, name=name)


def read_data(
    path: PathLike,
    names: Iterable[str] | PathLike | None = None,
    *,
    no_data: float = -99.0,
) -> Dataset:
    """Read a data file; auto-load ``<stem>.var`` / ``<stem>.dvn`` if ``names`` is not given."""
    path = Path(path)
    if names is None:
        for ext in (".var", ".VAR", ".dvn", ".DVN"):
            cand = path.with_suffix(ext)
            if cand.exists():
                names = read_var_names(cand)
                break
    elif isinstance(names, (str, Path)):
        names = read_var_names(names)
    text = path.read_text(errors="replace")
    return build_dataset(parse_records(text), names, no_data=no_data, name=path.name)


def to_text(ds: Dataset, *, delimiter: str = " ", float_format: str = "%.6g") -> str:
    buf = _io.StringIO()
    ds.data.fillna(ds.no_data).to_csv(
        buf, sep=delimiter, header=False, index=False, float_format=float_format
    )
    return buf.getvalue()


def write_data(ds: Dataset, path: PathLike, *, write_names: bool = True) -> Path:
    path = Path(path)
    path.write_text(to_text(ds))
    if write_names:
        write_var_names(ds.columns, path.with_suffix(".var"))
    return path


def write_var_names(names: Iterable[str], path: PathLike) -> Path:
    path = Path(path)
    path.write_text("\n".join(names) + "\n")
    return path
