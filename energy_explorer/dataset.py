"""The in-memory data set (replaces the VB globals ``d()``, ``vn$()``, ``grpd()``, ``timeint``...).

Design notes
------------
* Values are stored as ``float64`` with ``NaN`` for missing data. The VB program used a sentinel
  (``ndflag``, default -99) everywhere; here the sentinel is converted to/from NaN only at the I/O
  boundary (``energy_explorer.io``). ``Dataset.no_data`` remembers the flag for round-tripping.
* Timestamp columns are recognised by name exactly like ``OpenVar`` in SUBS.BAS: ``yr``, ``mo``,
  ``dy``, ``hr``, ``mn`` (case-insensitive). The time interval is implied by how many are present.
* Groups (``grpd()`` in VB) are an integer array: 1 = group 1 (e.g. "Pre"), 2 = group 2 ("Post"),
  0 = ungrouped. ``group_names`` holds (g1name$, g2name$).
* ``Dataset`` is immutable-by-convention: transform functions return new instances, which keeps the
  library safe to use behind a web API.
"""

from __future__ import annotations

from dataclasses import dataclass, replace
from enum import IntEnum

import numpy as np
import pandas as pd

TIME_FIELDS = ("yr", "mo", "dy", "hr", "mn")


class TimeInterval(IntEnum):
    """Mirror of VB ``timeint`` (count of yr/mo/dy/hr fields present)."""

    NONE = 0
    YEARLY = 1
    MONTHLY = 2
    DAILY = 3
    HOURLY = 4


@dataclass(frozen=True)
class Dataset:
    data: pd.DataFrame
    no_data: float = -99.0
    groups: np.ndarray | None = None
    group_names: tuple[str, str] = ("G1", "G2")
    name: str = ""

    def __post_init__(self) -> None:
        if self.groups is not None and len(self.groups) != len(self.data):
            raise ValueError("groups must have one entry per record")

    # ------------------------------------------------------------------ columns / time fields
    @property
    def columns(self) -> list[str]:
        return list(self.data.columns)

    @property
    def n_records(self) -> int:
        return len(self.data)

    @property
    def time_fields(self) -> dict[str, str]:
        """Map canonical time-field name ('yr', 'mo', ...) to the actual column name."""
        found: dict[str, str] = {}
        for col in self.data.columns:
            key = str(col).strip().lower()
            if key in TIME_FIELDS and key not in found:
                found[key] = col
        return found

    @property
    def interval(self) -> TimeInterval:
        tf = self.time_fields
        return TimeInterval(sum(1 for k in ("yr", "mo", "dy", "hr") if k in tf))

    @property
    def value_columns(self) -> list[str]:
        tf = set(self.time_fields.values())
        return [c for c in self.data.columns if c not in tf]

    def column(self, name: str) -> pd.Series:
        try:
            return self.data[name]
        except KeyError:
            raise KeyError(f"Unknown variable {name!r}. Available: {self.columns}") from None

    # ------------------------------------------------------------------ dates
    def serial_dates(self) -> np.ndarray:
        """VB ``SerDate``: yyyymmddhh as a number (missing parts are 0)."""
        tf = self.time_fields
        if "yr" not in tf:
            raise ValueError("Data set has no 'Yr' column; date operations are unavailable.")

        def part(key: str) -> np.ndarray:
            if key in tf and self.interval >= ("yr", "mo", "dy", "hr").index(key) + 1:
                return self.data[tf[key]].fillna(0).to_numpy(dtype=float)
            return np.zeros(len(self.data))

        return part("yr") * 1e6 + part("mo") * 1e4 + part("dy") * 1e2 + part("hr")

    def timestamps(self) -> pd.Series:
        """Best-effort pandas timestamps for plotting (hour 24 rolls to next day)."""
        tf = self.time_fields
        if "yr" not in tf:
            return pd.Series(pd.RangeIndex(len(self.data)), index=self.data.index)
        yr = self.data[tf["yr"]]
        mo = self.data[tf["mo"]] if "mo" in tf else 1
        dy = self.data[tf["dy"]] if "dy" in tf else 1
        base = pd.to_datetime(
            pd.DataFrame({"year": yr, "month": mo, "day": dy}), errors="coerce"
        )
        if "hr" in tf:
            base = base + pd.to_timedelta(self.data[tf["hr"]].fillna(0), unit="h")
        if "mn" in tf:
            base = base + pd.to_timedelta(self.data[tf["mn"]].fillna(0), unit="m")
        return base

    # ------------------------------------------------------------------ groups
    @property
    def is_grouped(self) -> bool:
        return self.groups is not None

    def group_mask(self, group: int | None) -> np.ndarray:
        if group is None:
            return np.ones(len(self.data), dtype=bool)
        if self.groups is None:
            raise ValueError("No grouping has been defined.")
        return self.groups == group

    def group_number(self, group: int | str | None) -> int | None:
        """Accept 1/2 or a group name ('Pre'/'Post')."""
        if group is None or isinstance(group, (int, np.integer)):
            return None if group is None else int(group)
        lowered = [g.lower() for g in self.group_names]
        if group.lower() in lowered:
            return lowered.index(group.lower()) + 1
        raise ValueError(f"Unknown group {group!r}; groups are {self.group_names}")

    # ------------------------------------------------------------------ helpers
    def with_data(self, data: pd.DataFrame, **kw) -> Dataset:
        groups = kw.pop("groups", self.groups if len(data) == len(self.data) else None)
        return replace(self, data=data.reset_index(drop=True), groups=groups, **kw)

    def with_groups(self, groups: np.ndarray, names: tuple[str, str]) -> Dataset:
        return replace(self, groups=np.asarray(groups, dtype=int), group_names=tuple(names))

    def summary(self) -> dict:
        return {
            "name": self.name,
            "n_records": self.n_records,
            "columns": self.columns,
            "time_fields": self.time_fields,
            "interval": self.interval.name.lower(),
            "grouped": self.is_grouped,
            "group_names": list(self.group_names),
            "group_counts": (
                {
                    self.group_names[0]: int((self.groups == 1).sum()),
                    self.group_names[1]: int((self.groups == 2).sum()),
                }
                if self.groups is not None
                else None
            ),
            "no_data": self.no_data,
        }

    def __repr__(self) -> str:  # pragma: no cover - cosmetic
        return (
            f"Dataset(name={self.name!r}, records={self.n_records}, "
            f"interval={self.interval.name.lower()}, columns={self.columns})"
        )


def default_names(n: int) -> list[str]:
    """VB default variable names V01, V02, ..."""
    return [f"V{i:02d}" for i in range(1, n + 1)]


__all__ = ["TIME_FIELDS", "Dataset", "TimeInterval", "default_names"]
