"""Pydantic request/response models (API contract shared with the React client)."""

from __future__ import annotations

from typing import Any, Literal

from pydantic import BaseModel, Field


class DatasetSummary(BaseModel):
    id: str
    name: str
    n_records: int
    columns: list[str]
    time_fields: dict[str, str]
    interval: str
    grouped: bool
    group_names: list[str]
    group_counts: dict[str, int] | None
    no_data: float


class GroupRequest(BaseModel):
    mode: Literal["date", "dow"] = "date"
    start: str | None = Field(None, description="e.g. '1/1/2019' (daily), '1/2019' (monthly)")
    end: str | None = None
    names: tuple[str, str] = ("Pre", "Post")


class TimescaleRequest(BaseModel):
    to: Literal["daily", "monthly", "yearly"]
    averages: list[str] = Field(default_factory=list, description="variables to average; others are totalled")
    strict_no_data: bool = False


class CreateVariableRequest(BaseModel):
    name: str
    a: str
    op: Literal["+", "-", "*", "/"] | None = None
    b: str | None = None


class ModifyRequest(BaseModel):
    column: str
    modifier: str
    constant: float = 1.0


class FitRequest(BaseModel):
    model_type: str = Field(..., examples=["3PC"])
    y: str
    x: list[str] = Field(default_factory=list)
    group: int | str | None = None
    options: dict[str, Any] = Field(default_factory=dict)


class FitResponse(BaseModel):
    id: str
    dataset_id: str
    model: dict[str, Any]
    summary: str
    points: dict[str, list[float | None]]
    curve: dict[str, list[float | None]] | None


class SavingsRequest(BaseModel):
    dataset_id: str | None = None  # defaults to the dataset the model was fit on
    post_group: int | str = 2
    measurement_error_pct: float = 0.0
    autocorrelation_corrected: bool = False


class SimulationRequest(BaseModel):
    params: dict[str, Any] = Field(default_factory=dict)
    wbe: str = "WBE"
    toa: str = "Toa"
    woa: str | None = None
    hsol: str | None = None
