"""Storage seam for the API.

The API never touches globals: it keeps ``Dataset`` objects and fitted models behind the
``Store`` protocol. ``InMemoryStore`` is fine for a single-process dev server; swap in a
database/object-store backed implementation (datasets as Parquet, models as the JSON from
``BaselineModel.to_dict()``) without changing the routers.
"""

from __future__ import annotations

import threading
import uuid
from typing import Protocol

from energy_explorer import BaselineModel, Dataset


class Store(Protocol):
    def put_dataset(self, ds: Dataset, parent: str | None = None) -> str: ...
    def get_dataset(self, ds_id: str) -> Dataset: ...
    def replace_dataset(self, ds_id: str, ds: Dataset) -> None: ...
    def list_datasets(self) -> dict[str, Dataset]: ...
    def put_model(self, model: BaselineModel, dataset_id: str) -> str: ...
    def get_model(self, model_id: str) -> tuple[BaselineModel, str]: ...


class NotFound(KeyError):
    pass


class InMemoryStore:
    def __init__(self) -> None:
        self._lock = threading.Lock()
        self._datasets: dict[str, Dataset] = {}
        self._models: dict[str, tuple[BaselineModel, str]] = {}

    def put_dataset(self, ds: Dataset, parent: str | None = None) -> str:
        ds_id = uuid.uuid4().hex[:12]
        with self._lock:
            self._datasets[ds_id] = ds
        return ds_id

    def get_dataset(self, ds_id: str) -> Dataset:
        try:
            return self._datasets[ds_id]
        except KeyError:
            raise NotFound(f"dataset {ds_id}") from None

    def replace_dataset(self, ds_id: str, ds: Dataset) -> None:
        self.get_dataset(ds_id)
        with self._lock:
            self._datasets[ds_id] = ds

    def list_datasets(self) -> dict[str, Dataset]:
        return dict(self._datasets)

    def put_model(self, model: BaselineModel, dataset_id: str) -> str:
        model_id = uuid.uuid4().hex[:12]
        with self._lock:
            self._models[model_id] = (model, dataset_id)
        return model_id

    def get_model(self, model_id: str) -> tuple[BaselineModel, str]:
        try:
            return self._models[model_id]
        except KeyError:
            raise NotFound(f"model {model_id}") from None


_store: Store = InMemoryStore()


def get_store() -> Store:
    """FastAPI dependency; override in tests or deployments via ``app.dependency_overrides``."""
    return _store
