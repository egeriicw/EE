"""energy_explorer — Python port of ENERGY EXPLORER.

ENERGY EXPLORER was written in Visual Basic 6 by Kelly Kissock and colleagues at the
Department of Mechanical Engineering, University of Dayton. This package re-implements its
analysis engine (data handling, change-point regression, savings / uncertainty, Lean Energy
Analysis and the simplified two-zone HVAC simulation) as plain, UI-free Python so that it can be
used from notebooks, scripts, and the FastAPI service in ``api/``.

Quick start::

    import energy_explorer as ee

    ds = ee.read_data("data/sample_daily.dat")           # picks up sample_daily.var
    ds = ee.group_by_date(ds, end="12/31/2019", names=("Pre", "Post"))
    model = ee.fit(ds, "3PC", y="CW", x=["Temp"], group=1)
    print(model.summary())
    result = ee.calculate_savings(model, ds, measurement_error_pct=5)
"""

from energy_explorer.dataset import Dataset, TimeInterval
from energy_explorer.grouping import group_by_date, group_by_day_of_week
from energy_explorer.io import (
    read_data,
    read_data_text,
    read_var_names,
    write_data,
    write_var_names,
)
from energy_explorer.modeling import MODEL_TYPES, fit, select_xy
from energy_explorer.models import (
    BaselineModel,
    BinFitModel,
    FivePModel,
    FourPModel,
    LinearModel,
    MeanModel,
    MedianModel,
    ThreePModel,
    model_from_dict,
)
from energy_explorer.savings import (
    LeanEnergyResult,
    SavingsResult,
    autocorrelated_savings_uncertainty,
    calculate_savings,
    create_model_projection,
    lean_energy_analysis,
)
from energy_explorer.simulation import SystemParams, simulate
from energy_explorer.timescale import average_intermittent, change_timescale
from energy_explorer.variables import MODIFIERS, create_variable, delete_variable, modify_values

__version__ = "0.1.0"

__all__ = [
    "MODEL_TYPES",
    "MODIFIERS",
    "BaselineModel",
    "BinFitModel",
    "Dataset",
    "FivePModel",
    "FourPModel",
    "LeanEnergyResult",
    "LinearModel",
    "MeanModel",
    "MedianModel",
    "SavingsResult",
    "SystemParams",
    "ThreePModel",
    "TimeInterval",
    "autocorrelated_savings_uncertainty",
    "average_intermittent",
    "calculate_savings",
    "change_timescale",
    "create_model_projection",
    "create_variable",
    "delete_variable",
    "fit",
    "group_by_date",
    "group_by_day_of_week",
    "lean_energy_analysis",
    "model_from_dict",
    "modify_values",
    "read_data",
    "read_data_text",
    "read_var_names",
    "select_xy",
    "simulate",
    "write_data",
    "write_var_names",
]
