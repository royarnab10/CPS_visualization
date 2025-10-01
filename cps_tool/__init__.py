"""Public API for the CPS tooling package."""
from __future__ import annotations

from .calendar import WorkCalendar, parse_weekend
from .calculator import calculate_schedule
from .csv_loader import load_tasks_from_csv

__all__ = [
    "WorkCalendar",
    "parse_weekend",
    "calculate_schedule",
    "load_tasks_from_csv",
    "convert_mpp_to_csv",
]


def convert_mpp_to_csv(*args, **kwargs):
    from .mpp_converter import convert_mpp_to_csv as _convert

    return _convert(*args, **kwargs)
