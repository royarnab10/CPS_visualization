"""Core dataclasses used by the CPS tooling package."""
from __future__ import annotations

from dataclasses import dataclass, field
from datetime import datetime
from typing import Iterable, List, Optional


@dataclass(slots=True)
class DependencySpec:
    """Describe a dependency relationship between two tasks."""

    predecessor_uid: int
    relation_type: str = "FS"
    lag_days: float = 0.0

    def as_tuple(self) -> tuple[int, str, float]:
        return self.predecessor_uid, self.relation_type, self.lag_days


@dataclass(slots=True)
class TaskSpec:
    """Normalized task information extracted from the Microsoft Project file."""

    uid: int
    name: str
    duration_days: float
    dependencies: List[DependencySpec] = field(default_factory=list)
    is_milestone: bool = False
    outline_level: Optional[int] = None
    constraint_type: Optional[str] = None
    constraint_date: Optional[datetime] = None
    calendar_name: Optional[str] = None
    original_start: Optional[datetime] = None
    original_finish: Optional[datetime] = None

    def iter_dependency_tuples(self) -> Iterable[tuple[int, str, float]]:
        for dependency in self.dependencies:
            yield dependency.as_tuple()


@dataclass(slots=True)
class ScheduledTask:
    """Computed schedule attributes for a single task."""

    spec: TaskSpec
    earliest_start: datetime
    earliest_finish: datetime
    latest_start: datetime
    latest_finish: datetime
    total_float_hours: float

    @property
    def is_critical(self) -> bool:
        return abs(self.total_float_hours) < 1e-4


@dataclass(slots=True)
class ScheduleResult:
    """Container for the calculated CPS schedule."""

    project_start: datetime
    project_finish: datetime
    tasks: List[ScheduledTask]

    def critical_path(self) -> List[ScheduledTask]:
        return [task for task in self.tasks if task.is_critical]

    def to_rows(self) -> List[dict[str, str]]:
        rows: List[dict[str, str]] = []
        for task in self.tasks:
            rows.append(
                {
                    "uid": str(task.spec.uid),
                    "name": task.spec.name,
                    "earliest_start": task.earliest_start.isoformat(),
                    "earliest_finish": task.earliest_finish.isoformat(),
                    "latest_start": task.latest_start.isoformat(),
                    "latest_finish": task.latest_finish.isoformat(),
                    "total_float_hours": f"{task.total_float_hours:.3f}",
                    "is_critical": "yes" if task.is_critical else "no",
                    "duration_days": f"{task.spec.duration_days:.3f}",
                    "is_milestone": "yes" if task.spec.is_milestone else "no",
                    "constraint_type": task.spec.constraint_type or "",
                    "constraint_date": task.spec.constraint_date.isoformat()
                    if task.spec.constraint_date
                    else "",
                }
            )
        return rows
