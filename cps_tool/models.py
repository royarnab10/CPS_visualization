"""Core dataclasses used by the CPS tooling package."""
from __future__ import annotations

import csv
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
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
class CycleResolution:
    """Describe how a dependency cycle was resolved."""

    cycle_task_uids: List[int]
    cycle_task_names: List[str]
    removed_from_task_uid: int
    removed_from_task_name: str
    removed_dependency: DependencySpec

    def formatted_cycle(self) -> str:
        """Return a user-friendly representation of the cycle."""

        segments = [
            f"{uid} ({name})"
            for uid, name in zip(self.cycle_task_uids, self.cycle_task_names)
        ]
        return " -> ".join(segments)


@dataclass(slots=True)
class DependencyIssue:
    """Describe a dependency that had to be removed due to invalid data."""

    task_uid: int
    task_name: str
    dependency: DependencySpec
    reason: str

    def formatted_issue(self) -> str:
        relation = (self.dependency.relation_type or "FS").upper()
        lag = f"{self.dependency.lag_days:g}"
        return (
            f"Task {self.task_uid} ({self.task_name}) <- "
            f"{self.dependency.predecessor_uid} [{relation} lag {lag} days]: {self.reason}"
        )


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
    cycle_resolutions: List[CycleResolution] = field(default_factory=list)
    dependency_issues: List[DependencyIssue] = field(default_factory=list)
    cycle_adjusted_csv: Optional[str] = None

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

    def to_csv(self, destination: Path | str) -> Path:
        """Write the calculated schedule to ``destination`` as CSV."""

        rows = self.to_rows()
        fieldnames = list(rows[0].keys()) if rows else [
            "uid",
            "name",
            "earliest_start",
            "earliest_finish",
            "latest_start",
            "latest_finish",
            "total_float_hours",
            "is_critical",
            "duration_days",
            "is_milestone",
            "constraint_type",
            "constraint_date",
        ]

        destination_path = Path(destination)
        destination_path.parent.mkdir(parents=True, exist_ok=True)

        with destination_path.open("w", newline="", encoding="utf-8") as handle:
            writer = csv.DictWriter(handle, fieldnames=fieldnames)
            writer.writeheader()
            if rows:
                writer.writerows(rows)

        return destination_path
