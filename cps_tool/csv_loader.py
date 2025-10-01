"""CSV conversion helpers for CPS data."""
from __future__ import annotations

import csv
from datetime import datetime
from pathlib import Path
from typing import List

from .models import DependencySpec, TaskSpec

DATE_FORMATS = [
    "%Y-%m-%dT%H:%M:%S",
    "%Y-%m-%dT%H:%M",
    "%Y-%m-%d",
]


def _parse_datetime(value: str | None) -> datetime | None:
    if not value:
        return None
    for fmt in DATE_FORMATS:
        try:
            return datetime.strptime(value, fmt)
        except ValueError:
            continue
    try:
        return datetime.fromisoformat(value)
    except ValueError as exc:  # pragma: no cover - guard clause
        raise ValueError(f"Invalid datetime value: {value}") from exc


def _parse_bool(value: str | None) -> bool:
    if value is None:
        return False
    return value.strip().lower() in {"1", "true", "yes", "y"}


def _parse_dependencies(value: str) -> List[DependencySpec]:
    dependencies: List[DependencySpec] = []
    if not value:
        return dependencies
    for chunk in value.split(";"):
        chunk = chunk.strip()
        if not chunk:
            continue
        parts = [part.strip() for part in chunk.split(":")]
        if len(parts) == 1:
            predecessor = int(parts[0])
            dependencies.append(DependencySpec(predecessor_uid=predecessor))
            continue
        if len(parts) != 3:
            raise ValueError(
                "Each dependency entry must be formatted as 'UID:TYPE:LAG_DAYS'."
            )
        predecessor, relation, lag = parts
        dependencies.append(
            DependencySpec(
                predecessor_uid=int(predecessor),
                relation_type=relation or "FS",
                lag_days=float(lag or 0.0),
            )
        )
    return dependencies


def load_tasks_from_csv(path: str | Path) -> List[TaskSpec]:
    csv_path = Path(path)
    if not csv_path.exists():
        raise FileNotFoundError(csv_path)
    tasks: List[TaskSpec] = []
    with csv_path.open("r", newline="", encoding="utf-8-sig") as handle:
        reader = csv.DictReader(handle)
        for row in reader:
            uid = int(row.get("uid") or row.get("UID") or row.get("Unique ID"))
            name = row.get("name") or row.get("Name") or row.get("Task Name") or ""
            duration_value = row.get("duration_days") or row.get("DurationDays")
            if duration_value is None:
                duration_value = row.get("Duration") or "0"
            duration_days = float(duration_value)
            task = TaskSpec(
                uid=uid,
                name=name,
                duration_days=duration_days,
                dependencies=_parse_dependencies(row.get("predecessors", "")),
                is_milestone=_parse_bool(row.get("is_milestone", "false")),
                outline_level=int(row.get("outline_level"))
                if row.get("outline_level")
                else None,
                constraint_type=(row.get("constraint_type") or None),
                constraint_date=_parse_datetime(row.get("constraint_date", "")),
                calendar_name=row.get("calendar", None) or row.get("Calendar"),
                original_start=_parse_datetime(row.get("start", "")),
                original_finish=_parse_datetime(row.get("finish", "")),
            )
            if not task.dependencies:
                task.dependencies = _parse_dependencies(row.get("Predecessors", ""))
            tasks.append(task)
    tasks.sort(key=lambda task: task.uid)
    return tasks
