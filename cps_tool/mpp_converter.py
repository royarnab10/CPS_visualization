"""Utilities to convert Microsoft Project schedules into CSV."""
from __future__ import annotations

import csv
from datetime import datetime
from functools import lru_cache
from pathlib import Path
from typing import List, Tuple

from .models import DependencySpec, TaskSpec


_IMPORT_HELP = (
    "mpxj is required to convert Microsoft Project schedules. "
    "Install the package with its optional Java dependencies as documented in the README."
)


@lru_cache(maxsize=1)
def _load_mpxj() -> Tuple[object, object]:
    """Lazily import the mpxj modules that are required for conversion."""

    try:  # pragma: no cover - import resolution depends on mpxj version
        from mpxj.enums import TimeUnit  # type: ignore[attr-defined]
    except ModuleNotFoundError as exc:  # pragma: no cover - missing dependency
        try:
            from mpxj import TimeUnit  # type: ignore[no-redef]
        except ModuleNotFoundError as inner_exc:  # pragma: no cover - missing dependency
            raise ImportError(_IMPORT_HELP) from inner_exc
    except ImportError:  # pragma: no cover - fallback for older releases
        try:
            from mpxj import TimeUnit  # type: ignore[no-redef]
        except ModuleNotFoundError as exc:  # pragma: no cover - missing dependency
            raise ImportError(_IMPORT_HELP) from exc

    try:
        from mpxj.reader import UniversalProjectReader
    except ModuleNotFoundError as exc:  # pragma: no cover - missing dependency
        raise ImportError(_IMPORT_HELP) from exc

    return TimeUnit, UniversalProjectReader


def _duration_to_days(duration) -> float:
    TimeUnit, _ = _load_mpxj()
    if duration is None:
        return 0.0
    converted = duration.convert(TimeUnit.DAYS)
    return float(converted.duration)


def _extract_dependencies(task) -> List[DependencySpec]:
    dependencies: List[DependencySpec] = []
    for relation in task.predecessors or []:
        predecessor = relation.source_task or relation.target_task
        if predecessor is None:
            continue
        lag = relation.lag
        dependencies.append(
            DependencySpec(
                predecessor_uid=int(predecessor.unique_id),
                relation_type=str(relation.type.name if relation.type else "FS"),
                lag_days=_duration_to_days(lag),
            )
        )
    return dependencies


def _safe_datetime(value) -> datetime | None:
    return value if isinstance(value, datetime) else None


def extract_tasks_from_mpp(path: str | Path, include_summary: bool = False) -> List[TaskSpec]:
    _, UniversalProjectReader = _load_mpxj()
    project = UniversalProjectReader().read(str(path))
    tasks: List[TaskSpec] = []
    for task in project.tasks:
        if task is None:
            continue
        if not include_summary and task.summary:
            continue
        duration_days = _duration_to_days(task.duration)
        tasks.append(
            TaskSpec(
                uid=int(task.unique_id),
                name=str(task.name or ""),
                duration_days=duration_days,
                dependencies=_extract_dependencies(task),
                is_milestone=bool(task.milestone),
                outline_level=int(task.outline_level) if task.outline_level is not None else None,
                constraint_type=str(task.constraint_type.name)
                if getattr(task, "constraint_type", None)
                else None,
                constraint_date=_safe_datetime(getattr(task, "constraint_date", None)),
                calendar_name=getattr(task.calendar, "name", None),
                original_start=_safe_datetime(task.start),
                original_finish=_safe_datetime(task.finish),
            )
        )
    tasks.sort(key=lambda item: item.uid)
    return tasks


def convert_mpp_to_csv(path: str | Path, output: str | Path, include_summary: bool = False) -> Path:
    tasks = extract_tasks_from_mpp(path, include_summary=include_summary)
    csv_path = Path(output)
    csv_path.parent.mkdir(parents=True, exist_ok=True)
    headers = [
        "uid",
        "name",
        "duration_days",
        "is_milestone",
        "outline_level",
        "constraint_type",
        "constraint_date",
        "calendar",
        "predecessors",
        "start",
        "finish",
    ]
    with csv_path.open("w", newline="", encoding="utf-8") as handle:
        writer = csv.DictWriter(handle, fieldnames=headers)
        writer.writeheader()
        for task in tasks:
            predecessor_value = ";".join(
                f"{dep.predecessor_uid}:{dep.relation_type}:{dep.lag_days}" for dep in task.dependencies
            )
            writer.writerow(
                {
                    "uid": task.uid,
                    "name": task.name,
                    "duration_days": f"{task.duration_days:.3f}",
                    "is_milestone": "yes" if task.is_milestone else "no",
                    "outline_level": task.outline_level if task.outline_level is not None else "",
                    "constraint_type": task.constraint_type or "",
                    "constraint_date": task.constraint_date.isoformat()
                    if task.constraint_date
                    else "",
                    "calendar": task.calendar_name or "",
                    "predecessors": predecessor_value,
                    "start": task.original_start.isoformat() if task.original_start else "",
                    "finish": task.original_finish.isoformat() if task.original_finish else "",
                }
            )
    return csv_path
