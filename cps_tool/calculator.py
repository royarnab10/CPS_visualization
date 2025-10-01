"""Critical Path Schedule calculator based on CSV task definitions."""
from __future__ import annotations

import csv
import heapq
from collections import defaultdict
from dataclasses import replace
from datetime import datetime
from pathlib import Path
from typing import Dict, Iterable, Iterator, List, Mapping, MutableMapping, Sequence, Tuple

from .calendar import WorkCalendar
from .models import (
    CycleResolution,
    DependencyIssue,
    DependencySpec,
    ScheduleResult,
    ScheduledTask,
    TaskSpec,
)


def _ensure_task_lookup(tasks: Sequence[TaskSpec]) -> Dict[int, TaskSpec]:
    lookup = {}
    for task in tasks:
        if task.uid in lookup:
            raise ValueError(f"Duplicate task UID detected: {task.uid}")
        lookup[task.uid] = task
    if not lookup:
        raise ValueError("No tasks were provided for CPS calculation.")
    return lookup


def _remove_missing_dependencies(
    tasks: Sequence[TaskSpec], lookup: Mapping[int, TaskSpec]
) -> List[DependencyIssue]:
    issues: List[DependencyIssue] = []
    for task in tasks:
        cleaned: List[DependencySpec] = []
        for dependency in task.dependencies:
            if dependency.predecessor_uid in lookup:
                cleaned.append(dependency)
                continue
            issues.append(
                DependencyIssue(
                    task_uid=task.uid,
                    task_name=task.name,
                    dependency=DependencySpec(
                        predecessor_uid=dependency.predecessor_uid,
                        relation_type=dependency.relation_type,
                        lag_days=dependency.lag_days,
                    ),
                    reason="Missing predecessor task",
                )
            )
        task.dependencies = cleaned
    return issues


def _build_successors(tasks: Sequence[TaskSpec]) -> Dict[int, List[tuple[int, DependencySpec]]]:
    successors: Dict[int, List[tuple[int, DependencySpec]]] = defaultdict(list)
    for task in tasks:
        for dependency in task.dependencies:
            successors[dependency.predecessor_uid].append((task.uid, dependency))
    return successors


def _clone_tasks(tasks: Iterable[TaskSpec]) -> List[TaskSpec]:
    cloned: List[TaskSpec] = []
    for task in tasks:
        cloned.append(
            replace(
                task,
                dependencies=[
                    DependencySpec(
                        predecessor_uid=dependency.predecessor_uid,
                        relation_type=dependency.relation_type,
                        lag_days=dependency.lag_days,
                    )
                    for dependency in task.dependencies
                ],
            )
        )
    return cloned


def _format_dependency(dependency: DependencySpec) -> str:
    relation = (dependency.relation_type or "FS").upper()
    lag = f"{dependency.lag_days:g}"
    return f"{dependency.predecessor_uid}:{relation}:{lag}"


def _tasks_to_rows(tasks: Sequence[TaskSpec]) -> List[dict[str, str]]:
    rows: List[dict[str, str]] = []
    for task in tasks:
        predecessors = ";".join(
            _format_dependency(dependency) for dependency in task.dependencies
        )
        rows.append(
            {
                "uid": str(task.uid),
                "name": task.name,
                "duration_days": f"{task.duration_days:.3f}",
                "is_milestone": "yes" if task.is_milestone else "no",
                "outline_level": str(task.outline_level) if task.outline_level else "",
                "constraint_type": task.constraint_type or "",
                "constraint_date": task.constraint_date.isoformat()
                if task.constraint_date
                else "",
                "calendar": task.calendar_name or "",
                "predecessors": predecessors,
                "start": task.original_start.isoformat()
                if task.original_start
                else "",
                "finish": task.original_finish.isoformat()
                if task.original_finish
                else "",
            }
        )
    return rows


def _write_cycle_adjusted_tasks(tasks: Sequence[TaskSpec]) -> str:
    output_dir = Path("outputs")
    output_dir.mkdir(parents=True, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_path = output_dir / f"cycle_resolved_cps_{timestamp}.csv"
    rows = _tasks_to_rows(tasks)
    with output_path.open("w", newline="", encoding="utf-8") as handle:
        if not rows:
            return str(output_path)
        writer = csv.DictWriter(handle, fieldnames=list(rows[0].keys()))
        writer.writeheader()
        writer.writerows(rows)
    return str(output_path)


def _find_cycle(tasks: Sequence[TaskSpec]) -> Tuple[List[int], int, DependencySpec] | None:
    successors = _build_successors(tasks)
    visited: set[int] = set()
    on_stack: set[int] = set()

    for task in tasks:
        if task.uid in visited:
            continue

        traversal_stack: List[tuple[int, Iterator[tuple[int, DependencySpec]]]] = []
        path: List[int] = []

        visited.add(task.uid)
        on_stack.add(task.uid)
        traversal_stack.append((task.uid, iter(successors.get(task.uid, []))))
        path.append(task.uid)

        while traversal_stack:
            current_uid, iterator = traversal_stack[-1]
            try:
                successor_uid, dependency = next(iterator)
            except StopIteration:
                traversal_stack.pop()
                on_stack.remove(current_uid)
                path.pop()
                continue

            if successor_uid not in visited:
                visited.add(successor_uid)
                on_stack.add(successor_uid)
                traversal_stack.append(
                    (successor_uid, iter(successors.get(successor_uid, [])))
                )
                path.append(successor_uid)
                continue

            if successor_uid in on_stack:
                try:
                    cycle_start_index = path.index(successor_uid)
                except ValueError:  # pragma: no cover - defensive guard
                    cycle_start_index = 0
                cycle = path[cycle_start_index:] + [successor_uid]
                return cycle, successor_uid, dependency

    return None


def _resolve_cycles(tasks: List[TaskSpec]) -> tuple[List[CycleResolution], str | None]:
    if not tasks:
        return [], None

    lookup = {task.uid: task for task in tasks}
    resolutions: List[CycleResolution] = []

    while True:
        cycle_info = _find_cycle(tasks)
        if not cycle_info:
            break
        cycle_uids, successor_uid, dependency = cycle_info
        successor_task = lookup.get(successor_uid)
        if successor_task is None:
            break
        removed_dependency = None
        relation_upper = (dependency.relation_type or "FS").upper()
        new_dependencies: List[DependencySpec] = []
        for existing in successor_task.dependencies:
            if (
                removed_dependency is None
                and existing.predecessor_uid == dependency.predecessor_uid
                and (existing.relation_type or "FS").upper() == relation_upper
                and abs(existing.lag_days - dependency.lag_days) < 1e-9
            ):
                removed_dependency = DependencySpec(
                    predecessor_uid=existing.predecessor_uid,
                    relation_type=existing.relation_type,
                    lag_days=existing.lag_days,
                )
                continue
            new_dependencies.append(existing)
        if removed_dependency is None:  # pragma: no cover - defensive guard
            break
        successor_task.dependencies = new_dependencies
        cycle_names = [
            lookup[uid].name if uid in lookup else f"Task {uid}"
            for uid in cycle_uids
        ]
        resolutions.append(
            CycleResolution(
                cycle_task_uids=cycle_uids,
                cycle_task_names=cycle_names,
                removed_from_task_uid=successor_task.uid,
                removed_from_task_name=successor_task.name,
                removed_dependency=removed_dependency,
            )
        )

    csv_path = None
    if resolutions:
        csv_path = _write_cycle_adjusted_tasks(tasks)
    return resolutions, csv_path


def _topological_order(tasks: Sequence[TaskSpec]) -> List[int]:
    indegree: Dict[int, int] = {task.uid: 0 for task in tasks}
    for task in tasks:
        for dependency in task.dependencies:
            indegree[task.uid] = indegree.get(task.uid, 0) + 1
    ready = [uid for uid, degree in indegree.items() if degree == 0]
    heapq.heapify(ready)
    order: List[int] = []
    successors = _build_successors(tasks)
    while ready:
        uid = heapq.heappop(ready)
        order.append(uid)
        for successor_uid, _ in successors.get(uid, []):
            indegree[successor_uid] -= 1
            if indegree[successor_uid] == 0:
                heapq.heappush(ready, successor_uid)
    if len(order) != len(tasks):  # pragma: no cover - guard clause
        raise ValueError("Unable to establish a valid task ordering (cycle detected).")
    return order


def _infer_project_start(tasks: Sequence[TaskSpec]) -> datetime | None:
    candidates: List[datetime] = []
    for task in tasks:
        if task.constraint_date:
            candidates.append(task.constraint_date)
        if task.original_start:
            candidates.append(task.original_start)
    if not candidates:
        return None
    return min(candidates)


def _apply_constraints(
    spec: TaskSpec,
    start: datetime,
    finish: datetime,
    calendar: WorkCalendar,
) -> tuple[datetime, datetime]:
    constraint_type = (spec.constraint_type or "").upper()
    constraint_date = spec.constraint_date
    if constraint_type in {"MUST_START_ON", "MSO"} and constraint_date:
        start = calendar.align_start(constraint_date)
        finish = calendar.add_work_duration(start, spec.duration_days)
    elif constraint_type in {"START_NO_EARLIER_THAN", "SNET"} and constraint_date:
        start = max(start, calendar.align_start(constraint_date))
        finish = calendar.add_work_duration(start, spec.duration_days)
    elif constraint_type in {"MUST_FINISH_ON", "MFO"} and constraint_date:
        finish = calendar.align_finish(constraint_date)
        start = calendar.subtract_work_duration(finish, spec.duration_days)
    elif constraint_type in {"FINISH_NO_EARLIER_THAN", "FNET"} and constraint_date:
        finish_candidate = calendar.align_finish(constraint_date)
        start = max(start, calendar.subtract_work_duration(finish_candidate, spec.duration_days))
        finish = calendar.add_work_duration(start, spec.duration_days)
    return start, finish


def _forward_pass(
    order: Sequence[int],
    tasks: Mapping[int, TaskSpec],
    calendar: WorkCalendar,
    project_start: datetime,
) -> Dict[int, ScheduledTask]:
    scheduled: Dict[int, ScheduledTask] = {}
    for uid in order:
        spec = tasks[uid]
        start = calendar.align_start(project_start)
        finish = calendar.add_work_duration(start, spec.duration_days)
        for dependency in spec.dependencies:
            predecessor_task = scheduled.get(dependency.predecessor_uid)
            if predecessor_task is None:
                raise ValueError(
                    f"Task {spec.uid} references missing predecessor {dependency.predecessor_uid}."
                )
            relation = (dependency.relation_type or "FS").upper()
            if relation == "FS":
                candidate = calendar.add_work_duration(
                    predecessor_task.earliest_finish, dependency.lag_days
                )
                start = max(start, calendar.align_start(candidate))
                finish = calendar.add_work_duration(start, spec.duration_days)
            elif relation == "SS":
                candidate = calendar.add_work_duration(
                    predecessor_task.earliest_start, dependency.lag_days
                )
                start = max(start, calendar.align_start(candidate))
                finish = calendar.add_work_duration(start, spec.duration_days)
            elif relation == "FF":
                candidate_finish = calendar.add_work_duration(
                    predecessor_task.earliest_finish, dependency.lag_days
                )
                candidate_finish = calendar.align_finish(candidate_finish)
                start = calendar.subtract_work_duration(candidate_finish, spec.duration_days)
                start = calendar.align_start(start)
                finish = calendar.add_work_duration(start, spec.duration_days)
            elif relation == "SF":
                candidate_finish = calendar.add_work_duration(
                    predecessor_task.earliest_start, dependency.lag_days
                )
                candidate_finish = calendar.align_finish(candidate_finish)
                start = calendar.subtract_work_duration(candidate_finish, spec.duration_days)
                start = calendar.align_start(start)
                finish = calendar.add_work_duration(start, spec.duration_days)
            else:
                raise ValueError(f"Unsupported dependency type: {dependency.relation_type}")
        start, finish = _apply_constraints(spec, start, finish, calendar)
        scheduled[uid] = ScheduledTask(
            spec=spec,
            earliest_start=start,
            earliest_finish=finish,
            latest_start=start,
            latest_finish=finish,
            total_float_hours=0.0,
        )
    return scheduled


def _backward_pass(
    order: Sequence[int],
    scheduled: MutableMapping[int, ScheduledTask],
    successors: Mapping[int, List[tuple[int, DependencySpec]]],
    calendar: WorkCalendar,
) -> None:
    project_finish = max(task.earliest_finish for task in scheduled.values())
    for uid in reversed(order):
        task = scheduled[uid]
        successor_records = successors.get(uid, [])
        if successor_records:
            candidate_finishes: List[datetime] = []
            for successor_uid, dependency in successor_records:
                successor_task = scheduled[successor_uid]
                relation = (dependency.relation_type or "FS").upper()
                if relation == "FS":
                    candidate_finish = calendar.subtract_work_duration(
                        successor_task.latest_start, dependency.lag_days
                    )
                elif relation == "SS":
                    candidate_start = calendar.subtract_work_duration(
                        successor_task.latest_start, dependency.lag_days
                    )
                    candidate_finish = calendar.add_work_duration(
                        candidate_start, task.spec.duration_days
                    )
                elif relation == "FF":
                    candidate_finish = calendar.subtract_work_duration(
                        successor_task.latest_finish, dependency.lag_days
                    )
                elif relation == "SF":
                    candidate_start = calendar.subtract_work_duration(
                        successor_task.latest_finish, dependency.lag_days
                    )
                    candidate_finish = calendar.add_work_duration(
                        candidate_start, task.spec.duration_days
                    )
                else:
                    candidate_finish = calendar.subtract_work_duration(
                        successor_task.latest_start, dependency.lag_days
                    )
                candidate_finish = calendar.align_finish(candidate_finish)
                candidate_finishes.append(candidate_finish)
            latest_finish_limit = min(candidate_finishes)
        else:
            latest_finish_limit = project_finish
        latest_start = calendar.subtract_work_duration(
            latest_finish_limit, task.spec.duration_days
        )
        latest_start = calendar.align_start(latest_start)
        latest_finish = calendar.add_work_duration(latest_start, task.spec.duration_days)
        latest_finish = min(latest_finish, latest_finish_limit)
        task.latest_start = latest_start
        task.latest_finish = latest_finish
        task.total_float_hours = calendar.work_hours_between(
            task.earliest_start, task.latest_start
        )


def calculate_schedule(
    task_specs: Iterable[TaskSpec],
    project_start: datetime | None = None,
    calendar: WorkCalendar | None = None,
) -> ScheduleResult:
    """Compute earliest and latest dates for the provided tasks."""

    tasks = _clone_tasks(list(task_specs))
    lookup = _ensure_task_lookup(tasks)
    dependency_issues = _remove_missing_dependencies(tasks, lookup)
    cycle_resolutions, cycle_csv_path = _resolve_cycles(tasks)
    successors = _build_successors(tasks)
    order = _topological_order(tasks)
    if calendar is None:
        calendar = WorkCalendar()
    if project_start is None:
        project_start = _infer_project_start(tasks) or datetime.now()
    project_start = calendar.align_start(project_start)
    scheduled = _forward_pass(order, lookup, calendar, project_start)
    _backward_pass(order, scheduled, successors, calendar)
    project_finish = max(task.latest_finish for task in scheduled.values())
    ordered_tasks = [scheduled[uid] for uid in order]
    return ScheduleResult(
        project_start=project_start,
        project_finish=project_finish,
        tasks=ordered_tasks,
        cycle_resolutions=cycle_resolutions,
        cycle_adjusted_csv=cycle_csv_path,
        dependency_issues=dependency_issues,
    )
