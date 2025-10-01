"""Critical Path Schedule calculator based on CSV task definitions."""
from __future__ import annotations

import heapq
from collections import defaultdict
from datetime import datetime
from typing import Dict, Iterable, List, Mapping, MutableMapping, Sequence

from .calendar import WorkCalendar
from .models import DependencySpec, ScheduleResult, ScheduledTask, TaskSpec


def _ensure_task_lookup(tasks: Sequence[TaskSpec]) -> Dict[int, TaskSpec]:
    lookup = {}
    for task in tasks:
        if task.uid in lookup:
            raise ValueError(f"Duplicate task UID detected: {task.uid}")
        lookup[task.uid] = task
    if not lookup:
        raise ValueError("No tasks were provided for CPS calculation.")
    return lookup


def _build_successors(tasks: Sequence[TaskSpec]) -> Dict[int, List[tuple[int, DependencySpec]]]:
    successors: Dict[int, List[tuple[int, DependencySpec]]] = defaultdict(list)
    for task in tasks:
        for dependency in task.dependencies:
            successors[dependency.predecessor_uid].append((task.uid, dependency))
    return successors


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
            relation = dependency.relation_type.upper() or "FS"
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
                relation = dependency.relation_type.upper() or "FS"
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

    tasks = list(task_specs)
    lookup = _ensure_task_lookup(tasks)
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
    )
