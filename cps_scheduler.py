"""Compute CPM schedule with advanced dependency handling."""
from __future__ import annotations

import argparse
import csv
import datetime as dt
import re
from dataclasses import dataclass
from typing import Dict, Iterable, List, Optional, Sequence, Tuple
from xml.etree import ElementTree as ET
from zipfile import ZipFile

# Calendar constants
WORK_HOURS_PER_DAY = 8.0
WORK_DAYS_PER_WEEK = 5.0
WORK_DAYS_PER_MONTH = 20.0

SUPPORTED_DEP_TYPES = {"FS", "SS", "FF", "SF"}


@dataclass
class Dependency:
    pred_id: str
    dep_type: str
    lag_hours: float

    def normalized(self) -> str:
        if abs(self.lag_hours) < 1e-9:
            lag_text = ""
        else:
            sign = "+" if self.lag_hours > 0 else "-"
            lag_text = f"{sign}{abs(self.lag_hours):.3f}h"
        return f"{self.pred_id}{self.dep_type}{lag_text}" if lag_text else f"{self.pred_id}{self.dep_type}"


@dataclass
class Task:
    task_id: str
    name: str
    rtask_name: str
    level: str
    base_duration_hours: float
    predecessors: List[Dependency]
    effective_duration_hours: float = 0.0
    es: float = 0.0
    ef: float = 0.0
    ls: float = 0.0
    lf: float = 0.0
    total_slack: float = 0.0
    free_slack: float = 0.0

    def normalize_predecessors(self) -> str:
        return "; ".join(dep.normalized() for dep in self.predecessors)


class WorkbookReader:
    """Thin wrapper around the existing workbook loader."""

    def __init__(self, path: str, sheet_name: str = "Sheet1") -> None:
        self.path = path
        self.sheet_name = sheet_name

    def read_rows(self, required_headers: Sequence[str]) -> List[Dict[str, str]]:
        from cps_preprocessor import _read_workbook

        with open(self.path, "rb") as handle:
            workbook = _read_workbook(handle.read())

        rows: List[Dict[str, str]] = []
        header_lookup = {key.lower(): value for key, value in workbook.header_lookup.items()}
        for required in required_headers:
            if required.lower() not in header_lookup:
                raise ValueError(f"Workbook is missing required column: {required}")
        for row in workbook.rows:
            result: Dict[str, str] = {}
            for required in required_headers:
                normalized = required.lower()
                header = header_lookup.get(normalized, required)
                result[required] = row.get(header, "")
            rows.append(result)
        return rows

    # Legacy helper methods are intentionally omitted; parsing is delegated to
    # the existing preprocessing loader for reliability on large workbooks.


class WorkingCalendar:
    def __init__(self, project_start: dt.datetime) -> None:
        self.project_start = self._align(project_start)
        self.day_start_time = self.project_start.time()

    def _align(self, moment: dt.datetime) -> dt.datetime:
        aligned = moment
        while aligned.weekday() >= 5:
            aligned = dt.datetime.combine(aligned.date() + dt.timedelta(days=1), aligned.time())
        day_start = dt.datetime.combine(aligned.date(), aligned.time())
        day_end = day_start + dt.timedelta(hours=WORK_HOURS_PER_DAY)
        if aligned < day_start:
            aligned = day_start
        if aligned >= day_end:
            aligned = self._next_workday_start(day_end)
        return aligned

    def _next_workday_start(self, moment: dt.datetime) -> dt.datetime:
        date = moment.date()
        next_date = date + dt.timedelta(days=1)
        while next_date.weekday() >= 5:
            next_date += dt.timedelta(days=1)
        return dt.datetime.combine(next_date, self.day_start_time)

    def add_work_hours(self, hours: float) -> dt.datetime:
        current = self.project_start
        return self.add_from(current, hours)

    def add_from(self, moment: dt.datetime, hours: float) -> dt.datetime:
        current = self._align(moment)
        remaining = max(hours, 0.0)
        while remaining > 1e-9:
            day_start = dt.datetime.combine(current.date(), self.day_start_time)
            day_end = day_start + dt.timedelta(hours=WORK_HOURS_PER_DAY)
            if current < day_start:
                current = day_start
            if current >= day_end:
                current = self._next_workday_start(current)
                continue
            available = (day_end - current).total_seconds() / 3600.0
            if remaining <= available + 1e-9:
                current += dt.timedelta(hours=remaining)
                remaining = 0.0
            else:
                remaining -= available
                current = self._next_workday_start(day_end)
        return current

    def hours_to_datetime(self, hours: float) -> dt.datetime:
        return self.add_from(self.project_start, hours)


class ScheduleBuilder:
    def __init__(self, calendar: WorkingCalendar) -> None:
        self.calendar = calendar

    def build(self, tasks: Dict[str, Task]) -> None:
        order = self._topological_order(tasks)
        self._forward_pass(order, tasks)
        self._backward_pass(order, tasks)
        self._compute_slack(tasks)

    def _forward_pass(self, order: Sequence[str], tasks: Dict[str, Task]) -> None:
        for task_id in order:
            task = tasks[task_id]
            es_candidate = 0.0
            for dep in task.predecessors:
                if dep.pred_id not in tasks:
                    raise ValueError(f"Predecessor {dep.pred_id} referenced by {task_id} is missing.")
                pred = tasks[dep.pred_id]
                if dep.dep_type == "FS":
                    candidate = pred.ef + dep.lag_hours
                elif dep.dep_type == "SS":
                    candidate = pred.es + dep.lag_hours
                elif dep.dep_type == "FF":
                    candidate = pred.ef + dep.lag_hours - task.effective_duration_hours
                else:  # SF
                    candidate = pred.es + dep.lag_hours - task.effective_duration_hours
                es_candidate = max(es_candidate, candidate)
            task.es = max(es_candidate, 0.0)
            task.ef = task.es + task.effective_duration_hours

    def _backward_pass(self, order: Sequence[str], tasks: Dict[str, Task]) -> None:
        project_finish = max((tasks[task_id].ef for task_id in order), default=0.0)
        successors: Dict[str, List[Tuple[str, Dependency]]] = {task_id: [] for task_id in tasks}
        for task in tasks.values():
            for dep in task.predecessors:
                if dep.pred_id in successors:
                    successors[dep.pred_id].append((task.task_id, dep))
        for task_id in reversed(order):
            task = tasks[task_id]
            duration = task.effective_duration_hours
            lf_candidates: List[float] = []
            ls_candidates: List[float] = []
            for succ_id, dep in successors.get(task_id, []):
                succ = tasks[succ_id]
                if dep.dep_type == "FS":
                    lf_candidates.append(succ.es - dep.lag_hours)
                elif dep.dep_type == "SS":
                    ls_candidates.append(succ.es - dep.lag_hours)
                elif dep.dep_type == "FF":
                    lf_candidates.append(succ.ef - dep.lag_hours)
                else:  # SF
                    ls_candidates.append(succ.ef - dep.lag_hours)
            lf = min(lf_candidates) if lf_candidates else project_finish
            ls = min(ls_candidates) if ls_candidates else lf - duration
            ls = min(ls, lf - duration)
            lf = min(lf, ls + duration)
            task.lf = lf
            task.ls = ls

    def _compute_slack(self, tasks: Dict[str, Task]) -> None:
        for task in tasks.values():
            task.total_slack = task.ls - task.es
            slack_candidates: List[float] = []
            for succ_id, dep in self._successors_for(task, tasks):
                succ = tasks[succ_id]
                if dep.dep_type == "FS":
                    slack_candidates.append(succ.es - (task.ef + dep.lag_hours))
                elif dep.dep_type == "SS":
                    slack_candidates.append(succ.es - (task.es + dep.lag_hours))
                elif dep.dep_type == "FF":
                    slack_candidates.append(succ.ef - (task.ef + dep.lag_hours))
                else:
                    slack_candidates.append(succ.ef - (task.es + dep.lag_hours))
            task.free_slack = min(slack_candidates) if slack_candidates else task.total_slack

    @staticmethod
    def _successors_for(task: Task, tasks: Dict[str, Task]) -> List[Tuple[str, Dependency]]:
        successors: List[Tuple[str, Dependency]] = []
        for other in tasks.values():
            for dep in other.predecessors:
                if dep.pred_id == task.task_id:
                    successors.append((other.task_id, dep))
        return successors

    def _topological_order(self, tasks: Dict[str, Task]) -> List[str]:
        adjacency: Dict[str, List[str]] = {task_id: [] for task_id in tasks}
        indegree: Dict[str, int] = {task_id: 0 for task_id in tasks}
        for task in tasks.values():
            for dep in task.predecessors:
                if dep.pred_id not in tasks:
                    raise ValueError(f"Predecessor {dep.pred_id} referenced by {task.task_id} is missing.")
                adjacency[dep.pred_id].append(task.task_id)
                indegree[task.task_id] += 1
        queue: List[str] = [task_id for task_id, deg in indegree.items() if deg == 0]
        order: List[str] = []
        while queue:
            current = queue.pop()
            order.append(current)
            for neighbor in adjacency[current]:
                indegree[neighbor] -= 1
                if indegree[neighbor] == 0:
                    queue.append(neighbor)
        if len(order) != len(tasks):
            raise ValueError("Cycle detected in dependency graph.")
        return order


def parse_duration(value: str) -> float:
    if value is None:
        raise ValueError("Missing duration value.")
    text = value.strip()
    if not text:
        raise ValueError("Empty duration string.")
    if text.endswith("?"):
        text = text[:-1]
    match = re.fullmatch(r"([0-9]*\.?[0-9]+)\s*(e?)(d|w|mo|h)", text.lower())
    if not match:
        raise ValueError(f"Invalid duration string: {value!r}")
    magnitude = float(match.group(1))
    unit = match.group(3)
    if unit == "h":
        return magnitude
    if unit == "d":
        return magnitude * WORK_HOURS_PER_DAY
    if unit == "w":
        return magnitude * WORK_DAYS_PER_WEEK * WORK_HOURS_PER_DAY
    if unit == "mo":
        return magnitude * WORK_DAYS_PER_MONTH * WORK_HOURS_PER_DAY
    raise ValueError(f"Unsupported duration unit: {unit}")


def parse_lag(text: str) -> float:
    cleaned = text.replace(" ", "")
    match = re.fullmatch(r"([+-])([0-9]*\.?[0-9]+)(d|w|mo|h)", cleaned.lower())
    if not match:
        raise ValueError(f"Invalid lag specification: {text!r}")
    sign = -1.0 if match.group(1) == "-" else 1.0
    magnitude = float(match.group(2))
    unit = match.group(3)
    hours = parse_duration(f"{magnitude}{unit}")
    return sign * hours


def parse_dependencies(raw: str) -> List[Dependency]:
    if not raw:
        return []
    dependencies: List[Dependency] = []
    tokens = re.split(r"[;,\n]+", raw)
    for token in tokens:
        stripped = token.strip()
        if not stripped:
            continue
        match = re.fullmatch(r"([A-Za-z0-9_.-]+?)(?:\s*(FS|SS|FF|SF))?(?:\s*([+-].+))?", stripped, re.IGNORECASE)
        if not match:
            raise ValueError(f"Unable to parse predecessor token: {token!r}")
        pred_id = match.group(1)
        dep_type = match.group(2).upper() if match.group(2) else "FS"
        if dep_type not in SUPPORTED_DEP_TYPES:
            raise ValueError(f"Unsupported dependency type {dep_type!r} for predecessor {pred_id}")
        lag_text = match.group(3) or ""
        lag_hours = parse_lag(lag_text) if lag_text else 0.0
        dependencies.append(Dependency(pred_id=pred_id, dep_type=dep_type, lag_hours=lag_hours))
    return dependencies


def compute_effective_durations(tasks: Dict[str, Task], targeted_levels: Iterable[str]) -> None:
    targeted = {level.lower() for level in targeted_levels}
    for task in tasks.values():
        base = task.base_duration_hours
        if task.level.lower() in targeted:
            unique_preds = {dep.pred_id for dep in task.predecessors}
            sum_pred = sum(tasks[pred_id].effective_duration_hours for pred_id in unique_preds if pred_id in tasks)
            task.effective_duration_hours = max(base, sum_pred)
        else:
            task.effective_duration_hours = base


def load_tasks(
    rows: Sequence[Dict[str, str]],
    targeted_levels: Iterable[str],
    ignore_missing: bool,
) -> Tuple[Dict[str, Task], List[str]]:
    tasks: Dict[str, Task] = {}
    for row in rows:
        task_id = (row.get("TaskId") or "").strip()
        if not task_id:
            raise ValueError("Encountered a task without TaskId.")
        if task_id in tasks:
            raise ValueError(f"Duplicate TaskId detected: {task_id}")
        name = (row.get("Task Name") or "").strip()
        rtask_name = (row.get("Rtask Name") or name).strip()
        if not name:
            raise ValueError(f"Task {task_id} is missing a name.")
        level = (row.get("Task Level") or "").strip()
        if not level:
            raise ValueError(f"Task {task_id} is missing Task Level.")
        base_duration = (row.get("Base Duration") or "").strip()
        base_hours = parse_duration(base_duration)
        predecessors = parse_dependencies(row.get("Predecessors IDs", ""))
        tasks[task_id] = Task(
            task_id=task_id,
            name=name,
            rtask_name=rtask_name,
            level=level,
            base_duration_hours=base_hours,
            predecessors=predecessors,
        )
    missing: set[str] = set()
    for task in tasks.values():
        filtered: List[Dependency] = []
        for dep in task.predecessors:
            if dep.pred_id not in tasks:
                missing.add(dep.pred_id)
            else:
                filtered.append(dep)
        task.predecessors = filtered
    if missing and not ignore_missing:
        missing_list = sorted(missing)
        raise ValueError(
            "Unresolved predecessor references detected: " + ", ".join(missing_list)
        )
    compute_effective_durations(tasks, targeted_levels)
    return tasks, sorted(missing)


def export_csv(path: str, headers: Sequence[str], rows: Sequence[Sequence[str]]) -> None:
    with open(path, "w", newline="", encoding="utf-8") as handle:
        writer = csv.writer(handle)
        writer.writerow(headers)
        for row in rows:
            writer.writerow(row)


def export_xlsx(path: str, headers: Sequence[str], rows: Sequence[Sequence[str]]) -> None:
    from xml.sax.saxutils import escape

    def cell_xml(col_letter: str, row_index: int, value: str) -> str:
        return (
            f"<c r=\"{col_letter}{row_index}\" t=\"str\"><v>{escape(value)}</v></c>"
        )

    def row_xml(index: int, values: Sequence[str]) -> str:
        cells = []
        for col_idx, value in enumerate(values):
            col_letter = column_letter(col_idx + 1)
            cells.append(cell_xml(col_letter, index, value))
        cells_xml = "".join(cells)
        return f"<row r=\"{index}\">{cells_xml}</row>"

    def column_letter(n: int) -> str:
        result = ""
        while n > 0:
            n, remainder = divmod(n - 1, 26)
            result = chr(65 + remainder) + result
        return result

    sheet_rows = [row_xml(1, headers)]
    for idx, row in enumerate(rows, start=2):
        sheet_rows.append(row_xml(idx, [str(value) for value in row]))
    sheet_data = "".join(sheet_rows)
    worksheet_xml = (
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
        "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">"
        f"<sheetData>{sheet_data}</sheetData>"
        "</worksheet>"
    )

    workbook_xml = (
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
        "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" "
        "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">"
        "<sheets><sheet name=\"Schedule\" sheetId=\"1\" r:id=\"rId1\"/></sheets>"
        "</workbook>"
    )

    workbook_rels_xml = (
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
        "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
        "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet1.xml\"/>"
        "<Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"styles.xml\"/>"
        "</Relationships>"
    )

    styles_xml = (
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
        "<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">"
        "<fonts count=\"1\"><font><sz val=\"11\"/><name val=\"Calibri\"/></font></fonts>"
        "<fills count=\"1\"><fill><patternFill patternType=\"none\"/></fill></fills>"
        "<borders count=\"1\"><border/></borders>"
        "<cellStyleXfs count=\"1\"><xf/></cellStyleXfs>"
        "<cellXfs count=\"1\"><xf/></cellXfs>"
        "</styleSheet>"
    )

    content_types_xml = (
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
        "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">"
        "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>"
        "<Default Extension=\"xml\" ContentType=\"application/xml\"/>"
        "<Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>"
        "<Override PartName=\"/xl/worksheets/sheet1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>"
        "<Override PartName=\"/xl/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml\"/>"
        "</Types>"
    )

    root_rels_xml = (
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
        "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
        "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.xml\"/>"
        "</Relationships>"
    )

    with ZipFile(path, "w") as zf:
        zf.writestr("[Content_Types].xml", content_types_xml)
        zf.writestr("_rels/.rels", root_rels_xml)
        zf.writestr("xl/workbook.xml", workbook_xml)
        zf.writestr("xl/_rels/workbook.xml.rels", workbook_rels_xml)
        zf.writestr("xl/worksheets/sheet1.xml", worksheet_xml)
        zf.writestr("xl/styles.xml", styles_xml)


def format_hours(hours: float) -> str:
    return f"{hours:.3f}"


def main(argv: Optional[Sequence[str]] = None) -> None:
    parser = argparse.ArgumentParser(description="Compute CPM schedule from Excel workbook.")
    parser.add_argument("--input", default="cps_rules_level_4.xlsx", help="Path to input workbook.")
    parser.add_argument("--sheet", default="Sheet1", help="Worksheet name to process.")
    parser.add_argument("--project-start", required=True, help="Project start datetime (ISO 8601).")
    parser.add_argument(
        "--effective-levels",
        default="L2,L3,L4",
        help="Comma-separated Task Level labels that use the Level-n duration rule.",
    )
    parser.add_argument("--output-csv", default="schedule_output.csv", help="Path for CSV export.")
    parser.add_argument("--output-xlsx", default="schedule_output.xlsx", help="Path for XLSX export.")
    parser.add_argument(
        "--ignore-missing",
        action="store_true",
        help="Ignore predecessor references to tasks that are absent from the sheet.",
    )
    args = parser.parse_args(argv)

    try:
        project_start = dt.datetime.fromisoformat(args.project_start)
    except ValueError as exc:
        raise SystemExit(f"Invalid project start datetime: {args.project_start!r}") from exc

    required_headers = [
        "TaskId",
        "Task Name",
        "Task Level",
        "Predecessors IDs",
        "Base Duration",
    ]
    reader = WorkbookReader(args.input, args.sheet)
    rows = reader.read_rows(required_headers)
    effective_levels = [level.strip() for level in args.effective_levels.split(",") if level.strip()]
    if not effective_levels:
        effective_levels = []
    tasks, missing_refs = load_tasks(rows, effective_levels, args.ignore_missing)

    calendar = WorkingCalendar(project_start)
    builder = ScheduleBuilder(calendar)
    builder.build(tasks)

    headers = [
        "TaskId",
        "Rtask Name",
        "Task Name",
        "Task Level",
        "BaseDurationHours",
        "EffectiveDurationHours",
        "Normalized Predecessors",
        "ES (h)",
        "EF (h)",
        "LS (h)",
        "LF (h)",
        "TotalSlack (h)",
        "FreeSlack (h)",
        "IsCritical",
        "StartDate",
        "FinishDate",
    ]

    rows_out: List[List[str]] = []
    if missing_refs and args.ignore_missing:
        print(
            "Warning: Ignoring predecessor references to missing task IDs: "
            + ", ".join(missing_refs)
        )

    for task in tasks.values():
        start_dt = calendar.hours_to_datetime(task.es)
        finish_dt = calendar.hours_to_datetime(task.ef)
        rows_out.append(
            [
                task.task_id,
                task.rtask_name,
                task.name,
                task.level,
                format_hours(task.base_duration_hours),
                format_hours(task.effective_duration_hours),
                task.normalize_predecessors(),
                format_hours(task.es),
                format_hours(task.ef),
                format_hours(task.ls),
                format_hours(task.lf),
                format_hours(task.total_slack),
                format_hours(task.free_slack),
                "Yes" if task.total_slack <= 1e-9 else "No",
                start_dt.isoformat(sep=" "),
                finish_dt.isoformat(sep=" "),
            ]
        )

    export_csv(args.output_csv, headers, rows_out)
    export_xlsx(args.output_xlsx, headers, rows_out)


if __name__ == "__main__":
    main()
