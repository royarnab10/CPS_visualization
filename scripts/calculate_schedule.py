"""Run the CPS calculator and export the schedule as an XLSX workbook."""
from __future__ import annotations

import argparse
import csv
import sys
from datetime import datetime, time
from pathlib import Path
from typing import Dict, Iterable, List, Sequence

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from cps_tool.calendar import WorkCalendar, parse_weekend
from cps_tool.calculator import calculate_schedule
from cps_tool.csv_loader import load_tasks_from_csv
from cps_tool.models import ScheduledTask
from cps_tool.xlsx import write_xlsx


BASE_HEADERS = [
    "TaskId",
    "Responsible Sub-team",
    "Responsible Function",
    "Approval Needed",
    "Rtask Name",
    "Task Level",
    "Predecessors IDs",
    "Successors IDs",
    "SCOPE",
    "Base Duration",
]

SCHEDULE_HEADERS = [
    "Duration (days)",
    "Earliest Start",
    "Earliest Finish",
    "Latest Start",
    "Latest Finish",
    "Total Float (hours)",
    "Is Critical",
    "Constraint Type",
    "Constraint Date",
    "Original Start",
    "Original Finish",
]


DATETIME_PATTERNS = ["%Y-%m-%d", "%Y-%m-%dT%H:%M", "%Y-%m-%dT%H:%M:%S"]


def _parse_time(value: str) -> time:
    try:
        return datetime.strptime(value, "%H:%M").time()
    except ValueError as exc:  # pragma: no cover - defensive parsing
        msg = "Time values must use HH:MM format."
        raise argparse.ArgumentTypeError(msg) from exc


def _parse_datetime(value: str | None) -> datetime | None:
    if not value:
        return None
    try:
        return datetime.fromisoformat(value)
    except ValueError:
        pass
    for pattern in DATETIME_PATTERNS:
        try:
            return datetime.strptime(value, pattern)
        except ValueError:
            continue
    raise argparse.ArgumentTypeError(
        "Datetime values must use ISO-8601 format (YYYY-MM-DD or YYYY-MM-DDTHH:MM[:SS])."
    )


def _read_metadata(csv_path: Path) -> Dict[str, Dict[str, str]]:
    metadata: Dict[str, Dict[str, str]] = {}
    if not csv_path.exists():
        return metadata
    with csv_path.open("r", newline="", encoding="utf-8-sig") as handle:
        reader = csv.DictReader(handle)
        for row in reader:
            uid_raw = (
                row.get("TaskId")
                or row.get("Task ID")
                or row.get("uid")
                or row.get("UID")
                or row.get("Unique ID")
            )
            if not uid_raw:
                continue
            uid_text = str(uid_raw).strip()
            if not uid_text:
                continue
            try:
                uid_key = str(int(float(uid_text)))
            except ValueError:
                uid_key = uid_text
            metadata[uid_key] = {key: value for key, value in row.items()}
    return metadata


def _format_datetime(dt: datetime | None) -> str:
    if dt is None:
        return ""
    return dt.strftime("%Y-%m-%d %H:%M")


def _metadata_lookup(row: Dict[str, str] | None, *keys: str, default: str = "") -> str:
    if not row:
        return default
    for key in keys:
        if key in row and row[key]:
            return str(row[key])
    return default


def _build_successors(tasks: Sequence[ScheduledTask]) -> Dict[int, List[int]]:
    mapping: Dict[int, List[int]] = {task.spec.uid: [] for task in tasks}
    for task in tasks:
        for dependency in task.spec.dependencies:
            mapping.setdefault(dependency.predecessor_uid, []).append(task.spec.uid)
    for successors in mapping.values():
        successors.sort()
    return mapping


def _compose_rows(
    tasks: Sequence[ScheduledTask],
    metadata: Dict[str, Dict[str, str]],
) -> Iterable[List[str]]:
    successors = _build_successors(tasks)
    for task in tasks:
        uid_text = str(task.spec.uid)
        row_meta = metadata.get(uid_text)

        predecessors_override = _metadata_lookup(
            row_meta,
            "Predecessors IDs",
            "Predecessor IDs",
            "Predecessors",
        )
        if predecessors_override:
            predecessors_display = predecessors_override
        else:
            predecessors_display = ",".join(
                str(dep.predecessor_uid) for dep in task.spec.dependencies
            )

        successors_override = _metadata_lookup(
            row_meta, "Successors IDs", "Successor IDs", "Successors"
        )
        if successors_override:
            successors_display = successors_override
        else:
            successors_display = ",".join(
                str(uid) for uid in successors.get(task.spec.uid, [])
            )

        base_duration = _metadata_lookup(
            row_meta,
            "Base Duration",
            "Calculated Duration (days)",
            "Duration",
            "duration_days",
        )
        if not base_duration:
            base_duration = f"{task.spec.duration_days:g}d"

        outline = _metadata_lookup(row_meta, "Task Level", "Level", default="")
        if not outline and task.spec.outline_level is not None:
            outline = f"L{task.spec.outline_level}"

        yield [
            _metadata_lookup(row_meta, "TaskId", "Task ID", "uid", "UID", default=uid_text),
            _metadata_lookup(row_meta, "Responsible Sub-team", "Sub-team"),
            _metadata_lookup(row_meta, "Responsible Function", "Function"),
            _metadata_lookup(row_meta, "Approval Needed"),
            _metadata_lookup(row_meta, "Rtask Name", "Task Name", "name", default=task.spec.name),
            outline,
            predecessors_display,
            successors_display,
            _metadata_lookup(row_meta, "SCOPE", "Scope"),
            base_duration,
            f"{task.spec.duration_days:.3f}",
            _format_datetime(task.earliest_start),
            _format_datetime(task.earliest_finish),
            _format_datetime(task.latest_start),
            _format_datetime(task.latest_finish),
            f"{task.total_float_hours:.2f}",
            "YES" if task.is_critical else "NO",
            task.spec.constraint_type or "",
            _format_datetime(task.spec.constraint_date),
            _metadata_lookup(
                row_meta,
                "Original Start",
                "Start",
                "start",
                "Planned Start",
            ),
            _metadata_lookup(
                row_meta,
                "Original Finish",
                "Finish",
                "finish",
                "Planned Finish",
            ),
        ]


def run_calculation(args: argparse.Namespace) -> Path:
    csv_path = Path(args.csv_path)
    metadata = _read_metadata(csv_path)

    tasks = load_tasks_from_csv(csv_path)
    weekend = parse_weekend(args.weekend or [])
    calendar = WorkCalendar(
        workday_start=_parse_time(args.workday_start),
        workday_end=_parse_time(args.workday_end),
        weekend_days=weekend,
    )
    project_start = _parse_datetime(args.project_start)

    result = calculate_schedule(tasks, project_start=project_start, calendar=calendar)

    headers = BASE_HEADERS + SCHEDULE_HEADERS
    rows = list(_compose_rows(result.tasks, metadata))
    sheet_name = args.sheet_name or "Schedule"
    return write_xlsx(args.output, headers, rows, sheet_name=sheet_name)


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Calculate a CPS schedule and export it as an XLSX workbook."
    )
    parser.add_argument("csv_path", type=Path, help="Path to the source task CSV file")
    parser.add_argument(
        "--output",
        type=Path,
        required=True,
        help="Destination XLSX path",
    )
    parser.add_argument(
        "--sheet-name",
        dest="sheet_name",
        help="Optional sheet name for the workbook",
    )
    parser.add_argument(
        "--project-start",
        dest="project_start",
        help="Override project start (ISO-8601 datetime)",
    )
    parser.add_argument(
        "--workday-start",
        default="08:00",
        help="Start of the workday (HH:MM)",
    )
    parser.add_argument(
        "--workday-end",
        default="17:00",
        help="End of the workday (HH:MM)",
    )
    parser.add_argument(
        "--weekend",
        nargs="*",
        default=None,
        help="Weekend days, e.g. 'sat sun' or '5 6'",
    )
    return parser


def main(argv: Sequence[str] | None = None) -> None:
    parser = build_parser()
    args = parser.parse_args(argv)
    output_path = run_calculation(args)
    print(f"Schedule written to {output_path}")


if __name__ == "__main__":
    main()

