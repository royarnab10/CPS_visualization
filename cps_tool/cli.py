"""Command line interface for the CPS tooling."""
from __future__ import annotations

import argparse
from datetime import datetime, time
from pathlib import Path
from typing import Sequence

from .calendar import WorkCalendar, parse_weekend
from .calculator import calculate_schedule
from .csv_loader import load_tasks_from_csv


TIME_FORMAT = "%H:%M"
DATETIME_HINTS = ["%Y-%m-%d", "%Y-%m-%dT%H:%M", "%Y-%m-%dT%H:%M:%S"]


def _parse_time(value: str) -> time:
    try:
        return datetime.strptime(value, TIME_FORMAT).time()
    except ValueError as exc:
        raise argparse.ArgumentTypeError(
            f"Time values must use HH:MM format; received '{value}'."
        ) from exc


def _parse_datetime(value: str) -> datetime:
    try:
        return datetime.fromisoformat(value)
    except ValueError:
        pass
    for pattern in DATETIME_HINTS:
        try:
            return datetime.strptime(value, pattern)
        except ValueError:
            continue
    raise argparse.ArgumentTypeError(
        "Datetime values must use ISO-8601 format (YYYY-MM-DD or YYYY-MM-DDTHH:MM[:SS])."
    )


def _handle_convert(args: argparse.Namespace) -> None:
    from .mpp_converter import convert_mpp_to_csv

    output = convert_mpp_to_csv(
        path=args.mpp_path,
        output=args.output,
        include_summary=args.include_summary,
    )
    print(f"Extracted {output}")


def _handle_calculate(args: argparse.Namespace) -> None:
    tasks = load_tasks_from_csv(args.csv_path)
    weekend = parse_weekend(args.weekend or [])
    calendar = WorkCalendar(
        workday_start=_parse_time(args.workday_start),
        workday_end=_parse_time(args.workday_end),
        weekend_days=weekend,
    )
    project_start = _parse_datetime(args.project_start) if args.project_start else None
    result = calculate_schedule(tasks, project_start=project_start, calendar=calendar)
    print("CPS calculation complete:")
    print(f"  Calendar: {calendar.describe()}")
    print(f"  Tasks processed: {len(result.tasks)}")
    print(f"  Project start:  {result.project_start.isoformat()}")
    print(f"  Project finish: {result.project_finish.isoformat()}")
    duration_hours = calendar.work_hours_between(
        result.project_start, result.project_finish
    )
    if calendar.hours_per_day:
        duration_days = duration_hours / calendar.hours_per_day
        print(f"  Working duration: {duration_hours:.2f} hours ({duration_days:.2f} days)")
    critical_count = sum(1 for task in result.tasks if task.is_critical)
    print(f"  Critical tasks: {critical_count}")
    if result.cycle_resolutions:
        print("  Dependency cycles detected and resolved:")
        for resolution in result.cycle_resolutions:
            print(f"    Cycle: {resolution.formatted_cycle()}")
            dependency = resolution.removed_dependency
            relation = (dependency.relation_type or "FS").upper()
            lag = f"{dependency.lag_days:g}"
            print(
                "      Removed dependency: "
                f"{resolution.removed_from_task_uid} ({resolution.removed_from_task_name}) "
                f"<- {dependency.predecessor_uid} [{relation} lag {lag} days]"
            )
        if result.cycle_adjusted_csv:
            print(f"  Cycle-adjusted CPS written to {result.cycle_adjusted_csv}")
    if result.dependency_issues:
        print("  Invalid dependencies removed:")
        for issue in result.dependency_issues:
            print(f"    {issue.formatted_issue()}")
    if args.output:
        output_path = Path(args.output)
        result.to_csv(output_path)
        print(f"  Detailed schedule written to {output_path}")


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="CPS tooling utilities")
    subparsers = parser.add_subparsers(dest="command", required=True)

    convert_parser = subparsers.add_parser(
        "convert", help="Convert a Microsoft Project .mpp file to CSV"
    )
    convert_parser.add_argument("mpp_path", type=Path, help="Path to the .mpp file")
    convert_parser.add_argument(
        "output", type=Path, help="Destination CSV file path"
    )
    convert_parser.add_argument(
        "--include-summary",
        action="store_true",
        help="Include summary rows when exporting tasks",
    )
    convert_parser.set_defaults(func=_handle_convert)

    calculate_parser = subparsers.add_parser(
        "calculate", help="Calculate a critical path schedule from CSV"
    )
    calculate_parser.add_argument("csv_path", type=Path, help="Path to the task CSV file")
    calculate_parser.add_argument(
        "--project-start",
        dest="project_start",
        help="Override project start (ISO-8601 datetime)",
    )
    calculate_parser.add_argument(
        "--workday-start", default="08:00", help="Start of the workday (HH:MM)",
    )
    calculate_parser.add_argument(
        "--workday-end", default="17:00", help="End of the workday (HH:MM)",
    )
    calculate_parser.add_argument(
        "--weekend",
        nargs="*",
        default=None,
        help="Weekend days, e.g. 'sat sun' or '5 6'",
    )
    calculate_parser.add_argument(
        "--output",
        type=Path,
        help="Optional CSV destination for the calculated schedule",
    )
    calculate_parser.set_defaults(func=_handle_calculate)

    return parser


def main(argv: Sequence[str] | None = None) -> None:
    parser = build_parser()
    args = parser.parse_args(argv)
    args.func(args)


if __name__ == "__main__":
    main()
