from __future__ import annotations

import csv
import re
import sys
from pathlib import Path
from typing import Iterable

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

import cps_preprocessor

OUTPUT_HEADERS = [
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

_DURATION_PATTERN = re.compile(r"(-?\d+(?:\.\d+)?)")
_SEPARATOR_RE = re.compile(r"[;,\n]+")


def _parse_duration(value: str | None) -> float:
    if not value:
        return 0.0
    match = _DURATION_PATTERN.search(value)
    if not match:
        return 0.0
    return float(match.group(1))


def _parse_outline(level: str | None) -> int | None:
    if not level:
        return None
    match = re.search(r"(\d+)", level)
    if not match:
        return None
    try:
        return int(match.group(1))
    except ValueError:
        return None


def _normalise_dependency_types(raw: str | None, count: int) -> list[str]:
    if not raw:
        return ["FS" for _ in range(count)]
    tokens = [token.strip().upper() for token in _SEPARATOR_RE.split(raw) if token.strip()]
    if not tokens:
        return ["FS" for _ in range(count)]
    if len(tokens) < count:
        tokens.extend([tokens[-1] if tokens else "FS"] * (count - len(tokens)))
    return tokens[:count]


def _parse_predecessors(raw: str | None) -> list[str]:
    if not raw:
        return []
    return [token.strip() for token in _SEPARATOR_RE.split(raw) if token.strip()]


def _format_predecessors(row: dict[str, str]) -> str:
    predecessors = _parse_predecessors(row.get("Predecessors IDs"))
    types = _normalise_dependency_types(row.get("Dependency Type"), len(predecessors))
    parts = []
    for predecessor, relation in zip(predecessors, types):
        try:
            uid = int(predecessor)
        except ValueError:
            continue
        parts.append(f"{uid}:{relation}:0")
    return ";".join(parts)


def _iter_rows(path: Path) -> Iterable[dict[str, str]]:
    with path.open("rb") as handle:
        data = handle.read()
    rows, _ = cps_preprocessor.preprocess_excel(data)
    for row in rows:
        task_id = row.get("TaskId")
        name = row.get("Task Name")
        if not task_id or not name:
            continue
        try:
            uid = int(task_id)
        except ValueError:
            continue
        duration_days = _parse_duration(row.get("Calculated Duration (days)"))
        if duration_days == 0.0:
            duration_days = _parse_duration(row.get("Base Duration"))
        outline_level = _parse_outline(row.get("Task Level"))
        predecessors = _format_predecessors(row)
        is_milestone = "yes" if abs(duration_days) < 1e-6 else "no"
        yield {
            "uid": str(uid),
            "name": name,
            "duration_days": f"{duration_days:.3f}",
            "is_milestone": is_milestone,
            "outline_level": str(outline_level) if outline_level is not None else "",
            "constraint_type": "",
            "constraint_date": "",
            "calendar": "",
            "predecessors": predecessors,
            "start": "",
            "finish": "",
        }


def export_sample(csv_path: Path, workbook_path: Path) -> None:
    csv_path.parent.mkdir(parents=True, exist_ok=True)
    rows = list(_iter_rows(workbook_path))
    with csv_path.open("w", newline="", encoding="utf-8") as handle:
        writer = csv.DictWriter(handle, fieldnames=OUTPUT_HEADERS)
        writer.writeheader()
        writer.writerows(rows)


if __name__ == "__main__":
    output_path = Path("outputs/generic_cps_sample.csv")
    workbook = Path("cps_rules_level_4.xlsx")
    export_sample(output_path, workbook)
    print(f"Wrote {output_path}")
