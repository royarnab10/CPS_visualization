"""Utilities for parsing task dependency workbooks used by the DAG explorer."""
from __future__ import annotations

from dataclasses import dataclass
from typing import Dict, List, Tuple

from cps_preprocessor import _read_workbook, _stringify, WorkbookData


@dataclass
class TaskDependencyPayload:
    """Normalized representation of the dependency workbook."""

    records: List[Dict[str, str]]
    headers: Tuple[str, ...]


def load_task_dependency_records(data: bytes) -> TaskDependencyPayload:
    """Load and normalize rows from a CPS task dependency workbook.

    The explorer expects string values with leading/trailing whitespace trimmed.  The
    workbook headers are preserved in their original order so clients can re-create a
    tabular view when needed.
    """

    workbook = _read_workbook(data)
    records = _normalize_rows(workbook)
    headers = tuple(workbook.headers)
    return TaskDependencyPayload(records=records, headers=headers)


def _normalize_rows(workbook: WorkbookData) -> List[Dict[str, str]]:
    normalized: List[Dict[str, str]] = []
    headers = list(workbook.headers)
    for row in workbook.rows:
        record: Dict[str, str] = {}
        for header in headers:
            record[header] = _normalize_value(row.get(header, ""))
        if any(value for value in record.values()):
            normalized.append(record)
    return normalized


def _normalize_value(value: object) -> str:
    text = _stringify(value)
    return text.strip()
