"""Preprocess CPS Excel workbooks to restore implicit dependencies."""
from __future__ import annotations

import io
import json
import math
import re
import sys
from dataclasses import dataclass
from typing import Dict, Iterable, List, Optional, Tuple
from xml.etree import ElementTree as ET
from zipfile import ZipFile


@dataclass
class WorkbookData:
    headers: List[str]
    rows: List[Dict[str, str]]
    header_lookup: Dict[str, str]
    indent_by_row: List[int]


@dataclass
class TaskDurationInfo:
    """Track duration context for a single task."""

    row: Dict[str, str]
    level: Optional[int]
    base_days: Optional[float]
    successors: List[str]
    dependency_types: List[str]
    predecessors: List[str]
    computed: Optional[float] = None
    child_sum: float = 0.0


@dataclass
class PreprocessMetadata:
    added_predecessors: int = 0
    added_successors: int = 0

    @property
    def cleaned(self) -> bool:
        return (self.added_predecessors + self.added_successors) > 0

    def to_dict(self) -> Dict[str, int | bool]:
        return {
            "addedPredecessors": self.added_predecessors,
            "addedSuccessors": self.added_successors,
            "cleaned": self.cleaned,
        }


def preprocess_excel(data: bytes) -> Tuple[List[Dict[str, str]], Dict[str, int | bool]]:
    """Parse and clean an uploaded Excel workbook."""
    workbook = _read_workbook(data)
    _, dependency_metadata = _fill_missing_dependencies(workbook)
    duration_metadata = _enrich_with_durations(workbook)
    metadata = dependency_metadata.to_dict()
    metadata.update(duration_metadata)
    return workbook.rows, metadata


# ---------------------------------------------------------------------------
# Workbook parsing helpers


def _read_workbook(data: bytes) -> WorkbookData:
    with ZipFile(io.BytesIO(data)) as zf:
        shared_strings = _read_shared_strings(zf)
        style_indents = _read_style_indents(zf)
        sheet_path = "xl/worksheets/sheet1.xml"
        if sheet_path not in zf.namelist():
            raise ValueError("Unable to locate the first worksheet in the workbook.")
        with zf.open(sheet_path) as sheet_stream:
            headers, header_lookup, rows, indent_by_row = _parse_sheet_stream(
                sheet_stream, shared_strings, style_indents
            )

    if not headers or not rows:
        raise ValueError("The uploaded workbook does not contain any rows.")

    return WorkbookData(headers=headers, rows=rows, header_lookup=header_lookup, indent_by_row=indent_by_row)


def _parse_sheet_stream(
    stream: io.BufferedReader, shared_strings: List[str], style_indents: Dict[int, int]
) -> Tuple[List[str], Dict[str, str], List[Dict[str, str]], List[int]]:
    context = ET.iterparse(stream, events=("start", "end"))
    _, root = next(context)

    headers: List[str] = []
    header_lookup: Dict[str, str] = {}
    rows: List[Dict[str, str]] = []
    indent_by_row: List[int] = []
    task_name_index: Optional[int] = None

    for event, element in context:
        if event != "end" or _local_name(element.tag) != "row":
            continue

        row_number = int(element.attrib.get("r", "0"))
        cells: Dict[int, Tuple[str, int]] = {}

        for cell in element:
            if _local_name(cell.tag) != "c":
                continue
            reference = cell.attrib.get("r", "")
            column_letter = _column_from_reference(reference)
            if column_letter is None:
                continue
            column_index = _column_index(column_letter)
            style_index = int(cell.attrib.get("s", "0"))
            indent = style_indents.get(style_index, 0)
            value = _read_cell_value(cell, shared_strings)
            cells[column_index] = (value, indent)

        if row_number == 1:
            if cells:
                max_index = max(cells.keys())
                headers = ["" for _ in range(max_index + 1)]
                for index in range(max_index + 1):
                    value = cells.get(index, ("", 0))[0]
                    header_value = value or ""
                    if header_value:
                        normalized = _normalize_key(header_value)
                        if normalized == "rtask name":
                            header_value = "Task Name"
                            normalized = "task name"
                        headers[index] = header_value
                        header_lookup[normalized] = header_value
                    else:
                        headers[index] = ""
                task_name_header = _resolve_header(header_lookup, ("task name", "name"))
                task_name_index = headers.index(task_name_header) if task_name_header else None
        elif headers:
            row_dict: Dict[str, str] = {}
            indent = 0
            for index, (value, indent_candidate) in cells.items():
                if index >= len(headers):
                    continue
                header = headers[index]
                if not header:
                    continue
                if task_name_index is not None and index == task_name_index:
                    indent = indent_candidate
                row_dict[header] = _stringify(value)
            rows.append(row_dict)
            indent_by_row.append(indent)

        element.clear()

    root.clear()
    return headers, header_lookup, rows, indent_by_row

def _read_shared_strings(zf: ZipFile) -> List[str]:
    if "xl/sharedStrings.xml" not in zf.namelist():
        return []
    with zf.open("xl/sharedStrings.xml") as shared_file:
        tree = ET.parse(shared_file)
    root = tree.getroot()
    namespace = {"a": _namespace(root)}
    strings: List[str] = []
    for si in root.findall("a:si", namespace):
        text_parts = [t.text or "" for t in si.findall(".//a:t", namespace)]
        strings.append("".join(text_parts))
    return strings


def _read_style_indents(zf: ZipFile) -> Dict[int, int]:
    if "xl/styles.xml" not in zf.namelist():
        return {}
    with zf.open("xl/styles.xml") as styles_file:
        tree = ET.parse(styles_file)
    root = tree.getroot()
    namespace = {"a": _namespace(root)}
    mapping: Dict[int, int] = {}
    cell_xfs = root.find("a:cellXfs", namespace)
    if cell_xfs is None:
        return mapping
    for index, xf in enumerate(cell_xfs.findall("a:xf", namespace)):
        alignment = xf.find("a:alignment", namespace)
        if alignment is not None and "indent" in alignment.attrib:
            try:
                mapping[index] = int(alignment.attrib["indent"])
            except ValueError:
                continue
    return mapping


def _read_cell_value(cell: ET.Element, shared_strings: List[str]) -> str:
    cell_type = cell.attrib.get("t")
    if cell_type == "s":
        index_text = _value_element_text(cell)
        if not index_text:
            return ""
        try:
            return shared_strings[int(index_text)]
        except (ValueError, IndexError):
            return ""
    if cell_type == "inlineStr":
        text_parts = [node.text or "" for node in cell.iter() if _local_name(node.tag) == "t"]
        return "".join(text_parts)
    value_text = _value_element_text(cell)
    return value_text or ""


# ---------------------------------------------------------------------------
# Dependency restoration


def _fill_missing_dependencies(workbook: WorkbookData) -> Tuple[List[Dict[str, str]], PreprocessMetadata]:
    id_header = _resolve_header(workbook.header_lookup, ("taskid", "task id", "id"))
    name_header = _resolve_header(workbook.header_lookup, ("task name", "rtask name", "name"))
    level_header = _resolve_header(workbook.header_lookup, ("task level", "level"))
    predecessors_header = _resolve_header(
        workbook.header_lookup, ("predecessors ids", "predecessor ids", "predecessors")
    )
    successors_header = _resolve_header(
        workbook.header_lookup, ("successors ids", "successor ids", "successors")
    )

    if not id_header or not name_header:
        raise ValueError("The workbook must include 'Task ID' and 'Task Name' columns.")

    metadata = PreprocessMetadata()
    stack: List[Dict[str, object]] = []

    for index, row in enumerate(workbook.rows):
        task_id = _stringify(row.get(id_header, ""))
        task_name = _stringify(row.get(name_header, ""))
        if not task_id or not task_name:
            continue

        level_value = row.get(level_header, "") if level_header else ""
        level_number = _parse_level(level_value)
        predecessors = _split_ids(row.get(predecessors_header, "")) if predecessors_header else []
        successors = _split_ids(row.get(successors_header, "")) if successors_header else []
        current_indent = workbook.indent_by_row[index] if index < len(workbook.indent_by_row) else 0

        while stack and stack[-1]["indent"] >= current_indent:
            stack.pop()

        parent = stack[-1] if stack else None
        if parent:
            parent_level = parent.get("level")
            if _is_valid_parent(level_number, parent_level):
                if predecessors_header:
                    parent_predecessors: List[str] = parent["predecessors"]  # type: ignore[assignment]
                    if task_id not in parent_predecessors:
                        parent_predecessors.append(task_id)
                        parent_row: Dict[str, str] = parent["row"]  # type: ignore[assignment]
                        parent_row[predecessors_header] = _join_ids(parent_predecessors)
                        metadata.added_predecessors += 1
                if successors_header:
                    if parent["id"] not in successors:
                        successors.append(parent["id"])
                        row[successors_header] = _join_ids(successors)
                        metadata.added_successors += 1

        stack.append(
            {
                "id": task_id,
                "indent": current_indent,
                "level": level_number,
                "row": row,
                "predecessors": predecessors,
                "successors": successors,
            }
        )

    return workbook.rows, metadata


def _is_valid_parent(child_level: Optional[int], parent_level: Optional[int]) -> bool:
    if child_level is None or parent_level is None:
        return True
    return child_level == parent_level + 1


# ---------------------------------------------------------------------------
# Duration computation


def _enrich_with_durations(workbook: WorkbookData) -> Dict[str, object]:
    """Normalize duration fields and calculate aggregate task durations."""

    id_header = _resolve_header(workbook.header_lookup, ("taskid", "task id", "id"))
    level_header = _resolve_header(workbook.header_lookup, ("task level", "level"))
    base_header = _resolve_header(
        workbook.header_lookup,
        ("base duration", "duration", "duration (days)", "task duration"),
    )

    if not id_header or not base_header:
        return {}

    successors_header = _resolve_header(
        workbook.header_lookup, ("successors ids", "successor ids", "successors")
    )
    predecessors_header = _resolve_header(
        workbook.header_lookup, ("predecessors ids", "predecessor ids", "predecessors")
    )
    dependency_header = _resolve_header(
        workbook.header_lookup, ("dependency type", "dependency types")
    )

    calc_header = "Calculated Duration (days)"
    if calc_header not in workbook.header_lookup:
        workbook.header_lookup[_normalize_key(calc_header)] = calc_header
        if calc_header not in workbook.headers:
            workbook.headers.append(calc_header)

    tasks: Dict[str, TaskDurationInfo] = {}
    normalized_count = 0

    for row in workbook.rows:
        task_id = _stringify(row.get(id_header, ""))
        if not task_id:
            continue

        original_value = row.get(base_header, "")
        updated_value, base_days, changed = _normalize_duration_value(original_value)
        if changed or updated_value is not None:
            row[base_header] = updated_value or ""
        if changed:
            normalized_count += 1

        level = _parse_level(row.get(level_header, "") if level_header else "")
        successors = _split_ids(row.get(successors_header, "")) if successors_header else []
        predecessors = (
            _split_ids(row.get(predecessors_header, "")) if predecessors_header else []
        )
        dependency_types = (
            [_normalize_dependency_type(value) for value in _split_ids(row.get(dependency_header, ""))]
            if dependency_header
            else []
        )

        tasks[task_id] = TaskDurationInfo(
            row=row,
            level=level,
            base_days=base_days,
            successors=successors,
            dependency_types=dependency_types,
            predecessors=predecessors,
        )

    if not tasks:
        return {}

    visiting: set[str] = set()

    def compute_total(task_id: str) -> float:
        info = tasks.get(task_id)
        if info is None:
            return 0.0
        if info.computed is not None:
            return info.computed
        if task_id in visiting:
            # Cycle detected. Fall back to the task's own duration to avoid recursion loops.
            info.computed = info.base_days or 0.0
            return info.computed

        visiting.add(task_id)
        total_children = 0.0
        if info.successors:
            for index, successor_id in enumerate(info.successors):
                child_total = compute_total(successor_id)
                dependency_type = _dependency_type_for(info.dependency_types, index)
                total_children += _apply_dependency_contribution(dependency_type, child_total)
        visiting.remove(task_id)

        if info.successors:
            info.child_sum = total_children
            info.computed = total_children
        else:
            info.child_sum = info.base_days or 0.0
            info.computed = info.base_days or 0.0

        return info.computed

    for task_identifier in tasks:
        compute_total(task_identifier)

    removed_parent_durations = 0
    for task_identifier, info in tasks.items():
        row = info.row
        if info.computed is not None:
            row[calc_header] = _format_duration_days(info.computed) if info.computed else "0d"
        if (
            info.successors
            and info.base_days is not None
            and info.level in {2, 3}
            and not _duration_close(info.base_days, info.child_sum)
        ):
            row[base_header] = ""
            info.base_days = None
            removed_parent_durations += 1

    root_ids = [task_id for task_id, info in tasks.items() if not info.predecessors]
    if not root_ids:
        min_level = min(
            (info.level for info in tasks.values() if info.level is not None),
            default=None,
        )
        if min_level is not None:
            root_ids = [task_id for task_id, info in tasks.items() if info.level == min_level]

    total_duration = sum(tasks[root_id].computed or 0.0 for root_id in root_ids)

    metadata: Dict[str, object] = {}
    if total_duration:
        metadata["totalDurationDays"] = round(total_duration, 2)
        metadata["totalDurationWeeks"] = round(total_duration / 7.0, 2)
        metadata["totalDurationDisplay"] = _format_duration_days(total_duration)
    else:
        metadata["totalDurationDays"] = 0.0
        metadata["totalDurationWeeks"] = 0.0
        metadata["totalDurationDisplay"] = "0d"

    metadata["durationColumn"] = calc_header
    metadata["clearedParentDurations"] = removed_parent_durations
    metadata["normalizedDurations"] = normalized_count

    return metadata


def _apply_dependency_contribution(dependency_type: str, duration: float) -> float:
    """Determine how a dependency contributes to its predecessor's duration."""

    if duration <= 0:
        return 0.0
    if dependency_type in {"FF", "SS"}:
        # Finish-to-finish and start-to-start usually overlap with their predecessors.
        # Treat them as direct contributions without additional weighting for now.
        return duration
    return duration


def _normalize_dependency_type(value: str) -> str:
    text = _stringify(value).upper()
    if text in {"FS", "FF", "SS", "SF"}:
        return text
    return "FS"


def _dependency_type_for(values: List[str], index: int) -> str:
    if not values:
        return "FS"
    if 0 <= index < len(values):
        candidate = values[index]
        return candidate if candidate else "FS"
    return values[-1] if values[-1] else "FS"


def _normalize_duration_value(value: object) -> Tuple[Optional[str], Optional[float], bool]:
    original = _stringify(value)
    if not original:
        return None, None, False

    stripped = original.strip()
    changed = stripped != original
    if stripped.endswith("?"):
        stripped = stripped[:-1].strip()
        changed = True

    days = _parse_duration_to_days(stripped)
    if days is not None:
        formatted = _format_duration_days(days)
        if formatted == original and not changed:
            return formatted, days, False
        return formatted, days, True

    if changed:
        return stripped, None, True
    return None, None, False


def _parse_duration_to_days(text: str) -> Optional[float]:
    if not text:
        return None
    match = re.fullmatch(
        r"(?i)\s*(\d+(?:\.\d+)?)\s*(d|day|days|w|week|weeks)?\s*",
        text,
    )
    if not match:
        return None
    value = float(match.group(1))
    unit = match.group(2).lower() if match.group(2) else "d"
    if unit.startswith("w"):
        value *= 7.0
    return value


def _format_duration_days(days: float) -> str:
    if math.isclose(days, round(days), rel_tol=1e-9, abs_tol=1e-9):
        return f"{int(round(days))}d"
    text = f"{days:.2f}".rstrip("0").rstrip(".")
    return f"{text}d"


def _duration_close(first: float, second: float, tolerance: float = 0.01) -> bool:
    return math.isclose(first, second, rel_tol=0.0, abs_tol=tolerance)


# ---------------------------------------------------------------------------
# Utility helpers


def _local_name(tag: str) -> str:
    return tag.rsplit('}', 1)[-1] if '}' in tag else tag


def _value_element_text(cell: ET.Element) -> str:
    for child in cell:
        if _local_name(child.tag) == 'v' and child.text:
            return child.text
    return ""


def _namespace(element: ET.Element) -> str:
    if element.tag.startswith("{"):
        return element.tag.split("}", 1)[0].strip("{")
    return ""


def _column_from_reference(reference: str) -> Optional[str]:
    match = re.match(r"([A-Z]+)", reference)
    return match.group(1) if match else None


def _column_index(column: str) -> int:
    index = 0
    for char in column:
        index = index * 26 + (ord(char) - 64)
    return index - 1


def _normalize_key(value: str) -> str:
    return re.sub(r"[^a-z0-9]+", " ", value.strip().lower())


def _resolve_header(mapping: Dict[str, str], options: Iterable[str]) -> Optional[str]:
    for option in options:
        normalized = _normalize_key(option)
        if normalized in mapping:
            return mapping[normalized]
    return None


def _stringify(value: object) -> str:
    if value is None:
        return ""
    if isinstance(value, str):
        return value.strip()
    if isinstance(value, (int, float)):
        if isinstance(value, float) and value.is_integer():
            return str(int(value))
        return str(value)
    return str(value)


def _split_ids(value: object) -> List[str]:
    text = _stringify(value)
    if not text:
        return []
    parts = re.split(r"[,;\n]", text)
    return [part.strip() for part in parts if part.strip()]


def _join_ids(values: List[str]) -> str:
    return ", ".join(dict.fromkeys(values))


def _parse_level(value: object) -> Optional[int]:
    text = _stringify(value)
    if not text:
        return None
    if text.isdigit():
        return int(text)
    match = re.search(r"(\d+)", text)
    if match:
        return int(match.group(1))
    return None


# ---------------------------------------------------------------------------
# CLI helper for manual execution


def main() -> None:
    import argparse

    parser = argparse.ArgumentParser(description="Clean CPS workbooks by inferring hierarchy dependencies.")
    parser.add_argument("input", help="Path to the source .xlsx workbook")
    parser.add_argument("output", nargs="?", help="Optional JSON file for the cleaned rows")
    args = parser.parse_args()

    with open(args.input, "rb") as source:
        records, metadata = preprocess_excel(source.read())

    if args.output:
        with open(args.output, "w", encoding="utf-8") as destination:
            json.dump({"records": records, "metadata": metadata}, destination, indent=2)
    else:
        json.dump({"records": records, "metadata": metadata}, sys.stdout, indent=2)
        sys.stdout.write("\n")


if __name__ == "__main__":
    main()
