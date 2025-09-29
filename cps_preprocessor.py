"""Preprocess CPS Excel workbooks to restore implicit dependencies."""
from __future__ import annotations

import io
import json
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
    cleaned_rows, metadata = _fill_missing_dependencies(workbook)
    return cleaned_rows, metadata.to_dict()


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
            if (predecessors_header and not predecessors) and _is_valid_parent(level_number, parent_level):
                predecessors.append(parent["id"])  # type: ignore[arg-type]
                row[predecessors_header] = _join_ids(predecessors)
                metadata.added_predecessors += 1
            if successors_header and not parent["successors"] and _is_valid_parent(level_number, parent_level):
                parent_successors: List[str] = parent["successors"]  # type: ignore[assignment]
                parent_successors.append(task_id)
                parent_row: Dict[str, str] = parent["row"]  # type: ignore[assignment]
                parent_row[successors_header] = _join_ids(parent_successors)
                metadata.added_successors += 1

        stack.append(
            {
                "id": task_id,
                "indent": current_indent,
                "level": level_number,
                "row": row,
                "successors": successors,
            }
        )

    return workbook.rows, metadata


def _is_valid_parent(child_level: Optional[int], parent_level: Optional[int]) -> bool:
    if child_level is None or parent_level is None:
        return True
    return child_level == parent_level + 1


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
