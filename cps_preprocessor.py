"""Preprocess CPS Excel workbooks to restore implicit dependencies."""
from __future__ import annotations

import io
import json
import re
import sys
from dataclasses import dataclass
from typing import Dict, Iterable, List, Optional, Tuple
from xml.etree import ElementTree as ET
from xml.sax.saxutils import escape
from zipfile import ZIP_DEFLATED, ZipFile


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


@dataclass
class PreprocessResult:
    """Structured response returned by :func:`preprocess_excel_with_workbook`."""

    workbook: WorkbookData
    rows: List[Dict[str, str]]
    metadata: Dict[str, int | bool]
    excel_bytes: bytes


def preprocess_excel_with_workbook(data: bytes) -> PreprocessResult:
    """Parse, clean, and serialize an uploaded Excel workbook."""

    workbook = _read_workbook(data)
    _ensure_schedule_compatibility(workbook)
    _, dependency_metadata = _fill_missing_dependencies(workbook)
    metadata = dependency_metadata.to_dict()
    excel_bytes = _serialize_workbook(workbook)
    return PreprocessResult(
        workbook=workbook,
        rows=workbook.rows,
        metadata=metadata,
        excel_bytes=excel_bytes,
    )


def preprocess_excel(data: bytes) -> Tuple[List[Dict[str, str]], Dict[str, int | bool]]:
    """Parse and clean an uploaded Excel workbook.

    This shim preserves the legacy return signature for callers that only
    require the JSON representation of the cleaned rows.
    """

    result = preprocess_excel_with_workbook(data)
    return result.rows, result.metadata


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


def _ensure_schedule_compatibility(workbook: WorkbookData) -> None:
    """Augment workbooks exported from the scheduler with legacy dependency fields."""

    normalized_header = _resolve_header(
        workbook.header_lookup,
        ("normalized predecessors", "normalized predecessor"),
    )
    if not normalized_header:
        return

    predecessor_fields = _ensure_headers(
        workbook,
        ("Predecessors IDs", "Predecessor IDs", "Predecessors"),
    )
    dependency_fields = _ensure_headers(
        workbook,
        ("Dependency Type", "Dependency Types"),
    )

    for row in workbook.rows:
        ids, types = _parse_normalized_dependencies(row.get(normalized_header, ""))
        if not ids:
            continue
        joined_ids = ", ".join(ids)
        joined_types = ", ".join(types)

        for field in predecessor_fields:
            if not _stringify(row.get(field, "")):
                row[field] = joined_ids

        if joined_types:
            for field in dependency_fields:
                if not _stringify(row.get(field, "")):
                    row[field] = joined_types


def _register_header(workbook: WorkbookData, header: str) -> str:
    if header not in workbook.headers:
        workbook.headers.append(header)
    workbook.header_lookup[_normalize_key(header)] = header
    return header


def _ensure_headers(workbook: WorkbookData, names: Iterable[str]) -> List[str]:
    resolved: List[str] = []
    for name in names:
        existing = _resolve_header(workbook.header_lookup, (name,))
        resolved.append(existing or _register_header(workbook, name))
    # Preserve insertion order while removing duplicates
    return list(dict.fromkeys(resolved))


def _parse_normalized_dependencies(value: object) -> Tuple[List[str], List[str]]:
    text = _stringify(value)
    if not text:
        return [], []

    entries = re.split(r"[,;\n]", text)
    ids: List[str] = []
    types: List[str] = []
    seen: set[str] = set()

    for entry in entries:
        candidate = entry.strip()
        if not candidate:
            continue

        type_match = re.search(r"(FS|SS|FF|SF)", candidate, re.IGNORECASE)
        if type_match:
            id_part = candidate[: type_match.start()]
            dependency_type = type_match.group().upper()
        else:
            id_part = candidate
            dependency_type = "FS"

        predecessor_id = _stringify(id_part)
        predecessor_id = re.sub(r"[^A-Za-z0-9_.-]+", "", predecessor_id)
        if not predecessor_id or predecessor_id in seen:
            continue

        seen.add(predecessor_id)
        ids.append(predecessor_id)
        types.append(dependency_type)

    return ids, types


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

    if not predecessors_header:
        predecessors_header = _register_header(workbook, "Predecessors IDs")
    if not successors_header:
        successors_header = _register_header(workbook, "Successors IDs")

    if not id_header or not name_header:
        raise ValueError("The workbook must include 'Task ID' and 'Task Name' columns.")

    metadata = PreprocessMetadata()
    stack: List[Dict[str, object]] = []
    level_one_tasks: List[Dict[str, object]] = []

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

        task_info = {
            "id": task_id,
            "indent": current_indent,
            "level": level_number,
            "row": row,
            "predecessors": predecessors,
            "successors": successors,
        }
        stack.append(task_info)

        if level_number == 1:
            level_one_tasks.append(task_info)

    _link_level_one_tasks(level_one_tasks, predecessors_header, successors_header, metadata)

    return workbook.rows, metadata


def _link_level_one_tasks(
    tasks: List[Dict[str, object]],
    predecessors_header: str,
    successors_header: str,
    metadata: PreprocessMetadata,
) -> None:
    if len(tasks) < 2:
        return

    sorted_tasks = sorted(tasks, key=lambda task: _task_id_sort_key(task["id"]))

    for previous, current in zip(sorted_tasks, sorted_tasks[1:]):
        previous_id = _stringify(previous["id"])
        current_id = _stringify(current["id"])
        if not previous_id or not current_id or previous_id == current_id:
            continue

        current_predecessors: List[str] = current["predecessors"]  # type: ignore[assignment]
        if previous_id not in current_predecessors:
            current_predecessors.append(previous_id)
            current_row: Dict[str, str] = current["row"]  # type: ignore[assignment]
            current_row[predecessors_header] = _join_ids(current_predecessors)
            metadata.added_predecessors += 1

        previous_successors: List[str] = previous["successors"]  # type: ignore[assignment]
        if current_id not in previous_successors:
            previous_successors.append(current_id)
            previous_row: Dict[str, str] = previous["row"]  # type: ignore[assignment]
            previous_row[successors_header] = _join_ids(previous_successors)
            metadata.added_successors += 1


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


def _task_id_sort_key(task_id: object) -> Tuple[int, str]:
    text = _stringify(task_id)
    if not text:
        return sys.maxsize, ""
    if text.isdigit():
        return int(text), text
    match = re.search(r"(\d+)", text)
    if match:
        try:
            return int(match.group(1)), text
        except ValueError:
            pass
    return sys.maxsize, text


# ---------------------------------------------------------------------------
# Workbook serialization


def _serialize_workbook(workbook: WorkbookData) -> bytes:
    headers = _unique_headers(workbook.headers)
    sheet_rows: List[str] = []

    if headers:
        sheet_rows.append(_build_row_xml(1, headers))

    start_index = 2 if headers else 1
    for index, row in enumerate(workbook.rows, start=start_index):
        values = [_stringify(row.get(header, "")) for header in headers]
        sheet_rows.append(_build_row_xml(index, values))

    sheet_data = "".join(sheet_rows)
    worksheet_xml = (
        "<?xml version=\"1.0\" encoding=\"UTF-8\"?>"
        "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">"
        f"<sheetData>{sheet_data}</sheetData>"
        "</worksheet>"
    )

    workbook_xml = (
        "<?xml version=\"1.0\" encoding=\"UTF-8\"?>"
        "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" "
        "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">"
        "<sheets><sheet name=\"Tasks\" sheetId=\"1\" r:id=\"rId1\"/></sheets>"
        "</workbook>"
    )

    workbook_rels_xml = (
        "<?xml version=\"1.0\" encoding=\"UTF-8\"?>"
        "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
        "<Relationship Id=\"rId1\" "
        "Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" "
        "Target=\"worksheets/sheet1.xml\"/>"
        "<Relationship Id=\"rId2\" "
        "Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" "
        "Target=\"styles.xml\"/>"
        "</Relationships>"
    )

    package_rels_xml = (
        "<?xml version=\"1.0\" encoding=\"UTF-8\"?>"
        "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
        "<Relationship Id=\"rId1\" "
        "Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" "
        "Target=\"xl/workbook.xml\"/>"
        "</Relationships>"
    )

    styles_xml = (
        "<?xml version=\"1.0\" encoding=\"UTF-8\"?>"
        "<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">"
        "<fonts count=\"1\"><font/></fonts>"
        "<fills count=\"1\"><fill><patternFill patternType=\"none\"/></fill></fills>"
        "<borders count=\"1\"><border/></borders>"
        "<cellStyleXfs count=\"1\"><xf/></cellStyleXfs>"
        "<cellXfs count=\"1\"><xf xfId=\"0\"/></cellXfs>"
        "<cellStyles count=\"1\"><cellStyle name=\"Normal\" xfId=\"0\" builtinId=\"0\"/></cellStyles>"
        "</styleSheet>"
    )

    content_types_xml = (
        "<?xml version=\"1.0\" encoding=\"UTF-8\"?>"
        "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">"
        "<Default Extension=\"rels\" "
        "ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>"
        "<Default Extension=\"xml\" ContentType=\"application/xml\"/>"
        "<Override PartName=\"/xl/workbook.xml\" "
        "ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>"
        "<Override PartName=\"/xl/worksheets/sheet1.xml\" "
        "ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>"
        "<Override PartName=\"/xl/styles.xml\" "
        "ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml\"/>"
        "</Types>"
    )

    stream = io.BytesIO()
    with ZipFile(stream, "w", ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", content_types_xml)
        zf.writestr("_rels/.rels", package_rels_xml)
        zf.writestr("xl/workbook.xml", workbook_xml)
        zf.writestr("xl/_rels/workbook.xml.rels", workbook_rels_xml)
        zf.writestr("xl/worksheets/sheet1.xml", worksheet_xml)
        zf.writestr("xl/styles.xml", styles_xml)

    return stream.getvalue()


def _build_row_xml(row_number: int, values: List[str]) -> str:
    cells: List[str] = []
    for index, value in enumerate(values):
        if value == "":
            continue
        column = _column_letter(index)
        reference = f"{column}{row_number}"
        escaped = escape(value, {"\n": "&#10;", "\r": "&#13;"})
        cell = (
            f"<c r=\"{reference}\" t=\"inlineStr\">"
            f"<is><t>{escaped}</t></is>"
            "</c>"
        )
        cells.append(cell)

    return f"<row r=\"{row_number}\">{''.join(cells)}</row>"


def _unique_headers(headers: Iterable[str]) -> List[str]:
    seen: set[str] = set()
    ordered: List[str] = []
    for header in headers:
        if not header or header in seen:
            continue
        seen.add(header)
        ordered.append(header)
    return ordered


def _column_letter(index: int) -> str:
    index += 1
    letters: List[str] = []
    while index:
        index, remainder = divmod(index - 1, 26)
        letters.append(chr(65 + remainder))
    return "".join(reversed(letters))


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
