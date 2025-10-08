"""Minimal HTTP server for the CPS visualizer with preprocessing support."""
from __future__ import annotations

import argparse
import base64
import json
from functools import partial
from http import HTTPStatus
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from typing import Tuple
import cgi

from cps_preprocessor import preprocess_excel_with_workbook
from dag_loader import load_task_dependency_records

WEBAPP_DIR = Path(__file__).resolve().parent / "webapp"


class CPSRequestHandler(SimpleHTTPRequestHandler):
    """Serve static assets and expose a preprocessing API."""

    def __init__(self, *args, directory: str | None = None, **kwargs) -> None:
        super().__init__(*args, directory=directory or str(WEBAPP_DIR), **kwargs)

    def do_POST(self) -> None:  # noqa: N802 - part of the HTTP handler API
        if self.path == "/api/preprocess":
            self._handle_preprocess()
        elif self.path == "/api/dag/upload":
            self._handle_dag_upload()
        else:
            self.send_error(HTTPStatus.NOT_FOUND, "Endpoint not found")

    def _handle_preprocess(self) -> None:
        upload = self._extract_uploaded_file()
        if upload is None:
            return

        file_bytes, file_name = upload
        try:
            result = preprocess_excel_with_workbook(file_bytes)
        except ValueError as exc:
            self.send_error(HTTPStatus.BAD_REQUEST, str(exc))
            return

        payload_dict = {
            "records": result.rows,
            "metadata": result.metadata,
            "excel": _build_excel_payload(
                result.excel_bytes,
                file_name,
            ),
        }

        payload = json.dumps(payload_dict).encode("utf-8")
        self.send_response(HTTPStatus.OK)
        self.send_header("Content-Type", "application/json")
        self.send_header("Content-Length", str(len(payload)))
        self.end_headers()
        self.wfile.write(payload)

    def _handle_dag_upload(self) -> None:
        upload = self._extract_uploaded_file()
        if upload is None:
            return

        file_bytes, _ = upload
        try:
            payload = load_task_dependency_records(file_bytes)
        except ValueError as exc:
            self.send_error(HTTPStatus.BAD_REQUEST, str(exc))
            return

        response_dict = {
            "records": payload.records,
            "headers": list(payload.headers),
        }

        encoded = json.dumps(response_dict).encode("utf-8")
        self.send_response(HTTPStatus.OK)
        self.send_header("Content-Type", "application/json")
        self.send_header("Content-Length", str(len(encoded)))
        self.end_headers()
        self.wfile.write(encoded)

    def _extract_uploaded_file(self) -> tuple[bytes, str | None] | None:
        content_type = self.headers.get("Content-Type", "")
        ctype, pdict = cgi.parse_header(content_type)
        if ctype != "multipart/form-data":
            self.send_error(HTTPStatus.BAD_REQUEST, "Expected multipart/form-data upload")
            return None

        boundary = pdict.get("boundary")
        if not boundary:
            self.send_error(HTTPStatus.BAD_REQUEST, "Missing multipart boundary")
            return None

        pdict["boundary"] = boundary.encode()
        pdict["CONTENT-LENGTH"] = int(self.headers.get("Content-Length", "0"))

        form = cgi.FieldStorage(
            fp=self.rfile,
            headers=self.headers,
            environ={
                "REQUEST_METHOD": "POST",
                "CONTENT_TYPE": content_type,
            },
            keep_blank_values=True,
        )

        file_item = form["file"] if "file" in form else None
        if file_item is None or not getattr(file_item, "file", None):
            self.send_error(HTTPStatus.BAD_REQUEST, "Upload must include a 'file' field")
            return None

        file_bytes = file_item.file.read()
        file_name = getattr(file_item, "filename", None)
        return file_bytes, file_name

    def log_message(self, format: str, *args) -> None:
        # Prefix log entries to make server output easier to follow.
        super().log_message("[CPS] " + format, *args)


def serve(address: Tuple[str, int]) -> None:
    handler = partial(CPSRequestHandler, directory=str(WEBAPP_DIR))
    with ThreadingHTTPServer(address, handler) as httpd:
        host, port = httpd.server_address
        print(f"Serving CPS visualizer on http://{host}:{port}")
        print("Press Ctrl+C to stop.")
        try:
            httpd.serve_forever()
        except KeyboardInterrupt:
            print("\nShutting down...")


def main() -> None:
    parser = argparse.ArgumentParser(description="Serve the CPS visualizer with preprocessing support")
    parser.add_argument("--host", default="0.0.0.0", help="Host interface to bind (default: 0.0.0.0)")
    parser.add_argument("--port", type=int, default=8000, help="Port to listen on (default: 8000)")
    args = parser.parse_args()
    serve((args.host, args.port))


if __name__ == "__main__":
    main()


def _build_excel_payload(excel_bytes: bytes, original_filename: str | None) -> dict[str, str]:
    content_type = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    filename = _derive_processed_filename(original_filename)
    encoded = base64.b64encode(excel_bytes).decode("ascii")
    return {
        "filename": filename,
        "contentType": content_type,
        "data": encoded,
    }


def _derive_processed_filename(original_filename: str | None) -> str:
    if not original_filename:
        return "cps_preprocessed.xlsx"
    sanitized = Path(original_filename).name
    stem = Path(sanitized).stem or "cps_preprocessed"
    return f"{stem}_preprocessed.xlsx"
