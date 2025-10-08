import json
import threading
import unittest
from functools import partial
from http.client import HTTPConnection
from http.server import ThreadingHTTPServer
from pathlib import Path

from dag_loader import load_task_dependency_records
from server import CPSRequestHandler, WEBAPP_DIR


def _load_workbook_bytes() -> bytes:
    return Path("amy_cps_new_only_tasks.xlsx").read_bytes()


class TaskDependencyUploadTests(unittest.TestCase):
    def test_load_task_dependency_records(self) -> None:
        payload = load_task_dependency_records(_load_workbook_bytes())
        self.assertEqual(payload.headers[0], "TaskID")
        self.assertEqual(len(payload.records), 916)

        first = payload.records[0]
        self.assertEqual(first["TaskID"], "18")
        self.assertEqual(first["Predecessors"], "1004")
        self.assertEqual(first["Successors"], "21")
        self.assertEqual(first["Current SIMPL Phase (visibility)"], "Launch")
        # Ensure whitespace is trimmed from text fields.
        self.assertTrue(first["Task Name"].endswith("(date needed)"))

    def test_dag_upload_endpoint(self) -> None:
        handler = partial(CPSRequestHandler, directory=str(WEBAPP_DIR))
        server = ThreadingHTTPServer(("127.0.0.1", 0), handler)
        thread = threading.Thread(target=server.serve_forever, daemon=True)
        thread.start()

        try:
            host, port = server.server_address
            boundary = "----PyTestBoundary0X"
            body_prefix = (
                f"--{boundary}\r\n"
                'Content-Disposition: form-data; name="file"; filename="amy_cps_new_only_tasks.xlsx"\r\n'
                "Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet\r\n\r\n"
            ).encode("utf-8")
            body_suffix = f"\r\n--{boundary}--\r\n".encode("utf-8")
            body = body_prefix + _load_workbook_bytes() + body_suffix

            headers = {
                "Content-Type": f"multipart/form-data; boundary={boundary}",
                "Content-Length": str(len(body)),
            }

            connection = HTTPConnection(host, port)
            connection.request("POST", "/api/dag/upload", body=body, headers=headers)
            response = connection.getresponse()
            data = response.read()
            connection.close()

            self.assertEqual(response.status, 200, data)

            payload = json.loads(data.decode("utf-8"))
            self.assertIn("records", payload)
            self.assertEqual(len(payload["records"]), 916)
            self.assertEqual(payload["records"][0]["TaskID"], "18")
        finally:
            server.shutdown()
            server.server_close()
            thread.join()


if __name__ == "__main__":
    unittest.main()
