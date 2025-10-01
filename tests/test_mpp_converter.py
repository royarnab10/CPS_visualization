import importlib
import sys
import types
import unittest
from datetime import datetime
from pathlib import Path
from unittest import mock


class MpxjImportTests(unittest.TestCase):
    def setUp(self) -> None:
        # Ensure we reload the module with a clean cache between tests.
        if "cps_tool.mpp_converter" in sys.modules:
            module = sys.modules["cps_tool.mpp_converter"]
            try:
                module._load_mpxj.cache_clear()  # type: ignore[attr-defined]
            except AttributeError:
                pass
            del sys.modules["cps_tool.mpp_converter"]

    def tearDown(self) -> None:
        if "cps_tool.mpp_converter" in sys.modules:
            module = sys.modules["cps_tool.mpp_converter"]
            try:
                module._load_mpxj.cache_clear()  # type: ignore[attr-defined]
            except AttributeError:
                pass
            del sys.modules["cps_tool.mpp_converter"]

    def test_extract_tasks_requires_mpxj(self) -> None:
        mpp_converter = importlib.import_module("cps_tool.mpp_converter")
        with self.assertRaises(ImportError) as ctx:
            mpp_converter.extract_tasks_from_mpp(Path("dummy.mpp"))
        self.assertIn("mpxj is required", str(ctx.exception))

    def test_extract_tasks_with_stubbed_mpxj(self) -> None:
        fake_timeunit = types.SimpleNamespace(DAYS="days")

        class FakeDuration:
            def __init__(self, value: float) -> None:
                self._value = value

            def convert(self, unit: object) -> types.SimpleNamespace:
                return types.SimpleNamespace(duration=self._value)

        class FakeTask:
            def __init__(self) -> None:
                self.unique_id = 1
                self.name = "Example"
                self.duration = FakeDuration(5.5)
                self.predecessors = []
                self.milestone = False
                self.outline_level = 2
                self.constraint_type = types.SimpleNamespace(name="ASAP")
                self.constraint_date = datetime(2023, 1, 1, 8, 0)
                self.calendar = types.SimpleNamespace(name="Standard")
                self.start = datetime(2023, 1, 1, 8, 0)
                self.finish = datetime(2023, 1, 2, 17, 0)
                self.summary = False

        class FakeProject:
            def __init__(self) -> None:
                self.tasks = [FakeTask()]

        class FakeReader:
            def read(self, path: str) -> FakeProject:
                self.path = path
                return FakeProject()

        fake_reader_class = FakeReader

        with mock.patch("cps_tool.mpp_converter._load_mpxj", return_value=(fake_timeunit, fake_reader_class)):
            mpp_converter = importlib.import_module("cps_tool.mpp_converter")
            tasks = mpp_converter.extract_tasks_from_mpp(Path("dummy.mpp"))

        self.assertEqual(len(tasks), 1)
        task = tasks[0]
        self.assertEqual(task.uid, 1)
        self.assertEqual(task.name, "Example")
        self.assertAlmostEqual(task.duration_days, 5.5)
        self.assertEqual(task.outline_level, 2)
        self.assertEqual(task.constraint_type, "ASAP")
        self.assertIsNotNone(task.constraint_date)
        self.assertEqual(task.calendar_name, "Standard")

    def test_convert_mpp_to_csv_uses_stubbed_reader(self) -> None:
        fake_timeunit = types.SimpleNamespace(DAYS="days")

        class FakeDuration:
            def __init__(self, value: float) -> None:
                self._value = value

            def convert(self, unit: object) -> types.SimpleNamespace:
                return types.SimpleNamespace(duration=self._value)

        class FakeTask:
            def __init__(self) -> None:
                self.unique_id = 1
                self.name = "Example"
                self.duration = FakeDuration(1.0)
                self.predecessors = []
                self.milestone = True
                self.outline_level = 1
                self.constraint_type = None
                self.constraint_date = None
                self.calendar = types.SimpleNamespace(name=None)
                self.start = None
                self.finish = None
                self.summary = False

        class FakeProject:
            def __init__(self) -> None:
                self.tasks = [FakeTask()]

        class FakeReader:
            def read(self, path: str) -> FakeProject:
                return FakeProject()

        fake_reader_class = FakeReader

        with mock.patch("cps_tool.mpp_converter._load_mpxj", return_value=(fake_timeunit, fake_reader_class)):
            mpp_converter = importlib.import_module("cps_tool.mpp_converter")
            output_path = Path("tests/tmp_output.csv")
            try:
                csv_path = mpp_converter.convert_mpp_to_csv(Path("dummy.mpp"), output_path)
                self.assertTrue(csv_path.exists())
                contents = output_path.read_text(encoding="utf-8").splitlines()
                self.assertEqual(contents[0], "uid,name,duration_days,is_milestone,outline_level,constraint_type,constraint_date,calendar,predecessors,start,finish")
                self.assertIn("Example", contents[1])
            finally:
                if output_path.exists():
                    output_path.unlink()


if __name__ == "__main__":  # pragma: no cover
    unittest.main()
