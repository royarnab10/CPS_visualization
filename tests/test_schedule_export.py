import csv
import sys
from datetime import datetime
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from cps_tool.models import ScheduleResult, ScheduledTask, TaskSpec


def _example_task(uid: int) -> ScheduledTask:
    spec = TaskSpec(uid=uid, name=f"Task {uid}", duration_days=1.5)
    start = datetime(2024, 1, uid, 8, 0)
    finish = datetime(2024, 1, uid, 17, 0)
    return ScheduledTask(
        spec=spec,
        earliest_start=start,
        earliest_finish=finish,
        latest_start=start,
        latest_finish=finish,
        total_float_hours=0.0,
    )


def test_schedule_result_to_csv(tmp_path):
    task = _example_task(2)
    result = ScheduleResult(
        project_start=datetime(2024, 1, 2, 8, 0),
        project_finish=datetime(2024, 1, 2, 17, 0),
        tasks=[task],
    )

    output_path = tmp_path / "schedule.csv"
    written_path = result.to_csv(output_path)

    assert written_path == output_path
    assert output_path.exists()

    with output_path.open(newline="", encoding="utf-8") as handle:
        reader = csv.DictReader(handle)
        rows = list(reader)

    assert reader.fieldnames is not None
    assert {"earliest_start", "earliest_finish", "latest_start", "latest_finish"}.issubset(
        set(reader.fieldnames)
    )
    assert len(rows) == 1
    assert rows[0]["uid"] == "2"
    assert rows[0]["name"] == "Task 2"
