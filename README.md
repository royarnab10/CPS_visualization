# Critical Path Schedule Visualizer

An interactive, browser-based tool for validating complex critical path schedules with non-technical stakeholders. Upload an Excel workbook that lists tasks, milestones, and dependencies to explore the network, expand milestones on demand, and highlight the relationships that drive your critical path.

## Features

- **Excel upload** – Drag in a schedule exported from your planning tool (the first worksheet is used).
- **Collapsible exploration** – Tasks stay grouped by level and scope until you decide to reveal them.
- **Interactive dependency graph** – Visualize predecessor/successor relationships, auto-include nearby context, and highlight task metadata.
- **Dependency typing** – Add a `Dependency Type` column (FS/SS/SF/FF) to distinguish relationships; otherwise Finish-to-Start (FS) is assumed.
- **Stakeholder-friendly UI** – Styled for projection or screen sharing; no coding experience required.

A lightweight sample dataset is included in [`webapp/data/sample_schedule.json`](webapp/data/sample_schedule.json) and the repository ships with the stakeholder workbook `CPS_Rules_comm_LPA_Readiness.xlsx` for real-world testing.

---

## 1. Prerequisites

No runtime installation is required beyond a modern web browser. To keep setup lightweight on macOS or Windows, use the built-in Python interpreter (already available on macOS and easily installed via the Microsoft Store on Windows) to serve the static files. Any alternative static HTTP server (Node, Ruby, Go, etc.) works just as well.

### 1.1 (Optional) create a Python virtual environment

If you plan to use the bundled preprocessing server or the `cps_tool` command-line utilities, isolate the Python dependencies in a virtual environment and install the requirements once:

```bash
python -m venv .venv
source .venv/bin/activate  # On Windows use: .venv\Scripts\activate
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
```

This installs the [`mpxj`](https://www.mpxj.org/) parser used by the CLI to process Microsoft Project files. Reactivate the environment (`source .venv/bin/activate`) whenever you return to the project.

---

## 2. Quick start

1. **Clone the repository** (or download the ZIP).
2. **Open a terminal** in the project root.
3. **Start the preprocessing server** (serves the web UI and cleans uploads):

   ```bash
   python server.py --port 8000
   ```

   The server streams and cleans uploaded workbooks before handing them to the front-end. Large workbooks may take a few seconds to process while hierarchy-based dependencies are restored.

4. **Open the app** in your preferred browser at `http://localhost:8000`.
5. **Load data**:
   - Click **“Load sample data”** to experiment with the bundled JSON schedule, or
   - Use **“Upload Excel (.xlsx)”** and select `CPS_Rules_comm_LPA_Readiness.xlsx` (or your own workbook).

The visualization updates instantly after a successful upload—no rebuild steps are required.

---

## 3. Preparing your Excel workbook

The first sheet of the workbook is parsed. A header row is required. The following columns are recognized (case-insensitive, additional columns are ignored):

| Column | Required? | Description |
| --- | --- | --- |
| `TaskId` | ✅ | Unique identifier used to link dependencies. Stored as text so leading zeros are preserved. |
| `Task Name` | ✅ | Human-friendly task or milestone name. |
| `Task Level` | ✅ | Hierarchical level label (e.g., `L2`, `L3`, …). Lower numbers represent higher-level milestones. |
| `Responsible Sub-team` | Optional | Displayed in the task list for quick ownership review. |
| `Responsible Function` | Optional | Additional metadata shown in the task detail drawer. |
| `Predecessors IDs` | Optional | Comma/semicolon/new-line separated list of predecessor task IDs. Each creates a directed edge in the graph. |
| `Successors IDs` | Optional | Not required for the graph, but displayed in the task detail drawer for reference. |
| `Dependency Type` | Optional | Comma-separated list that matches the predecessor order. Supported values: `FS`, `SS`, `SF`, `FF`. Defaults to `FS` when omitted. |
| `SCOPE` | Optional | Used to group tasks in the collapsible tree. Defaults to `Unscoped`. |
| `Learning_Plan` | Optional | Boolean flag (`TRUE`/`FALSE`, `Yes`/`No`, `1`/`0`). Highlights the task in the detail drawer. |
| `Critical` | Optional | Boolean flag. Critical tasks receive a red badge in the detail drawer. |

> ℹ️ The bundled workbook `CPS_Rules_comm_LPA_Readiness.xlsx` matches this layout. You can duplicate the sheet, add a `Dependency Type` column, and tailor it for your programme without modifying the application code.

---

## 4. Using the visualizer

1. **Levels & scopes panel** – Tasks are grouped by `Task Level` (e.g., L2 milestones) and then by `SCOPE`. Each level/scope pair is collapsible. Checking a task reveals it in the network. Leaving it unchecked keeps the milestone collapsed.
2. **Level toggles** – Pills above the hierarchy toggle an entire level. When only some tasks of a level are displayed, the toggle shows a partial (indeterminate) state.
3. **Graph controls**:
   - **Fit to screen** recenters and scales the layout.
   - **Clear selection** empties the detail drawer.
   - **Auto-include context** adds the direct predecessors and successors of selected tasks so that dependencies are never orphaned.
4. **Task detail drawer** – Click a node to surface its metadata, predecessor/successor IDs, and status badges (critical, learning plan).
5. **Dependency cleaning** – When an Excel file is uploaded the app pauses briefly while the Python pipeline inspects indentation and fills in missing predecessor/successor links.
6. **Dependency table** – Lists every predecessor → successor relationship for the currently visible tasks, including dependency type.

Tip: For complex reviews, start with only L2 milestones selected, then progressively expand lower levels or specific scopes while discussing with stakeholders.

---

## 5. Customisation tips

- **Change the colour palette** – Update `levelPalette` near the top of [`webapp/app.js`](webapp/app.js) to align with your brand or portfolio.
- **Default selection logic** – By default, only the lowest-numbered level is selected after upload (typically your top milestones). Adjust the logic in `prepareDataset` if your organisation uses a different convention.
- **Alternative hosting** – Any static web host (GitHub Pages, Netlify, Azure Static Web Apps, S3) can serve the contents of `webapp/`. Upload the folder as-is—no build pipeline is required.

---

## 6. Troubleshooting

| Symptom | Resolution |
| --- | --- |
| The file uploads but no graph is displayed. | Ensure the sheet has a header row, that `TaskId` values are unique, and that at least one task is selected in the Levels panel. |
| Dependencies show `External task`. | The predecessor ID exists in the dependency list but not in the uploaded sheet. Either add the missing task or correct the ID. |
| Dependency types all read `FS`. | Add a `Dependency Type` column with comma-separated values that mirror the `Predecessors IDs` sequence for each row (e.g., `FS,SS`). |

For further enhancements, fork the repository and tailor the UI/logic in `webapp/app.js`—no backend code is involved.

---

## 7. MPP extraction & Python critical path calculator

This repository now bundles a lightweight command-line application that can extract task data from a Microsoft Project (`.mpp`) schedule and reproduce the critical path schedule calculation entirely in Python. The tool is split into two parts:

- **`convert`** – reads an `.mpp` file (such as the included `Generic CPS for I-O.mpp`) and emits a normalized CSV file with durations, dependencies, constraints, and baseline dates.
- **`calculate`** – consumes the CSV output (or any CSV that follows the same column layout) and computes earliest/latest start and finish dates, task float, and the overall programme duration using a simple working-day calendar.

### 7.1. Prerequisites

1. Python 3.10 or newer.
2. The [`mpxj`](https://www.mpxj.org/) package (used to parse `.mpp` files). Install it via the project requirements:

   ```bash
   python -m pip install -r requirements.txt
   ```

   > ℹ️ Only the `convert` command requires `mpxj`. The calculator works with CSV files alone.

### 7.2. Command-line usage

The CLI is exposed through `python -m cps_tool.cli`. Use `--help` at any time to discover available flags.

```bash
# Export the bundled sample schedule to CSV
python -m cps_tool.cli convert 'Generic CPS for I-O.mpp' output/tasks.csv

# Recalculate the critical path using an 08:00–17:00 workday (Mon–Fri)
python -m cps_tool.cli calculate output/tasks.csv \
  --project-start 2023-01-02T08:00 \
  --output output/cps_schedule.csv
```

Both commands print a short summary to stdout. When `--output` is provided the calculator writes a CSV containing earliest/latest dates, float, and a critical-path flag for every task. If dependency cycles are discovered they are reported during execution, the blocking dependency is removed automatically, and a cycle-adjusted CPS CSV is written alongside the summary.

### 7.3. Embedding as a Python module

All functionality is also available programmatically:

```python
from datetime import datetime

from cps_tool import (
    WorkCalendar,
    calculate_schedule,
    convert_mpp_to_csv,
    load_tasks_from_csv,
)

# Step 1 – extract tasks from Microsoft Project (optional if you already have CSV)
csv_path = convert_mpp_to_csv("Generic CPS for I-O.mpp", "output/tasks.csv")

# Step 2 – load tasks and run the calculator
tasks = load_tasks_from_csv(csv_path)
calendar = WorkCalendar()  # defaults to 08:00–17:00, Monday–Friday
result = calculate_schedule(tasks, project_start=datetime(2023, 1, 2, 8, 0), calendar=calendar)

print("Project finish:", result.project_finish.isoformat())
for task in result.critical_path():
    print("Critical:", task.spec.uid, task.spec.name)
if result.cycle_resolutions:
    print("Resolved cycles:")
    for resolution in result.cycle_resolutions:
        print("  ", resolution.formatted_cycle())
    if result.cycle_adjusted_csv:
        print("  Modified CPS written to", result.cycle_adjusted_csv)
```

The `ScheduleResult.to_rows()` helper returns dictionaries that can be written back to CSV (the CLI uses the same method). The calculator honours Finish-to-Start, Start-to-Start, Finish-to-Finish, and Start-to-Finish dependencies, plus common constraint types such as “Must Start On” and “Finish No Earlier Than”.

Call `ScheduleResult.to_csv(path)` to persist the calculated schedule directly to disk. The generated CSV lists each task with its earliest/latest start and finish timestamps, total float, and critical-path flag so the data can be analysed in other tools.
