# Critical Path Schedule Visualizer

An interactive, browser-based tool for validating complex critical path schedules with non-technical stakeholders. Upload an Excel workbook that lists tasks, milestones, and dependencies to explore the network, expand milestones on demand, and highlight the relationships that drive your critical path.

## Features

- **Excel upload** – Drag in a schedule exported from your planning tool (the first worksheet is used).
- **Collapsible exploration** – Tasks stay grouped by level and scope until you decide to reveal them.
- **Interactive dependency graph** – Visualize predecessor/successor relationships, auto-include nearby context, and highlight task metadata.
- **Task dependency explorer** – Filter by SIMPL phase, Commercial/Technical Lego Block, and scope definition to inspect DAGs derived from `amy_new_cps_only_tasks.xlsx`.
- **Dependency typing** – Add a `Dependency Type` column (FS/SS/SF/FF) to distinguish relationships; otherwise Finish-to-Start (FS) is assumed.
- **Stakeholder-friendly UI** – Styled for projection or screen sharing; no coding experience required.
- **Download cleaned workbook** – Retrieve the Excel file produced by the preprocessing service for archival or further review.

A lightweight sample dataset is included in [`webapp/data/sample_schedule.json`](webapp/data/sample_schedule.json) and the repository ships with the stakeholder workbook `CPS_Rules_comm_LPA_Readiness.xlsx` for real-world testing.

---

## 1. Prerequisites

No runtime installation is required beyond a modern web browser. To keep setup lightweight on macOS or Windows, use the built-in Python interpreter (already available on macOS and easily installed via the Microsoft Store on Windows) to serve the static files. Any alternative static HTTP server (Node, Ruby, Go, etc.) works just as well.

---

## 2. Quick start

1. **Clone the repository** (or download the ZIP).
2. **Open a terminal** in the project root.
3. **Start the preprocessing server** (serves the web UI and cleans uploads):

   - **macOS (Terminal):**

     ```bash
     python3 server.py --port 8000
     ```

   - **Windows (PowerShell or Command Prompt):**

     ```powershell
     py server.py --port 8000
     ```

   The server streams and cleans uploaded workbooks before handing them to the front-end. Large workbooks may take a few seconds to process while hierarchy-based dependencies are restored.

4. **Open the app** in your preferred browser at `http://localhost:8000`.
5. **Choose an experience**:
   - Visit `/index.html` for the original schedule explorer (upload XLSX, expand milestones, inspect dependencies).
   - Visit `/dag.html` for the CPS Task Dependency Explorer that renders the DAG sourced from `amy_new_cps_only_tasks.xlsx`.
6. **Load data** in the schedule explorer (optional):
   - Click **“Load sample data”** to experiment with the bundled JSON schedule, or
   - Use **“Upload Excel (.xlsx)”** and select `CPS_Rules_comm_LPA_Readiness.xlsx` (or your own workbook).

The visualization updates instantly after a successful upload—no rebuild steps are required.

---

## 3. CPS Task Dependency Explorer

The dedicated DAG view at [`/dag.html`](webapp/dag.html) ships with the processed dataset [`webapp/data/amy_new_cps_tasks.json`](webapp/data/amy_new_cps_tasks.json), a direct export of `amy_new_cps_only_tasks.xlsx`. It mirrors the styling and interaction patterns of the original app while tailoring the workflow to SIMPL phase and Lego Block analysis:

1. **Select a SIMPL phase** – The first dropdown filters rows by the `Current SIMPL Phase (visibility)` column.
2. **Pick a Commercial/Technical Lego Block** – The second dropdown narrows the tasks to a single Lego Block within the chosen phase.
3. **Choose a scope definition** – The final dropdown isolates a `Scope Definition (Arnab)` value. The graph then renders a directed acyclic graph (DAG) containing the scoped tasks plus their immediate predecessors and successors.

Key behaviours:

- **Cross-block highlighting** – Any dependency node whose `Commercial/Technical Lego Block` differs from the active selection is outlined in orange to flag ownership or scope gaps.
- **Context preservation** – Related nodes remain visible (with a softer colour palette) so upstream/downstream impacts stay in view.
- **Task definition card** – Selecting a node surfaces the full row details (team, responsible person, timing, scope metadata, etc.).
- **Graph controls** – Use **Fit to screen** to re-centre the layout and **Reset filters** to start another exploration path.

To substitute your own CPS dataset, replace `amy_new_cps_only_tasks.xlsx`, regenerate the JSON via the helper script of your choice, and refresh the page. As long as the column headers remain unchanged the explorer will pick up the new data automatically.

---

## 4. Preparing your Excel workbook

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

## 5. Using the visualizer

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

## 6. Customisation tips

- **Change the colour palette** – Update `levelPalette` near the top of [`webapp/app.js`](webapp/app.js) to align with your brand or portfolio.
- **Default selection logic** – By default, only the lowest-numbered level is selected after upload (typically your top milestones). Adjust the logic in `prepareDataset` if your organisation uses a different convention.
- **Alternative hosting** – Any static web host (GitHub Pages, Netlify, Azure Static Web Apps, S3) can serve the contents of `webapp/`. Upload the folder as-is—no build pipeline is required.

---

## 7. Troubleshooting

| Symptom | Resolution |
| --- | --- |
| The file uploads but no graph is displayed. | Ensure the sheet has a header row, that `TaskId` values are unique, and that at least one task is selected in the Levels panel. |
| Dependencies show `External task`. | The predecessor ID exists in the dependency list but not in the uploaded sheet. Either add the missing task or correct the ID. |
| Dependency types all read `FS`. | Add a `Dependency Type` column with comma-separated values that mirror the `Predecessors IDs` sequence for each row (e.g., `FS,SS`). |

For further enhancements, fork the repository and tailor the UI/logic in `webapp/app.js`—no backend code is involved.

---

## 8. Running the CPM scheduler

In addition to the interactive visualizer, the repository ships with `cps_scheduler.py`, a Critical Path Method (CPM) engine that reads the level-4 rules workbook and produces normalized CSV/XLSX schedule reports. The script understands Finish-to-Start (FS), Start-to-Start (SS), Finish-to-Finish (FF), and Start-to-Finish (SF) dependencies, including positive and negative lags expressed in days (`d`), weeks (`w`), or months (`mo`).

### Command-line usage

Run the scheduler from the project root with your desired project start date and output paths:

```bash
python cps_scheduler.py \
  --input cps_rules_level_4.xlsx \
  --project-start 2024-01-01T08:00 \
  --output-csv schedule_output.csv \
  --output-xlsx schedule_output.xlsx \
  --ignore-missing
```

- **`--project-start`** – ISO-8601 timestamp. Task start times align to the next working instant (8h/day, 5d/week, 20d/month).
- **`--ignore-missing`** – Optional. Suppresses errors for predecessors that reference IDs absent from the workbook; missing links are listed in the console.
- **`--effective-levels`** – Optional. Comma-separated Task Level labels that apply the “Level-n” duration rule (e.g., `--effective-levels L2,L3,L4`).
- **`--sheet`** – Optional. Choose a different worksheet name if your source workbook stores tasks on a sheet other than `Sheet1`.

Outputs include normalized durations (in hours), computed early/late dates, slack values, and ISO-formatted start/finish timestamps based on the working calendar.

### Notebook automation

For a guided workflow, open [`notebooks/run_cps_scheduler.ipynb`](notebooks/run_cps_scheduler.ipynb). The notebook exposes editable parameters for the workbook path, project start date, and output locations, then executes the scheduler and prints a preview of the resulting CSV. It is ideal for analysts who prefer reproducible, shareable runs without remembering command-line flags.
