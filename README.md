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

---

## 2. Quick start

1. **Clone the repository** (or download the ZIP).
2. **Open a terminal** in the project root.
3. **Start a lightweight static server**:

   ### macOS (built-in Python 3)
   ```bash
   python3 -m http.server 8000 --directory webapp
   ```

   ### Windows (PowerShell with Python installed)
   ```powershell
   py -3 -m http.server 8000 --directory webapp
   ```

   > 💡 If `py` is unavailable, try `python` instead. The command prints the URL you can open in the browser (`http://0.0.0.0:8000` or `http://localhost:8000`).

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
5. **Dependency table** – Lists every predecessor → successor relationship for the currently visible tasks, including dependency type.

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
