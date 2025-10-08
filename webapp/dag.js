import { buildTaskIndex, normalizeRecords, parseLinkedIds } from "./dag-data.js";

const EMPTY_TOKEN = "__EMPTY__";

const dom = {
  fileInput: document.getElementById("dag-file-input"),
  loadDefault: document.getElementById("dag-load-default"),
  download: document.getElementById("dag-download-json"),
  processingNotice: document.getElementById("dag-processing-notice"),
  phase: document.getElementById("phase-filter"),
  lego: document.getElementById("lego-filter"),
  scope: document.getElementById("scope-filter"),
  fit: document.getElementById("dag-fit"),
  reset: document.getElementById("dag-reset"),
  graph: document.getElementById("dag-cy"),
  graphEmpty: document.getElementById("graph-empty"),
  taskCard: document.getElementById("task-card"),
};

const detailFields = [
  ["Task Name", "Task Name"],
  ["TaskID", "Task ID"],
  ["Sub Team", "Sub Team"],
  ["Resp Function", "Responsible Function"],
  ["Resp Person", "Responsible Person"],
  ["PACE", "PACE"],
  ["LP", "LP"],
  ["Duration", "Duration"],
  ["Start", "Start"],
  ["Finish", "Finish"],
  ["Predecessors", "Predecessors"],
  ["Successors", "Successors"],
  ["Scope Definition (Arnab)", "Scope Definition (Arnab)"],
  ["Current SIMPL Phase (visibility)", "Current SIMPL Phase"],
  ["VoS (visibility)", "VoS"],
  ["Commercial/Technical Lego Block", "Commercial/Technical Lego Block"],
  ["Lego Block (Arnab)", "Lego Block (Arnab)"],
  ["Body of Evidence (both - Arnab - 1)", "Body of Evidence"],
  ["Bucket of Work (visibility - old)", "Bucket of Work"],
];

let rawRecords = [];
let tasks = [];
let tasksById = new Map();
let cyInstance = null;
let selectedTaskId = null;
let downloadUrl = null;
let downloadFilename = "cps_tasks.json";

init();

async function init() {
  attachEventListeners();
  updateDownloadLink();
  await loadDefaultDataset(false);
}

async function loadDefaultDataset(showSpinner = true) {
  try {
    if (showSpinner) {
      setProcessingState("Loading default dataset…", true);
    }
    const response = await fetch("data/amy_new_cps_tasks.json");
    if (!response.ok) {
      throw new Error(`Failed to load dataset (status ${response.status})`);
    }
    const records = await response.json();
    setDataset(records, { filename: "amy_new_cps_tasks.json" });
  } catch (error) {
    console.error(error);
    showGraphMessage("Unable to load the CPS task data. Refresh the page to try again.");
  } finally {
    if (showSpinner) {
      setProcessingState("", false);
    }
  }
}

function setDataset(records, options = {}) {
  rawRecords = normalizeRecords(records);
  const { tasks: nextTasks, tasksById: nextMap } = buildTaskIndex(rawRecords);
  tasks = nextTasks;
  tasksById = nextMap;
  downloadFilename = options.filename || "cps_tasks.json";
  populatePhaseOptions();
  resetFilters();
  updateDownloadLink();
}

async function handleFileUpload(event) {
  const input = event.target;
  const file = input && input.files && input.files[0] ? input.files[0] : null;
  if (!file) {
    return;
  }

  const displayName = file.name || "Uploaded workbook";
  setProcessingState(`Processing ${displayName}…`, true);

  try {
    const formData = new FormData();
    formData.append("file", file);
    const response = await fetch("/api/dag/upload", {
      method: "POST",
      body: formData,
    });

    if (!response.ok) {
      const errorText = await response.text();
      const message = errorText || `Upload failed with status ${response.status}`;
      throw new Error(message);
    }

    const payload = await response.json();
    const records = Array.isArray(payload.records) ? payload.records : [];
    const safeStem = sanitizeFilenameStem(displayName.replace(/\.xlsx?$/i, ""));
    const downloadName = `${safeStem || "cps_tasks"}.json`;
    setDataset(records, { filename: downloadName });
  } catch (error) {
    console.error(error);
    showGraphMessage(
      "Unable to process the uploaded workbook. Please verify the format and try again.",
    );
  } finally {
    if (input) {
      input.value = "";
    }
    setProcessingState("", false);
  }
}

function setProcessingState(message, active) {
  const notice = dom.processingNotice;
  const text = message || "Processing task dependency workbook…";
  if (notice) {
    if (active) {
      notice.textContent = text;
      notice.hidden = false;
    } else {
      notice.hidden = true;
    }
  }
  if (dom.fileInput) {
    dom.fileInput.disabled = active;
  }
  if (dom.loadDefault) {
    dom.loadDefault.disabled = active;
  }
}

function updateDownloadLink() {
  if (!dom.download) {
    return;
  }

  if (downloadUrl) {
    URL.revokeObjectURL(downloadUrl);
    downloadUrl = null;
  }

  if (!rawRecords.length) {
    dom.download.href = "";
    dom.download.removeAttribute("download");
    dom.download.setAttribute("aria-disabled", "true");
    dom.download.classList.add("disabled");
    return;
  }

  const blob = new Blob([JSON.stringify(rawRecords, null, 2)], {
    type: "application/json",
  });
  downloadUrl = URL.createObjectURL(blob);
  dom.download.href = downloadUrl;
  dom.download.download = downloadFilename;
  dom.download.setAttribute("aria-disabled", "false");
  dom.download.classList.remove("disabled");
}

function sanitizeFilenameStem(value) {
  if (!value) {
    return "";
  }
  return value.replace(/[^A-Za-z0-9_.-]+/g, "_");
}

function populatePhaseOptions() {
  const options = new Set();
  for (const task of tasks) {
    if (task.phase) {
      options.add(task.phase);
    }
  }

  const sorted = Array.from(options).sort(localeCompare);
  const placeholder = '<option value="" selected>Select a phase…</option>';
  dom.phase.innerHTML = placeholder + sorted.map((value) => buildOption(value)).join("");
  dom.phase.disabled = sorted.length === 0;
}

function updateLegoOptions() {
  const phaseValue = getSelectedValue(dom.phase);
  const options = new Set();
  for (const task of tasks) {
    if (phaseValue != null && task.phase !== phaseValue) {
      continue;
    }
    if (task.lego) {
      options.add(task.lego);
    }
  }
  const sorted = Array.from(options).sort(localeCompare);
  const placeholder = '<option value="" selected>Select a lego block…</option>';
  dom.lego.innerHTML = placeholder + sorted.map((value) => buildOption(value)).join("");
  dom.lego.disabled = phaseValue == null || sorted.length === 0;
  dom.lego.value = "";
  dom.scope.innerHTML = '<option value="" selected>Select a scope…</option>';
  dom.scope.disabled = true;
}

function updateScopeOptions() {
  const phaseValue = getSelectedValue(dom.phase);
  const legoValue = getSelectedValue(dom.lego);
  if (legoValue == null) {
    dom.scope.innerHTML = '<option value="" selected>Select a scope…</option>';
    dom.scope.disabled = true;
    return;
  }
  const options = new Set();
  for (const task of tasks) {
    if (phaseValue != null && task.phase !== phaseValue) {
      continue;
    }
    if (legoValue != null && task.lego !== legoValue) {
      continue;
    }
    options.add(task.scope || "");
  }
  const sorted = Array.from(options).sort(localeCompare);
  const placeholder = '<option value="" selected>Select a scope…</option>';
  dom.scope.innerHTML = placeholder + sorted.map((value) => buildOption(value)).join("");
  dom.scope.disabled = sorted.length === 0;
  dom.scope.value = "";
}

function attachEventListeners() {
  if (dom.fileInput) {
    dom.fileInput.addEventListener("change", handleFileUpload);
  }

  if (dom.loadDefault) {
    dom.loadDefault.addEventListener("click", () => {
      loadDefaultDataset();
    });
  }

  dom.phase.addEventListener("change", () => {
    updateLegoOptions();
    showGraphMessage("Select a Commercial/Technical Lego Block to continue.");
    clearSelection();
    updateGraph();
  });

  dom.lego.addEventListener("change", () => {
    updateScopeOptions();
    showGraphMessage("Select a scope definition to visualize the graph.");
    clearSelection();
    updateGraph();
  });

  dom.scope.addEventListener("change", () => {
    clearSelection();
    updateGraph();
  });

  dom.fit.addEventListener("click", () => {
    if (cyInstance) {
      cyInstance.fit(undefined, 60);
    }
  });

  dom.reset.addEventListener("click", () => {
    resetFilters();
  });
}

function resetFilters() {
  dom.phase.value = "";
  updateLegoOptions();
  dom.lego.value = "";
  dom.lego.disabled = true;
  dom.scope.innerHTML = '<option value="" disabled selected>Select a scope…</option>';
  dom.scope.disabled = true;
  clearGraph();
  showGraphMessage("Select a SIMPL phase, Commercial/Technical Lego Block, and scope definition to explore the dependency graph.");
  clearSelection();
}

function updateGraph() {
  const phaseValue = getSelectedValue(dom.phase);
  const legoValue = getSelectedValue(dom.lego);
  const scopeValue = getSelectedValue(dom.scope);

  if (phaseValue == null) {
    clearGraph();
    showGraphMessage("Select a SIMPL phase to begin exploring the graph.");
    return;
  }

  if (legoValue == null) {
    clearGraph();
    showGraphMessage("Select a Commercial/Technical Lego Block to continue.");
    return;
  }

  if (scopeValue == null) {
    clearGraph();
    showGraphMessage("Select a scope definition to visualize the graph.");
    return;
  }

  const focusTasks = tasks.filter((task) => {
    if (phaseValue != null && task.phase !== phaseValue) {
      return false;
    }
    if (legoValue != null && task.lego !== legoValue) {
      return false;
    }
    if (scopeValue != null && task.scope !== scopeValue) {
      return false;
    }
    return true;
  });

  if (focusTasks.length === 0) {
    clearGraph();
    showGraphMessage("No tasks match the selected filters.");
    return;
  }

  ensureCy();

  const focusIds = new Set(focusTasks.map((task) => task.id));
  const visited = new Set(focusIds);
  const queue = Array.from(focusIds);

  while (queue.length > 0) {
    const currentId = queue.shift();
    const task = tasksById.get(currentId);
    if (!task) {
      continue;
    }
    const neighbors = [...task.predecessors, ...task.successors];
    for (const neighbor of neighbors) {
      if (!tasksById.has(neighbor) || visited.has(neighbor)) {
        continue;
      }
      visited.add(neighbor);
      queue.push(neighbor);
    }
  }

  const elements = [];
  const focusEdges = new Set();

  for (const id of visited) {
    const task = tasksById.get(id);
    if (!task) {
      continue;
    }
    const classes = [];
    if (focusIds.has(id)) {
      classes.push("focus");
    } else {
      classes.push("context");
    }
    if (task.lego !== legoValue) {
      classes.push("lego-mismatch");
    }

    elements.push({
      group: "nodes",
      data: {
        id: task.id,
        label: formatNodeLabel(task),
      },
      classes: classes.join(" "),
    });
  }

  const edges = new Set();
  for (const id of visited) {
    const task = tasksById.get(id);
    if (!task) {
      continue;
    }
    for (const predecessor of task.predecessors) {
      if (!visited.has(predecessor)) {
        continue;
      }
      const key = `${predecessor}->${id}`;
      if (edges.has(key)) {
        continue;
      }
      edges.add(key);
      const isFocusEdge = focusIds.has(predecessor) && focusIds.has(id);
      if (isFocusEdge) {
        focusEdges.add(key);
      }
      elements.push({
        group: "edges",
        data: {
          id: key,
          source: predecessor,
          target: id,
        },
        classes: isFocusEdge ? "focus-edge" : "",
      });
    }
    for (const successor of task.successors) {
      if (!visited.has(successor)) {
        continue;
      }
      const key = `${id}->${successor}`;
      if (edges.has(key)) {
        continue;
      }
      edges.add(key);
      const isFocusEdge = focusIds.has(id) && focusIds.has(successor);
      if (isFocusEdge) {
        focusEdges.add(key);
      }
      elements.push({
        group: "edges",
        data: {
          id: key,
          source: id,
          target: successor,
        },
        classes: isFocusEdge ? "focus-edge" : "",
      });
    }
  }

  cyInstance.elements().remove();
  cyInstance.add(elements);

  const roots = Array.from(focusIds).filter((id) => {
    const predecessors = tasksById.get(id)?.predecessors || [];
    return predecessors.filter((pred) => visited.has(pred)).length === 0;
  });

  const layout = cyInstance.layout({
    name: "breadthfirst",
    directed: true,
    padding: 40,
    spacingFactor: 1.2,
    roots: roots.length ? roots : undefined,
  });
  layout.run();
  cyInstance.fit(undefined, 80);

  if (selectedTaskId && visited.has(selectedTaskId)) {
    cyInstance.$id(selectedTaskId).select();
  } else {
    clearSelection();
  }

  dom.graphEmpty.hidden = true;
}

function ensureCy() {
  if (cyInstance) {
    return;
  }

  cyInstance = cytoscape({
    container: dom.graph,
    elements: [],
    wheelSensitivity: 0.2,
    style: [
      {
        selector: "core",
        style: {
          "active-bg-color": "#bfdbfe",
          "selection-box-color": "#2563eb",
        },
      },
      {
        selector: "node",
        style: {
          "background-color": "#94a3b8",
          "border-color": "#94a3b8",
          "border-width": 2,
          color: "#0f172a",
          "font-size": "10px",
          "font-weight": 500,
          label: "data(label)",
          padding: "8px",
          shape: "round-rectangle",
          "text-max-width": 180,
          "text-outline-width": 0,
          "text-valign": "center",
          "text-halign": "center",
          "text-wrap": "wrap",
        },
      },
      {
        selector: "node.focus",
        style: {
          "background-color": "#2563eb",
          "border-color": "#1d4ed8",
          color: "#ffffff",
          "font-weight": 700,
          "shadow-blur": 18,
          "shadow-color": "rgba(37, 99, 235, 0.28)",
          "shadow-offset-x": 0,
          "shadow-offset-y": 4,
        },
      },
      {
        selector: "node.context",
        style: {
          "background-color": "#dbeafe",
          "border-color": "#60a5fa",
          color: "#1e293b",
        },
      },
      {
        selector: "node.lego-mismatch",
        style: {
          "background-color": "#fb923c",
          "border-color": "#f97316",
          color: "#0f172a",
          "font-weight": 600,
        },
      },
      {
        selector: "node.focus.lego-mismatch",
        style: {
          "background-color": "#f43f5e",
          "border-color": "#e11d48",
          color: "#ffffff",
        },
      },
      {
        selector: "edge",
        style: {
          "curve-style": "bezier",
          width: 2,
          "line-color": "#94a3b8",
          "target-arrow-shape": "triangle",
          "target-arrow-color": "#94a3b8",
        },
      },
      {
        selector: "edge.focus-edge",
        style: {
          width: 3,
          "line-color": "#2563eb",
          "target-arrow-color": "#2563eb",
        },
      },
      {
        selector: "node:selected",
        style: {
          "border-color": "#facc15",
          "border-width": 4,
          "background-color": "#fbbf24",
          color: "#0f172a",
        },
      },
    ],
  });

  cyInstance.on("tap", (event) => {
    if (event.target === cyInstance) {
      cyInstance.elements().unselect();
    }
  });

  cyInstance.on("select", "node", (event) => {
    selectedTaskId = event.target.id();
    const task = tasksById.get(selectedTaskId) || null;
    renderTaskCard(task);
  });

  cyInstance.on("unselect", "node", () => {
    if (cyInstance.$("node:selected").length === 0) {
      selectedTaskId = null;
      renderTaskCard(null);
    }
  });
}

function formatNodeLabel(task) {
  const lines = [task.id];
  if (task.name) {
    lines.push(task.name);
  }
  if (task.raw["Resp Function"]) {
    lines.push(task.raw["Resp Function"]);
  }
  return lines.join("\n");
}

function localeCompare(a, b) {
  return a.localeCompare(b, undefined, { sensitivity: "base" });
}

function buildOption(value) {
  const optionValue = value === "" ? EMPTY_TOKEN : value;
  const label = value === "" ? "Unspecified" : value;
  return `<option value="${escapeHtml(optionValue)}">${escapeHtml(label)}</option>`;
}

function getSelectedValue(select) {
  const raw = select.value;
  if (!raw) {
    return null;
  }
  if (raw === EMPTY_TOKEN) {
    return "";
  }
  return raw;
}

function escapeHtml(value) {
  return value
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

function renderTaskCard(task) {
  if (!task) {
    dom.taskCard.classList.add("empty");
    dom.taskCard.innerHTML = "<p>Select a node from the graph to view its full task definition.</p>";
    return;
  }

  dom.taskCard.classList.remove("empty");
  const heading = document.createElement("h3");
  heading.textContent = `${task.id} · ${task.name}`;

  const list = document.createElement("div");
  list.className = "task-details-list";

  for (const [field, label] of detailFields) {
    const value = task.raw[field] || "—";
    const row = document.createElement("div");
    row.className = "task-details-row";

    const keyEl = document.createElement("div");
    keyEl.className = "task-details-key";
    keyEl.textContent = label;

    const valueEl = document.createElement("div");
    valueEl.className = "task-details-value";
    valueEl.textContent = value;

    row.append(keyEl, valueEl);
    list.append(row);
  }

  dom.taskCard.innerHTML = "";
  dom.taskCard.append(heading, list);
}

function clearGraph() {
  if (cyInstance) {
    cyInstance.elements().remove();
  }
  dom.graphEmpty.hidden = false;
}

function showGraphMessage(message) {
  if (dom.graphEmpty) {
    dom.graphEmpty.textContent = message;
    dom.graphEmpty.hidden = false;
  }
}

function clearSelection() {
  if (cyInstance) {
    cyInstance.elements().unselect();
  }
  selectedTaskId = null;
  renderTaskCard(null);
}

window.addEventListener("beforeunload", () => {
  if (downloadUrl) {
    URL.revokeObjectURL(downloadUrl);
  }
});
