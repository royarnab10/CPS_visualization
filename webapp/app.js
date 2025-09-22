const levelPalette = [
  '#38bdf8',
  '#a855f7',
  '#f97316',
  '#22c55e',
  '#facc15',
  '#ec4899',
  '#6366f1'
];

let dataset = null;
let selectedTaskIds = new Set();
let hierarchySelections = new Map();
let graphSelections = new Set();
let cyInstance = null;
let showFullGraph = false;
let selectedTask = null;

const dom = {
  fileInput: document.getElementById('file-input'),
  loadSample: document.getElementById('load-sample'),
  hierarchy: document.getElementById('hierarchy'),
  levelControls: document.getElementById('level-controls'),
  taskDetails: document.getElementById('task-details'),
  dependencyTable: document.getElementById('dependency-table'),
  dependencyPanel: document.getElementById('dependency-panel'),
  dependencySummaryLabel: document.getElementById('dependency-summary-label'),
  dependencySummaryCount: document.getElementById('dependency-summary-count'),
  dependencyFocus: document.getElementById('dependency-focus'),
  fitGraph: document.getElementById('fit-graph'),
  resetSelection: document.getElementById('reset-selection'),
  legend: document.getElementById('legend'),
  showFullGraph: document.getElementById('show-full-graph')
};

dom.fileInput.addEventListener('change', handleFileInput);
dom.loadSample.addEventListener('click', loadSampleData);
dom.showFullGraph.addEventListener('change', () => {
  showFullGraph = dom.showFullGraph.checked;
  updateGraph();
});
dom.fitGraph.addEventListener('click', () => cyInstance && cyInstance.fit(undefined, 40));
dom.resetSelection.addEventListener('click', () => {
  if (cyInstance) {
    cyInstance.elements().unselect();
  }
  hierarchySelections.clear();
  graphSelections.clear();
  rebuildSelectedTaskIds();
  selectedTask = null;
  renderLevelControls();
  renderHierarchy();
  renderDependencies();
  updateGraph();
  renderTaskDetails();
  if (dom.dependencyPanel) {
    dom.dependencyPanel.open = false;
  }
});

async function handleFileInput(event) {
  const file = event.target.files?.[0];
  if (!file) return;
  const arrayBuffer = await file.arrayBuffer();
  const workbook = XLSX.read(arrayBuffer, { type: 'array' });
  const sheetName = workbook.SheetNames[0];
  const worksheet = workbook.Sheets[sheetName];
  const records = XLSX.utils.sheet_to_json(worksheet, { defval: '' });
  prepareDataset(records, file.name);
}

async function loadSampleData() {
  const response = await fetch('data/sample_schedule.json');
  const records = await response.json();
  prepareDataset(records, 'Sample schedule');
}

function prepareDataset(records, label) {
  if (!records || records.length === 0) {
    alert('The provided worksheet is empty.');
    return;
  }

  const tasks = records.map(normalizeRecord).filter(Boolean);
  const tasksById = new Map(tasks.map((task) => [task.id, task]));
  const dependencies = buildDependencies(tasks);
  const levels = Array.from(new Set(tasks.map((t) => t.level))).sort((a, b) => a - b);

  selectedTaskIds = new Set();
  hierarchySelections = new Map();
  graphSelections = new Set();
  selectedTask = null;

  dataset = {
    label,
    tasks,
    tasksById,
    dependencies,
    levels,
    modified: false
  };

  if (dom.dependencyPanel) {
    dom.dependencyPanel.open = false;
  }

  if (dom.showFullGraph) {
    dom.showFullGraph.checked = false;
  }
  showFullGraph = false;

  renderLegend();
  renderLevelControls();
  renderHierarchy();
  renderDependencies();
  updateGraph();
  renderTaskDetails();
}

function normalizeRecord(record) {
  const mapped = {};
  for (const [key, value] of Object.entries(record)) {
    if (!key) continue;
    mapped[key.trim()] = typeof value === 'string' ? value.trim() : value;
  }

  const id = mapped['TaskId'] || mapped['Task ID'] || mapped['ID'];
  const name = mapped['Task Name'] || mapped['Name'];
  if (!id || !name) {
    return null;
  }

  const levelLabel = mapped['Task Level'] || mapped['Level'] || '';
  const level = parseLevel(levelLabel);
  const predecessorsRaw = mapped['Predecessors IDs'] || mapped['Predecessor IDs'] || mapped['Predecessors'] || '';
  const successorsRaw = mapped['Successors IDs'] || mapped['Successor IDs'] || mapped['Successors'] || '';
  const dependencyTypeRaw = mapped['Dependency Type'] || mapped['Dependency Types'] || '';

  return {
    id: String(id),
    name,
    level,
    levelLabel: levelLabel || `L${level}`,
    team: mapped['Responsible Sub-team'] || mapped['Sub-team'] || '',
    function: mapped['Responsible Function'] || mapped['Function'] || '',
    scope: mapped['SCOPE'] || mapped['Scope'] || 'Unscoped',
    learningPlan: normalizeBoolean(mapped['Learning_Plan'] ?? mapped['Learning Plan'] ?? ''),
    critical: normalizeBoolean(mapped['Critical'] ?? mapped['Is Critical'] ?? ''),
    predecessors: splitIds(predecessorsRaw),
    successors: splitIds(successorsRaw),
    dependencyTypes: splitIds(dependencyTypeRaw).map((type) => (type ? type.toUpperCase() : 'FS')),
    raw: mapped
  };
}

function parseLevel(label) {
  if (typeof label === 'number') return label;
  const match = typeof label === 'string' ? label.match(/(\d+)/) : null;
  return match ? Number(match[1]) : 1;
}

function splitIds(value) {
  if (!value) return [];
  if (Array.isArray(value)) return value.map((v) => String(v).trim()).filter(Boolean);
  return String(value)
    .split(/[,\n;]/)
    .map((part) => part.trim())
    .filter(Boolean);
}

function normalizeBoolean(value) {
  if (typeof value === 'boolean') return value;
  const normalized = String(value || '').toLowerCase();
  if (!normalized) return false;
  return ['true', 'yes', 'y', '1'].includes(normalized);
}

function buildDependencies(tasks) {
  const index = new Map(tasks.map((t) => [t.id, t]));
  const dependencies = [];

  tasks.forEach((task) => {
    task.predecessors.forEach((predId, indexPosition) => {
      const type = task.dependencyTypes[indexPosition] || task.dependencyTypes[0] || 'FS';
      dependencies.push({
        source: String(predId),
        target: task.id,
        type: type.toUpperCase(),
        targetLevel: task.level,
        sourceLevel: index.get(predId)?.level ?? null
      });
    });
  });

  return dependencies;
}

function renderLegend() {
  if (!dataset) {
    dom.legend.innerHTML = '';
    return;
  }

  const uniqueLevels = dataset.levels;
  const items = uniqueLevels
    .map((level, idx) => {
      const color = levelColor(level, idx);
      return `<span class="legend-item"><span class="legend-swatch" style="background:${color}"></span>Level ${level}</span>`;
    })
    .join('');

  const context = `<span class="legend-item"><span class="legend-swatch" style="background:transparent;border:2px dashed rgba(148,163,184,0.8);"></span>Connected context</span>`;
  const dimmed = '<span class="legend-item"><span class="legend-swatch" style="background:rgba(148,163,184,0.4);border:1px solid rgba(148,163,184,0.6);"></span>Dimmed outside focus</span>';

  dom.legend.innerHTML = items + context + dimmed;
}

function renderLevelControls() {
  if (!dataset) {
    dom.levelControls.innerHTML = '';
    return;
  }

  const entries = Array.from(hierarchySelections.entries()).sort((a, b) => a[0] - b[0]);
  const networkEntries = Array.from(graphSelections.values());

  if (entries.length === 0 && networkEntries.length === 0) {
    dom.levelControls.innerHTML = '<p class="empty-state">Select a Level 2 milestone to begin, then expand follow-on work from the network.</p>';
    return;
  }

  const fragment = document.createDocumentFragment();

  if (entries.length > 0) {
    const hierarchyWrapper = document.createElement('div');
    hierarchyWrapper.className = 'selection-path';

    entries.forEach(([level, taskId]) => {
      const task = dataset.tasksById.get(taskId);
      if (!task) return;

      const chip = document.createElement('button');
      chip.type = 'button';
      chip.className = 'path-chip';
      chip.innerHTML = `<span>L${level}</span><strong>${task.name}</strong>`;
      chip.addEventListener('click', () => clearHierarchyFrom(level));
      hierarchyWrapper.appendChild(chip);
    });

    fragment.appendChild(hierarchyWrapper);
  }

  if (networkEntries.length > 0) {
    const networkWrapper = document.createElement('div');
    networkWrapper.className = 'selection-path';

    networkEntries
      .map((taskId) => ({ id: taskId, task: dataset.tasksById.get(taskId) }))
      .sort((a, b) => {
        const aName = a.task?.name || a.id;
        const bName = b.task?.name || b.id;
        return aName.localeCompare(bName);
      })
      .forEach(({ id, task }) => {
        const chip = document.createElement('button');
        chip.type = 'button';
        chip.className = 'path-chip';
        const label = task ? task.name : `External ${id}`;
        chip.innerHTML = `<span>Graph</span><strong>${label}</strong>`;
        chip.addEventListener('click', () => toggleGraphSelection(id));
        networkWrapper.appendChild(chip);
      });

    fragment.appendChild(networkWrapper);
  }

  dom.levelControls.innerHTML = '';
  dom.levelControls.appendChild(fragment);
}

function renderHierarchy() {
  if (!dataset) {
    dom.hierarchy.innerHTML = '';
    return;
  }

  const hierarchyLevels = getHierarchyLevels();
  const fragment = document.createDocumentFragment();

  hierarchyLevels.forEach((level, index) => {
    const column = document.createElement('section');
    column.className = 'hierarchy-column';

    const header = document.createElement('header');
    const title = document.createElement('h3');
    title.textContent = `Level ${level}`;
    const message = document.createElement('p');
    message.textContent =
      index === 0
        ? `Choose a Level ${level} task to begin exploring.`
        : 'Select a task to spotlight it in the network. Continue deeper by choosing nodes directly in the visualization.';
    header.append(title, message);
    column.appendChild(header);

    const tasksAtLevel = dataset.tasks.filter((task) => task.level === level);
    const previousLevel = index > 0 ? hierarchyLevels[index - 1] : null;
    const parentSelection = previousLevel ? hierarchySelections.get(previousLevel) : null;

    let directTasks = [];
    if (index === 0) {
      directTasks = tasksAtLevel;
    } else if (parentSelection) {
      directTasks = tasksAtLevel.filter((task) => task.predecessors.includes(parentSelection));
    }

    const directTaskIds = new Set(directTasks.map((task) => task.id));
    const relatedEntries = [];

    if (selectedTaskIds.size > 0) {
      tasksAtLevel.forEach((task) => {
        if (directTaskIds.has(task.id)) return;
        const dependsOnSelected = task.predecessors.some((pred) => selectedTaskIds.has(pred));
        const feedsSelected = task.successors.some((succ) => selectedTaskIds.has(succ));
        if (!dependsOnSelected && !feedsSelected) return;

        let relationship = '';
        if (dependsOnSelected && feedsSelected) {
          relationship = 'Connected to selected path';
        } else if (dependsOnSelected) {
          relationship = 'Depends on selected task';
        } else {
          relationship = 'Feeds selected task';
        }
        relatedEntries.push({ task, relationship });
      });
    }

    const hasDirect = directTasks.length > 0;
    const hasRelated = relatedEntries.length > 0;

    if (!hasDirect && !hasRelated) {
      if (index > 0 && !parentSelection && selectedTaskIds.size === 0) {
        return;
      }
      const empty = document.createElement('div');
      empty.className = 'empty-state soft';
      if (index === 0) {
        empty.textContent = 'No tasks available for this level.';
      } else if (parentSelection) {
        empty.textContent = 'No downstream tasks for this milestone.';
      } else if (selectedTaskIds.size > 0) {
        empty.textContent = 'No cross-level tasks connect to the current path.';
      } else {
        empty.textContent = 'Select a task in the previous level to reveal its dependencies.';
      }
      column.appendChild(empty);
      fragment.appendChild(column);
      return;
    }

    if (hasDirect) {
      const directGroup = document.createElement('div');
      directGroup.className = 'hierarchy-group';
      if (index > 0) {
        const parentTask = parentSelection ? dataset.tasksById.get(parentSelection) : null;
        const label = document.createElement('h4');
        label.className = 'hierarchy-group-title';
        label.textContent = parentTask
          ? `Direct successors of ${parentTask.name}`
          : 'Direct successors';
        directGroup.appendChild(label);
      }
      directGroup.appendChild(buildTaskList(directTasks, level));
      column.appendChild(directGroup);
    }

    if (hasRelated) {
      const relatedGroup = document.createElement('div');
      relatedGroup.className = 'hierarchy-group related';
      const label = document.createElement('h4');
      label.className = 'hierarchy-group-title';
      label.textContent = 'Cross-level connections';
      relatedGroup.appendChild(label);

      const relationshipByTask = new Map(
        relatedEntries.map(({ task, relationship }) => [task.id, relationship])
      );
      relatedGroup.appendChild(
        buildTaskList(
          relatedEntries.map((entry) => entry.task),
          level,
          (task) => ({ relationship: relationshipByTask.get(task.id) })
        )
      );
      column.appendChild(relatedGroup);
    }

    fragment.appendChild(column);
  });

  dom.hierarchy.innerHTML = '';
  dom.hierarchy.appendChild(fragment);
}

function buildTaskList(tasks, level, optionsResolver) {
  const list = document.createElement('div');
  list.className = 'hierarchy-task-list';

  tasks
    .slice()
    .sort((a, b) => a.name.localeCompare(b.name))
    .forEach((task) => {
      const options = typeof optionsResolver === 'function' ? optionsResolver(task) : undefined;
      const card = createTaskCard(task, level, options);
      list.appendChild(card);
    });

  return list;
}

function createTaskCard(task, level, options = {}) {
  const card = document.createElement('button');
  card.type = 'button';
  card.className = 'hierarchy-card';

  if (hierarchySelections.get(level) === task.id) {
    card.classList.add('selected');
  }
  if (graphSelections.has(task.id)) {
    card.classList.add('network-selected');
  }
  if (options?.relationship) {
    card.classList.add('related');
  }

  const title = document.createElement('span');
  title.className = 'title';
  title.textContent = task.name;

  const meta = document.createElement('span');
  meta.className = 'meta';
  meta.textContent = `#${task.id} · ${task.team || 'Unassigned'}`;

  card.append(title, meta);

  if (options?.relationship) {
    const badge = document.createElement('span');
    badge.className = 'relationship-badge';
    badge.textContent = options.relationship;
    card.appendChild(badge);
  }

  if (graphSelections.has(task.id) && hierarchySelections.get(level) !== task.id) {
    const focusBadge = document.createElement('span');
    focusBadge.className = 'relationship-badge network';
    focusBadge.textContent = 'Graph focus';
    card.appendChild(focusBadge);
  }

  card.addEventListener('click', () => handleHierarchySelection(level, task));

  return card;
}

function getHierarchyLevels() {
  if (!dataset) return [];
  const sorted = dataset.levels.slice().sort((a, b) => a - b);
  const levelTwoIndex = sorted.indexOf(2);
  const relevant = levelTwoIndex >= 0 ? sorted.slice(levelTwoIndex) : sorted;
  return relevant.slice(0, 2);
}

function handleHierarchySelection(level, task) {
  if (!dataset) return;

  const currentlySelected = hierarchySelections.get(level);
  if (currentlySelected === task.id) {
    clearHierarchyFrom(level);
    return;
  }

  if (graphSelections.size > 0) {
    graphSelections.clear();
  }

  hierarchySelections.set(level, task.id);

  const hierarchyLevels = getHierarchyLevels();
  const levelIndex = hierarchyLevels.indexOf(level);
  hierarchyLevels.slice(levelIndex + 1).forEach((lvl) => hierarchySelections.delete(lvl));

  rebuildSelectedTaskIds();
  selectedTask = task;
  renderLevelControls();
  renderHierarchy();
  renderDependencies();
  updateGraph();
  renderTaskDetails();
}

function clearHierarchyFrom(level) {
  if (!dataset) return;
  const hierarchyLevels = getHierarchyLevels();
  const levelIndex = hierarchyLevels.indexOf(level);
  if (levelIndex === -1) return;

  hierarchyLevels.slice(levelIndex).forEach((lvl) => hierarchySelections.delete(lvl));
  graphSelections.clear();
  rebuildSelectedTaskIds();

  if (selectedTask && !selectedTaskIds.has(selectedTask.id)) {
    selectedTask = null;
  }

  renderLevelControls();
  renderHierarchy();
  renderDependencies();
  updateGraph();
  renderTaskDetails();
}

function rebuildSelectedTaskIds() {
  selectedTaskIds = new Set([...hierarchySelections.values(), ...graphSelections.values()]);
}

function toggleGraphSelection(taskId) {
  if (isInHierarchy(taskId)) {
    renderLevelControls();
    renderHierarchy();
    renderDependencies();
    updateGraph();
    renderTaskDetails();
    return;
  }

  if (graphSelections.has(taskId)) {
    graphSelections.delete(taskId);
  } else {
    graphSelections.add(taskId);
  }
  rebuildSelectedTaskIds();

  if (selectedTask && selectedTask.id === taskId && !selectedTaskIds.has(taskId)) {
    selectedTask = null;
  }

  renderLevelControls();
  renderHierarchy();
  renderDependencies();
  updateGraph();
  renderTaskDetails();
}

function isInHierarchy(taskId) {
  for (const value of hierarchySelections.values()) {
    if (value === taskId) return true;
  }
  return false;
}

function updateGraph() {
  const container = document.getElementById('cy');
  if (!dataset || !container) return;

  const nodes = new Map();
  const edges = [];
  const hasExplicitFocus = Boolean(selectedTask);
  const focusIds = hasExplicitFocus ? new Set([selectedTask.id]) : new Set(selectedTaskIds);
  const visible = new Set();
  const highlightNodes = new Set(focusIds);
  const highlightEdges = new Set();
  const hasFocus = focusIds.size > 0;

  if (showFullGraph && !hasFocus) {
    dataset.tasks.forEach((task) => visible.add(task.id));
    dataset.dependencies.forEach((dep) => {
      visible.add(dep.source);
      visible.add(dep.target);
    });
  } else {
    focusIds.forEach((id) => visible.add(id));
    if (hasFocus) {
      dataset.dependencies.forEach((dep) => {
        if (focusIds.has(dep.source)) {
          visible.add(dep.target);
        }
        if (focusIds.has(dep.target)) {
          visible.add(dep.source);
        }
      });
    }
  }

  if (hasFocus) {
    dataset.dependencies.forEach((dep) => {
      if (focusIds.has(dep.source) || focusIds.has(dep.target)) {
        const edgeId = `${dep.source}->${dep.target}`;
        highlightEdges.add(edgeId);
        highlightNodes.add(dep.source);
        highlightNodes.add(dep.target);
      }
    });
  }

  const ensureNode = (taskId) => {
    if (nodes.has(taskId)) return;
    const task = dataset.tasksById.get(taskId);
    const isFocus = focusIds.has(taskId);
    const isHighlighted = highlightNodes.has(taskId);
    const isDimmed = showFullGraph && hasFocus && !isHighlighted;

    if (task) {
      nodes.set(taskId, {
        data: {
          id: taskId,
          label: formatNodeLabel(task.name, taskId),
          color: levelColor(task.level),
          level: task.level,
          scope: task.scope,
          isFocus,
          isNeighbor: !isFocus && isHighlighted,
          isDimmed
        }
      });
    } else {
      nodes.set(taskId, {
        data: {
          id: taskId,
          label: formatNodeLabel(`External ${taskId}`, taskId),
          color: '#94a3b8',
          level: 0,
          scope: 'External',
          isFocus: false,
          isNeighbor: isHighlighted,
          isDimmed,
          isExternal: true
        }
      });
    }
  };

  visible.forEach((taskId) => ensureNode(taskId));

  dataset.dependencies.forEach((dep) => {
    if (!visible.has(dep.source) || !visible.has(dep.target)) {
      return;
    }

    const edgeId = `${dep.source}->${dep.target}`;
    const isHighlighted = highlightEdges.has(edgeId);

    if (hasFocus && !isHighlighted) {
      return;
    }

    if (!showFullGraph && !hasFocus) {
      return;
    }

    ensureNode(dep.source);
    ensureNode(dep.target);

    edges.push({
      data: {
        id: edgeId,
        source: dep.source,
        target: dep.target,
        type: dep.type,
        isDimmed: showFullGraph && hasFocus && !isHighlighted
      }
    });
  });

  const elements = [...nodes.values(), ...edges];

  if (!cyInstance) {
    cyInstance = cytoscape({
      container,
      elements,
      layout: { name: 'cose', animate: false },
      style: [
        {
          selector: 'node',
          style: {
            'background-color': 'data(color)',
            'background-opacity': 0.92,
            'label': 'data(label)',
            'color': '#0f172a',
            'font-size': '12px',
            'font-weight': 600,
            'text-wrap': 'wrap',
            'text-max-width': 160,
            'text-valign': 'center',
            'text-halign': 'center',
            'shape': 'round-rectangle',
            'padding': '12px',
            'border-width': 3,
            'border-color': 'rgba(15,23,42,0.2)',
            'shadow-blur': 12,
            'shadow-color': 'rgba(15,23,42,0.25)',
            'shadow-offset-x': 0,
            'shadow-offset-y': 4,
            'text-outline-width': 2,
            'text-outline-color': 'rgba(255,255,255,0.85)'
          }
        },
        {
          selector: 'node[?isNeighbor]',
          style: {
            'border-style': 'dashed',
            'opacity': 0.75,
            'background-opacity': 0.5,
            'color': '#1f2937'
          }
        },
        {
          selector: 'node[?isDimmed]',
          style: {
            'opacity': 0.28,
            'background-opacity': 0.18,
            'color': '#475569',
            'text-opacity': 0.45,
            'border-color': 'rgba(148,163,184,0.35)',
            'shadow-blur': 0,
            'text-outline-opacity': 0
          }
        },
        {
          selector: 'node:selected',
          style: {
            'border-width': 4,
            'border-color': '#facc15',
            'shadow-color': 'rgba(250, 204, 21, 0.45)',
            'shadow-blur': 18
          }
        },
        {
          selector: 'edge',
          style: {
            'width': 3,
            'line-color': 'rgba(148, 163, 184, 0.6)',
            'target-arrow-color': 'rgba(71, 85, 105, 0.9)',
            'target-arrow-shape': 'triangle',
            'curve-style': 'unbundled-bezier',
            'label': 'data(type)',
            'font-size': '9px',
            'color': '#cbd5f5',
            'text-background-color': 'rgba(15,23,42,0.85)',
            'text-background-opacity': 1,
            'text-background-padding': 3,
            'text-background-shape': 'roundrectangle'
          }
        },
        {
          selector: 'edge[?isDimmed]',
          style: {
            'opacity': 0.2,
            'line-color': 'rgba(148,163,184,0.25)',
            'target-arrow-color': 'rgba(148,163,184,0.35)',
            'text-opacity': 0
          }
        },
        {
          selector: 'edge[type = "FF"]',
          style: {
            'line-style': 'dashed',
            'target-arrow-shape': 'tee'
          }
        },
        {
          selector: 'edge[type = "SS"]',
          style: {
            'line-style': 'dotted',
            'source-arrow-shape': 'circle',
            'source-arrow-color': 'rgba(71,85,105,0.9)',
            'arrow-scale': 1.2
          }
        },
        {
          selector: 'edge[type = "SF"]',
          style: {
            'line-style': 'dashed',
            'source-arrow-shape': 'triangle',
            'source-arrow-color': 'rgba(71,85,105,0.9)',
            'target-arrow-shape': 'tee'
          }
        }
      ]
    });

    cyInstance.on('tap', 'node', (evt) => {
      const node = evt.target;
      const nodeId = node.id();
      const task = dataset.tasksById.get(nodeId);
      selectedTask = task || { id: nodeId, placeholder: true };
      toggleGraphSelection(nodeId);
    });

    cyInstance.on('tap', 'edge', handleEdgeTap);

    cyInstance.on('tap', (evt) => {
      if (evt.target === cyInstance) {
        selectedTask = null;
        renderTaskDetails();
      }
    });
  } else {
    cyInstance.json({ elements });
  }

  cyInstance.layout({ name: 'cose', animate: false, padding: 60, nodeRepulsion: 9000 }).run();
}

function levelColor(level, idx) {
  const paletteIndex = typeof idx === 'number' ? idx % levelPalette.length : (level - 1) % levelPalette.length;
  return levelPalette[paletteIndex];
}

function formatNodeLabel(name, id) {
  const wrapped = wrapText(name);
  return `${wrapped}\n#${id}`;
}

function wrapText(text, maxLineLength = 18) {
  if (!text) return '';
  const words = String(text).split(/\s+/).filter(Boolean);
  if (words.length === 0) return '';

  const lines = [];
  let currentLine = '';

  words.forEach((word) => {
    const candidate = currentLine ? `${currentLine} ${word}` : word;
    if (candidate.length > maxLineLength && currentLine) {
      lines.push(currentLine);
      currentLine = word;
    } else {
      currentLine = candidate;
    }
  });

  if (currentLine) {
    lines.push(currentLine);
  }

  return lines.join('\n');
}

function renderTaskDetails() {
  if (!selectedTask) {
    dom.taskDetails.classList.add('empty');
    dom.taskDetails.innerHTML = '<p>Select a node in the network to see its attributes here.</p>';
    return;
  }

  dom.taskDetails.classList.remove('empty');

  if (selectedTask.placeholder) {
    dom.taskDetails.innerHTML = `
      <h3>External dependency</h3>
      <p>The task <strong>${selectedTask.id}</strong> exists in the dependency chain but is not present in the uploaded schedule.</p>
    `;
    return;
  }

  const task = dataset?.tasksById.get(selectedTask.id);
  if (!task) {
    dom.taskDetails.innerHTML = `
      <h3>Task unavailable</h3>
      <p>The selected task is no longer part of the working dataset.</p>
    `;
    return;
  }

  selectedTask = task;

  const form = buildTaskEditForm(task);
  dom.taskDetails.innerHTML = '';
  dom.taskDetails.appendChild(form);
}

function buildTaskEditForm(task) {
  const form = document.createElement('form');
  form.className = 'task-edit-form';
  form.noValidate = true;

  form.innerHTML = `
    <div class="task-form-header">
      <div>
        <h3>${escapeHtml(task.name)}</h3>
        <div class="task-form-meta">Task ID <strong>#${escapeHtml(task.id)}</strong> · Level ${escapeHtml(
          task.levelLabel
        )}</div>
      </div>
      <div class="task-form-actions">
        <button type="submit" class="save-button" data-role="save" disabled>Save changes</button>
        <button type="button" class="delete-button" data-role="delete">Delete task</button>
      </div>
    </div>
    <div class="task-form-grid">
      <label class="task-field">
        <span>Task name</span>
        <input name="taskName" type="text" value="${escapeHtml(task.name)}" />
      </label>
      <label class="task-field">
        <span>Responsible sub-team</span>
        <input name="responsibleTeam" type="text" value="${escapeHtml(task.team || '')}" />
      </label>
      <label class="task-field">
        <span>Responsible function</span>
        <input name="responsibleFunction" type="text" value="${escapeHtml(task.function || '')}" />
      </label>
      <label class="task-field">
        <span>Scope</span>
        <input name="scope" type="text" value="${escapeHtml(task.scope || '')}" />
      </label>
      <label class="task-field">
        <span>Level label</span>
        <input name="levelLabel" type="text" value="${escapeHtml(task.levelLabel)}" />
      </label>
      <label class="task-field">
        <span>Level number</span>
        <input name="level" type="number" min="1" value="${escapeHtml(task.level)}" />
      </label>
      <div class="task-boolean-group">
        <label class="task-checkbox">
          <input name="critical" type="checkbox" ${task.critical ? 'checked' : ''} />
          <span>Critical path task</span>
        </label>
        <label class="task-checkbox">
          <input name="learningPlan" type="checkbox" ${task.learningPlan ? 'checked' : ''} />
          <span>Has learning plan</span>
        </label>
      </div>
      <label class="task-field">
        <span>Predecessors (comma separated)</span>
        <input name="predecessors" type="text" value="${escapeHtml(task.predecessors.join(', '))}" />
      </label>
      <label class="task-field">
        <span>Dependency types (match predecessors order)</span>
        <input name="dependencyTypes" type="text" value="${escapeHtml(task.dependencyTypes.join(', '))}" />
        <small class="task-field-hint">Defaults to FS when left blank.</small>
      </label>
      <label class="task-field">
        <span>Successors (comma separated)</span>
        <input name="successors" type="text" value="${escapeHtml(task.successors.join(', '))}" />
      </label>
    </div>
  `;

  const saveButton = form.querySelector('[data-role="save"]');
  const deleteButton = form.querySelector('[data-role="delete"]');
  const initialState = getTaskFormState(form);

  const updateSaveState = () => {
    const currentState = getTaskFormState(form);
    const dirty = hasTaskFormChanges(initialState, currentState);
    saveButton.disabled = !dirty;
    saveButton.classList.toggle('dirty', dirty);
  };

  form.addEventListener('input', updateSaveState);
  form.addEventListener('change', updateSaveState);

  form.addEventListener('submit', (event) => {
    event.preventDefault();
    const nextState = getTaskFormState(form);
    if (!hasTaskFormChanges(initialState, nextState)) {
      return;
    }
    applyTaskEdits(task.id, nextState);
  });

  deleteButton.addEventListener('click', (event) => {
    event.preventDefault();
    requestTaskDeletion(task.id);
  });

  updateSaveState();
  return form;
}

function getTaskFormState(form) {
  const formData = new FormData(form);
  return {
    taskName: (formData.get('taskName') || '').toString().trim(),
    responsibleTeam: (formData.get('responsibleTeam') || '').toString().trim(),
    responsibleFunction: (formData.get('responsibleFunction') || '').toString().trim(),
    scope: (formData.get('scope') || '').toString().trim(),
    levelLabel: (formData.get('levelLabel') || '').toString().trim(),
    level: (formData.get('level') || '').toString().trim(),
    critical: form.querySelector('input[name="critical"]').checked,
    learningPlan: form.querySelector('input[name="learningPlan"]').checked,
    predecessors: (formData.get('predecessors') || '').toString().trim(),
    dependencyTypes: (formData.get('dependencyTypes') || '').toString().trim(),
    successors: (formData.get('successors') || '').toString().trim()
  };
}

function hasTaskFormChanges(initialState, currentState) {
  return Object.keys(initialState).some((key) => initialState[key] !== currentState[key]);
}

function renderDependencies() {
  if (!dom.dependencyTable) return;

  if (!dataset) {
    dom.dependencyTable.innerHTML = '';
    if (dom.dependencyFocus) {
      dom.dependencyFocus.innerHTML = '';
    }
    if (dom.dependencySummaryLabel) {
      dom.dependencySummaryLabel.textContent = 'Dependencies';
    }
    if (dom.dependencySummaryCount) {
      dom.dependencySummaryCount.textContent = '';
    }
    if (dom.dependencyPanel) {
      dom.dependencyPanel.open = false;
    }
    return;
  }

  const focusId = selectedTask?.id ?? null;
  const hasExplicitFocus = Boolean(selectedTask);

  if (dom.dependencySummaryLabel) {
    if (hasExplicitFocus) {
      const labelName = selectedTask.placeholder
        ? `External ${selectedTask.id}`
        : `${selectedTask.name} (#${selectedTask.id})`;
      dom.dependencySummaryLabel.textContent = `Dependencies for ${labelName}`;
    } else {
      dom.dependencySummaryLabel.textContent = 'Dependencies';
    }
  }

  if (dom.dependencyFocus) {
    if (!selectedTask) {
      dom.dependencyFocus.innerHTML =
        '<p class="empty-state soft">Select a task in the hierarchy or network to preview its dependency chain.</p>';
    } else if (selectedTask.placeholder) {
      dom.dependencyFocus.innerHTML = `
        <div class="dependency-focus-card">
          <div class="dependency-focus-header">
            <span class="dependency-focus-title">External dependency</span>
            <span class="dependency-focus-id">#${selectedTask.id}</span>
          </div>
          <p class="dependency-focus-note">This task appears in the dependency chain but is not part of the uploaded schedule.</p>
        </div>
      `;
    } else {
      const badges = [];
      if (selectedTask.critical) {
        badges.push('<span class="status-pill critical">Critical</span>');
      }
      if (selectedTask.learningPlan) {
        badges.push('<span class="status-pill learning">Learning plan</span>');
      }
      dom.dependencyFocus.innerHTML = `
        <div class="dependency-focus-card">
          <div class="dependency-focus-header">
            <span class="dependency-focus-title">${selectedTask.name}</span>
            <span class="dependency-focus-id">#${selectedTask.id}</span>
          </div>
          <div class="dependency-focus-meta">${selectedTask.team || 'Unassigned'} · Level ${selectedTask.levelLabel}</div>
          <div class="dependency-focus-meta">${selectedTask.scope}</div>
          ${badges.length ? `<div class="dependency-focus-badges">${badges.join(' ')}</div>` : ''}
        </div>
      `;
    }
  }

  const rows = dataset.dependencies
    .filter((dep) => {
      if (focusId) {
        return dep.source === focusId || dep.target === focusId;
      }
      return selectedTaskIds.has(dep.target) || selectedTaskIds.has(dep.source);
    })
    .map((dep) => {
      const predecessor = dataset.tasksById.get(dep.source);
      const successor = dataset.tasksById.get(dep.target);
      const sourceFocused = focusId ? dep.source === focusId : selectedTaskIds.has(dep.source);
      const targetFocused = focusId ? dep.target === focusId : selectedTaskIds.has(dep.target);
      return {
        sourceId: dep.source,
        sourceName: predecessor?.name || 'External task',
        targetId: dep.target,
        targetName: successor?.name || 'External task',
        type: dep.type,
        sourceFocused,
        targetFocused
      };
    });

  if (dom.dependencySummaryCount) {
    dom.dependencySummaryCount.textContent = rows.length
      ? `${rows.length} link${rows.length === 1 ? '' : 's'}`
      : 'No links';
  }

  if (rows.length === 0) {
    dom.dependencyTable.innerHTML = '<p class="empty-state soft">No dependencies to display for the current focus.</p>';
    return;
  }

  const table = `
    <table>
      <thead>
        <tr>
          <th>Predecessor</th>
          <th>Dependency</th>
          <th>Successor</th>
        </tr>
      </thead>
      <tbody>
        ${rows
          .map(
            (row) => `
            <tr class="${row.sourceFocused || row.targetFocused ? 'focused-row' : ''}">
              <td><strong>${row.sourceId}</strong> · ${row.sourceName}${
                row.sourceFocused ? ' <span class="dep-badge">In focus</span>' : ''
              }</td>
              <td>${row.type}</td>
              <td><strong>${row.targetId}</strong> · ${row.targetName}${
                row.targetFocused ? ' <span class="dep-badge">In focus</span>' : ''
              }</td>
            </tr>
          `
          )
          .join('')}
      </tbody>
    </table>
  `;

  dom.dependencyTable.innerHTML = table;
}

function applyTaskEdits(taskId, state) {
  if (!dataset) return false;
  const task = dataset.tasksById.get(taskId);
  if (!task) return false;

  const name = state.taskName;
  if (!name) {
    alert('Task name cannot be empty.');
    return false;
  }

  const levelNumber = Number(state.level) || parseLevel(state.levelLabel || task.levelLabel);
  const levelLabel = state.levelLabel || `L${levelNumber}`;

  const predecessorIds = splitIds(state.predecessors);
  const dependencyTypesRaw = splitIds(state.dependencyTypes).map((type) => type.toUpperCase());
  const dependencyTypes = predecessorIds.map((_, idx) => {
    const resolved = dependencyTypesRaw[idx] || dependencyTypesRaw[0] || 'FS';
    return resolved || 'FS';
  });

  const successorIds = splitIds(state.successors);

  const oldPredecessors = [...task.predecessors];
  const oldSuccessors = [...task.successors];

  task.name = name;
  task.team = state.responsibleTeam;
  task.function = state.responsibleFunction;
  task.scope = state.scope || 'Unscoped';
  task.level = Math.max(1, levelNumber || task.level || 1);
  task.levelLabel = levelLabel;
  task.critical = Boolean(state.critical);
  task.learningPlan = Boolean(state.learningPlan);
  task.predecessors = predecessorIds;
  task.dependencyTypes = dependencyTypes;
  task.successors = successorIds;

  updateTaskRaw(task);
  syncPredecessorRelationships(task, oldPredecessors, predecessorIds);
  syncSuccessorRelationships(task, oldSuccessors, successorIds);

  afterDatasetMutation(`Task "${task.name}" updated.`);
  return true;
}

function syncPredecessorRelationships(task, previous, next) {
  const nextSet = new Set(next);

  previous.forEach((predId) => {
    if (nextSet.has(predId)) return;
    const predecessor = dataset.tasksById.get(predId);
    if (!predecessor) return;
    predecessor.successors = predecessor.successors.filter((id) => id !== task.id);
    updateTaskRaw(predecessor);
  });

  next.forEach((predId) => {
    const predecessor = dataset.tasksById.get(predId);
    if (!predecessor) return;
    if (!predecessor.successors.includes(task.id)) {
      predecessor.successors.push(task.id);
      updateTaskRaw(predecessor);
    }
  });
}

function syncSuccessorRelationships(task, previous, next) {
  const nextSet = new Set(next);

  previous.forEach((succId) => {
    if (nextSet.has(succId)) return;
    const successor = dataset.tasksById.get(succId);
    if (!successor) return;
    for (let i = successor.predecessors.length - 1; i >= 0; i -= 1) {
      if (successor.predecessors[i] === task.id) {
        successor.predecessors.splice(i, 1);
        successor.dependencyTypes.splice(i, 1);
      }
    }
    updateTaskRaw(successor);
  });

  next.forEach((succId) => {
    const successor = dataset.tasksById.get(succId);
    if (!successor) return;
    if (!successor.predecessors.includes(task.id)) {
      successor.predecessors.push(task.id);
      successor.dependencyTypes.push(task.dependencyTypes[0] || 'FS');
      updateTaskRaw(successor);
    }
  });
}

function requestTaskDeletion(taskId) {
  if (!dataset) return;
  const task = dataset.tasksById.get(taskId);
  if (!task) return;

  const affectedIds = new Set();
  dataset.tasks.forEach((entry) => {
    if (entry.predecessors.includes(taskId) || entry.successors.includes(taskId)) {
      affectedIds.add(entry.id);
    }
  });

  const affectedList = Array.from(affectedIds)
    .map((id) => ` - ${describeTaskById(id)}`)
    .join('\n');

  const approval = window.confirm(
    `Deleting ${describeTaskById(taskId)} will remove the task and all of its dependencies.` +
      (affectedList ? `\n\nDependent tasks impacted:\n${affectedList}` : '\n\nNo other tasks reference this node.') +
      '\n\nDo you approve this deletion?'
  );

  if (!approval) return;
  deleteTaskById(taskId);
}

function deleteTaskById(taskId) {
  const index = dataset.tasks.findIndex((task) => task.id === taskId);
  if (index === -1) return;
  const [removedTask] = dataset.tasks.splice(index, 1);
  dataset.tasksById.delete(taskId);

  dataset.tasks.forEach((task) => {
    let changed = false;
    for (let i = task.predecessors.length - 1; i >= 0; i -= 1) {
      if (task.predecessors[i] === taskId) {
        task.predecessors.splice(i, 1);
        task.dependencyTypes.splice(i, 1);
        changed = true;
      }
    }
    const successorIndex = task.successors.indexOf(taskId);
    if (successorIndex !== -1) {
      task.successors.splice(successorIndex, 1);
      changed = true;
    }
    if (changed) {
      updateTaskRaw(task);
    }
  });

  if (selectedTask && selectedTask.id === taskId) {
    selectedTask = null;
  }

  hierarchySelections.forEach((value, level) => {
    if (value === taskId) {
      hierarchySelections.delete(level);
    }
  });
  graphSelections.delete(taskId);
  rebuildSelectedTaskIds();

  afterDatasetMutation(`Task "${removedTask.name}" deleted.`);
}

function handleEdgeTap(evt) {
  if (!dataset) return;
  const edge = evt.target;
  const sourceId = edge.data('source');
  const targetId = edge.data('target');

  const affectedIds = new Set([targetId]);
  dataset.tasks.forEach((task) => {
    if (task.predecessors.includes(targetId)) {
      affectedIds.add(task.id);
    }
  });

  const affectedList = Array.from(affectedIds)
    .map((id) => ` - ${describeTaskById(id)}`)
    .join('\n');

  const approval = window.confirm(
    `Removing the dependency ${describeTaskById(sourceId)} → ${describeTaskById(targetId)} will detach the following tasks:` +
      (affectedList ? `\n\n${affectedList}` : '\n\nNo downstream tasks are linked to this dependency.') +
      '\n\nDo you approve this change?'
  );

  if (!approval) return;
  removeDependency(sourceId, targetId);
}

function removeDependency(sourceId, targetId) {
  const targetTask = dataset.tasksById.get(targetId);
  if (!targetTask) return;

  const sourceTask = dataset.tasksById.get(sourceId);
  let removed = false;

  for (let i = targetTask.predecessors.length - 1; i >= 0; i -= 1) {
    if (targetTask.predecessors[i] === sourceId) {
      targetTask.predecessors.splice(i, 1);
      targetTask.dependencyTypes.splice(i, 1);
      removed = true;
    }
  }
  updateTaskRaw(targetTask);

  if (sourceTask) {
    const filtered = sourceTask.successors.filter((id) => id !== targetId);
    if (filtered.length !== sourceTask.successors.length) {
      sourceTask.successors = filtered;
      updateTaskRaw(sourceTask);
      removed = true;
    }
  }

  if (removed) {
    afterDatasetMutation(`Removed dependency ${describeTaskById(sourceId)} → ${describeTaskById(targetId)}.`);
  }
}

function describeTaskById(taskId) {
  if (!dataset) return `Task ${taskId}`;
  const task = dataset.tasksById.get(taskId);
  if (!task) {
    return `External task ${taskId}`;
  }
  return `${task.name} (#${task.id})`;
}

function afterDatasetMutation(actionDescription) {
  if (!dataset) return;
  dataset.modified = true;
  refreshDatasetIndexes();
  cleanupSelections();
  rebuildSelectedTaskIds();

  renderLegend();
  renderLevelControls();
  renderHierarchy();
  renderDependencies();
  updateGraph();
  renderTaskDetails();

  promptForDownload(actionDescription);
}

function refreshDatasetIndexes() {
  dataset.tasksById = new Map(dataset.tasks.map((task) => [task.id, task]));
  dataset.dependencies = buildDependencies(dataset.tasks);
  dataset.levels = Array.from(new Set(dataset.tasks.map((task) => task.level))).sort((a, b) => a - b);
}

function cleanupSelections() {
  hierarchySelections.forEach((taskId, level) => {
    if (!dataset.tasksById.has(taskId)) {
      hierarchySelections.delete(level);
    }
  });

  graphSelections = new Set([...graphSelections].filter((taskId) => dataset.tasksById.has(taskId)));

  if (selectedTask && !selectedTask.placeholder) {
    const updatedTask = dataset.tasksById.get(selectedTask.id);
    if (updatedTask) {
      selectedTask = updatedTask;
    } else {
      selectedTask = null;
    }
  }
}

function promptForDownload(actionDescription) {
  if (!dataset) return;
  const label = dataset.label || 'schedule';
  const approval = window.confirm(
    `${actionDescription}\n\nA temporary working copy of "${label}" has been updated.` +
      '\nSelect OK to download the latest Excel file now, or Cancel to continue editing.'
  );
  if (approval) {
    exportDatasetToExcel();
  }
}

function exportDatasetToExcel() {
  if (!dataset) return;
  const rows = dataset.tasks.map((task) => ({
    'Task ID': task.id,
    'Task Name': task.name,
    'Task Level': task.levelLabel || `L${task.level}`,
    'Responsible Sub-team': task.team || '',
    'Responsible Function': task.function || '',
    Scope: task.scope || '',
    Learning_Plan: task.learningPlan ? 'Yes' : 'No',
    Critical: task.critical ? 'Yes' : 'No',
    'Predecessors IDs': task.predecessors.join(', '),
    'Successors IDs': task.successors.join(', '),
    'Dependency Types': task.dependencyTypes.join(', ')
  }));

  const worksheet = XLSX.utils.json_to_sheet(rows);
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, 'Tasks');
  XLSX.writeFile(workbook, generateDownloadFilename());
}

function generateDownloadFilename() {
  const label = dataset?.label || 'schedule';
  const safe = label.replace(/[^a-z0-9]+/gi, '_').replace(/^_+|_+$/g, '');
  return `${safe || 'schedule'}_edited.xlsx`;
}

function escapeHtml(value) {
  return String(value ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function updateTaskRaw(task) {
  if (!task.raw) {
    task.raw = {};
  }
  setRawField(task.raw, ['Task Name', 'Name'], task.name || '');
  setRawField(task.raw, ['TaskId', 'Task ID', 'ID'], task.id || '');
  setRawField(task.raw, ['Task Level', 'Level'], task.levelLabel || `L${task.level}`);
  setRawField(task.raw, ['Responsible Sub-team', 'Sub-team'], task.team || '');
  setRawField(task.raw, ['Responsible Function', 'Function'], task.function || '');
  setRawField(task.raw, ['SCOPE', 'Scope'], task.scope || '');
  setRawField(task.raw, ['Learning_Plan', 'Learning Plan'], task.learningPlan ? 'Yes' : 'No');
  setRawField(task.raw, ['Critical', 'Is Critical'], task.critical ? 'Yes' : 'No');
  setRawField(task.raw, ['Predecessors IDs', 'Predecessor IDs', 'Predecessors'], task.predecessors.join(', '));
  setRawField(task.raw, ['Successors IDs', 'Successor IDs', 'Successors'], task.successors.join(', '));
  setRawField(task.raw, ['Dependency Type', 'Dependency Types'], task.dependencyTypes.join(', '));
}

function setRawField(raw, keys, value) {
  for (const key of keys) {
    if (Object.prototype.hasOwnProperty.call(raw, key)) {
      raw[key] = value;
      return;
    }
  }
  raw[keys[0]] = value;
}

window.addEventListener('resize', () => {
  if (cyInstance) {
    cyInstance.resize();
  }
});

loadSampleData();
