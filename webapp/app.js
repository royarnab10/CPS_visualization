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
let cyInstance = null;
let includeNeighbors = true;
let selectedTask = null;

const dom = {
  fileInput: document.getElementById('file-input'),
  loadSample: document.getElementById('load-sample'),
  hierarchy: document.getElementById('hierarchy'),
  levelControls: document.getElementById('level-controls'),
  taskDetails: document.getElementById('task-details'),
  dependencyTable: document.getElementById('dependency-table'),
  autoExpand: document.getElementById('auto-expand'),
  fitGraph: document.getElementById('fit-graph'),
  resetSelection: document.getElementById('reset-selection'),
  legend: document.getElementById('legend')
};

dom.fileInput.addEventListener('change', handleFileInput);
dom.loadSample.addEventListener('click', loadSampleData);
dom.autoExpand.addEventListener('change', () => {
  includeNeighbors = dom.autoExpand.checked;
  updateGraph();
});
dom.fitGraph.addEventListener('click', () => cyInstance && cyInstance.fit(undefined, 40));
dom.resetSelection.addEventListener('click', () => {
  if (cyInstance) {
    cyInstance.elements().unselect();
  }
  hierarchySelections.clear();
  rebuildSelectedTaskIds();
  selectedTask = null;
  renderLevelControls();
  renderHierarchy();
  renderDependencies();
  updateGraph();
  renderTaskDetails();
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
  selectedTask = null;

  dataset = {
    label,
    tasks,
    tasksById,
    dependencies,
    levels
  };

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

  const context = `<span class="legend-item"><span class="legend-swatch" style="background:transparent;border:2px dashed rgba(255,255,255,0.6);"></span>Auto-added context</span>`;

  dom.legend.innerHTML = items + context;
}

function renderLevelControls() {
  if (!dataset) {
    dom.levelControls.innerHTML = '';
    return;
  }

  const entries = Array.from(hierarchySelections.entries()).sort((a, b) => a[0] - b[0]);

  if (entries.length === 0) {
    dom.levelControls.innerHTML = '<p class="empty-state">Select a starting task to explore downstream milestones.</p>';
    return;
  }

  const wrapper = document.createElement('div');
  wrapper.className = 'selection-path';

  entries.forEach(([level, taskId]) => {
    const task = dataset.tasksById.get(taskId);
    if (!task) return;

    const chip = document.createElement('button');
    chip.type = 'button';
    chip.className = 'path-chip';
    chip.innerHTML = `<span>L${level}</span><strong>${task.name}</strong>`;
    chip.addEventListener('click', () => clearHierarchyFrom(level));
    wrapper.appendChild(chip);
  });

  dom.levelControls.innerHTML = '';
  dom.levelControls.appendChild(wrapper);
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
        : 'Select a task to drill into the next level. Cross-level links appear when available.';
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

  card.addEventListener('click', () => handleHierarchySelection(level, task));

  return card;
}

function getHierarchyLevels() {
  if (!dataset) return [];
  const sorted = dataset.levels.slice().sort((a, b) => a - b);
  const levelTwoIndex = sorted.indexOf(2);
  if (levelTwoIndex >= 0) {
    return sorted.slice(levelTwoIndex);
  }
  return sorted;
}

function handleHierarchySelection(level, task) {
  if (!dataset) return;

  const currentlySelected = hierarchySelections.get(level);
  if (currentlySelected === task.id) {
    clearHierarchyFrom(level);
    return;
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
  selectedTaskIds = new Set(hierarchySelections.values());
}

function updateGraph() {
  const container = document.getElementById('cy');
  if (!dataset || !container) return;

  const nodes = new Map();
  const edges = [];
  const visible = new Set(selectedTaskIds);
  const neighborIds = new Set();

  if (includeNeighbors && selectedTaskIds.size > 0) {
    dataset.dependencies.forEach((dep) => {
      if (selectedTaskIds.has(dep.target) && !selectedTaskIds.has(dep.source)) {
        neighborIds.add(dep.source);
      }
      if (selectedTaskIds.has(dep.source) && !selectedTaskIds.has(dep.target)) {
        neighborIds.add(dep.target);
      }
    });
  }

  neighborIds.forEach((taskId) => visible.add(taskId));

  visible.forEach((taskId) => {
    const task = dataset.tasksById.get(taskId);
    if (task) {
      const color = levelColor(task.level);
      nodes.set(taskId, {
        data: {
          id: taskId,
          label: formatNodeLabel(task.name, taskId),
          color,
          level: task.level,
          scope: task.scope,
          isContext: !selectedTaskIds.has(taskId)
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
          isContext: true
        }
      });
    }
  });

  dataset.dependencies.forEach((dep) => {
    if (!visible.has(dep.source) || !visible.has(dep.target)) return;

    const touchesSelected =
      selectedTaskIds.has(dep.source) || selectedTaskIds.has(dep.target);
    if (!touchesSelected) return;

    if (!nodes.has(dep.source)) {
      nodes.set(dep.source, {
        data: {
          id: dep.source,
          label: formatNodeLabel(`External ${dep.source}`, dep.source),
          color: '#94a3b8',
          level: dep.sourceLevel || 0,
          scope: 'External',
          isContext: true
        }
      });
    }
    if (!nodes.has(dep.target)) {
      nodes.set(dep.target, {
        data: {
          id: dep.target,
          label: formatNodeLabel(`External ${dep.target}`, dep.target),
          color: '#94a3b8',
          level: dep.targetLevel || 0,
          scope: 'External',
          isContext: true
        }
      });
    }
    edges.push({
      data: {
        id: `${dep.source}->${dep.target}`,
        source: dep.source,
        target: dep.target,
        type: dep.type
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
          selector: 'node[?isContext]',
          style: {
            'border-style': 'dashed',
            'opacity': 0.65,
            'background-opacity': 0.4,
            'color': '#1f2937'
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
      const task = dataset.tasksById.get(node.id());
      selectedTask = task || { id: node.id(), placeholder: true };
      renderTaskDetails();
    });

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

  const critical = selectedTask.critical
    ? '<span class="status-pill critical">Critical</span>'
    : '<span class="status-pill">Non-critical</span>';
  const learning = selectedTask.learningPlan
    ? '<span class="status-pill learning">Learning plan</span>'
    : '';

  dom.taskDetails.innerHTML = `
    <h3>${selectedTask.name}</h3>
    <p>${critical} ${learning}</p>
    <dl>
      <dt>Task ID</dt><dd>${selectedTask.id}</dd>
      <dt>Level</dt><dd>${selectedTask.levelLabel}</dd>
      <dt>Sub-team</dt><dd>${selectedTask.team || '—'}</dd>
      <dt>Function</dt><dd>${selectedTask.function || '—'}</dd>
      <dt>Scope</dt><dd>${selectedTask.scope}</dd>
      <dt>Predecessors</dt><dd>${formatList(selectedTask.predecessors)}</dd>
      <dt>Successors</dt><dd>${formatList(selectedTask.successors)}</dd>
    </dl>
  `;
}

function formatList(items) {
  if (!items || items.length === 0) return '—';
  return items.join(', ');
}

function renderDependencies() {
  if (!dataset) {
    dom.dependencyTable.innerHTML = '';
    return;
  }

  const rows = dataset.dependencies
    .filter((dep) => selectedTaskIds.has(dep.target))
    .map((dep) => {
      const predecessor = dataset.tasksById.get(dep.source);
      const successor = dataset.tasksById.get(dep.target);
      return {
        sourceId: dep.source,
        sourceName: predecessor?.name || 'External task',
        targetId: dep.target,
        targetName: successor?.name || 'External task',
        type: dep.type,
        critical: successor?.critical || predecessor?.critical || false
      };
    });

  if (rows.length === 0) {
    dom.dependencyTable.innerHTML = '<p>No dependencies to display for the current selection.</p>';
    return;
  }

  const header = `
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
            <tr>
              <td><strong>${row.sourceId}</strong> · ${row.sourceName}</td>
              <td>${row.type}</td>
              <td><strong>${row.targetId}</strong> · ${row.targetName}</td>
            </tr>
          `
          )
          .join('')}
      </tbody>
    </table>
  `;

  dom.dependencyTable.innerHTML = header;
}

window.addEventListener('resize', () => {
  if (cyInstance) {
    cyInstance.resize();
  }
});

loadSampleData();
