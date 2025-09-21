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
let levelState = new Map();
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
    selectedTask = null;
    renderTaskDetails();
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
  const minLevel = Math.min(...levels);

  selectedTaskIds = new Set(tasks.filter((t) => t.level === minLevel).map((t) => t.id));
  levelState = new Map(levels.map((lvl) => [lvl, lvl === minLevel ? 'all' : 'none']));

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

  dom.levelControls.innerHTML = '';

  dataset.levels.forEach((level, idx) => {
    if (!levelState.has(level)) {
      levelState.set(level, 'none');
    }
    const state = levelState.get(level);
    const pill = document.createElement('label');
    pill.className = `level-pill ${state === 'all' ? 'active' : ''}`;

    const checkbox = document.createElement('input');
    checkbox.type = 'checkbox';
    checkbox.checked = state === 'all';
    checkbox.indeterminate = state === 'partial';
    checkbox.addEventListener('change', () => {
      if (checkbox.checked) {
        levelState.set(level, 'all');
        dataset.tasks
          .filter((task) => task.level === level)
          .forEach((task) => selectedTaskIds.add(task.id));
      } else {
        levelState.set(level, 'none');
        dataset.tasks
          .filter((task) => task.level === level)
          .forEach((task) => selectedTaskIds.delete(task.id));
      }
      renderLevelControls();
      renderHierarchy();
      renderDependencies();
      updateGraph();
    });

    const swatch = document.createElement('span');
    swatch.className = 'legend-swatch';
    swatch.style.background = levelColor(level, idx);

    pill.append(checkbox, swatch, document.createTextNode(`L${level}`));
    dom.levelControls.appendChild(pill);
  });
}

function renderHierarchy() {
  if (!dataset) {
    dom.hierarchy.innerHTML = '';
    return;
  }

  const levels = new Map();
  dataset.tasks.forEach((task) => {
    if (!levels.has(task.level)) {
      levels.set(task.level, new Map());
    }
    const scopes = levels.get(task.level);
    if (!scopes.has(task.scope)) {
      scopes.set(task.scope, []);
    }
    scopes.get(task.scope).push(task);
  });

  const fragments = [];
  const sortedLevels = Array.from(levels.keys()).sort((a, b) => a - b);

  sortedLevels.forEach((level) => {
    const scopes = levels.get(level);
    const scopeEntries = Array.from(scopes.entries()).sort((a, b) => a[0].localeCompare(b[0]));

    const levelDetails = document.createElement('details');
    levelDetails.open = levelState.get(level) !== 'none';

    const selectedCount = dataset.tasks.filter((task) => task.level === level && selectedTaskIds.has(task.id)).length;
    const totalCount = dataset.tasks.filter((task) => task.level === level).length;

    levelDetails.innerHTML = `<summary>Level ${level} &middot; ${selectedCount}/${totalCount} selected</summary>`;

    scopeEntries.forEach(([scope, tasks]) => {
      const scopeDetails = document.createElement('details');
      scopeDetails.className = 'scope-block';
      scopeDetails.open = tasks.some((task) => selectedTaskIds.has(task.id));
      const selectedInScope = tasks.filter((task) => selectedTaskIds.has(task.id)).length;
      scopeDetails.innerHTML = `<summary class="scope-header">${scope} <span class="info">${selectedInScope}/${tasks.length}</span></summary>`;

      const list = document.createElement('div');
      list.className = 'task-list';

      tasks
        .slice()
        .sort((a, b) => a.name.localeCompare(b.name))
        .forEach((task) => {
          const item = document.createElement('div');
          item.className = 'task-item';

          const label = document.createElement('label');
          const checkbox = document.createElement('input');
          checkbox.type = 'checkbox';
          checkbox.checked = selectedTaskIds.has(task.id);
          checkbox.addEventListener('change', () => {
            if (checkbox.checked) {
              selectedTaskIds.add(task.id);
            } else {
              selectedTaskIds.delete(task.id);
              if (selectedTask?.id === task.id) {
                selectedTask = null;
                renderTaskDetails();
              }
            }
            recalculateLevelState(task.level);
            renderLevelControls();
            renderHierarchy();
            renderDependencies();
            updateGraph();
          });

          label.appendChild(checkbox);
          label.appendChild(document.createTextNode(`${task.name}`));

          const info = document.createElement('div');
          info.className = 'info';
          info.textContent = `#${task.id} · ${task.team || 'Unassigned'}`;

          item.append(label, info);
          list.appendChild(item);
        });

      scopeDetails.appendChild(list);
      levelDetails.appendChild(scopeDetails);
    });

    fragments.push(levelDetails);
  });

  dom.hierarchy.innerHTML = '';
  fragments.forEach((fragment) => dom.hierarchy.appendChild(fragment));
}

function recalculateLevelState(level) {
  if (!dataset) return;
  const tasks = dataset.tasks.filter((task) => task.level === level);
  const selectedCount = tasks.filter((task) => selectedTaskIds.has(task.id)).length;
  let state = 'none';
  if (selectedCount === tasks.length) {
    state = 'all';
  } else if (selectedCount > 0) {
    state = 'partial';
  }
  levelState.set(level, state);
}

function updateGraph() {
  const container = document.getElementById('cy');
  if (!dataset || !container) return;

  const nodes = new Map();
  const edges = [];
  const visible = new Set(selectedTaskIds);

  if (includeNeighbors) {
    dataset.dependencies.forEach((dep) => {
      if (visible.has(dep.target)) {
        visible.add(dep.source);
      }
      if (visible.has(dep.source)) {
        visible.add(dep.target);
      }
    });
  }

  visible.forEach((taskId) => {
    const task = dataset.tasksById.get(taskId);
    if (task) {
      const color = levelColor(task.level);
      nodes.set(taskId, {
        data: {
          id: taskId,
          label: `${task.name} ( ${taskId} )`,
          color,
          level: task.level,
          scope: task.scope,
          isContext: false
        }
      });
    } else {
      nodes.set(taskId, {
        data: {
          id: taskId,
          label: `External ${taskId}`,
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
    if (!nodes.has(dep.source)) {
      nodes.set(dep.source, {
        data: {
          id: dep.source,
          label: `External ${dep.source}`,
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
          label: `External ${dep.target}`,
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
            'label': 'data(label)',
            'color': '#fff',
            'font-size': '10px',
            'text-wrap': 'wrap',
            'text-max-width': 120,
            'text-valign': 'center',
            'text-halign': 'center',
            'border-width': 2,
            'border-color': 'rgba(255,255,255,0.65)'
          }
        },
        {
          selector: 'node[?isContext]',
          style: {
            'border-style': 'dashed',
            'opacity': 0.65,
            'background-opacity': 0.25,
            'color': '#cbd5f5'
          }
        },
        {
          selector: 'node:selected',
          style: {
            'border-width': 4,
            'border-color': '#facc15'
          }
        },
        {
          selector: 'edge',
          style: {
            'width': 2,
            'line-color': '#cbd5f5',
            'target-arrow-color': '#cbd5f5',
            'target-arrow-shape': 'triangle',
            'curve-style': 'bezier',
            'label': 'data(type)',
            'font-size': '8px',
            'color': '#e2e8f0',
            'text-background-color': 'rgba(15,23,42,0.75)',
            'text-background-opacity': 1,
            'text-background-padding': 2
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
            'source-arrow-color': '#cbd5f5',
            'arrow-scale': 1.2
          }
        },
        {
          selector: 'edge[type = "SF"]',
          style: {
            'line-style': 'dashed',
            'source-arrow-shape': 'triangle',
            'source-arrow-color': '#cbd5f5',
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

  cyInstance.layout({ name: 'cose', animate: false, padding: 40 }).run();
}

function levelColor(level, idx) {
  const paletteIndex = typeof idx === 'number' ? idx % levelPalette.length : (level - 1) % levelPalette.length;
  return levelPalette[paletteIndex];
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
