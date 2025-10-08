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
let isProcessingUpload = false;

const BASE_DURATION_HEADERS = ['Base Duration', 'Duration', 'Duration (days)', 'Task Duration', 'BaseDuration'];
const CALCULATED_DURATION_HEADERS = [
  'Calculated Duration (days)',
  'Calculated Duration',
  'Effective Duration',
  'EffectiveDuration'
];
const SCHEDULE_BASE_DURATION_HEADERS = ['BaseDurationHours', 'Base Duration Hours', 'Base Duration (h)'];
const SCHEDULE_EFFECTIVE_DURATION_HEADERS = [
  'EffectiveDurationHours',
  'Effective Duration Hours',
  'Schedule Duration Hours'
];
const SCHEDULE_ES_HEADERS = ['ES (h)', 'ES Hours', 'ES'];
const SCHEDULE_EF_HEADERS = ['EF (h)', 'EF Hours', 'EF'];
const SCHEDULE_LS_HEADERS = ['LS (h)', 'LS Hours', 'LS'];
const SCHEDULE_LF_HEADERS = ['LF (h)', 'LF Hours', 'LF'];
const SCHEDULE_TOTAL_SLACK_HEADERS = ['TotalSlack (h)', 'Total Slack (h)', 'TotalSlackHours', 'Total Slack'];
const SCHEDULE_FREE_SLACK_HEADERS = ['FreeSlack (h)', 'Free Slack (h)', 'FreeSlackHours', 'Free Slack'];
const SCHEDULE_CRITICAL_HEADERS = ['IsCritical', 'Critical', 'On Critical Path'];
const SCHEDULE_START_HEADERS = ['StartDate', 'Start Date', 'Start'];
const SCHEDULE_FINISH_HEADERS = ['FinishDate', 'Finish Date', 'Finish'];
const HOURS_PER_DAY = 8;

const dom = {
  fileInput: document.getElementById('file-input'),
  loadSample: document.getElementById('load-sample'),
  downloadDataset: document.getElementById('download-dataset'),
  processingNotice: document.getElementById('processing-notice'),
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
  showFullGraph: document.getElementById('show-full-graph'),
  durationSummary: document.getElementById('duration-summary'),
  totalDurationDisplay: document.getElementById('total-duration-display'),
  totalDurationDetail: document.getElementById('total-duration-detail'),
  selectedDurationDisplay: document.getElementById('selected-duration-display'),
  selectedDurationDetail: document.getElementById('selected-duration-detail'),
  pathDurationDisplay: document.getElementById('path-duration-display'),
  pathDurationDetail: document.getElementById('path-duration-detail')
};

renderDurationSummary();

dom.fileInput.addEventListener('change', handleFileInput);
dom.loadSample.addEventListener('click', loadSampleData);
if (dom.downloadDataset) {
  dom.downloadDataset.addEventListener('click', downloadCurrentWorkbook);
}
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

function setProcessingState(message, active) {
  isProcessingUpload = active;
  if (!dom.processingNotice) {
    if (dom.fileInput) dom.fileInput.disabled = active;
    if (dom.loadSample) dom.loadSample.disabled = active;
    updateDownloadButton();
    return;
  }
  const displayMessage = message || 'Cleaning schedule dependencies…';
  if (active) {
    dom.processingNotice.textContent = displayMessage;
    dom.processingNotice.hidden = false;
  } else {
    dom.processingNotice.hidden = true;
  }
  if (dom.fileInput) dom.fileInput.disabled = active;
  if (dom.loadSample) dom.loadSample.disabled = active;
  updateDownloadButton();
}

async function preprocessWorkbook(file) {
  const formData = new FormData();
  formData.append('file', file);
  const response = await fetch('/api/preprocess', {
    method: 'POST',
    body: formData
  });
  if (!response.ok) {
    const text = await response.text();
    throw new Error(text || `Preprocessing failed with status ${response.status}`);
  }
  return response.json();
}

function extractScheduleInfo(payload) {
  if (!payload || typeof payload !== 'object') {
    return { records: [], metadata: null };
  }

  const schedule = payload.schedule || payload.scheduleRecords || payload.scheduleResult || null;

  if (Array.isArray(schedule)) {
    return { records: schedule, metadata: payload.scheduleMetadata || null };
  }

  if (schedule && Array.isArray(schedule.records)) {
    return { records: schedule.records, metadata: schedule.metadata || null };
  }

  if (schedule && Array.isArray(schedule.rows)) {
    return { records: schedule.rows, metadata: schedule.metadata || null };
  }

  if (schedule && Array.isArray(schedule.tasks)) {
    return { records: schedule.tasks, metadata: schedule.metadata || null };
  }

  if (Array.isArray(payload.scheduleRows)) {
    return { records: payload.scheduleRows, metadata: payload.scheduleMetadata || null };
  }

  return { records: [], metadata: payload.scheduleMetadata || null };
}

function parseExcelPayload(excelPayload, fallbackName = '') {
  if (!excelPayload || typeof excelPayload !== 'object') {
    return null;
  }

  const base64 =
    typeof excelPayload.data === 'string'
      ? excelPayload.data
      : typeof excelPayload.base64 === 'string'
        ? excelPayload.base64
        : '';

  if (!base64) {
    return null;
  }

  const contentType =
    typeof excelPayload.contentType === 'string' && excelPayload.contentType
      ? excelPayload.contentType
      : 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
  const filename =
    typeof excelPayload.filename === 'string' && excelPayload.filename
      ? excelPayload.filename
      : buildProcessedFilename(fallbackName);

  try {
    const binary = atob(base64);
    const length = binary.length;
    const buffer = new Uint8Array(length);
    for (let i = 0; i < length; i += 1) {
      buffer[i] = binary.charCodeAt(i);
    }
    const blob = new Blob([buffer], { type: contentType });
    return { filename, contentType, blob };
  } catch (error) {
    console.error('Unable to decode Excel payload', error);
    return null;
  }
}

async function handleFileInput(event) {
  const file = event.target.files?.[0];
  if (!file) return;

  try {
    setProcessingState(`Cleaning ${file.name}…`, true);
    const payload = await preprocessWorkbook(file);
    const records = payload.records || payload.rows || [];
    const scheduleInfo = extractScheduleInfo(payload);
    const processedExcel = parseExcelPayload(payload.excel, file.name);
    if (!Array.isArray(records) || records.length === 0) {
      throw new Error('The preprocessing pipeline returned no rows.');
    }
    prepareDataset(records, file.name, {
      cleaned: Boolean(payload.metadata?.cleaned),
      metadata: payload.metadata || null,
      schedule: scheduleInfo.records,
      scheduleMetadata: scheduleInfo.metadata,
      excelFile: processedExcel
    });
  } catch (error) {
    console.error('Failed to preprocess workbook', error);
    alert('Unable to clean the uploaded file. Ensure the preprocessing service is running and try again.');
  } finally {
    setProcessingState('', false);
    if (event.target) {
      event.target.value = '';
    }
  }
}

async function loadSampleData() {
  const response = await fetch('data/sample_schedule.json');
  const records = await response.json();
  prepareDataset(records, 'Sample schedule', { cleaned: false, excelFile: null });
}

function prepareDataset(records, label, options = {}) {
  if (!records || records.length === 0) {
    alert('The provided worksheet is empty.');
    return;
  }

  const rawRows = Array.isArray(records) ? records.map((row) => ({ ...row })) : [];
  const scheduleRecords = Array.isArray(options.schedule) ? options.schedule : [];
  const fallbackScheduleEntries = rawRows.map((row) => normalizeScheduleRow(row)).filter(Boolean);
  const scheduleById = new Map(fallbackScheduleEntries.map((entry) => [entry.id, entry]));

  scheduleRecords.forEach((record) => {
    const normalized = normalizeScheduleRow(record);
    if (!normalized) return;
    const existing = scheduleById.get(normalized.id);
    scheduleById.set(normalized.id, existing ? { ...existing, ...normalized } : normalized);
  });

  const normalizedSchedule = Array.from(scheduleById.values());

  const tasks = records.map((record) => normalizeRecord(record, scheduleById)).filter(Boolean);
  const scheduleSummary = computeScheduleSummary(
    normalizedSchedule,
    options.scheduleMetadata || null,
    tasks
  );
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
    modified: false,
    preprocessed: Boolean(options.cleaned),
    rawRows,
    processedExcel: options.excelFile || null,
    metadata: options.metadata || null,
    schedule: normalizedSchedule,
    scheduleSummary,
    scheduleById
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
  renderDurationSummary();
  updateDownloadButton();
}

function normalizeRecord(record, scheduleById = null) {
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

  const scheduleEntry = scheduleById ? scheduleById.get(String(id)) : null;
  const levelLabel = mapped['Task Level'] || mapped['Level'] || '';
  const level = parseLevel(levelLabel);
  const predecessorsRaw =
    mapped['Predecessors IDs'] || mapped['Predecessor IDs'] || mapped['Predecessors'] || '';
  const successorsRaw = mapped['Successors IDs'] || mapped['Successor IDs'] || mapped['Successors'] || '';
  const dependencyTypeRaw = mapped['Dependency Type'] || mapped['Dependency Types'] || '';
  const normalizedDependencyRaw =
    mapped['Normalized Predecessors'] || mapped['Normalized Predecessor'] || '';
  const normalizedDependencies = parseNormalizedDependencies(normalizedDependencyRaw);
  const predecessorIds = splitIds(predecessorsRaw);
  const successorIds = splitIds(successorsRaw);
  const dependencyTypeList = splitIds(dependencyTypeRaw)
    .map((type) => (type ? type.toUpperCase() : 'FS'))
    .filter(Boolean);

  const baseDurationValue = readFirstAvailable(mapped, BASE_DURATION_HEADERS);
  const calculatedDurationValue = readFirstAvailable(mapped, CALCULATED_DURATION_HEADERS);
  const scheduleBaseValue = readFirstAvailable(mapped, SCHEDULE_BASE_DURATION_HEADERS);
  const scheduleEffectiveValue = readFirstAvailable(mapped, SCHEDULE_EFFECTIVE_DURATION_HEADERS);

  const baseDuration = parseDurationValue(baseDurationValue);
  const calculatedDuration = parseDurationValue(calculatedDurationValue);

  const scheduleBase = scheduleEntry?.baseDurationHours ?? parseHoursValue(scheduleBaseValue);
  const scheduleEffective = scheduleEntry?.effectiveDurationHours ?? parseHoursValue(scheduleEffectiveValue);

  let baseDetail = '';
  let calculatedDetail = '';

  if (typeof scheduleBase === 'number' && !Number.isNaN(scheduleBase)) {
    const derived = describeDurationFromHours(scheduleBase);
    if (derived) {
      baseDuration.display = derived.display;
      baseDuration.days = derived.days;
      baseDuration.hasValue = true;
      baseDetail = `Schedule base duration · ${derived.detail}`;
    }
  }

  if (typeof scheduleEffective === 'number' && !Number.isNaN(scheduleEffective)) {
    const derived = describeDurationFromHours(scheduleEffective);
    if (derived) {
      calculatedDuration.display = derived.display;
      calculatedDuration.days = derived.days;
      calculatedDuration.hasValue = true;
      calculatedDetail = `Effective CPM duration · ${derived.detail}`;
    }
  }

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
    predecessors: predecessorIds.length > 0 ? predecessorIds : normalizedDependencies.ids,
    successors: successorIds,
    dependencyTypes:
      dependencyTypeList.length > 0 ? dependencyTypeList : normalizedDependencies.types,
    baseDurationDisplay: baseDuration.display,
    baseDurationDays: baseDuration.days,
    hasBaseDuration: baseDuration.hasValue,
    baseDurationDetail: baseDetail,
    calculatedDurationDisplay: calculatedDuration.display,
    calculatedDurationDays: calculatedDuration.days,
    hasCalculatedDuration: calculatedDuration.hasValue,
    calculatedDurationDetail: calculatedDetail,
    raw: mapped,
    schedule: scheduleEntry || buildScheduleFallback(scheduleBase, scheduleEffective, record)
  };
}

function normalizeScheduleRow(record) {
  if (!record || typeof record !== 'object') return null;

  const mapped = {};
  for (const [key, value] of Object.entries(record)) {
    if (!key) continue;
    mapped[key.trim()] = typeof value === 'string' ? value.trim() : value;
  }

  const id = mapped['TaskId'] || mapped['Task ID'] || mapped['ID'];
  if (!id) {
    return null;
  }

  const baseHours = parseHoursValue(readFirstAvailable(mapped, SCHEDULE_BASE_DURATION_HEADERS));
  const effectiveHours = parseHoursValue(readFirstAvailable(mapped, SCHEDULE_EFFECTIVE_DURATION_HEADERS));
  const esHours = parseHoursValue(readFirstAvailable(mapped, SCHEDULE_ES_HEADERS));
  const efHours = parseHoursValue(readFirstAvailable(mapped, SCHEDULE_EF_HEADERS));
  const lsHours = parseHoursValue(readFirstAvailable(mapped, SCHEDULE_LS_HEADERS));
  const lfHours = parseHoursValue(readFirstAvailable(mapped, SCHEDULE_LF_HEADERS));
  const totalSlackHours = parseHoursValue(readFirstAvailable(mapped, SCHEDULE_TOTAL_SLACK_HEADERS));
  const freeSlackHours = parseHoursValue(readFirstAvailable(mapped, SCHEDULE_FREE_SLACK_HEADERS));
  const startRaw = readFirstAvailable(mapped, SCHEDULE_START_HEADERS);
  const finishRaw = readFirstAvailable(mapped, SCHEDULE_FINISH_HEADERS);
  const startDate = parseDateValue(startRaw);
  const finishDate = parseDateValue(finishRaw);
  const startDisplay = formatScheduleDate(startDate, startRaw);
  const finishDisplay = formatScheduleDate(finishDate, finishRaw);
  const isCritical = normalizeBoolean(readFirstAvailable(mapped, SCHEDULE_CRITICAL_HEADERS));

  return {
    id: String(id),
    baseDurationHours: baseHours,
    effectiveDurationHours: effectiveHours,
    esHours,
    efHours,
    lsHours,
    lfHours,
    totalSlackHours,
    freeSlackHours,
    isCritical,
    startDate,
    finishDate,
    startDateDisplay: startDisplay,
    finishDateDisplay: finishDisplay,
    startDateRaw: startRaw || '',
    finishDateRaw: finishRaw || '',
    raw: mapped
  };
}

function buildScheduleFallback(baseHours, effectiveHours, record) {
  const hasScheduleData =
    (typeof baseHours === 'number' && !Number.isNaN(baseHours)) ||
    (typeof effectiveHours === 'number' && !Number.isNaN(effectiveHours));
  if (hasScheduleData) {
    return normalizeScheduleRow(record) || {
      id: record?.TaskId || record?.['Task ID'] || null,
      baseDurationHours: baseHours ?? null,
      effectiveDurationHours: effectiveHours ?? null
    };
  }
  return null;
}

function parseHoursValue(value) {
  if (value === undefined || value === null) return null;
  if (typeof value === 'number' && Number.isFinite(value)) {
    return value;
  }
  const text = String(value).trim();
  if (!text) return null;
  const normalized = text.replace(/\s+/g, '').toLowerCase();
  const simple = Number(normalized);
  if (!Number.isNaN(simple)) {
    return simple;
  }

  const match = normalized.match(/^(-?\d+(?:\.\d+)?)(h|d|w|mo)?$/);
  if (!match) {
    return null;
  }
  const magnitude = Number(match[1]);
  if (Number.isNaN(magnitude)) {
    return null;
  }
  const unit = match[2] || 'h';
  if (unit === 'h') return magnitude;
  if (unit === 'd') return magnitude * HOURS_PER_DAY;
  if (unit === 'w') return magnitude * HOURS_PER_DAY * 5;
  if (unit === 'mo') return magnitude * HOURS_PER_DAY * 20;
  return null;
}

function normalizeScheduleHours(hours, fallbackDays) {
  if (typeof hours === 'number' && !Number.isNaN(hours)) {
    return hours;
  }
  if (typeof fallbackDays === 'number' && !Number.isNaN(fallbackDays)) {
    return fallbackDays * HOURS_PER_DAY;
  }
  return null;
}

function hoursToDays(hours) {
  if (typeof hours !== 'number' || Number.isNaN(hours)) return null;
  return hours / HOURS_PER_DAY;
}

function describeDurationFromHours(hours) {
  if (typeof hours !== 'number' || Number.isNaN(hours)) return null;
  const days = hoursToDays(hours);
  const display = typeof days === 'number' ? formatDurationLabel(days) : '';
  return {
    days,
    display,
    detail: `${formatNumber(days)} days (${formatNumber(hours)} hours)`
  };
}

function formatHoursDetail(hours) {
  if (typeof hours !== 'number' || Number.isNaN(hours)) return '';
  const days = hoursToDays(hours);
  if (typeof days === 'number' && !Number.isNaN(days)) {
    return `${formatNumber(days)} days (${formatNumber(hours)} hours)`;
  }
  return `${formatNumber(hours)} hours`;
}

function describeAggregateDuration(label, hours) {
  if (!label) return '';
  if (typeof hours !== 'number' || Number.isNaN(hours)) return '';
  const detail = formatHoursDetail(hours);
  const days = hoursToDays(hours);
  const segments = [];
  if (detail) {
    segments.push(`${label} ${detail}`);
  } else {
    segments.push(label);
  }
  if (typeof days === 'number' && !Number.isNaN(days)) {
    segments.push(formatWeeksLabel(days));
  }
  return segments.join(' · ');
}

function parseDateValue(value) {
  if (!value) return null;
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return new Date(value.getTime());
  }
  const text = String(value).trim();
  if (!text) return null;
  const parsed = new Date(text);
  if (Number.isNaN(parsed.getTime())) {
    return null;
  }
  return parsed;
}

function resolveFirstDate(candidates) {
  if (!Array.isArray(candidates)) return null;
  for (const candidate of candidates) {
    const parsed = parseDateValue(candidate);
    if (parsed) {
      return parsed;
    }
  }
  return null;
}

function formatScheduleDate(date, fallback = '') {
  if (date instanceof Date && !Number.isNaN(date.getTime())) {
    const year = date.getFullYear();
    const month = String(date.getMonth() + 1).padStart(2, '0');
    const day = String(date.getDate()).padStart(2, '0');
    const hours = String(date.getHours()).padStart(2, '0');
    const minutes = String(date.getMinutes()).padStart(2, '0');
    return `${year}-${month}-${day} ${hours}:${minutes}`;
  }
  if (fallback && typeof fallback === 'string') {
    return fallback.trim();
  }
  return '';
}

function computeScheduleSummary(entries, metadata = null, tasks = []) {
  const scheduleEntries = Array.isArray(entries) ? entries : [];
  const taskList = Array.isArray(tasks) ? tasks : [];

  const metadataStart = metadata?.projectStartDate
    ? parseDateValue(metadata.projectStartDate)
    : metadata?.projectStart
    ? parseDateValue(metadata.projectStart)
    : null;
  const metadataFinish = metadata?.projectFinishDate
    ? parseDateValue(metadata.projectFinishDate)
    : metadata?.projectFinish
    ? parseDateValue(metadata.projectFinish)
    : null;

  const metadataHours =
    typeof metadata?.totalDurationHours === 'number' && !Number.isNaN(metadata.totalDurationHours)
      ? metadata.totalDurationHours
      : null;
  const metadataDays =
    typeof metadata?.totalDurationDays === 'number' && !Number.isNaN(metadata.totalDurationDays)
      ? metadata.totalDurationDays
      : metadataHours !== null
      ? hoursToDays(metadataHours)
      : null;

  const l1Tasks = taskList.filter((task) => task && task.level === 1);
  let primaryLevel = null;
  taskList.forEach((task) => {
    if (!task || typeof task.level !== 'number' || Number.isNaN(task.level)) {
      return;
    }
    primaryLevel = primaryLevel === null ? task.level : Math.min(primaryLevel, task.level);
  });
  const primaryTasks =
    l1Tasks.length > 0
      ? l1Tasks
      : primaryLevel !== null
      ? taskList.filter((task) => task && task.level === primaryLevel)
      : taskList;

  let baseHoursTotal = 0;
  let effectiveHoursTotal = 0;
  let hasBaseHours = false;
  let hasEffectiveHours = false;
  let projectStartDate = null;
  let projectFinishDate = null;

  primaryTasks.forEach((task) => {
    if (!task) return;
    const schedule = task.schedule || null;

    const baseHours = normalizeScheduleHours(schedule?.baseDurationHours, task.baseDurationDays);
    if (typeof baseHours === 'number' && !Number.isNaN(baseHours)) {
      baseHoursTotal += baseHours;
      hasBaseHours = true;
    }

    const effectiveHours = normalizeScheduleHours(
      schedule?.effectiveDurationHours,
      task.calculatedDurationDays
    );
    if (typeof effectiveHours === 'number' && !Number.isNaN(effectiveHours)) {
      effectiveHoursTotal += effectiveHours;
      hasEffectiveHours = true;
    }

    const startDate = resolveFirstDate([
      schedule?.startDate,
      schedule?.startDateDisplay,
      schedule?.startDateRaw,
      task.raw?.StartDate,
      task.raw?.['Start Date'],
      task.raw?.Start
    ]);
    if (startDate) {
      projectStartDate = projectStartDate && projectStartDate <= startDate ? projectStartDate : startDate;
    }

    const finishDate = resolveFirstDate([
      schedule?.finishDate,
      schedule?.finishDateDisplay,
      schedule?.finishDateRaw,
      task.raw?.FinishDate,
      task.raw?.['Finish Date'],
      task.raw?.Finish
    ]);
    if (finishDate) {
      projectFinishDate = projectFinishDate && projectFinishDate >= finishDate ? projectFinishDate : finishDate;
    }
  });

  if (!projectStartDate || !projectFinishDate) {
    scheduleEntries.forEach((entry) => {
      if (!entry) return;
      const startDate = resolveFirstDate([
        entry.startDate,
        entry.startDateDisplay,
        entry.startDateRaw
      ]);
      if (startDate) {
        projectStartDate = projectStartDate && projectStartDate <= startDate ? projectStartDate : startDate;
      }
      const finishDate = resolveFirstDate([
        entry.finishDate,
        entry.finishDateDisplay,
        entry.finishDateRaw
      ]);
      if (finishDate) {
        projectFinishDate = projectFinishDate && projectFinishDate >= finishDate ? projectFinishDate : finishDate;
      }
    });
  }

  const baseHoursValue = hasBaseHours ? baseHoursTotal : null;
  const effectiveHoursValue = hasEffectiveHours ? effectiveHoursTotal : null;

  const baseDaysValue =
    typeof baseHoursValue === 'number' && !Number.isNaN(baseHoursValue)
      ? hoursToDays(baseHoursValue)
      : null;
  const effectiveDaysValue =
    typeof effectiveHoursValue === 'number' && !Number.isNaN(effectiveHoursValue)
      ? hoursToDays(effectiveHoursValue)
      : null;

  const totalDurationHours =
    typeof effectiveHoursValue === 'number' && !Number.isNaN(effectiveHoursValue)
      ? effectiveHoursValue
      : typeof baseHoursValue === 'number' && !Number.isNaN(baseHoursValue)
      ? baseHoursValue
      : metadataHours;

  const totalDurationDays =
    typeof totalDurationHours === 'number' && !Number.isNaN(totalDurationHours)
      ? hoursToDays(totalDurationHours)
      : metadataDays;

  const totalDurationDisplay =
    (typeof effectiveDaysValue === 'number' && !Number.isNaN(effectiveDaysValue)
      ? formatDurationLabel(effectiveDaysValue)
      : '') ||
    (typeof baseDaysValue === 'number' && !Number.isNaN(baseDaysValue)
      ? formatDurationLabel(baseDaysValue)
      : '') ||
    (metadata?.totalDurationDisplay && metadata.totalDurationDisplay.trim()) ||
    (typeof totalDurationDays === 'number' && !Number.isNaN(totalDurationDays)
      ? formatDurationLabel(totalDurationDays)
      : '');

  const startDateValue = projectStartDate || metadataStart;
  const finishDateValue = projectFinishDate || metadataFinish;

  const summary = {
    totalDurationHours,
    totalDurationDays,
    totalDurationDisplay,
    totalBaseDurationHours: baseHoursValue,
    totalBaseDurationDays: baseDaysValue,
    totalBaseDurationDisplay:
      typeof baseDaysValue === 'number' && !Number.isNaN(baseDaysValue)
        ? formatDurationLabel(baseDaysValue)
        : '',
    totalEffectiveDurationHours: effectiveHoursValue,
    totalEffectiveDurationDays: effectiveDaysValue,
    totalEffectiveDurationDisplay:
      typeof effectiveDaysValue === 'number' && !Number.isNaN(effectiveDaysValue)
        ? formatDurationLabel(effectiveDaysValue)
        : '',
    projectStartDate: startDateValue,
    projectFinishDate: finishDateValue,
    projectStartDisplay: formatScheduleDate(
      startDateValue,
      metadata?.projectStartDisplay || metadata?.projectStart || metadata?.projectStartDate || ''
    ),
    projectFinishDisplay: formatScheduleDate(
      finishDateValue,
      metadata?.projectFinishDisplay || metadata?.projectFinish || metadata?.projectFinishDate || ''
    ),
    metadata: metadata || null,
    primaryLevel
  };

  return summary;
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

function parseNormalizedDependencies(value) {
  if (value === undefined || value === null) {
    return { ids: [], types: [] };
  }

  const text = String(value).trim();
  if (!text) {
    return { ids: [], types: [] };
  }

  const ids = [];
  const types = [];
  const seen = new Set();

  text.split(/[;,\n]/)
    .map((part) => part.trim())
    .filter(Boolean)
    .forEach((entry) => {
      const match = entry.match(/(FS|FF|SS|SF)$/i);
      let dependencyType = 'FS';
      let idPortion = entry;
      if (match && typeof match.index === 'number') {
        dependencyType = match[1].toUpperCase();
        idPortion = entry.slice(0, match.index);
      }
      const sanitized = idPortion.replace(/[^A-Za-z0-9_.-]+/g, '').trim();
      if (!sanitized || seen.has(sanitized)) {
        return;
      }
      seen.add(sanitized);
      ids.push(sanitized);
      types.push(dependencyType);
    });

  return { ids, types };
}

function normalizeBoolean(value) {
  if (typeof value === 'boolean') return value;
  const normalized = String(value || '').toLowerCase();
  if (!normalized) return false;
  return ['true', 'yes', 'y', '1'].includes(normalized);
}

function readFirstAvailable(record, keys) {
  for (const key of keys) {
    if (!key) continue;
    const value = record[key];
    if (value === undefined || value === null) continue;
    const text = typeof value === 'string' ? value.trim() : value;
    if (typeof text === 'string' && text === '') continue;
    return value;
  }
  return '';
}

function parseDurationValue(value) {
  if (value === undefined || value === null) {
    return { display: '', days: null, hasValue: false };
  }
  const text = String(value).trim();
  if (!text) {
    return { display: '', days: null, hasValue: false };
  }
  const cleaned = text.replace(/\?/g, '').trim();
  if (!cleaned) {
    return { display: '', days: null, hasValue: false };
  }

  const normalized = cleaned.replace(/\s+/g, '').toLowerCase();
  let days = null;

  if (/^-?\d+(?:\.\d+)?$/.test(normalized)) {
    days = Number(normalized);
  } else if (/^-?\d+(?:\.\d+)?d$/.test(normalized)) {
    days = Number(normalized.slice(0, -1));
  } else if (/^-?\d+(?:\.\d+)?w$/.test(normalized)) {
    days = Number(normalized.slice(0, -1)) * 7;
  } else {
    const wordMatch = normalized.match(/^(-?\d+(?:\.\d+)?)(days?|weeks?)$/);
    if (wordMatch) {
      const valueNum = Number(wordMatch[1]);
      if (!Number.isNaN(valueNum)) {
        days = wordMatch[2].startsWith('week') ? valueNum * 7 : valueNum;
      }
    }
  }

  const parsedDays = typeof days === 'number' && !Number.isNaN(days) ? days : null;
  let display = cleaned.replace(/\s+/g, '');
  if (parsedDays !== null && !/[a-z]$/i.test(display)) {
    display = formatDurationLabel(parsedDays);
  }

  return { display, days: parsedDays, hasValue: true };
}

function formatNumber(value) {
  if (typeof value !== 'number' || Number.isNaN(value)) return '0';
  return Number(value.toFixed(2)).toString();
}

function formatDurationLabel(days) {
  if (typeof days !== 'number' || Number.isNaN(days)) return '';
  return `${formatNumber(days)}d`;
}

function formatWeeksLabel(days) {
  if (typeof days !== 'number' || Number.isNaN(days)) return '0 weeks';
  return `${formatNumber(days / 7)} weeks`;
}

function durationsClose(a, b) {
  if (typeof a !== 'number' || typeof b !== 'number') return false;
  return Math.abs(a - b) <= 0.01;
}

function describeSingleDuration(days, display, hasValue) {
  if (typeof days === 'number' && !Number.isNaN(days)) {
    return {
      display: display && display.trim() ? display : formatDurationLabel(days),
      detail: `${formatNumber(days)} days · ${formatWeeksLabel(days)}`
    };
  }
  if (hasValue && display && display.trim()) {
    return {
      display: display.trim(),
      detail: `Recorded as ${display.trim()}`
    };
  }
  return { display: '', detail: '' };
}

function buildDurationSummaryForTask(task, options = {}) {
  const includeLabel = Boolean(options.includeLabel);
  const label = includeLabel ? `${task.name} (#${task.id})` : '';

  const hasCalculated = Boolean(
    task.hasCalculatedDuration ||
      (task.calculatedDurationDisplay && task.calculatedDurationDisplay.trim()) ||
      typeof task.calculatedDurationDays === 'number'
  );
  const hasBase = Boolean(
    task.hasBaseDuration ||
      (task.baseDurationDisplay && task.baseDurationDisplay.trim()) ||
      typeof task.baseDurationDays === 'number'
  );

  const calcInfo = describeSingleDuration(
    task.calculatedDurationDays,
    task.calculatedDurationDisplay,
    hasCalculated
  );
  if (task.calculatedDurationDetail) {
    calcInfo.detail = calcInfo.detail
      ? `${calcInfo.detail} · ${task.calculatedDurationDetail}`
      : task.calculatedDurationDetail;
  }
  const baseInfo = describeSingleDuration(task.baseDurationDays, task.baseDurationDisplay, hasBase);
  if (task.baseDurationDetail) {
    baseInfo.detail = baseInfo.detail
      ? `${baseInfo.detail} · ${task.baseDurationDetail}`
      : task.baseDurationDetail;
  }

  if (hasCalculated && calcInfo.display) {
    const detailParts = [];
    detailParts.push(includeLabel ? `Rolled-up for ${label}` : 'Rolled-up duration');
    if (calcInfo.detail) {
      detailParts.push(calcInfo.detail);
    }
    if (hasBase && baseInfo.display) {
      if (
        typeof task.calculatedDurationDays === 'number' &&
        typeof task.baseDurationDays === 'number' &&
        durationsClose(task.calculatedDurationDays, task.baseDurationDays)
      ) {
        detailParts.push('Matches base duration');
      } else if (baseInfo.detail) {
        detailParts.push(
          includeLabel ? `Base for ${label} · ${baseInfo.detail}` : `Base duration · ${baseInfo.detail}`
        );
      } else {
        detailParts.push(includeLabel ? `Base for ${label}` : 'Base duration');
      }
    }
    appendScheduleDetails(detailParts, task.schedule);
    return {
      display: calcInfo.display,
      detail: detailParts.join(' · ')
    };
  }

  if (hasBase && baseInfo.display) {
    const detailParts = [];
    detailParts.push(includeLabel ? `Base for ${label}` : 'Base duration');
    if (baseInfo.detail) {
      detailParts.push(baseInfo.detail);
    }
    appendScheduleDetails(detailParts, task.schedule);
    return {
      display: baseInfo.display,
      detail: detailParts.join(' · ')
    };
  }

  const fallbackDetails = [];
  if (calcInfo.detail) fallbackDetails.push(calcInfo.detail);
  if (baseInfo.detail) fallbackDetails.push(baseInfo.detail);
  appendScheduleDetails(fallbackDetails, task.schedule);

  return {
    display: '—',
    detail:
      fallbackDetails.length > 0
        ? fallbackDetails.join(' · ')
        : includeLabel
        ? `No duration data recorded for ${label}.`
        : 'No duration data recorded for this task.'
  };
}

function appendScheduleDetails(parts, schedule) {
  if (!Array.isArray(parts) || !schedule) return;
  const detailSegments = [];
  if (schedule.startDateDisplay) {
    detailSegments.push(`Start ${schedule.startDateDisplay}`);
  } else if (schedule.startDateRaw) {
    detailSegments.push(`Start ${schedule.startDateRaw}`);
  }
  if (schedule.finishDateDisplay) {
    detailSegments.push(`Finish ${schedule.finishDateDisplay}`);
  } else if (schedule.finishDateRaw) {
    detailSegments.push(`Finish ${schedule.finishDateRaw}`);
  }
  if (typeof schedule.totalSlackHours === 'number' && !Number.isNaN(schedule.totalSlackHours)) {
    detailSegments.push(`Total slack ${formatHoursDetail(schedule.totalSlackHours)}`);
  }
  if (typeof schedule.freeSlackHours === 'number' && !Number.isNaN(schedule.freeSlackHours)) {
    detailSegments.push(`Free slack ${formatHoursDetail(schedule.freeSlackHours)}`);
  }
  if (schedule.isCritical) {
    detailSegments.push('Critical path task');
  }
  if (detailSegments.length > 0) {
    parts.push(detailSegments.join(' · '));
  }
}

function computeScheduleTotalDays() {
  if (
    dataset?.scheduleSummary &&
    typeof dataset.scheduleSummary.totalDurationDays === 'number' &&
    !Number.isNaN(dataset.scheduleSummary.totalDurationDays)
  ) {
    return dataset.scheduleSummary.totalDurationDays;
  }
  if (!dataset || !dataset.tasks || dataset.tasks.length === 0) {
    return null;
  }

  let rootIds = dataset.tasks
    .filter((task) => task.predecessors.length === 0)
    .map((task) => task.id);

  if (rootIds.length === 0) {
    let minLevel = Infinity;
    dataset.tasks.forEach((task) => {
      if (typeof task.level === 'number') {
        minLevel = Math.min(minLevel, task.level);
      }
    });
    if (minLevel !== Infinity) {
      rootIds = dataset.tasks.filter((task) => task.level === minLevel).map((task) => task.id);
    }
  }

  if (rootIds.length === 0) {
    return null;
  }

  let total = 0;
  let hasValue = false;

  rootIds.forEach((taskId) => {
    const task = dataset.tasksById.get(taskId);
    if (!task) return;
    let value = null;
    if (typeof task.calculatedDurationDays === 'number' && !Number.isNaN(task.calculatedDurationDays)) {
      value = task.calculatedDurationDays;
    } else if (typeof task.baseDurationDays === 'number' && !Number.isNaN(task.baseDurationDays)) {
      value = task.baseDurationDays;
    }
    if (value !== null) {
      total += value;
      hasValue = true;
    }
  });

  return hasValue ? total : null;
}

function describeSelectedTaskDuration() {
  if (!dataset || !selectedTask || selectedTask.placeholder) {
    return {
      display: '—',
      detail: 'Select a task to inspect its duration.'
    };
  }

  const task = dataset.tasksById.get(selectedTask.id);
  if (!task) {
    return {
      display: '—',
      detail: 'The selected task is no longer part of this dataset.'
    };
  }

  return buildDurationSummaryForTask(task);
}

function describeCurrentPathDuration() {
  if (!dataset || hierarchySelections.size === 0) {
    return {
      display: '—',
      detail: `Choose a Level ${getPrimaryLevel()} milestone to review its rolled-up timing.`
    };
  }

  const entries = Array.from(hierarchySelections.entries()).sort((a, b) => a[0] - b[0]);
  const [, taskId] = entries[entries.length - 1];
  const task = dataset.tasksById.get(taskId);

  if (!task) {
    return {
      display: '—',
      detail: 'The selected path is no longer available in this dataset.'
    };
  }

  return buildDurationSummaryForTask(task, { includeLabel: true });
}

function renderDurationSummary() {
  if (
    !dom.totalDurationDisplay ||
    !dom.totalDurationDetail ||
    !dom.selectedDurationDisplay ||
    !dom.selectedDurationDetail ||
    !dom.pathDurationDisplay ||
    !dom.pathDurationDetail
  ) {
    return;
  }

  if (!dataset) {
    dom.totalDurationDisplay.textContent = '—';
    dom.totalDurationDetail.textContent = 'Upload a schedule to calculate the overall timeline.';
    dom.selectedDurationDisplay.textContent = '—';
    dom.selectedDurationDetail.textContent = 'Select a task to inspect its duration.';
    dom.pathDurationDisplay.textContent = '—';
    dom.pathDurationDetail.textContent = 'Choose a top-level milestone to review its rolled-up timing.';
    return;
  }

  const scheduleSummary = dataset.scheduleSummary || dataset.metadata?.scheduleSummary || null;
  const hasDurationData = dataset.tasks.some(
    (task) =>
      typeof task.calculatedDurationDays === 'number' ||
      typeof task.baseDurationDays === 'number' ||
      (task.calculatedDurationDisplay && task.calculatedDurationDisplay.trim()) ||
      (task.baseDurationDisplay && task.baseDurationDisplay.trim())
  );

  let totalDays =
    typeof scheduleSummary?.totalDurationDays === 'number'
      ? scheduleSummary.totalDurationDays
      : typeof dataset.metadata?.totalDurationDays === 'number'
      ? dataset.metadata.totalDurationDays
      : null;

  if (totalDays === null || Number.isNaN(totalDays)) {
    totalDays = computeScheduleTotalDays();
  }

  const totalDisplay =
    (scheduleSummary?.totalDurationDisplay && scheduleSummary.totalDurationDisplay.trim()) ||
    (dataset.metadata?.totalDurationDisplay && dataset.metadata.totalDurationDisplay.trim()) ||
    (totalDays !== null && !Number.isNaN(totalDays) ? formatDurationLabel(totalDays) : '—');

  dom.totalDurationDisplay.textContent = totalDisplay || '—';
  const totalParts = [];
  if (
    typeof scheduleSummary?.totalEffectiveDurationHours === 'number' &&
    !Number.isNaN(scheduleSummary.totalEffectiveDurationHours)
  ) {
    totalParts.push(describeAggregateDuration('Effective', scheduleSummary.totalEffectiveDurationHours));
  }
  if (
    typeof scheduleSummary?.totalBaseDurationHours === 'number' &&
    !Number.isNaN(scheduleSummary.totalBaseDurationHours)
  ) {
    totalParts.push(describeAggregateDuration('Base', scheduleSummary.totalBaseDurationHours));
  }
  const startDisplay = scheduleSummary?.projectStartDisplay;
  const finishDisplay = scheduleSummary?.projectFinishDisplay;
  if ((startDisplay && startDisplay.trim()) || (finishDisplay && finishDisplay.trim())) {
    const windowParts = [];
    if (startDisplay && startDisplay.trim()) {
      windowParts.push(`Start ${startDisplay.trim()}`);
    }
    if (finishDisplay && finishDisplay.trim()) {
      windowParts.push(`Finish ${finishDisplay.trim()}`);
    }
    if (windowParts.length > 0) {
      totalParts.push(windowParts.join(' · '));
    }
  }

  if (totalParts.length > 0) {
    dom.totalDurationDetail.textContent = totalParts.join(' · ');
  } else if (hasDurationData) {
    dom.totalDurationDetail.textContent = 'Unable to compute a rolled-up total from the provided durations.';
  } else {
    dom.totalDurationDetail.textContent = 'No duration data found in this dataset.';
  }

  const selectedSummary = describeSelectedTaskDuration();
  dom.selectedDurationDisplay.textContent = selectedSummary.display;
  dom.selectedDurationDetail.textContent = selectedSummary.detail;

  const pathSummary = describeCurrentPathDuration();
  dom.pathDurationDisplay.textContent = pathSummary.display;
  dom.pathDurationDetail.textContent = pathSummary.detail;
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

  const context =
    '<span class="legend-item"><span class="legend-swatch" style="background:transparent;border:2px dashed rgba(148,163,184,0.8);"></span>Connected context</span>';
  const dimmed =
    '<span class="legend-item"><span class="legend-swatch" style="background:rgba(148,163,184,0.4);border:1px solid rgba(148,163,184,0.6);"></span>Dimmed outside focus</span>';

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
    const primaryLevel = getPrimaryLevel();
    dom.levelControls.innerHTML = `<p class="empty-state">Select a Level ${primaryLevel} milestone to begin, then expand follow-on work from the network.</p>`;
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

  const rolledInfo = describeSingleDuration(
    task.calculatedDurationDays,
    task.calculatedDurationDisplay,
    task.hasCalculatedDuration
  );
  const baseInfo = describeSingleDuration(
    task.baseDurationDays,
    task.baseDurationDisplay,
    task.hasBaseDuration
  );
  const durationSource = rolledInfo.display ? { info: rolledInfo, type: 'rolled' } : baseInfo.display ? { info: baseInfo, type: 'base' } : null;

  if (durationSource) {
    const badge = document.createElement('span');
    badge.className = 'duration';
    badge.textContent = durationSource.info.display;
    const tooltipParts = [];
    if (durationSource.type === 'rolled') {
      tooltipParts.push(rolledInfo.detail || 'Rolled-up duration');
      if (baseInfo.display) {
        if (
          typeof task.calculatedDurationDays === 'number' &&
          typeof task.baseDurationDays === 'number' &&
          durationsClose(task.calculatedDurationDays, task.baseDurationDays)
        ) {
          tooltipParts.push('Matches base duration');
        } else if (baseInfo.detail) {
          tooltipParts.push(`Base ${baseInfo.detail}`);
        } else {
          tooltipParts.push(`Base ${baseInfo.display}`);
        }
      }
    } else if (durationSource.type === 'base') {
      tooltipParts.push(baseInfo.detail || 'Base duration');
    }
    badge.title = tooltipParts.filter(Boolean).join(' · ');
    card.appendChild(badge);
  }

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

function getPrimaryLevel() {
  if (!dataset) return 1;
  const summaryLevel = dataset.scheduleSummary?.primaryLevel;
  if (typeof summaryLevel === 'number' && !Number.isNaN(summaryLevel)) {
    return summaryLevel;
  }
  if (Array.isArray(dataset.levels) && dataset.levels.length > 0) {
    const sorted = dataset.levels.slice().sort((a, b) => a - b);
    return sorted[0] ?? 1;
  }
  return 1;
}

function getHierarchyLevels() {
  if (!dataset) return [];
  const sorted = dataset.levels.slice().sort((a, b) => a - b);
  return sorted;
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
            'color': '#0b1120',
            'font-size': '13px',
            'font-weight': 600,
            'line-height': '1.3',
            'text-wrap': 'wrap',
            'text-max-width': '170px',
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
            'text-background-color': 'rgba(255,255,255,0.94)',
            'text-background-opacity': 1,
            'text-background-padding': 6,
            'text-background-shape': 'roundrectangle'
          }
        },
        {
          selector: 'node[?isNeighbor]',
          style: {
            'border-style': 'dashed',
            'opacity': 0.75,
            'background-opacity': 0.5,
            'color': '#1f2937',
            'text-background-opacity': 0.9
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
            'text-outline-opacity': 0,
            'text-background-opacity': 0.45
          }
        },
        {
          selector: 'node:selected',
          style: {
            'border-width': 4,
            'border-color': '#facc15',
            'shadow-color': 'rgba(250, 204, 21, 0.45)',
            'shadow-blur': 18,
            'text-background-color': 'rgba(255,255,255,0.98)'
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
            'font-size': '10px',
            'color': '#0f172a',
            'text-wrap': 'wrap',
            'text-max-width': '120px',
            'text-rotation': 'autorotate',
            'text-background-color': 'rgba(255,255,255,0.95)',
            'text-background-opacity': 1,
            'text-background-padding': 4,
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
  if (!dom.taskDetails) return;
  if (!selectedTask) {
    dom.taskDetails.classList.add('empty');
    dom.taskDetails.innerHTML =
      '<p>Select a task from the hierarchy or network to edit its fields, remove it, or manage its dependencies.</p>';
    renderDurationSummary();
    return;
  }

  dom.taskDetails.classList.remove('empty');

  if (selectedTask.placeholder) {
    dom.taskDetails.innerHTML = `
      <h3>External dependency</h3>
      <p>The task <strong>${selectedTask.id}</strong> exists in the dependency chain but is not present in the uploaded schedule.</p>
    `;
    renderDurationSummary();
    return;
  }

  const task = dataset?.tasksById.get(selectedTask.id);
  if (!task) {
    dom.taskDetails.innerHTML = `
      <h3>Task unavailable</h3>
      <p>The selected task is no longer part of the working dataset.</p>
    `;
    renderDurationSummary();
    return;
  }

  selectedTask = task;

  const form = buildTaskEditForm(task);
  dom.taskDetails.innerHTML = '';
  dom.taskDetails.appendChild(form);
  renderDurationSummary();
}

function buildTaskEditForm(task) {
  const form = document.createElement('form');
  form.className = 'task-edit-form';
  form.noValidate = true;

  const rolledInfo = describeSingleDuration(
    task.calculatedDurationDays,
    task.calculatedDurationDisplay,
    task.hasCalculatedDuration
  );
  const baseInfo = describeSingleDuration(
    task.baseDurationDays,
    task.baseDurationDisplay,
    task.hasBaseDuration
  );

  const rolledDisplay = rolledInfo.display || '—';
  const rolledDetail = rolledInfo.detail || 'No rolled-up duration available.';
  const baseDisplay = baseInfo.display || '—';
  const baseDetail = baseInfo.detail || 'No base duration recorded.';
  const schedule = task.schedule || null;
  const startText = schedule?.startDateDisplay || schedule?.startDateRaw || '';
  const finishText = schedule?.finishDateDisplay || schedule?.finishDateRaw || '';
  const totalSlackText =
    typeof schedule?.totalSlackHours === 'number' && !Number.isNaN(schedule.totalSlackHours)
      ? formatHoursDetail(schedule.totalSlackHours)
      : '';
  const freeSlackText =
    typeof schedule?.freeSlackHours === 'number' && !Number.isNaN(schedule.freeSlackHours)
      ? formatHoursDetail(schedule.freeSlackHours)
      : '';
  const startDetail = schedule?.startDateDisplay || schedule?.startDateRaw ? 'Earliest schedule start' : '';
  const finishDetail = schedule?.finishDateDisplay || schedule?.finishDateRaw ? 'Projected finish' : '';
  const totalSlackDetail = totalSlackText ? 'Total slack relative to successors' : '';
  const freeSlackDetail = freeSlackText ? 'Free slack before impacting successors' : '';
  const criticalDetail = schedule
    ? schedule.isCritical
      ? 'This task is on the critical path.'
      : 'Not currently on the critical path.'
    : '';
  const hasScheduleDetails = Boolean(
    (startText && startText.trim()) ||
      (finishText && finishText.trim()) ||
      totalSlackText ||
      freeSlackText ||
      schedule?.isCritical
  );
  const scheduleBlock = schedule && hasScheduleDetails
    ? `
    <div class="task-schedule-overview">
      <div class="schedule-card">
        <span class="schedule-label">Start</span>
        <strong>${escapeHtml(startText || '—')}</strong>
        <small>${escapeHtml(startDetail || '')}</small>
      </div>
      <div class="schedule-card">
        <span class="schedule-label">Finish</span>
        <strong>${escapeHtml(finishText || '—')}</strong>
        <small>${escapeHtml(finishDetail || '')}</small>
      </div>
      <div class="schedule-card">
        <span class="schedule-label">Total slack</span>
        <strong>${escapeHtml(totalSlackText || '—')}</strong>
        <small>${escapeHtml(totalSlackDetail || '')}</small>
      </div>
      <div class="schedule-card">
        <span class="schedule-label">Free slack</span>
        <strong>${escapeHtml(freeSlackText || '—')}</strong>
        <small>${escapeHtml(freeSlackDetail || '')}</small>
      </div>
      <div class="schedule-card">
        <span class="schedule-label">Critical path</span>
        <strong>${schedule.isCritical ? 'Yes' : 'No'}</strong>
        <small>${escapeHtml(criticalDetail || '')}</small>
      </div>
    </div>
  `
    : '';

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
    <div class="task-duration-overview">
      <div class="duration-card">
        <span class="duration-label">Rolled-up duration</span>
        <strong>${escapeHtml(rolledDisplay)}</strong>
        <small>${escapeHtml(rolledDetail)}</small>
      </div>
      <div class="duration-card">
        <span class="duration-label">Base duration</span>
        <strong>${escapeHtml(baseDisplay)}</strong>
        <small>${escapeHtml(baseDetail)}</small>
      </div>
    </div>
    ${scheduleBlock}
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
        '<p class="empty-state soft">Select a task in the hierarchy or network to preview its dependency chain. Click edges in the graph to request removal of a link.</p>';
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
  dataset.preprocessed = false;
  dataset.rawRows = null;
  dataset.processedExcel = null;
  refreshDatasetIndexes();
  cleanupSelections();
  rebuildSelectedTaskIds();

  renderLegend();
  renderLevelControls();
  renderHierarchy();
  renderDependencies();
  updateGraph();
  renderTaskDetails();
  renderDurationSummary();

  updateDownloadButton();
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

function updateDownloadButton() {
  if (!dom.downloadDataset) return;
  const hasServerExcel = Boolean(dataset?.processedExcel);
  const hasRawRows = Boolean(dataset?.rawRows && dataset.rawRows.length);
  const hasTasks = Boolean(dataset?.tasks && dataset.tasks.length);
  const hasData = hasServerExcel || hasRawRows || hasTasks;
  dom.downloadDataset.disabled = !hasData || isProcessingUpload;
  if (!hasData) {
    dom.downloadDataset.textContent = 'Download processed Excel';
    return;
  }
  if (hasServerExcel && !dataset?.modified) {
    dom.downloadDataset.textContent = 'Download preprocessed Excel';
    return;
  }
  const showProcessedLabel = Boolean(!dataset?.modified && dataset?.preprocessed);
  dom.downloadDataset.textContent = showProcessedLabel ? 'Download processed Excel' : 'Download Excel';
}

function downloadCurrentWorkbook() {
  if (!dataset || (dom.downloadDataset && dom.downloadDataset.disabled)) return;
  if (!dataset.modified && dataset.processedExcel?.blob) {
    const filename = dataset.processedExcel.filename || generateDownloadFilename();
    downloadBlob(dataset.processedExcel.blob, filename);
    return;
  }
  if (!dataset.modified && dataset.rawRows && dataset.rawRows.length) {
    exportRowsToExcel(dataset.rawRows);
    return;
  }
  exportDatasetToExcel();
}

function promptForDownload(actionDescription) {
  if (!dataset) return;
  const label = dataset.label || 'schedule';
  const approval = window.confirm(
    `${actionDescription}\n\nA temporary working copy of "${label}" has been updated.` +
      '\nSelect OK to download the latest Excel file now, or Cancel to continue editing.'
  );
  if (approval) {
    downloadCurrentWorkbook();
  }
}

function exportDatasetToExcel() {
  if (!dataset) return;
  const rows = dataset.tasks.map((task) => {
    const row = task.raw ? { ...task.raw } : {};
    row['Task ID'] = task.id;
    row['TaskId'] = task.id;
    row['Task Name'] = task.name;
    row['Name'] = task.name;
    row['Task Level'] = task.levelLabel || `L${task.level}`;
    row['Level'] = task.levelLabel || `L${task.level}`;
    row['Responsible Sub-team'] = task.team || '';
    row['Sub-team'] = task.team || '';
    row['Responsible Function'] = task.function || '';
    row['Function'] = task.function || '';
    row.Scope = task.scope || '';
    row.SCOPE = task.scope || row.SCOPE || '';
    row.Learning_Plan = task.learningPlan ? 'Yes' : 'No';
    row['Learning Plan'] = task.learningPlan ? 'Yes' : 'No';
    row.Critical = task.critical ? 'Yes' : 'No';
    row['Is Critical'] = task.critical ? 'Yes' : 'No';
    row['Predecessors IDs'] = task.predecessors.join(', ');
    row['Predecessor IDs'] = task.predecessors.join(', ');
    row.Predecessors = task.predecessors.join(', ');
    row['Successors IDs'] = task.successors.join(', ');
    row['Successor IDs'] = task.successors.join(', ');
    row.Successors = task.successors.join(', ');
    row['Dependency Type'] = task.dependencyTypes.join(', ');
    row['Dependency Types'] = task.dependencyTypes.join(', ');
    return row;
  });

  exportRowsToExcel(rows);
}

function exportRowsToExcel(rows) {
  if (!rows || rows.length === 0) return;
  const worksheet = XLSX.utils.json_to_sheet(rows);
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, 'Tasks');
  XLSX.writeFile(workbook, generateDownloadFilename());
}

function downloadBlob(blob, filename) {
  if (!(blob instanceof Blob)) return;
  const link = document.createElement('a');
  const url = URL.createObjectURL(blob);
  link.href = url;
  link.download = filename || 'cps_preprocessed.xlsx';
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  URL.revokeObjectURL(url);
}

function generateDownloadFilename() {
  if (!dataset?.modified && dataset?.processedExcel?.filename) {
    return dataset.processedExcel.filename;
  }
  const label = dataset?.label || 'schedule';
  const safe = sanitizeLabel(label);
  return `${safe || 'schedule'}_edited.xlsx`;
}

function buildProcessedFilename(label) {
  const baseName = typeof label === 'string' ? label : '';
  const withoutExtension = baseName.replace(/\.[^.]+$/, '');
  const fallback = withoutExtension || dataset?.label || 'schedule';
  const safe = sanitizeLabel(fallback);
  return `${safe || 'schedule'}_preprocessed.xlsx`;
}

function sanitizeLabel(label) {
  return String(label || '')
    .trim()
    .replace(/[^a-z0-9]+/gi, '_')
    .replace(/^_+|_+$/g, '');
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
