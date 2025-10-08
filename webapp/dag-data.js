export function normalizeRecords(records) {
  if (!Array.isArray(records)) {
    return [];
  }

  return records
    .map((record) => normalizeRecord(record))
    .filter((record) => record !== null);
}

function normalizeRecord(record) {
  if (!record || typeof record !== "object") {
    return null;
  }

  const normalized = {};
  for (const [key, value] of Object.entries(record)) {
    if (typeof value === "string") {
      normalized[key] = value.trim();
    } else if (value == null) {
      normalized[key] = "";
    } else {
      normalized[key] = String(value);
    }
  }
  return normalized;
}

export function buildTaskIndex(records) {
  const tasks = [];
  const tasksById = new Map();

  for (const record of records) {
    const id = resolveTaskId(record);
    if (!id) {
      continue;
    }

    const task = {
      id,
      name: record["Task Name"] || record["Task"] || "Untitled task",
      phase:
        record["Current SIMPL Phase (visibility)"] ||
        record["Current SIMPL Phase"] ||
        "",
      lego:
        record["Commercial/Technical Lego Block"] ||
        record["Commercial Lego Block"] ||
        record["Technical Lego Block"] ||
        "",
      scope: record["Scope Definition (Arnab)"] || record["Scope"] || "",
      predecessors: parseLinkedIds(
        record.Predecessors || record["Predecessors IDs"] || "",
      ),
      successors: parseLinkedIds(
        record.Successors || record["Successors IDs"] || "",
      ),
      raw: record,
    };

    tasks.push(task);
    tasksById.set(task.id, task);
  }

  return { tasks, tasksById };
}

function resolveTaskId(record) {
  const candidates = [
    record.TaskID,
    record.TaskId,
    record["Task ID"],
    record.id,
    record.ID,
  ];

  for (const candidate of candidates) {
    if (candidate == null) {
      continue;
    }
    const text = String(candidate).trim();
    if (text) {
      return text;
    }
  }
  return "";
}

export function parseLinkedIds(value) {
  if (!value) {
    return [];
  }
  return String(value)
    .split(/[,\n;]/)
    .map((part) => part.trim())
    .filter(Boolean)
    .map((part) => {
      const match = part.match(/([A-Za-z0-9_.-]+)/);
      return match ? match[1] : null;
    })
    .filter(Boolean);
}
