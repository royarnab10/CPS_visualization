import test from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';

import { buildTaskIndex, normalizeRecords, parseLinkedIds } from '../webapp/dag-data.js';

const datasetUrl = new URL('../webapp/data/amy_new_cps_tasks.json', import.meta.url);
const dataset = JSON.parse(readFileSync(fileURLToPath(datasetUrl), 'utf8'));

const normalizedDataset = normalizeRecords(dataset);

test('normalizeRecords trims whitespace and preserves fields', () => {
  const sample = normalizeRecords([{ 'Task Name': ' Example ', TaskID: ' 42 ' }]);
  assert.equal(sample[0]['Task Name'], 'Example');
  assert.equal(sample[0].TaskID, '42');
});

test('buildTaskIndex creates lookup map with expected tasks', () => {
  const { tasks, tasksById } = buildTaskIndex(normalizedDataset);
  assert.equal(tasks.length, normalizedDataset.length);

  const task = tasksById.get('18');
  assert(task);
  assert.equal(task.name, 'XX Country Sales Samples Due to "Company Issuing Sales Samples to Trade" (date needed)');
  assert.deepEqual(task.predecessors, ['1004']);
  assert.deepEqual(task.successors, ['21']);
  assert.equal(task.phase, 'Launch');
});

test('parseLinkedIds extracts identifiers regardless of separators', () => {
  assert.deepEqual(parseLinkedIds('1004, 1005'), ['1004', '1005']);
  assert.deepEqual(parseLinkedIds('ABC-123 SS; XYZ-789'), ['ABC-123', 'XYZ-789']);
  assert.deepEqual(parseLinkedIds(''), []);
});
