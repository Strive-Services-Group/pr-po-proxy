'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const {
  groupPendingWorkItems,
  latestCompletedWorkItems,
  summarizePending
} = require('../src/shared/workflow');

const key = subject => (String(subject).match(/PR-\d+/) || [null])[0];

test('preserves parallel pending approvers instead of selecting one', () => {
  const grouped = groupPendingWorkItems([
    { Subject: 'PR-000001 Approval', Status: 'Pending', UserId: 'ayman.g', ElementId: 'A', DueDateTime: '2026-08-06T08:00:00Z' },
    { Subject: 'PR-000001 Approval', Status: 'Pending', UserId: 'patrick.s', ElementId: 'A', DueDateTime: '2026-08-06T08:00:00Z' }
  ], key, { A: 'Director approval' });
  const summary = summarizePending(grouped['PR-000001']);
  assert.deepEqual(summary.pendingApprovers, ['ayman.g', 'patrick.s']);
  assert.equal(summary.pendingApprover, 'ayman.g, patrick.s');
  assert.equal(summary.pendingWorkItems.length, 2);
});

test('deduplicates exact duplicate work items but retains distinct assignments', () => {
  const duplicate = { Subject: 'PR-000002 Approval', Status: 'Pending', UserId: 'user.a', ElementId: 'B', DueDateTime: '2026-08-06T09:00:00Z' };
  const grouped = groupPendingWorkItems([duplicate, { ...duplicate }], key, { B: 'Finance' });
  assert.equal(grouped['PR-000002'].length, 1);
});

test('uses the custom-header user only when no pending work item is available', () => {
  assert.deepEqual(summarizePending([], 'fallback.user').pendingApprovers, ['fallback.user']);
  assert.deepEqual(summarizePending([{ userId: 'live.user' }], 'fallback.user').pendingApprovers, ['live.user']);
});

test('keeps the latest completed work item per document', () => {
  const latest = latestCompletedWorkItems([
    { Subject: 'PR-000003 Approval', Status: 'Completed', UserId: 'old', DueDateTime: '2026-08-01T08:00:00Z' },
    { Subject: 'PR-000003 Approval', Status: 'Completed', UserId: 'new', DueDateTime: '2026-08-02T08:00:00Z' }
  ], key);
  assert.equal(latest['PR-000003'].UserId, 'new');
});
