'use strict';

function validDateValue(value) {
  const time = Date.parse(value || '');
  return Number.isFinite(time) ? time : Number.NEGATIVE_INFINITY;
}

function unique(values) {
  return [...new Set(values.filter(Boolean))];
}

function normalizeWorkItem(workItem, stepMap) {
  const elementId = workItem.ElementId || null;
  return {
    userId: workItem.UserId || null,
    elementId,
    stepName: elementId ? (stepMap[elementId] || null) : null,
    dueDateTime: workItem.DueDateTime || null
  };
}

function groupPendingWorkItems(workItems, keyFromSubject, stepMap) {
  const grouped = Object.create(null);
  for (const workItem of workItems || []) {
    if (workItem.Status !== 'Pending') continue;
    const documentKey = keyFromSubject(workItem.Subject);
    if (!documentKey) continue;
    if (!grouped[documentKey]) grouped[documentKey] = [];
    grouped[documentKey].push(normalizeWorkItem(workItem, stepMap));
  }

  for (const documentKey of Object.keys(grouped)) {
    const seen = new Set();
    grouped[documentKey] = grouped[documentKey]
      .filter(item => {
        const signature = [item.userId, item.elementId, item.dueDateTime].join('|');
        if (seen.has(signature)) return false;
        seen.add(signature);
        return true;
      })
      .sort((a, b) => validDateValue(a.dueDateTime) - validDateValue(b.dueDateTime));
  }
  return grouped;
}

function summarizePending(workItems, fallbackUserId) {
  const items = workItems || [];
  const pendingApprovers = unique(items.map(item => item.userId));
  if (!pendingApprovers.length && fallbackUserId) pendingApprovers.push(fallbackUserId);
  const stepNames = unique(items.map(item => item.stepName));
  const elementIds = unique(items.map(item => item.elementId));
  const dueDateTimes = unique(items.map(item => item.dueDateTime));
  return {
    pendingApprover: pendingApprovers.join(', ') || null,
    pendingApprovers,
    pendingWorkItems: items,
    stepName: stepNames[0] || null,
    stepNames,
    stepElementId: elementIds[0] || null,
    stepElementIds: elementIds,
    workItemDueDateTime: dueDateTimes[0] || null,
    workItemDueDateTimes: dueDateTimes
  };
}

function latestCompletedWorkItems(workItems, keyFromSubject) {
  const latest = Object.create(null);
  for (const workItem of workItems || []) {
    if (workItem.Status !== 'Completed') continue;
    const documentKey = keyFromSubject(workItem.Subject);
    if (!documentKey) continue;
    const previous = latest[documentKey];
    if (!previous || validDateValue(workItem.DueDateTime) > validDateValue(previous.DueDateTime)) {
      latest[documentKey] = workItem;
    }
  }
  return latest;
}

module.exports = {
  groupPendingWorkItems,
  latestCompletedWorkItems,
  summarizePending
};
