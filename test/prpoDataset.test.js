'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const {
  deriveClock,
  lifecycleStage,
  observationKey,
  openPO,
  refreshWithFallback
} = require('../src/shared/prpoDataset');

function context(options = {}) {
  return {
    numberKeys: new Map([['PO-1', ['scbm|PO-1']]]),
    capture: new Map(),
    packing: new Map(),
    ...options
  };
}

test('P1a separates authoritative stage from its clock', () => {
  const header = { PurchaseOrderNumber: 'PO-1', PurchaseOrderStatus: 'Backorder', DocumentApprovalStatus: 'Confirmed' };
  assert.deepEqual(lifecycleStage(header, 'scbm|PO-1', context()), {
    stage: 'Sent to supplier', reason: 'F_AND_O_ORDER_AND_APPROVAL_STATUS'
  });
});

test('draft and approved orders have a certain not-yet-sent stage', () => {
  for (const approval of ['Draft', 'Approved', 'InReview']) {
    const result = lifecycleStage({ PurchaseOrderNumber: 'PO-1', PurchaseOrderStatus: 'Backorder', DocumentApprovalStatus: approval }, 'scbm|PO-1', context());
    assert.equal(result.stage, 'Not yet sent');
  }
});

test('a packing-slip event advances an open order to Receipt posted', () => {
  const packing = new Map([['scbm|PO-1', [{ DocumentDate: '2026-09-07T10:00:00Z' }]]]);
  const result = lifecycleStage({ PurchaseOrderNumber: 'PO-1', PurchaseOrderStatus: 'Backorder', DocumentApprovalStatus: 'Confirmed' }, 'scbm|PO-1', context({ packing }));
  assert.equal(result.stage, 'Receipt posted');
});

test('final-workbook seeds remain explicitly labelled', () => {
  const existing = {
    ssg_currentstepname: 'Not yet sent',
    ssg_observedpendingsince: '2026-09-07T05:00:00Z',
    ssg_dataqualitynotes: JSON.stringify({
      clockProvenance: 'SEEDED_FROM_FINAL_WORKBOOK',
      workbookValue: '2026-09-07 09:00:00',
      workbookExportTimestamp: '2026-09-07T05:30:00Z'
    })
  };
  const clock = deriveClock({ liveStage: 'Not yet sent', liveEventTimestamp: null }, existing, '2026-09-07T16:00:00Z');
  assert.equal(clock.provenance, 'SEEDED_FROM_FINAL_WORKBOOK');
  assert.equal(clock.label, 'since (from last export)');
});

test('a post-cutover stage change gets a first-observed clock', () => {
  const existing = { ssg_currentstepname: 'Not yet sent', ssg_observedpendingsince: null, ssg_dataqualitynotes: '{}' };
  const clock = deriveClock({ liveStage: 'Sent to supplier', liveEventTimestamp: null }, existing, '2026-09-07T16:00:00Z');
  assert.equal(clock.provenance, 'PENDING_SINCE_FIRST_OBSERVED');
  assert.equal(clock.timestamp, '2026-09-07T16:00:00Z');
});

test('composite PO identity and open population are exact', () => {
  assert.equal(observationKey('SCBM', 'po-1'), 'ifahr-live|PO|scbm|PO-1');
  assert.equal(openPO('Backorder', 'Confirmed'), true);
  assert.equal(openPO('Invoiced', 'Confirmed'), false);
  assert.equal(openPO('Backorder', 'Rejected'), false);
});

test('refresh failures serve stale data only when a last-good revision exists', async () => {
  const failed = async () => { throw new Error('simulated F&O outage'); };
  const stale = await refreshWithFallback(failed, { revision: 'last-good', sourceState: 'LIVE' });
  assert.equal(stale.revision, 'last-good');
  assert.equal(stale.sourceState, 'STALE');
  assert.match(stale.refreshError, /simulated F&O outage/);
  await assert.rejects(refreshWithFallback(failed, null), /simulated F&O outage/);
});
