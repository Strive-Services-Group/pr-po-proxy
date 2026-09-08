'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const { buildItems } = require('../src/functions/prpoEmail');

function pr(status, number) {
  return {
    'Purchase requisition': number,
    'Status': status,
    'Step name': 'Unit prices updated in PR lines',
    'Total amount': 100,
    'Department': 'Building Services'
  };
}

test('temporary workbook fallback preserves the successful legacy PR population', () => {
  const items = buildItems([
    pr('In review', 'PR-1'),
    pr('Approved', 'PR-2'),
    pr('Draft', 'PR-3'),
    pr('Closed', 'PR-4'),
    pr('Rejected', 'PR-5'),
    pr('Cancelled', 'PR-6'),
    { ...pr('In review', 'PR-7'), 'Step name': 'Sourcing' }
  ], []);
  assert.deepEqual(items.map(item => item.ref), ['PR-1', 'PR-2', 'PR-3']);
  assert.ok(items.every(item => item.owner === 'dinesh.laxman'));
});
