'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const { buildItems } = require('../src/functions/prpoEmail');

function pr(status, number) {
  return {
    'Purchase requisition': number,
    'Status': status,
    'Step name': 'Sourcing',
    'Total amount': 100,
    'Department': 'Building Services'
  };
}

test('email PR population is exactly In review and Approved', () => {
  const items = buildItems([
    pr('In review', 'PR-1'),
    pr('Approved', 'PR-2'),
    pr('Draft', 'PR-3'),
    pr('Closed', 'PR-4'),
    pr('Rejected', 'PR-5'),
    pr('Cancelled', 'PR-6')
  ], []);
  assert.deepEqual(items.map(item => item.ref), ['PR-1', 'PR-2']);
});
