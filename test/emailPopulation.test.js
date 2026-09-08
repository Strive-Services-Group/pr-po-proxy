'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const XLSX = require('xlsx');
const workRule = require('../work-class-rule.json');
const { buildItems, buildDivision, buildXlsxBase64, groupByOwner, personalPool, buildPersonal, DIVS, parseXlsx } = require('../src/functions/prpoEmail');

function pr(status, number) {
  return {
    'Purchase requisition': number,
    'Status': status,
    'Step name': 'Unit prices updated in PR lines',
    'Stage reason code': 'ACTIVE_LINES_PRICED',
    'Total amount': 100,
    'Department': 'Building Services',
    'Location': 'TEST SITE',
    'Name': 'Test requisition',
    'Pending Approver/User': 'dinesh.laxman',
    'Preparer': 'dinesh.laxman',
    'Accepted By/Assign To': 'dinesh.laxman',
    'Created date': '2026-09-01T00:00:00Z'
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
    { ...pr('In review', 'PR-7'), 'Step name': 'Sourcing', 'Stage reason code': '' }
  ], []);
  assert.deepEqual(items.map(item => item.ref), ['PR-1', 'PR-2', 'PR-3']);
  assert.ok(items.every(item => item.owner === 'dinesh.laxman'));
});

test('every Stage reason code becomes the shared plain-English class', () => {
  for (const [code, rule] of Object.entries(workRule.classes)) {
    const row = pr('In review', `PR-${code}`);
    row['Stage reason code'] = code;
    if (rule.holderMode === 'preparer') row['Preparer'] = 'dinesh.laxman';
    const item = buildItems([row], [])[0];
    assert.equal(item.workClassCode, code);
    assert.equal(item.workClass, rule.label);
    assert.equal(item.workAction, rule.action);
  }
});

test('no named owner is excluded from personal mail and included in procurement digest', () => {
  const row = pr('In review', 'PR-NO-OWNER');
  row['Department'] = 'Surveying Services';
  row['Pending Approver/User'] = 'No named owner — no operations person mapped for Surveying Services';
  row['Preparer'] = row['Pending Approver/User'];
  row['Accepted By/Assign To'] = row['Pending Approver/User'];
  const items = buildItems([row], []);
  assert.equal(items[0].noNamedOwner, true);
  assert.equal(personalPool(items).length, 0);
  const procurement = buildDivision(DIVS.find(d => d.key === 'procurement'), items, {});
  assert.match(procurement.html, /No named owner/);
  assert.match(procurement.html, /Surveying Services/);
});

test('personal email groups by class then department and preserves Pending Client', () => {
  const rows = [
    { ...pr('In review', 'PR-BUILDING'), 'Department': 'Building Services', 'Pending Approver/User': 'dinesh.laxman' },
    { ...pr('In review', 'PR-LANDSCAPING'), 'Department': 'Landscaping Services', 'Pending Approver/User': 'dinesh.laxman' },
    { ...pr('In review', 'PR-RAISER'), 'Stage reason code': 'NO_CURRENT_WORK_ITEM', 'Step name': 'PurchReqReviewTask', 'Preparer': 'dinesh.laxman', 'Pending Approver/User': 'dinesh.laxman' }
  ];
  const grouped = groupByOwner(personalPool(buildItems(rows, [])));
  const out = buildPersonal(grouped.find(person => person.key === 'dinesh.laxman'), {});
  assert.deepEqual(out.classCounts, { ACTIVE_LINES_PRICED: 2, NO_CURRENT_WORK_ITEM: 1 });
  assert.deepEqual(out.sections[0].departments, [
    { department: 'Building Services', count: 1 },
    { department: 'Landscaping Services', count: 1 }
  ]);
  assert.match(out.html, /Shape of your PR queue/);
  assert.match(out.html, /Prices are in/);
  assert.match(out.html, /Pending Client/);
  assert.match(out.html, /PR-BUILDING/);
  assert.match(out.html, /PR-LANDSCAPING/);
  assert.doesNotMatch(out.html, /kept in the attached Excel/);
});

test('personal email sorts oldest first and largest first when ages tie', () => {
  const rows = [
    { ...pr('In review', 'PR-NEWER'), 'Created date': '2026-09-01T00:00:00Z', 'Total amount': 900 },
    { ...pr('In review', 'PR-OLDER'), 'Created date': '2026-01-01T00:00:00Z', 'Total amount': 10 },
    { ...pr('In review', 'PR-SMALLER'), 'Created date': '2026-06-01T00:00:00Z', 'Total amount': 50 },
    { ...pr('In review', 'PR-LARGER'), 'Created date': '2026-06-01T00:00:00Z', 'Total amount': 200 }
  ];
  const person = groupByOwner(personalPool(buildItems(rows, [])))[0];
  const html = buildPersonal(person, {}).html;
  assert.ok(html.indexOf('PR-OLDER') < html.indexOf('PR-LARGER'));
  assert.ok(html.indexOf('PR-LARGER') < html.indexOf('PR-SMALLER'));
  assert.ok(html.indexOf('PR-SMALLER') < html.indexOf('PR-NEWER'));
});

test('Pending Internal remains distinct from Pending Client', () => {
  const row = { ...pr('In review', 'PR-INTERNAL'), 'Step name': 'Quotation shared to Operations for confirmation' };
  const person = groupByOwner(personalPool(buildItems([row], [])))[0];
  const out = buildPersonal(person, {});
  assert.equal(out.fil[0].pendingSide, 'Pending Internal');
  assert.match(out.html, /Pending Internal/);
  assert.match(out.html, /PR-INTERNAL/);
});

test('email attachment exposes class code, class, action and age band', async () => {
  const item = buildItems([pr('In review', 'PR-ATTACHMENT')], [])[0];
  const base64 = await buildXlsxBase64([item], { key: 'personal' });
  const workbook = XLSX.read(Buffer.from(base64, 'base64'), { type: 'buffer' });
  const rows = XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]], { defval: null });
  assert.equal(rows[0]['Stage reason code'], 'ACTIVE_LINES_PRICED');
  assert.equal(rows[0]['Class of work'], workRule.classes.ACTIVE_LINES_PRICED.label);
  assert.equal(rows[0]['Age band'], '0–7');
});

test('parseXlsx accepts an added named column without a whitelist', () => {
  const sheet = XLSX.utils.aoa_to_sheet([['Purchase requisition', 'Stage reason code'], ['PR-1', 'ACTIVE_LINES_PRICED']]);
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, sheet, 'Sheet1');
  const parsed = parseXlsx(XLSX.write(workbook, { type: 'buffer', bookType: 'xlsx' }));
  assert.equal(parsed.length, 1);
  assert.equal(parsed[0]['Stage reason code'], 'ACTIVE_LINES_PRICED');
});
