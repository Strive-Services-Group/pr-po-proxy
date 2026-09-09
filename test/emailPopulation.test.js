'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const XLSX = require('xlsx');
const workRule = require('../work-class-rule.json');
const { buildItems, buildDivision, buildXlsxBase64, groupByOwner, personalPool, buildPersonal, DIVS, parseXlsx, applyDeliveryPolicy, freshnessWarning, userEmailMap } = require('../src/functions/prpoEmail');

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
  assert.equal(rows[0]['Age band'], `Raised · ${item.ageBand}`);
  assert.equal(rows[0]['Age basis'], 'Raised date');
});

test('parseXlsx accepts an added named column without a whitelist', () => {
  const sheet = XLSX.utils.aoa_to_sheet([['Purchase requisition', 'Stage reason code'], ['PR-1', 'ACTIVE_LINES_PRICED']]);
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, sheet, 'Sheet1');
  const parsed = parseXlsx(XLSX.write(workbook, { type: 'buffer', bookType: 'xlsx' }));
  assert.equal(parsed.length, 1);
  assert.equal(parsed[0]['Stage reason code'], 'ACTIVE_LINES_PRICED');
});

test('all-unpriced and partly-priced queues never imply one total covers every item', () => {
  const unpriced = { ...pr('In review', 'PR-UNPRICED'), 'Stage reason code': 'ACTIVE_LINES_NOT_FULLY_PRICED', 'Total amount': 0, 'Pending Approver/User': 'Adnan.Ullah' };
  const allUnpriced = buildPersonal(groupByOwner(personalPool(buildItems([unpriced], [])))[0], {});
  assert.match(allUnpriced.html, /1 still being priced/);
  assert.match(allUnpriced.html, /not yet priced/i);
  assert.doesNotMatch(allUnpriced.html, /totalling/i);

  const priced = { ...pr('In review', 'PR-PRICED'), 'Stage reason code': 'UNMAPPED_ELEMENT', 'Total amount': 250, 'Pending Approver/User': 'Adnan.Ullah' };
  const partial = buildPersonal(groupByOwner(personalPool(buildItems([unpriced, priced], [])))[0], {});
  assert.equal(partial.pricingQueueCount, 1);
  assert.equal(partial.pricedCount, 1);
  assert.equal(partial.value, 250);
  assert.match(partial.html, /2 items/);
  assert.match(partial.html, /1 still being priced/);
  assert.match(partial.html, /1 priced/);
  assert.match(partial.html, /AED 250 excl\. VAT/);
});

test('PR age wording distinguishes raised date from a distinct step date', () => {
  const same = { ...pr('In review', 'PR-SAME-CLOCK'), 'Created date': '2026-09-01T00:00:00Z', 'Step date and time': '2026-09-01T00:00:00Z' };
  const distinct = { ...pr('In review', 'PR-DISTINCT-CLOCK'), 'Created date': '2026-08-01T00:00:00Z', 'Step date and time': '2026-09-05T00:00:00Z' };
  const [sameItem, distinctItem] = buildItems([same, distinct], []);
  assert.equal(sameItem.clockBasis, 'raised');
  assert.match(sameItem.clockLabel, /^raised \d+ days ago$/);
  assert.equal(distinctItem.clockBasis, 'current step');
  assert.match(distinctItem.clockLabel, /^with you since 2026-09-05 \(\d+ days\)$/);
});

test('shared item names other active buyers and excludes an inactive username', () => {
  const row = { ...pr('In review', 'PR-SHARED-ACTIVE'), 'Stage reason code': 'ACTIVE_LINES_NOT_FULLY_PRICED', 'Pending Approver/User': 'Adnan.Ullah' };
  const rows = [row];
  Object.defineProperty(rows, 'routingMetadata', { value: [
    { 'Purchase requisition': 'PR-SHARED-ACTIVE', 'Source holder': 'Adnan.Ullah' },
    { 'Purchase requisition': 'PR-SHARED-ACTIVE', 'Source holder': 'Layusha.cleatus' },
    { 'Purchase requisition': 'PR-SHARED-ACTIVE', 'Source holder': 'roderick.red' }
  ] });
  const item = buildItems(rows, [])[0];
  assert.equal(item.sourceShared, true);
  assert.deepEqual(item.otherLiveBuyers, ['roderick.red']);
  assert.doesNotMatch(item.sharedLabel, /Layusha/i);
  const out = buildPersonal(groupByOwner(personalPool([item]))[0], {});
  assert.equal(out.sourceSharedCount, 1);
  assert.equal(out.sharedWithOtherActiveBuyers, 1);
  assert.match(out.html, /Shared with roderick\.red/);
});

test('inactive, unaddressed and addressed holders have one delivery route', () => {
  const rows = [
    { ...pr('In review', 'PR-INACTIVE'), 'Pending Approver/User': 'Layusha.cleatus', 'Department': 'Procurement', 'Stage reason code': 'ACTIVE_LINES_NOT_FULLY_PRICED' },
    { ...pr('In review', 'PR-NO-EMAIL'), 'Pending Approver/User': 'Sirinikhil', 'Department': 'Housekeeping Services', 'Stage reason code': 'UNMAPPED_ELEMENT' },
    { ...pr('In review', 'PR-ZAHEER'), 'Pending Approver/User': 'Zaheer Ahmed Ameer', 'Department': 'Accomodation Services', 'Stage reason code': 'UNMAPPED_ELEMENT' }
  ];
  const items = applyDeliveryPolicy(buildItems(rows, []));
  assert.equal(items[0].deliveryIssue, 'no active owner');
  assert.equal(items[0].noNamedOwner, true);
  assert.equal(items[1].deliveryIssue, 'no email address on file');
  assert.equal(items[1].noNamedOwner, true);
  assert.equal(items[2].deliveryIssue, undefined);
  assert.equal(userEmailMap()['zaheer.ahmed'], 'Zaheer.Ahmed@domus-housing.com');
  assert.deepEqual(groupByOwner(personalPool(items)).map(p=>p.user), ['Zaheer.Ahmed']);
  const procurement = buildDivision(DIVS.find(d=>d.key==='procurement'), items, {});
  assert.match(procurement.html, /Layusha\.cleatus/);
  assert.match(procurement.html, /no active owner/);
  assert.match(procurement.html, /Sirinikhil/);
  assert.match(procurement.html, /no email address on file/);
  assert.doesNotMatch(procurement.html, /PR-ZAHEER/);
  assert.equal(personalPool(items).length + items.filter(it=>it.doc==='PR'&&it.noNamedOwner).length, 3);
});

test('stale warning appears after six hours and stays absent when fresh', () => {
  const fresh = freshnessWarning('2026-09-09T04:00:00Z', '2026-09-09T09:59:59Z');
  const stale = freshnessWarning('2026-09-08T16:06:00Z', '2026-09-09T04:30:00Z');
  assert.equal(fresh, '');
  assert.equal(stale, 'These figures were prepared at 20:06 on 8 September and may not include work raised since.');

  const personalItems = applyDeliveryPolicy(buildItems([pr('In review', 'PR-STALE')], []));
  personalItems[0].freshnessWarning = stale;
  const personal = buildPersonal(groupByOwner(personalPool(personalItems))[0], {});
  assert.match(personal.html, /These figures were prepared at 20:06 on 8 September/);

  const noOwner = pr('In review', 'PR-STALE-DIVISION');
  noOwner['Pending Approver/User'] = 'Sirinikhil';
  noOwner['Stage reason code'] = 'UNMAPPED_ELEMENT';
  const divisionItems = applyDeliveryPolicy(buildItems([noOwner], []));
  divisionItems[0].freshnessWarning = fresh;
  const division = buildDivision(DIVS.find(d=>d.key==='procurement'), divisionItems, {});
  assert.doesNotMatch(division.html, /may not include work raised since/);
});
