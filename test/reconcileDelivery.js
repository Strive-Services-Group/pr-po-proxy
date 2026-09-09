'use strict';

const fs = require('fs');
const path = require('path');
const {
  buildItems, parseXlsx, applyDeliveryPolicy, personalPool, groupByOwner,
  buildPersonal, buildDivision, DIVS, freshnessWarning
} = require('../src/functions/prpoEmail');

const dashboardRoot = path.resolve(process.argv[2] || '.');
const outputDir = path.resolve(process.argv[3] || path.join(dashboardRoot, 'evidence'));
const state = JSON.parse(fs.readFileSync(path.join(dashboardRoot, 'legacy-email-workbook-state.json'), 'utf8'));
const dataset = JSON.parse(fs.readFileSync(path.join(dashboardRoot, 'evidence', 'nobody-missing-live-dataset.tmp.json'), 'utf8'));
const prRows = parseXlsx(fs.readFileSync(path.join(dashboardRoot, 'pr.xlsx')));
const poRows = parseXlsx(fs.readFileSync(path.join(dashboardRoot, 'po.xlsx')));
const items = applyDeliveryPolicy(buildItems(prRows, poRows));
const warning = freshnessWarning(state.datasetGeneratedAt, dataset.generatedAt);
for (const item of items) item.freshnessWarning = warning;
items.freshnessWarning = warning;

const prItems = items.filter(item => item.doc === 'PR');
const personal = personalPool(items);
const noNamed = prItems.filter(item => item.noNamedOwner);
const people = groupByOwner(personal);
const largest = people[0];
const procurement = buildDivision(DIVS.find(item => item.key === 'procurement'), items, {});
const largestEmail = buildPersonal(largest, {});
fs.mkdirSync(outputDir, { recursive: true });
fs.writeFileSync(path.join(outputDir, 'nobody-missing-largest-holder-email.html'), largestEmail.html);
fs.writeFileSync(path.join(outputDir, 'nobody-missing-procurement-email.html'), procurement.html);

const issues = {};
for (const item of noNamed) {
  const key = `${item.originalOwner || item.owner} | ${item.deliveryIssue || 'owner not recorded in F&O'}`;
  issues[key] = (issues[key] || 0) + 1;
}
const output = {
  datasetRevision: state.datasetRevision,
  datasetGeneratedAt: state.datasetGeneratedAt,
  liveDatasetGeneratedAt: dataset.generatedAt,
  freshnessWarning: warning,
  workbookAttributionRows: prRows.length,
  senderPrAttributionRows: prItems.length,
  namedPersonalEmailAttributions: personal.length,
  noNamedOwnerBlockAttributions: noNamed.length,
  sumOfDeliveryRoutes: personal.length + noNamed.length,
  neitherRoute: prItems.length - personal.length - noNamed.length,
  largestHolder: {
    user: largest.user,
    count: largest.items.length,
    address: require('../user-email-addresses.json')[largest.user] || null,
    sectionHeadings: largestEmail.sections,
  },
  zaheer: people.find(person => person.key === 'zaheer.ahmed') ? {
    count: people.find(person => person.key === 'zaheer.ahmed').items.length,
    address: require('../user-email-addresses.json')['Zaheer.Ahmed'],
  } : null,
  personalCounts: Object.fromEntries(people.map(person => [person.user, person.items.length])),
  noNamedByHolderAndReason: issues,
};
fs.writeFileSync(path.join(outputDir, 'nobody-missing-email-reconciliation.json'), JSON.stringify(output, null, 2) + '\n');
console.log(JSON.stringify(output, null, 2));

if (output.workbookAttributionRows !== output.sumOfDeliveryRoutes || output.neitherRoute !== 0) process.exitCode = 1;
