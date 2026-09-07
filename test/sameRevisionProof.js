'use strict';

const { getDataset, clearCache } = require('../src/shared/prpoDataset');
const { loadItems } = require('../src/functions/prpoEmail');
const fs = require('fs');

(async () => {
  clearCache();
  const dashboardDataset = await getDataset({ force: true });
  const emailItems = await loadItems();
  const proof = {
    provedAt: new Date().toISOString(),
    dashboardDatasetRevision: dashboardDataset.revision,
    emailDryRunDatasetRevision: emailItems.datasetRevision,
    sameRevision: dashboardDataset.revision === emailItems.datasetRevision,
    counts: {
      prHeaders: dashboardDataset.pr.count,
      poHeaders: dashboardDataset.po.count,
      poOpen: dashboardDataset.po.openCount,
      emailOpenItems: emailItems.length
    },
    sourceState: dashboardDataset.sourceState,
    sendsPerformed: 0
  };
  if (process.argv[2]) fs.writeFileSync(process.argv[2], JSON.stringify(proof, null, 2) + '\n');
  process.stdout.write(JSON.stringify(proof, null, 2) + '\n');
  if (!proof.sameRevision) process.exitCode = 1;
})().catch(error => {
  console.error(error);
  process.exitCode = 1;
});
