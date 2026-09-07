'use strict';

const { buildDataset } = require('../src/shared/prpoDataset');

buildDataset().then(dataset => {
  const openPr = dataset.pr.rows.filter(row => ['in review', 'approved'].includes(String(row.Status || '').toLowerCase()));
  const stages = {};
  for (const row of openPr) {
    const stage = row['Step name'] || '(blank)';
    stages[stage] = (stages[stage] || 0) + 1;
  }
  process.stdout.write(JSON.stringify({
    revision: dataset.revision,
    pr: dataset.pr.count,
    prOpen: openPr.length,
    prStages: stages,
    po: dataset.po.count,
    poOpen: dataset.po.openCount,
    stageNotEvidenced: dataset.po.stageNotEvidenced,
    clocks: dataset.po.clockCounts,
    pendingObservationWrites: dataset.pendingObservationWrites
  }, null, 2) + '\n');
}).catch(error => {
  console.error(error);
  process.exitCode = 1;
});
