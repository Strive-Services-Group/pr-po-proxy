'use strict';

const fs = require('fs');
const {
  deriveClock,
  observationKey,
  observationPayload,
  readObservation,
  upsertObservation
} = require('../src/shared/prpoDataset');

(async () => {
  const number = 'C04-STAGE-CHANGE-PROOF';
  const company = 'proof';
  const key = observationKey(company, number);
  const startedAt = new Date().toISOString();
  const initialPo = { number, legalEntity: company, liveStage: 'Sent to supplier', isReportable: false };
  const initialClock = { timestamp: null, provenance: 'NOT_RECORDED', label: 'since — not recorded', liveEvent: null };
  await upsertObservation(key, observationPayload(initialPo, initialClock, startedAt));
  const baseline = await readObservation(key);

  const changedAt = new Date(Date.now() + 10).toISOString();
  const changedPo = { number, legalEntity: company, liveStage: 'Receipt posted', isReportable: false };
  const changedClock = deriveClock(changedPo, baseline, changedAt);
  await upsertObservation(key, observationPayload(changedPo, changedClock, changedAt));
  const stored = await readObservation(key);
  const notes = JSON.parse(stored.ssg_dataqualitynotes);
  const proof = {
    provedAt: new Date().toISOString(),
    key,
    reportable: stored.ssg_isreportable,
    before: { stage: initialPo.liveStage, provenance: initialClock.provenance },
    after: {
      stage: stored.ssg_currentstepname,
      observedSince: stored.ssg_observedpendingsince,
      provenance: notes.clockProvenance,
      liveEvent: notes.liveEvent
    },
    passed: stored.ssg_isreportable === false &&
      stored.ssg_currentstepname === 'Receipt posted' &&
      notes.clockProvenance === 'PENDING_SINCE_FIRST_OBSERVED' &&
      stored.ssg_observedpendingsince === changedAt
  };
  if (process.argv[2]) fs.writeFileSync(process.argv[2], JSON.stringify(proof, null, 2) + '\n');
  process.stdout.write(JSON.stringify(proof, null, 2) + '\n');
  if (!proof.passed) process.exitCode = 1;
})().catch(error => {
  console.error(error);
  process.exitCode = 1;
});
