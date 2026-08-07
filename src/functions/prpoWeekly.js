/*
 * PR/PO WEEKLY SNAPSHOT — archives the approver-based bucket counts every Sunday so the
 * dashboard's Analysis tab can show true week-over-week deltas.
 *
 * TIMER: 0 15 18 * * 0  (Sunday 18:15 UTC = 22:15 Dubai; after buckets-weekly at 18:00)
 *   1. Fetches pr.xlsx / po.xlsx from the dashboard GitHub Pages (same source as the emails).
 *   2. buildItems() from prpoEmail.js -> per-bucket + per-division counts + exception counts.
 *   3. Appends {date, buckets, divisions, exceptions, totals} to weekly_snapshots.json in the
 *      PR-PO-Pipeline-Dashboard repo (GitHub contents API).
 *
 * HTTP test (authLevel function, add &code=KEY):
 *   /api/prpo-weekly?dryrun=1  -> compute + return the snapshot JSON, commit nothing
 *   /api/prpo-weekly?commit=1  -> compute AND commit (what the timer does)
 *
 * Env: PRPO_GH_TOKEN (or GH_TOKEN) needs Contents R/W on Strive-Services-Group/PR-PO-Pipeline-Dashboard.
 *      Optional: PRPO_SNAP_REPO, PRPO_SNAP_BRANCH.
 */
const { app } = require('@azure/functions');
const { buildItems, parseXlsx } = require('./prpoEmail.js');

const PR_URL = process.env.PRPO_PR_URL || 'https://strive-services-group.github.io/PR-PO-Pipeline-Dashboard/pr.xlsx';
const PO_URL = process.env.PRPO_PO_URL || 'https://strive-services-group.github.io/PR-PO-Pipeline-Dashboard/po.xlsx';
const SNAP_URL = 'https://strive-services-group.github.io/PR-PO-Pipeline-Dashboard/weekly_snapshots.json';

async function fetchBuf(url){ const r = await fetch(url + (url.includes('?')?'&':'?') + 't=' + Date.now()); if(!r.ok) throw new Error('fetch ' + r.status + ' ' + url); return Buffer.from(await r.arrayBuffer()); }

function computeSnapshot(items){
  const buckets = {}, divisions = {};
  let totalValue = 0;
  for (const it of items) {
    const bk = it.doc + ' ' + it.stage;
    buckets[bk] = (buckets[bk] || 0) + 1;
    divisions[it.div] = (divisions[it.div] || 0) + 1;
    totalValue += (it.value || 0);
  }
  const exceptions = {
    opsConfirmOverdue: items.filter(it => it.doc === 'PR' && it.stage === 'Operations to Confirm' && (it.age || 0) > 7).length,
    reassignedRejected: items.filter(it => it.stage === 'Re-Assigned/Rejected').length,
    noRfqOver7d: items.filter(it => it.doc === 'PR' && it.stage === 'Procurement' && (it.age || 0) > 7 && !String((it.raw && it.raw['Request for quotation case']) || '').trim()).length,
    approvedPoNotSent: items.filter(it => it.doc === 'PO' && ['Procurement','Finance','Director','CEO'].includes(it.stage) && ['Confirmed','Approved'].includes(String((it.raw && it.raw['Approval status']) || '')) && String((it.raw && it.raw['Purchase order status']) || '') === 'Open order').length,
    receivedNotInvoiced: items.filter(it => it.doc === 'PO' && it.stage === 'Pending Invoicing').length,
    breach7d: items.filter(it => (it.age || 0) > 7).length
  };
  return { date: new Date().toISOString().slice(0, 10), total: items.length, totalValue: Math.round(totalValue), buckets, divisions, exceptions };
}

async function loadExisting(){
  try { const r = await fetch(SNAP_URL + '?t=' + Date.now()); if (r.ok) { const j = await r.json(); if (j && Array.isArray(j.snapshots)) return j; } } catch (e) { /* first run */ }
  return { snapshots: [] };
}

async function commitSnapshots(obj){
  const token = process.env.PRPO_GH_TOKEN || process.env.GH_TOKEN;
  if (!token) throw new Error('PRPO_GH_TOKEN / GH_TOKEN not set');
  const repo = process.env.PRPO_SNAP_REPO || 'Strive-Services-Group/PR-PO-Pipeline-Dashboard';
  const branch = process.env.PRPO_SNAP_BRANCH || 'main';
  const api = 'https://api.github.com/repos/' + repo + '/contents/weekly_snapshots.json';
  const H = { Authorization: 'Bearer ' + token, Accept: 'application/vnd.github+json', 'User-Agent': 'prpo-weekly-fn', 'X-GitHub-Api-Version': '2022-11-28' };
  let sha;
  const lr = await fetch(api + '?ref=' + branch, { headers: H });
  if (lr.ok) { const f = await lr.json(); if (f && f.sha) sha = f.sha; }
  const body = { message: 'Weekly PR/PO bucket snapshot ' + new Date().toISOString().slice(0, 16) + 'Z', content: Buffer.from(JSON.stringify(obj, null, 2)).toString('base64'), branch };
  if (sha) body.sha = sha;
  const r = await fetch(api, { method: 'PUT', headers: { ...H, 'Content-Type': 'application/json' }, body: JSON.stringify(body) });
  if (!r.ok) { const j = await r.json().catch(() => ({})); throw new Error('github commit ' + r.status + ' ' + JSON.stringify(j).slice(0, 200)); }
  return true;
}

async function runSnapshot(doCommit, context){
  const [prBuf, poBuf] = await Promise.all([fetchBuf(PR_URL), fetchBuf(PO_URL)]);
  const items = buildItems(parseXlsx(prBuf), parseXlsx(poBuf));
  const snap = computeSnapshot(items);
  const existing = await loadExisting();
  existing.snapshots = existing.snapshots.filter(s => s.date !== snap.date); // one per day max
  existing.snapshots.push(snap);
  if (existing.snapshots.length > 60) existing.snapshots = existing.snapshots.slice(-60);
  let committed = false;
  if (doCommit) { committed = await commitSnapshots(existing); if (context) context.log('prpo-weekly committed snapshot ' + snap.date); }
  return { snapshot: snap, snapshotsStored: existing.snapshots.length, committed };
}

app.timer('prpo-weekly-snap', { schedule: '0 15 18 * * 0', handler: async (timer, context) => {
  try { await runSnapshot(true, context); } catch (e) { context.error('prpo-weekly FAILED: ' + e.message); }
}});

app.http('prpo-weekly', { methods: ['GET','OPTIONS'], authLevel: 'function', route: 'prpo-weekly', handler: async (request, context) => {
  try {
    const url = new URL(request.url);
    const doCommit = url.searchParams.get('commit') === '1';
    const out = await runSnapshot(doCommit, context);
    return { status: 200, jsonBody: out };
  } catch (e) { context.error('prpo-weekly failed:', e); return { status: 500, jsonBody: { error: e.message } }; }
}});

module.exports = { computeSnapshot };
