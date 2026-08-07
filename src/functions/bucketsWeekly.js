/*
 * WEEKLY ALL-PROJECT BUCKET SNAPSHOT — fully headless (no laptop, no one opening the dashboard).
 *
 * Every Sunday 22:00 Dubai (18:00 UTC) this records THIS ISO week's In-Unit bucket counts
 * (A+ / A / B / C / D / E / F) for EVERY flagship project and commits them into
 * buckets.json in the dashboard repo — so on Monday the dashboard shows the finished week
 * as "Last Week", and the week after as "Last Two Week". Movement therefore appears
 * automatically without anyone opening and filtering each project.
 *
 * The classification is a 1:1 port of the dashboard's classifyAt() (index.html); the opp and
 * booking queries mirror fetchProjectOpps() / fetchProjectBookings(); competitor visits come
 * from the same visitor.xlsx the dashboard reads; the unit universe comes from locations.json.
 *
 * HTTP test endpoints (Function App -> Functions -> buckets-weekly -> Get function URL):
 *   GET <url>&dryrun=1   -> compute per-project counts, commit NOTHING (compare with the dashboard)
 *   GET <url>&commit=1   -> compute AND commit buckets.json (what the Sunday timer does)
 *   GET <url>            -> same as dryrun (safe default)
 *
 * Uses the SAME Function App settings already in place for the telemetry email — no new secrets:
 *   TENANT_ID, CLIENT_ID, CLIENT_SECRET   (Azure AD app / client credentials)
 *   DV_RESOURCE = https://operations-ifahr-live.crm15.dynamics.com
 *   GH_TOKEN    = fine-grained PAT with Contents R/W on Strive-Services-Group/Visitor-Competitor-Dashboard
 *   (optional) GH_REPO, GH_BRANCH, DV_DIVISION_CANDOO, OPP_WON_GATE, VMS_URL, LOC_URL, BUCKETS_URL
 *
 * NOTE: this runs as the Function's APPLICATION user (client credentials). If that user lacks
 * read on opportunity / msdyn_functionallocation / bookableresourcebooking, the dryrun will show
 * per-project errors or zeroed counts — grant the app user a read role and re-test.
 */
const { app } = require('@azure/functions');
const XLSX = require('xlsx');

const DV_DIVISION_CANDOO = process.env.DV_DIVISION_CANDOO || '93bd43eb-0a44-ef11-a316-6045bd6a8335';
const OPP_WON_GATE = process.env.OPP_WON_GATE || 'Gate 04.13 Closed As Won';
const BASE = 'https://strive-services-group.github.io/Visitor-Competitor-Dashboard/';
const VMS_URL = process.env.VMS_URL || (BASE + 'visitor.xlsx');
const LOC_URL = process.env.LOC_URL || (BASE + 'locations.json');
const BUCKETS_URL = process.env.BUCKETS_URL || (BASE + 'buckets.json');
const F = '@OData.Community.Display.V1.FormattedValue';
const BKTS = ['A+', 'A', 'B', 'C', 'D', 'E', 'F'];

// EXACT copy of the dashboard's project selector whitelist (buildProjectMap WL).
const WL = ['SHORELINE 7AND8', 'BALQIS RESIDENCE', 'MAG 318', 'NORTH RESIDENCE', 'SOUTH RESIDENCE', 'THE8', 'BARTON HOUSE 1', 'EDEN APARTMENTS', 'FORTUNATO', 'I RISE TOWER', 'IRIS BLUE', 'MONREVE TOWER', 'THE PRISM TOWER', 'SEVEN PALM HOTEL AND APTS', 'RESIDENCE 110', 'BELGRAVIA SQUARE JVC', 'CAYAN BUSSINES CENTER', 'CAYAN BUSINESS CENTER', 'GOLDEN MILE', 'SWAY RESIDENCE', 'THE AUTOGRAPH', 'WILTON PARK RESIDENCES', 'J ONE TOWER', 'SPARKLE TOWERS', 'DOMUS AL GARHOUD', 'DOMUS INDIGO'];
const normName = s => String(s).toUpperCase().replace(/&/g, 'AND').replace(/[^A-Z0-9]/g, '');

/* ---- helpers copied 1:1 from index.html so counts match the dashboard exactly ---- */
function canonService(s) {
  const u = (s || '').toUpperCase();
  if (u.includes('POOL') || u.includes('SWIM')) return 'POOL';
  if (u.includes('MAINT') || u.includes('HANDYMAN') || u.includes('AMC')) return 'MAINT';
  if (u.includes('CLEAN') || u.includes('HOUSE')) return 'HOUSEKEEP';
  if (u.includes('LAUNDR')) return 'LAUNDRY';
  if (u.includes('FIT') && u.includes('OUT')) return 'FITOUT';
  if (u.includes('INSPECT') || u.includes('SURVEY')) return 'INSPECT';
  return 'OTHER';
}
function dvIsOurCompany(name) {
  const n = (name || '').trim().toLowerCase();
  return n.indexOf('candoo') === 0 || n.indexOf('s & c') === 0 || n.indexOf('s&c') === 0 || n.indexOf('strive') === 0;
}
function isExcludedUnit(up) {
  if (!up) return false;
  if (up.indexOf('BALQIS BLOCK C_') !== 0) return false;
  var tok = up.split('-').pop();
  if (/^G/.test(tok)) return true;
  var n = parseInt(tok, 10);
  if (!isNaN(n) && n >= 100 && n <= 599) return true;
  return false;
}
function getISOWeek(date) {
  const d = new Date(Date.UTC(date.getFullYear(), date.getMonth(), date.getDate()));
  const dn = d.getUTCDay() || 7;
  d.setUTCDate(d.getUTCDate() + 4 - dn);
  const yearStart = new Date(Date.UTC(d.getUTCFullYear(), 0, 1));
  const w = Math.ceil((((d - yearStart) / 86400000) + 1) / 7);
  return `${d.getUTCFullYear()}-W${String(w).padStart(2, '0')}`;
}
function fmtDate(d) { return d.toISOString().slice(0, 10); }

/* ---- OAuth (client credentials) + Dataverse GET (with formatted-value annotations) ---- */
async function getToken(scopeBase) {
  const body = new URLSearchParams({
    client_id: process.env.CLIENT_ID,
    client_secret: process.env.CLIENT_SECRET,
    grant_type: 'client_credentials',
    scope: scopeBase.replace(/\/+$/, '') + '/.default'
  });
  const r = await fetch(`https://login.microsoftonline.com/${process.env.TENANT_ID}/oauth2/v2.0/token`, {
    method: 'POST', headers: { 'Content-Type': 'application/x-www-form-urlencoded' }, body
  });
  const j = await r.json();
  if (!r.ok || !j.access_token) throw new Error('token ' + r.status + ' ' + (j.error_description || j.error || ''));
  return j.access_token;
}
async function dvGet(url, token) {
  const r = await fetch(url, {
    headers: {
      Authorization: 'Bearer ' + token, Accept: 'application/json',
      'OData-MaxVersion': '4.0', 'OData-Version': '4.0',
      Prefer: 'odata.include-annotations="*"'
    }
  });
  const j = await r.json().catch(() => ({}));
  if (!r.ok) throw new Error('odata ' + r.status + ' ' + JSON.stringify(j.error || j).slice(0, 200));
  return j;
}

/* ---- unit universe per project (locations.json) ---- */
async function loadUniverse() {
  const r = await fetch(LOC_URL + '?t=' + Date.now());
  if (!r.ok) throw new Error('locations.json ' + r.status);
  const arr = await r.json();
  const byProj = {};
  for (const l of (arr || [])) {
    const p = String(l.p || '').trim().toUpperCase(); if (!p) continue;
    const u = String(l.u || '').trim(); if (!u) continue;
    if (isExcludedUnit(u.toUpperCase())) continue;
    (byProj[p] = byProj[p] || []).push(u);
  }
  return byProj;
}

/* ---- competitor visits per unit (visitor.xlsx) — our companies + Dima excluded ---- */
async function loadVmsCompByUnit() {
  const r = await fetch(VMS_URL + '?t=' + Date.now());
  if (!r.ok) throw new Error('visitor.xlsx ' + r.status);
  const buf = Buffer.from(await r.arrayBuffer());
  const wb = XLSX.read(buf, { type: 'buffer', cellDates: false });
  const ws = wb.Sheets['FINAL'] || wb.Sheets[wb.SheetNames[0]];
  const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
  if (!aoa.length) return {};
  const hdr = aoa[0].map(h => String(h).trim());
  const iP = hdr.indexOf('Check In Purpose'), iC = hdr.indexOf('Company Name'), iU = hdr.indexOf('Building/ Unit'), iT = hdr.indexOf('Check In Type');
  const map = {};
  for (let i = 1; i < aoa.length; i++) {
    const row = aoa[i]; if (!row) continue;
    const comp = String(row[iC] || '').trim(); if (!comp) continue;
    if (iT >= 0 && String(row[iT] || '').trim().toLowerCase() !== 'unit visit') continue;
    if (/dima/i.test(comp)) continue;
    if (dvIsOurCompany(comp)) continue;
    const unit = String(row[iU] || '').trim(); if (!unit) continue;
    (map[unit] = map[unit] || new Set()).add(canonService(String(row[iP] || '')));
  }
  return map; // { unit: Set(services) }
}

/* ---- opportunities per project (port of fetchProjectOpps) -> { unit: {w,t,lw,wd:[]} } ---- */
async function oppsForProject(project, api, token) {
  const safe = String(project).replace(/'/g, "''");
  let pj = await dvGet(api + "/msdyn_functionallocations?$select=msdyn_functionallocationid&$filter=" + encodeURIComponent("msdyn_name eq '" + safe + "'"), token);
  let pids = (pj.value || []).map(x => x.msdyn_functionallocationid);
  if (!pids.length) {
    try {
      const pj2 = await dvGet(api + "/msdyn_functionallocations?$select=msdyn_functionallocationid&$filter=" + encodeURIComponent("startswith(msdyn_name,'" + safe + "')"), token);
      pids = (pj2.value || []).map(x => x.msdyn_functionallocationid);
    } catch (e) {}
  }
  if (!pids.length) return {};
  const selq = "$select=ssg_textdepartment,createdon,ssg_homeservicegates,_ssg_building_value,_ssg_unit_value";
  const oflt = "(" + pids.map(g => "_ssg_communityproject_value eq " + g).join(" or ") + ")";
  let u2 = api + "/opportunities?" + selq + "&$filter=" + encodeURIComponent(oflt);
  const recs = []; let guard = 0;
  while (u2 && guard < 25) { const j = await dvGet(u2, token); recs.push.apply(recs, j.value || []); u2 = j['@odata.nextLink'] || null; guard++; }
  const per = {};
  for (const o of recs) {
    const b = o['_ssg_building_value' + F] || '', un = o['_ssg_unit_value' + F] || '';
    if (!b && !un) continue;
    const key = b + '_' + un, gate = (o['ssg_homeservicegates' + F] || '').trim(), dept = (o.ssg_textdepartment || '').trim();
    let ds = ''; if (o.createdon) { const d = new Date(new Date(o.createdon).getTime() + 4 * 3600 * 1000); ds = d.toISOString().slice(0, 10); }
    if (!per[key]) per[key] = { w: 0, t: 0, lw: '', wd: new Set() };
    const e = per[key]; e.t++;
    if (gate === OPP_WON_GATE) { e.w++; if (ds && (!e.lw || ds > e.lw)) e.lw = ds; if (dept) e.wd.add(dept); }
  }
  for (const k in per) per[k].wd = [...per[k].wd];
  return per;
}

/* ---- Candoo bookings per project (port of fetchProjectBookings aggregate) -> { unit:{count,services:[]} } ---- */
async function bookingsForProject(project, api, token) {
  const safe = String(project).replace(/'/g, "''");
  let fl = await dvGet(api + "/msdyn_functionallocations?$select=msdyn_functionallocationid&$filter=" + encodeURIComponent("msdyn_name eq '" + safe + "'") + "&$top=25", token);
  let ids = (fl.value || []).map(x => x.msdyn_functionallocationid);
  if (!ids.length) {
    try {
      const fl2 = await dvGet(api + "/msdyn_functionallocations?$select=msdyn_functionallocationid&$filter=" + encodeURIComponent("startswith(msdyn_name,'" + safe + "')") + "&$top=25", token);
      ids = (fl2.value || []).map(x => x.msdyn_functionallocationid);
    } catch (e) {}
  }
  if (!ids.length) return {};
  const projValues = ids.map(g => '<value>' + g + '</value>').join('');
  const fx = "<fetch aggregate='true'><entity name='bookableresourcebooking'>" +
    "<attribute name='bookableresourcebookingid' aggregate='count' alias='cnt'/>" +
    "<link-entity name='msdyn_workorder' from='msdyn_workorderid' to='msdyn_workorder' link-type='inner' alias='wo'>" +
    "<attribute name='ssg_building' alias='bld' groupby='true'/><attribute name='ssg_unit' alias='unt' groupby='true'/><attribute name='ssg_department' alias='dep' groupby='true'/>" +
    "<filter type='and'><condition attribute='ssg_division' operator='eq' value='" + DV_DIVISION_CANDOO + "'/><condition attribute='ssg_project' operator='in'>" + projValues + "</condition></filter>" +
    "</link-entity>" +
    "<link-entity name='bookableresource' from='bookableresourceid' to='resource' link-type='inner' alias='res'><filter type='and'><condition attribute='name' operator='not-like' value='%Team Leader%'/><condition attribute='name' operator='not-like' value='%Team Member%'/></filter></link-entity>" +
    "<link-entity name='bookingstatus' from='bookingstatusid' to='bookingstatus' link-type='inner' alias='bs'><filter type='and'><condition attribute='name' operator='not-like' value='%ancel%'/></filter></link-entity>" +
    "</entity></fetch>";
  const ja = await dvGet(api + '/bookableresourcebookings?fetchXml=' + encodeURIComponent(fx), token);
  const map = {};
  for (const r of (ja.value || [])) {
    const b = (r['bld' + F] || '').trim(), un = (r['unt' + F] || '').trim();
    if (!b && !un) continue;
    const key = b + '_' + un;
    if (!map[key]) map[key] = { count: 0, services: new Set() };
    map[key].count += (r.cnt || 0);
    map[key].services.add(canonService(r['dep' + F] || ''));
  }
  for (const k in map) map[k].services = [...map[k].services];
  return map;
}

/* ---- classify a project's units -> {A+,A,B,C,D,E,F} counts (1:1 with classifyAt) ---- */
function classifyCounts(units, vmsMap, bookMap, oppMap, aging) {
  const counts = { 'A+': 0, 'A': 0, 'B': 0, 'C': 0, 'D': 0, 'E': 0, 'F': 0 };
  for (const u of units) {
    const vms = vmsMap[u] || null;                     // Set of competitor services (or null)
    const bk = bookMap[u] || { count: 0, services: [] };
    const opp = oppMap[u] || { w: 0, t: 0, lw: '', wd: [] };
    const hasScBookings = bk.count > 0;
    const hasCompVms = !!(vms && vms.size > 0);
    const hasWon = (opp.w || 0) > 0;
    const hasAnyOpp = (opp.t || 0) > 0;
    const hasScEngagement = hasScBookings || hasAnyOpp;
    const isAging = hasWon && opp.lw && opp.lw < aging;
    let bucket;
    if (!hasScEngagement && !hasCompVms) bucket = 'F';
    else if (!hasScEngagement && hasCompVms) bucket = 'E';
    else if (hasAnyOpp && !hasWon && !hasScBookings) bucket = 'D';
    else {
      const scServ = new Set(bk.services || []);
      (opp.wd || []).forEach(d => scServ.add(canonService(d)));
      if (!hasCompVms) bucket = 'A+';
      else {
        let overlap = false;
        for (const s of scServ) if (vms.has(s)) { overlap = true; break; }
        bucket = overlap ? 'A' : 'B';
      }
    }
    counts[bucket]++;
    if (isAging) counts['C']++;   // C is an overlay
  }
  return counts;
}

/* ---- buckets.json load (from Pages) + commit (GitHub contents API, GH_TOKEN) ---- */
async function loadBucketsJson() {
  try { const r = await fetch(BUCKETS_URL + '?t=' + Date.now()); if (r.ok) return (await r.json()) || {}; } catch (e) {}
  return {};
}
async function commitBucketsJson(obj) {
  const token = process.env.GH_TOKEN;
  if (!token) throw new Error('GH_TOKEN not set');
  const repo = process.env.GH_REPO || 'Strive-Services-Group/Visitor-Competitor-Dashboard';
  const branch = process.env.GH_BRANCH || 'main';
  const api = 'https://api.github.com/repos/' + repo + '/contents/buckets.json';
  const H = { Authorization: 'Bearer ' + token, Accept: 'application/vnd.github+json', 'User-Agent': 'buckets-weekly-fn', 'X-GitHub-Api-Version': '2022-11-28' };
  let sha;
  const lr = await fetch('https://api.github.com/repos/' + repo + '/contents/?ref=' + branch, { headers: H });
  if (lr.ok) { const list = await lr.json(); const f = (list || []).find(x => x.name === 'buckets.json'); if (f) sha = f.sha; }
  const body = { message: 'Weekly all-project bucket snapshot (server) ' + new Date().toISOString().slice(0, 16) + 'Z', content: Buffer.from(JSON.stringify(obj, null, 2)).toString('base64'), branch };
  if (sha) body.sha = sha;
  const r = await fetch(api, { method: 'PUT', headers: { ...H, 'Content-Type': 'application/json' }, body: JSON.stringify(body) });
  if (!r.ok) { const j = await r.json().catch(() => ({})); throw new Error('github commit ' + r.status + ' ' + JSON.stringify(j).slice(0, 200)); }
  return true;
}

/* ================= main ================= */
async function runWeeklyBuckets(context, opts) {
  opts = opts || {};
  const log = context ? context.log.bind(context) : () => {};
  const resource = (process.env.DV_RESOURCE || '').replace(/\/+$/, '');
  if (!resource) throw new Error('DV_RESOURCE app setting is not set');
  const api = resource + '/api/data/v9.2';
  const token = await getToken(resource);

  const byProj = await loadUniverse();
  const vmsMap = await loadVmsCompByUnit();

  const wlset = {}; WL.forEach(w => wlset[normName(w)] = 1);
  let projects = Object.keys(byProj).filter(p => wlset[normName(p)]).sort();
  if (byProj['BALQIS RESIDENCE'] && projects.indexOf('BALQIS RESIDENCE') < 0) projects.unshift('BALQIS RESIDENCE');

  const now = new Date();
  const wk = getISOWeek(now);
  const agingD = new Date(now); agingD.setDate(agingD.getDate() - 90);
  const aging = fmtDate(agingD);

  const perProject = {}, errors = {};
  for (const p of projects) {
    try {
      const units = byProj[p] || [];
      if (!units.length) { errors[p] = 'no units in locations.json'; continue; }
      const oppMap = await oppsForProject(p, api, token);
      const bookMap = await bookingsForProject(p, api, token);
      perProject[p] = classifyCounts(units, vmsMap, bookMap, oppMap, aging);
      log('  ' + p + ': ' + JSON.stringify(perProject[p]));
    } catch (e) { errors[p] = String(e && e.message || e).slice(0, 180); log('  !! ' + p + ': ' + errors[p]); }
  }

  let committed = false;
  if (!opts.dryrun && Object.keys(perProject).length) {
    const buckets = await loadBucketsJson();
    buckets.projects = buckets.projects || {};
    for (const p in perProject) { (buckets.projects[p] = buckets.projects[p] || {})[wk] = { counts: perProject[p], ts: new Date().toISOString() }; }
    committed = await commitBucketsJson(buckets);
  }
  return { week: wk, aging, projectsRecorded: Object.keys(perProject).length, committed, dryrun: !!opts.dryrun, perProject, errors };
}

/* ================= triggers ================= */
// Sunday 18:00 UTC = 22:00 Dubai — end of the ISO week, so Monday the dashboard shows it as "Last Week".
app.timer('buckets-weekly-run', {
  schedule: '0 0 18 * * 0',
  handler: async (timer, context) => {
    try {
      const out = await runWeeklyBuckets(context, {});
      context.log('weekly buckets:', JSON.stringify({ week: out.week, recorded: out.projectsRecorded, committed: out.committed, errors: Object.keys(out.errors) }));
    } catch (e) { context.error('weekly buckets FAILED:', e.message); throw e; }
  }
});

// Manual test: ?dryrun=1 (default) computes only; ?commit=1 computes AND commits.
app.http('buckets-weekly', {
  methods: ['GET'],
  authLevel: 'function',
  route: 'buckets-weekly',
  handler: async (request, context) => {
    try {
      const url = new URL(request.url);
      const dryrun = url.searchParams.get('commit') !== '1';
      const out = await runWeeklyBuckets(context, { dryrun });
      return { status: 200, jsonBody: out };
    } catch (e) { context.error('buckets-weekly failed:', e); return { status: 500, jsonBody: { error: e.message } }; }
  }
});

module.exports = { runWeeklyBuckets, classifyCounts, canonService, getISOWeek };
