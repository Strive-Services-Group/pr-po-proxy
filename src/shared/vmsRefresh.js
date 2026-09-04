/*
 * HEADLESS VMS REFRESH — port of Visitor-Competitor-Dashboard/clean_vms.py.
 *
 * Downloads the 6 VMS source workbooks straight from OneDrive via Microsoft Graph
 * (application permission Files.Read.All), applies the SAME cleaning rules as
 * clean_vms.py, and returns the merged records + a fresh visitor.xlsx workbook
 * buffer. Optionally commits that buffer to the GitHub repo so the dashboard
 * (GitHub Pages) shows the same data as the 9 AM email.
 *
 * Sources (drive IDs found via SharePoint search, stable unless the OneDrive is recreated):
 *   - Balqis Security's OneDrive  (sec_balqis_sahalahfm_com):
 *       /Desktop/POOL2026/Pool and beach Records/DAILY CONTRACTORS RECORDS/Visitors Details Balqis Residence.xlsx
 *       (a POOL2026 folder was inserted above "Pool and beach Records" on 27 Aug 2026, which
 *        404'd the old hard-coded path for 5 days without anyone noticing — see fetchSource:
 *        a path may now be a string OR an array of candidates, and if every candidate fails
 *        the workbook is located by name search, so a folder move self-heals)
 *   - Abdul Muqeet's OneDrive     (abdul_muqeet_sahalahfm_com):
 *       /VMS-DATA FILES/VMS-TH8.xlsx, VMS-AL HASEER.xlsx, VMS-AL NABAT.xlsx,
 *       /VMS-DATA FILES/VMS-NORTH RESIDENCE.xlsx, VMS-SOUTH RESIDENCE.xlsx
 *
 * Cleaning rules (identical to clean_vms.py):
 *   - keep purposes (the four telemetry service lines only, from 4 Sep 2026):
 *     LAUNDRY, CLEANING/CLEANERS + CLEANING/CLEANER, MAINTENANCE/HANDYMAN, FIT-OUT
 *     (compared uppercase, spaces stripped). INSPECTION and AMCCONTRACTORS were
 *     dropped on 4 Sep — they no longer count toward All Visitors or Share %.
 *   - "Unit Visit" check-ins only; Dima excluded everywhere
 *   - date from cell (Excel date or "DDMMYYYY hh:mm[:ss]" text), year 2024–2027
 *   - per-file de-dupe (date|purpose|unit|company) for the Abdul Muqeet files
 *   - SOUTH RESIDENCE carried forward from the previously published visitor.xlsx
 *     if its source file is missing/empty
 *   - output sheet 'FINAL' with the exact dashboard header
 */
const XLSX = require('xlsx');

// Balqis moved to a SharePoint list (Sahalah Visitor Log) on 1 Sep 2026; the OneDrive
// workbook stays configured as an automatic fallback until the list is proven in prod.
const BALQIS_SITE = process.env.VMS_BALQIS_SITE || 'ifares.sharepoint.com:/sites/SahalahVisitorLog';
const BALQIS_LIST = process.env.VMS_BALQIS_LIST || 'Visit Log';
const BALQIS_DRIVE = process.env.VMS_BALQIS_DRIVE || 'b!5ma3QhDyZ0GZsiXhkdXOuhs9QX6ol9RInUHGE6t7AIyt1DCpPIAcS6cBimNKf0JF';
const MUQEET_DRIVE = process.env.VMS_MUQEET_DRIVE || 'b!2jsyR69LVE63449EoxnFAaA1OVzt3PxKqmAZ7NtQgKH9TivpAV8ORZN5aAEIpdJm';

// [project default, driveId, path in drive, sheet names (null = all), dedupe, forceProject]
// forceProject=true: single-project files — ALWAYS use our project name, ignore the file's
// "Project Name" column (Abdul's restructured SOUTH file carried "NORTH RESIDENCE" in that
// column, which mislabelled all South visits as North). AL HASEER / AL NABAT keep the column
// because it maps them to SHORELINE 7AND8; Balqis keeps it as before.
const DEFAULT_SOURCES = [
  { kind: 'splist', project: 'BALQIS RESIDENCE', site: BALQIS_SITE, list: BALQIS_LIST,
    only: 'BALQIS',            // the list is Sahalah-wide: keep Balqis rows only, the rest still come from Abdul's files
    dedupe: false, forceProject: true,   // Balqis has never been de-duplicated (a company can legitimately re-visit a unit the same day); keep it that way
    fallback: { driveId: BALQIS_DRIVE, sheets: null, dedupe: false, forceProject: false,
                paths: ['/Desktop/POOL2026/Pool and beach Records/DAILY CONTRACTORS RECORDS/Visitors Details Balqis Residence.xlsx',
                        '/Desktop/Pool and beach Records/DAILY CONTRACTORS RECORDS/Visitors Details Balqis Residence.xlsx'] } },
  ['THE8', MUQEET_DRIVE, '/VMS-DATA FILES/VMS-TH8.xlsx', ["VMS-TH8-'26"], true, true],
  ['AL HASEER', MUQEET_DRIVE, '/VMS-DATA FILES/VMS-AL HASEER.xlsx', ["VMS-AL HASEER-'26"], true, false],
  ['AL NABAT', MUQEET_DRIVE, '/VMS-DATA FILES/VMS-AL NABAT.xlsx', ["VMS-AL NABAT-'26"], true, false],
  ['NORTH RESIDENCE', MUQEET_DRIVE, '/VMS-DATA FILES/VMS-NORTH RESIDENCE.xlsx', ["VMS-NORTH RESIDENCE-'26"], true, true],
  ['SOUTH RESIDENCE', MUQEET_DRIVE, '/VMS-DATA FILES/VMS-SOUTH RESIDENCE.xlsx', ["VMS-SOUTH RESIDENCE-'26"], true, true],
  ['MAG 318', MUQEET_DRIVE, '/VMS-DATA FILES/VMS-MAG 318.xlsx', ["VMS-MAG 318-'26"], true, true]
];
// a source is either the legacy tuple [proj, driveId, path, sheets, dedupe, forceProject]
// or an object; normalise to an object so the loop below stays readable.
function normalizeSource(s) {
  if (Array.isArray(s)) {
    return { kind: 'workbook', project: s[0], driveId: s[1], path: s[2], sheets: s[3], dedupe: s[4], forceProject: s[5] };
  }
  return Object.assign({ kind: 'workbook' }, s);
}
function sources() {
  try { if (process.env.VMS_SOURCES) return JSON.parse(process.env.VMS_SOURCES); } catch (e) {}
  return DEFAULT_SOURCES;
}

/* Telemetry is STRICTLY the four service lines — Laundry, Cleaning (Housekeeping),
   Maintenance/Handyman, Fit-Out (CK, 4 Sep 2026). Everything else is excluded:
   INSPECTION and AMC CONTRACTORS (legacy workbook purposes, dropped 4 Sep — this
   lowers historical All Visitors/Share % for every project, which is intended), and
   the app's other purposes (AMC service provider, Emergency call out, Swimming pool
   cleaning, Landscaping, Pest control, Move In/Out, Delivery, Guest).
   Both spellings of cleaning are kept: workbooks wrote "Cleaning/Cleaners" (plural),
   the Visitor App writes "Cleaning/Cleaner" (singular). norm() = uppercase, spaces stripped. */
const KEEP = new Set([
  'LAUNDRY', 'CLEANING/CLEANERS', 'CLEANING/CLEANER', 'MAINTENANCE/HANDYMAN', 'FIT-OUT'
]);
const HEADER = ['Check In Date', 'Check In Type', 'Check In Purpose', 'Company Name', 'Scope of work', 'Building/ Unit', 'Project Name'];
const norm = s => String(s == null ? '' : s).toUpperCase().replace(/ /g, '');

/* ---- Graph download: /drives/{id}/root:{path}:/content ---- */
async function graphDownload(driveId, path, token) {
  const url = 'https://graph.microsoft.com/v1.0/drives/' + driveId + '/root:' + encodeURI(path) + ':/content';
  const r = await fetch(url, { headers: { Authorization: 'Bearer ' + token } });
  if (!r.ok) {
    const j = await r.json().catch(() => ({}));
    throw new Error('graph download ' + r.status + ' for ' + path + ' ' + JSON.stringify(j.error || {}).slice(0, 200));
  }
  return Buffer.from(await r.arrayBuffer());
}

/* ---- locate a workbook by name when its folder has been moved or renamed ----
 * Building teams reorganise their OneDrive without telling anyone. Searching the drive for
 * the file name recovers it wherever it landed; we download by item id so no path rebuild
 * is needed, and log the new path so DEFAULT_SOURCES can be corrected at leisure. */
async function graphFindByName(driveId, fileName, token) {
  const q = fileName.replace(/'/g, "''");
  const url = 'https://graph.microsoft.com/v1.0/drives/' + driveId + "/root/search(q='" + encodeURIComponent(q) + "')"
    + '?$select=id,name,parentReference,lastModifiedDateTime&$top=50';
  const r = await fetch(url, { headers: { Authorization: 'Bearer ' + token } });
  if (!r.ok) throw new Error('graph search ' + r.status + ' for ' + fileName);
  const j = await r.json();
  const hits = (j.value || []).filter(x => x.name === fileName && !x.folder);
  if (!hits.length) throw new Error('graph search found nothing named ' + fileName);
  hits.sort((a, b) => String(b.lastModifiedDateTime || '').localeCompare(String(a.lastModifiedDateTime || '')));
  return hits[0];
}
async function graphDownloadById(driveId, itemId, token) {
  const r = await fetch('https://graph.microsoft.com/v1.0/drives/' + driveId + '/items/' + itemId + '/content',
    { headers: { Authorization: 'Bearer ' + token } });
  if (!r.ok) throw new Error('graph download by id ' + r.status + ' for item ' + itemId);
  return Buffer.from(await r.arrayBuffer());
}
/* try every configured path, then fall back to a name search. Throws only if both fail,
 * so the per-source catch still marks the project missing when the file is really gone. */
async function fetchSource(driveId, paths, token, log) {
  const list = Array.isArray(paths) ? paths : [paths];
  const errs = [];
  for (const cand of list) {
    try { return { buf: await graphDownload(driveId, cand, token), path: cand }; }
    catch (e) { errs.push(e.message); }
  }
  const fileName = String(list[0]).split('/').pop();
  const hit = await graphFindByName(driveId, fileName, token);
  const parent = (hit.parentReference && hit.parentReference.path) || '';
  const found = String(parent).replace(/^.*root:/, '') + '/' + hit.name; // /drive/root:/a/b and /drives/{id}/root:/a/b both -> /a/b
  log('  !! PATH MOVED: ' + fileName + ' not at configured path(s) [' + errs.join(' | ').slice(0, 200)
    + '] — found by search at ' + found + '. Update DEFAULT_SOURCES / VMS_SOURCES.');
  return { buf: await graphDownloadById(driveId, hit.id, token), path: found, movedFrom: list[0] };
}

/* ---- SharePoint list source (Balqis, from 1 Sep 2026) ----------------------
 * Graph: resolve site -> list -> page the items with $expand=fields. Field keys come
 * back as SharePoint INTERNAL names, which are not the display names and differ per
 * column (Check_x0020_In_x0020_Date, CheckInDate, Check_x0020_In_x0020_Date0 after a
 * rename...). So every lookup goes through nkey(): decode _xHHHH_ escapes, drop
 * punctuation and case. A renamed or re-created column keeps resolving. */
function nkey(s) {
  return String(s == null ? '' : s)
    .replace(/_x([0-9a-fA-F]{4})_/g, (mm, h) => String.fromCharCode(parseInt(h, 16)))
    .replace(/[^A-Za-z0-9]/g, '')
    .toUpperCase();
}
function fieldMap(f) { const m = {}; for (const k in f) m[nkey(k)] = f[k]; return m; }
function pickField(m) {
  for (let i = 1; i < arguments.length; i++) {
    const v = m[nkey(arguments[i])];
    if (v != null && String(v).trim() !== '') return v;
  }
  return '';
}
async function graphJson(url, token) {
  const r = await fetch(url, { headers: { Authorization: 'Bearer ' + token, Accept: 'application/json' } });
  if (!r.ok) {
    const j = await r.json().catch(() => ({}));
    throw new Error('graph ' + r.status + ' ' + url.replace('https://graph.microsoft.com/v1.0', '') + ' ' + JSON.stringify(j.error || {}).slice(0, 200));
  }
  return r.json();
}
/* Needs Graph application permission Sites.Read.All (or Sites.Selected granted on this
 * site). Files.Read.All alone is NOT enough — a 403 here is the tell. */
async function graphListRows(src, token, log) {
  const G = 'https://graph.microsoft.com/v1.0';
  const site = await graphJson(G + '/sites/' + src.site, token);
  const lists = await graphJson(G + '/sites/' + site.id + '/lists?$select=id,name,displayName&$top=200', token);
  const want = nkey(src.list);
  const hit = (lists.value || []).find(l => nkey(l.displayName) === want || nkey(l.name) === want);
  if (!hit) throw new Error('list "' + src.list + '" not found on ' + src.site + ' (lists seen: ' + (lists.value || []).map(l => l.displayName).join(', ').slice(0, 200) + ')');
  let url = G + '/sites/' + site.id + '/lists/' + hit.id + '/items?$expand=fields&$top=500';
  const items = []; let pages = 0;
  while (url && pages < 400) {
    const page = await graphJson(url, token);
    (page.value || []).forEach(it => { if (it.fields) items.push(it.fields); });
    url = page['@odata.nextLink']; pages++;
  }
  if (log) log('  ' + src.project + ': pulled ' + items.length + ' items from list "' + hit.displayName + '" (' + pages + ' page(s))');
  return { items, listName: hit.displayName, sampleKeys: items.length ? Object.keys(items[0]).slice(0, 40) : [] };
}
/* Same keep/exclude rules as cleanSource, applied to list items instead of sheet rows. */
function cleanListRows(proj, items, src, log, info) {
  const out = []; const seen = new Set();
  let dropped = 0, offProject = 0, timed = 0;
  const only = src.only ? nkey(src.only) : null;
  for (const f of items) {
    const m = fieldMap(f);
    const pur = String(pickField(m, 'Check In Purpose') || '').trim();
    if (!KEEP.has(norm(pur))) continue;
    if (String(pickField(m, 'Check In Type') || '').trim().toLowerCase() !== 'unit visit') continue;
    let comp = String(pickField(m, 'Company Name') || '').trim();
    if (/dima/i.test(comp)) continue; // Dima excluded everywhere
    const bu = String(pickField(m, 'Building/ Unit', 'Building/Unit', 'Building') || '').trim();
    const pnRaw = String(pickField(m, 'Project Name') || '').trim();
    // the list covers all Sahalah projects; take only the one this source owns so the
    // others keep coming from their own workbooks and nothing is counted twice
    if (only && nkey(pnRaw).indexOf(only) < 0 && nkey(bu).indexOf(only) < 0) { offProject++; continue; }
    const rawDate = pickField(m, 'Check In Date');
    const date = parseDateCell(rawDate);
    if (!date) { dropped++; continue; }
    if (typeof rawDate === 'string' && /T(?!00:00)/.test(rawDate)) timed++;
    comp = comp.toUpperCase();
    const scope = String(pickField(m, 'Scope of work') || '').trim();
    const unit = String(pickField(m, 'Unit') || '').trim();
    const pn = src.forceProject ? proj : (pnRaw || proj);
    if (src.dedupe) {
      const key = date.ymd + '|' + norm(pur) + '|' + (unit || bu).toUpperCase() + '|' + comp;
      if (seen.has(key)) continue;
      seen.add(key);
    }
    out.push([date.serial, 'Unit Visit', pur, comp, scope, bu, pn, date.ymd]);
  }
  if (info) { info.items = items.length; info.otherProjectRows = offProject; }
  if (log) {
    log('  ' + proj + ': ' + out.length + ' rows (dropped ' + dropped + ', other-project ' + offProject + ')');
    if (timed) log('  note: ' + timed + ' list rows carry a time of day on Check In Date; the UTC calendar date is used — check the day boundary if counts look shifted');
  }
  return out;
}

/* ---- date cell -> {ymd, serial} | null (same rules as clean_vms.parse_date) ----
 * serial = Excel date serial number (1900 system). We keep serials, never JS Date
 * objects, so the output workbook is byte-stable regardless of server timezone. */
function toSerial(y, m, d, hh, mi, ss) {
  return 25569 + Date.UTC(y, m - 1, d, hh, mi, ss) / 86400000; // 25569 = days 1900-01-01(excel epoch)..1970-01-01
}
function parseDateCell(v) {
  let y, m, d, hh = 0, mi = 0, ss = 0, serial = null;
  if (typeof v === 'number') {
    const pd = XLSX.SSF.parse_date_code(v);
    if (!pd) return null;
    y = pd.y; m = pd.m; d = pd.d; hh = pd.H || 0; mi = pd.M || 0; ss = Math.floor(pd.S || 0);
    serial = v; // pass the original serial through unchanged
  } else if (v instanceof Date) {
    y = v.getUTCFullYear(); m = v.getUTCMonth() + 1; d = v.getUTCDate();
    hh = v.getUTCHours(); mi = v.getUTCMinutes(); ss = v.getUTCSeconds();
  } else if (typeof v === 'string' && v.trim()) {
    // ISO first (SharePoint/Graph: 2026-08-30T00:00:00Z), then the VMS "DDMMYYYY hh:mm" text form
    const iso = v.trim().match(/^(\d{4})-(\d{2})-(\d{2})(?:[T ](\d{2}):(\d{2})(?::(\d{2}))?)?/);
    if (iso) {
      y = +iso[1]; m = +iso[2]; d = +iso[3]; hh = +(iso[4] || 0); mi = +(iso[5] || 0); ss = +(iso[6] || 0);
    } else {
      const mm = v.trim().match(/^(\d{2})(\d{2})(\d{4})(?:\s+(\d{1,2}):(\d{2})(?::(\d{2}))?)?/);
      if (!mm) return null;
      d = +mm[1]; m = +mm[2]; y = +mm[3]; hh = +(mm[4] || 0); mi = +(mm[5] || 0); ss = +(mm[6] || 0);
    }
    if (m < 1 || m > 12 || d < 1 || d > 31) return null;
  } else return null;
  if (y < 2024 || y > 2027) return null;
  if (serial == null) serial = toSerial(y, m, d, hh, mi, ss);
  const ymd = y + '-' + String(m).padStart(2, '0') + '-' + String(d).padStart(2, '0');
  return { ymd, serial };
}

function colIdx(headers, ...names) {
  const low = headers.map(h => String(h).trim().toLowerCase());
  for (const n of names) { const i = low.indexOf(n.toLowerCase()); if (i >= 0) return i; }
  return -1;
}

/* ---- clean one workbook (same logic/order as clean_vms.clean_source) ---- */
// sheet-name matching is normalized (curly vs straight apostrophe, case, spaces)
function normSheetName(s) { return String(s == null ? '' : s).replace(/[‘’]/g, "'").trim().toLowerCase(); }
function cleanSource(proj, buf, sheetNames, dedupe, log, info, forceProject) {
  const wb = XLSX.read(buf, { type: 'buffer', cellDates: false });
  let sns;
  if (sheetNames) {
    const wanted = sheetNames.map(normSheetName);
    sns = wb.SheetNames.filter(n => wanted.indexOf(normSheetName(n)) >= 0);
  } else {
    sns = wb.SheetNames;
  }
  if (info) { info.sheets = wb.SheetNames.slice(); info.matchedSheets = sns.slice(); info.bytes = buf.length; }
  const out = []; const seen = new Set(); let dropped = 0;
  for (const sn of sns) {
    if (wb.SheetNames.indexOf(sn) < 0) continue;
    const aoa = XLSX.utils.sheet_to_json(wb.Sheets[sn], { header: 1, defval: null });
    let hi = -1, hdr = null;
    for (let i = 0; i < aoa.length; i++) {
      const v = (aoa[i] || []).map(c => c == null ? '' : String(c));
      if (v.indexOf('Check In Date') >= 0) { hi = i; hdr = v; break; }
    }
    if (hi < 0) continue;
    const iD = colIdx(hdr, 'Check In Date'), iT = colIdx(hdr, 'Check In Type'), iP = colIdx(hdr, 'Check In Purpose');
    const iC = colIdx(hdr, 'Company Name', 'COMPANY NAME'), iS = colIdx(hdr, 'Scope of work');
    const iBU = colIdx(hdr, 'Building/ Unit', 'Building/Unit'), iPN = colIdx(hdr, 'Project Name', 'Project name'), iU = colIdx(hdr, 'Unit');
    for (let ri = hi + 1; ri < aoa.length; ri++) {
      const row = aoa[ri]; if (!row) continue;
      if (iD >= 0 && (row[iD] == null || row[iD] === '')) continue;
      if (!KEEP.has(norm(iP >= 0 ? row[iP] : ''))) continue;
      let comp = (iC >= 0 && row[iC] != null) ? String(row[iC]).trim() : '';
      if (/dima/i.test(comp)) continue; // Dima excluded everywhere
      const typ = (iT >= 0 && row[iT] != null) ? String(row[iT]).trim().toLowerCase() : '';
      if (typ !== 'unit visit') continue;
      const date = parseDateCell(row[iD]);
      if (!date) { dropped++; continue; }
      const pur = String(row[iP]).trim(); comp = comp.toUpperCase();
      const scope = (iS >= 0 && row[iS] != null) ? String(row[iS]).trim() : '';
      const bu = (iBU >= 0 && row[iBU] != null) ? String(row[iBU]).trim() : '';
      const pn = forceProject ? proj : ((iPN >= 0 && row[iPN] != null && String(row[iPN]).trim()) ? String(row[iPN]).trim() : proj);
      const unit = (iU >= 0 && row[iU] != null) ? String(row[iU]).trim() : '';
      if (dedupe) {
        const key = date.ymd + '|' + norm(pur) + '|' + unit.toUpperCase() + '|' + comp;
        if (seen.has(key)) continue;
        seen.add(key);
      }
      out.push([date.serial, 'Unit Visit', pur, comp, scope, bu, pn, date.ymd]); // ymd kept as 8th col for records; stripped before writing
    }
  }
  if (log) log('  ' + proj + ': ' + out.length + ' rows (dropped ' + dropped + ')');
  return out;
}

/* ---- previous published visitor.xlsx -> ALL rows (for the history merge) ---- */
function prevRowsAll(buf) {
  const wb = XLSX.read(buf, { type: 'buffer', cellDates: false });
  const ws = wb.Sheets['FINAL'] || wb.Sheets[wb.SheetNames[0]];
  const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null });
  const out = [];
  for (let i = 1; i < aoa.length; i++) {
    const r = aoa[i]; if (!r) continue;
    const date = parseDateCell(r[0]); if (!date) continue;
    out.push([date.serial, r[1] || 'Unit Visit', String(r[2] || ''), String(r[3] || ''), String(r[4] || ''), String(r[5] || ''), String(r[6] || ''), date.ymd]);
  }
  return out;
}

/*
 * Main entry. getToken(scopeBase) is passed in from telemetryEmail.js.
 * Returns { records, xlsxBuffer, counts, missing, total }:
 *   records = [{date, purpose, company, unit, project}] ready for buildEmail()
 */
async function refreshVms(getToken, vmsUrlFallback, context) {
  const log = context ? context.log.bind(context) : () => {};
  const token = await getToken('https://graph.microsoft.com');
  let rows = [];
  const missing = [];
  const moved = [];     // sources found somewhere other than their configured path
  const fellBack = [];  // list-backed sources that had to use their workbook fallback
  const stale = {};     // project -> last date we actually have, for sources that failed
  const perSource = {}; // diagnostics per source file (visible in ?refresh=1&dryrun=1)
  for (const raw of sources()) {
    const src = normalizeSource(raw);
    const proj = src.project;
    try {
      const info = {}; let cleaned, where;
      if (src.kind === 'splist') {
        try {
          const got = await graphListRows(src, token, log);
          cleaned = cleanListRows(proj, got.items, src, log, info);
          where = 'sharepoint list "' + got.listName + '" @ ' + src.site;
          info.sampleKeys = got.sampleKeys;
        } catch (e) {
          if (!src.fallback) throw e;
          // never let a list problem silently zero a project: fall back to the workbook and say so
          log('  !! SHAREPOINT LIST FAILED for ' + proj + ': ' + e.message + ' — falling back to the OneDrive workbook');
          const fb = src.fallback;
          const got = await fetchSource(fb.driveId, fb.paths, token, log);
          cleaned = cleanSource(proj, got.buf, fb.sheets, fb.dedupe, log, info, fb.forceProject);
          where = 'workbook ' + got.path + ' (FALLBACK — list unavailable)';
          info.listError = e.message.slice(0, 200);
          fellBack.push(proj);
        }
      } else {
        const got = await fetchSource(src.driveId, src.path, token, log);
        cleaned = cleanSource(proj, got.buf, src.sheets, src.dedupe, log, info, src.forceProject);
        where = got.path;
        if (got.movedFrom) { info.movedFrom = got.movedFrom; moved.push(proj); }
      }
      let maxD = ''; cleaned.forEach(r => { if (r[7] > maxD) maxD = r[7]; });
      perSource[proj] = Object.assign({ rows: cleaned.length, newestDate: maxD, source: where }, info);
      rows = rows.concat(cleaned);
    } catch (e) {
      log('  !! MISSING ' + proj + ': ' + e.message);
      perSource[proj] = { error: e.message.slice(0, 200) };
      missing.push(proj);
    }
  }
  // HISTORY MERGE (generalises the old SOUTH-only carry-forward): the building teams
  // periodically move old rows out of the working sheets (into DUPLICATES), which would
  // silently erase history from visitor.xlsx. For every project, keep rows from the
  // previous publish that are OLDER than the project's new data window; a project whose
  // source is missing/empty keeps ALL its previous rows.
  try {
    const pr = await fetch(vmsUrlFallback + '?t=' + Date.now());
    if (pr.ok) {
      const prevAll = prevRowsAll(Buffer.from(await pr.arrayBuffer()));
      const minNew = {};
      rows.forEach(r => { const p = String(r[6] || '').trim().toUpperCase(); if (!minNew[p] || r[7] < minNew[p]) minNew[p] = r[7]; });
      const mergedBy = {};
      prevAll.forEach(r => {
        const p = String(r[6] || '').trim().toUpperCase();
        if (!minNew[p] || r[7] < minNew[p]) {
          rows.push(r); mergedBy[p] = (mergedBy[p] || 0) + 1;
          if (!minNew[p] && (!stale[p] || r[7] > stale[p])) stale[p] = r[7]; // whole project is history
        }
      });
      if (Object.keys(mergedBy).length) {
        log('  history merged from previous publish: ' + JSON.stringify(mergedBy));
        Object.keys(mergedBy).forEach(p => { perSource[p + ' (history)'] = { rows: mergedBy[p], staleSince: stale[p] || null }; });
      }
    }
  } catch (e) { log('  history merge failed: ' + e.message); }
  // A source that failed to download is NOT a project with no visitors: every row above is
  // last-publish history, so the totals stay plausible while silently frozen. Balqis sat like
  // that for 5 days in Aug 2026. Say it out loud and hand the caller something to act on.
  if (missing.length) {
    log('  !! STALE PROJECTS: ' + missing.map(p => p + ' (carried forward, data ends ' + (stale[p] || 'unknown') + ')').join(', ')
      + ' — these figures are NOT current.');
  }
  if (moved.length) log('  note: sources relocated since last config update: ' + moved.join(', '));
  if (fellBack.length) log('  !! FELL BACK TO WORKBOOK (SharePoint list unreachable): ' + fellBack.join(', ')
    + ' — check the app registration has Graph Sites.Read.All.');
  if (!rows.length) throw new Error('VMS refresh produced 0 rows — keeping previous file');

  // build visitor.xlsx (sheet FINAL, same header). Dates are raw Excel serials with a
  // date number format so Excel/openpyxl show real dates and the dashboard parser
  // (XLSX numeric-serial branch) reads them exactly as before.
  const aoa = [HEADER].concat(rows.map(r => r.slice(0, 7)));
  const ws = XLSX.utils.aoa_to_sheet(aoa);
  for (let i = 1; i <= rows.length; i++) {
    const cell = ws['A' + (i + 1)];
    if (cell && cell.t === 'n') cell.z = 'dd/mm/yyyy hh:mm';
  }
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'FINAL');
  const xlsxBuffer = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx', compression: true });

  // records for the email — same shape AND same filters as the dashboard's parser
  // (it also drops rows with an empty Company Name at load time)
  const records = rows
    .filter(r => String(r[3] || '').trim())
    .map(r => ({
      date: r[7],
      purpose: String(r[2] || '').trim(),
      company: String(r[3] || '').trim(),
      unit: String(r[5] || '').trim(),
      project: String(r[6] || '').trim().toUpperCase()
    }));

  const counts = {};
  records.forEach(r => { counts[r.project] = (counts[r.project] || 0) + 1; });
  log('VMS refresh: ' + rows.length + ' rows total', JSON.stringify(counts));
  return { records, xlsxBuffer, counts, missing, stale, moved, fellBack, total: rows.length, perSource };
}

/* ---- commit visitor.xlsx to the dashboard repo (GitHub contents API) ---- */
async function commitVisitorXlsx(xlsxBuffer, context) {
  const log = context ? context.log.bind(context) : () => {};
  const token = process.env.GH_TOKEN;
  if (!token) { log('GH_TOKEN not set — skipping GitHub commit'); return false; }
  const repo = process.env.GH_REPO || 'Strive-Services-Group/Visitor-Competitor-Dashboard';
  const branch = process.env.GH_BRANCH || 'main';
  const api = 'https://api.github.com/repos/' + repo + '/contents/visitor.xlsx';
  const H = { Authorization: 'Bearer ' + token, Accept: 'application/vnd.github+json', 'User-Agent': 'telemetry-email-fn', 'X-GitHub-Api-Version': '2022-11-28' };
  // current sha via directory listing (works regardless of file size)
  let sha;
  const lr = await fetch('https://api.github.com/repos/' + repo + '/contents/?ref=' + branch, { headers: H });
  if (lr.ok) { const list = await lr.json(); const f = (list || []).find(x => x.name === 'visitor.xlsx'); if (f) sha = f.sha; }
  const body = { message: 'Auto VMS refresh (9 AM telemetry email) ' + new Date().toISOString().slice(0, 16) + 'Z', content: xlsxBuffer.toString('base64'), branch };
  if (sha) body.sha = sha;
  const r = await fetch(api, { method: 'PUT', headers: { ...H, 'Content-Type': 'application/json' }, body: JSON.stringify(body) });
  if (!r.ok) { const j = await r.json().catch(() => ({})); throw new Error('github commit ' + r.status + ' ' + JSON.stringify(j).slice(0, 200)); }
  log('visitor.xlsx committed to ' + repo + '@' + branch);
  return true;
}

module.exports = { refreshVms, commitVisitorXlsx, cleanSource, cleanListRows, graphListRows, nkey, parseDateCell, prevRowsAll };
