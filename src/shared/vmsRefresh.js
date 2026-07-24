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
 *       /Desktop/Pool and beach Records/DAILY CONTRACTORS RECORDS/Visitors Details Balqis Residence.xlsx
 *   - Abdul Muqeet's OneDrive     (abdul_muqeet_sahalahfm_com):
 *       /VMS-DATA FILES/VMS-TH8.xlsx, VMS-AL HASEER.xlsx, VMS-AL NABAT.xlsx,
 *       /VMS-DATA FILES/VMS-NORTH RESIDENCE.xlsx, VMS-SOUTH RESIDENCE.xlsx
 *
 * Cleaning rules (identical to clean_vms.py):
 *   - keep purposes: LAUNDRY, INSPECTION, CLEANING/CLEANERS, MAINTENANCE/HANDYMAN,
 *     FIT-OUT, AMCCONTRACTORS (compared uppercase, spaces stripped)
 *   - "Unit Visit" check-ins only; Dima excluded everywhere
 *   - date from cell (Excel date or "DDMMYYYY hh:mm[:ss]" text), year 2024–2027
 *   - per-file de-dupe (date|purpose|unit|company) for the Abdul Muqeet files
 *   - SOUTH RESIDENCE carried forward from the previously published visitor.xlsx
 *     if its source file is missing/empty
 *   - output sheet 'FINAL' with the exact dashboard header
 */
const XLSX = require('xlsx');

const BALQIS_DRIVE = process.env.VMS_BALQIS_DRIVE || 'b!5ma3QhDyZ0GZsiXhkdXOuhs9QX6ol9RInUHGE6t7AIyt1DCpPIAcS6cBimNKf0JF';
const MUQEET_DRIVE = process.env.VMS_MUQEET_DRIVE || 'b!2jsyR69LVE63449EoxnFAaA1OVzt3PxKqmAZ7NtQgKH9TivpAV8ORZN5aAEIpdJm';

// [project default, driveId, path in drive, sheet names (null = all), dedupe, forceProject]
// forceProject=true: single-project files — ALWAYS use our project name, ignore the file's
// "Project Name" column (Abdul's restructured SOUTH file carried "NORTH RESIDENCE" in that
// column, which mislabelled all South visits as North). AL HASEER / AL NABAT keep the column
// because it maps them to SHORELINE 7AND8; Balqis keeps it as before.
const DEFAULT_SOURCES = [
  ['BALQIS RESIDENCE', BALQIS_DRIVE, '/Desktop/Pool and beach Records/DAILY CONTRACTORS RECORDS/Visitors Details Balqis Residence.xlsx', null, false, false],
  ['THE8', MUQEET_DRIVE, '/VMS-DATA FILES/VMS-TH8.xlsx', ["VMS-TH8-'26"], true, true],
  ['AL HASEER', MUQEET_DRIVE, '/VMS-DATA FILES/VMS-AL HASEER.xlsx', ["VMS-AL HASEER-'26"], true, false],
  ['AL NABAT', MUQEET_DRIVE, '/VMS-DATA FILES/VMS-AL NABAT.xlsx', ["VMS-AL NABAT-'26"], true, false],
  ['NORTH RESIDENCE', MUQEET_DRIVE, '/VMS-DATA FILES/VMS-NORTH RESIDENCE.xlsx', ["VMS-NORTH RESIDENCE-'26"], true, true],
  ['SOUTH RESIDENCE', MUQEET_DRIVE, '/VMS-DATA FILES/VMS-SOUTH RESIDENCE.xlsx', ["VMS-SOUTH RESIDENCE-'26"], true, true],
  ['MAG 318', MUQEET_DRIVE, '/VMS-DATA FILES/VMS-MAG 318.xlsx', ["VMS-MAG 318-'26"], true, true]
];
function sources() {
  try { if (process.env.VMS_SOURCES) return JSON.parse(process.env.VMS_SOURCES); } catch (e) {}
  return DEFAULT_SOURCES;
}

const KEEP = new Set(['LAUNDRY', 'INSPECTION', 'CLEANING/CLEANERS', 'MAINTENANCE/HANDYMAN', 'FIT-OUT', 'AMCCONTRACTORS']);
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
    const mm = v.trim().match(/^(\d{2})(\d{2})(\d{4})(?:\s+(\d{1,2}):(\d{2})(?::(\d{2}))?)?/);
    if (!mm) return null;
    d = +mm[1]; m = +mm[2]; y = +mm[3]; hh = +(mm[4] || 0); mi = +(mm[5] || 0); ss = +(mm[6] || 0);
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
  const perSource = {}; // diagnostics per source file (visible in ?refresh=1&dryrun=1)
  for (const [proj, driveId, path, sheets, dedupe, forceProject] of sources()) {
    try {
      const buf = await graphDownload(driveId, path, token);
      const info = {};
      const cleaned = cleanSource(proj, buf, sheets, dedupe, log, info, forceProject);
      let maxD = ''; cleaned.forEach(r => { if (r[7] > maxD) maxD = r[7]; });
      perSource[proj] = { rows: cleaned.length, newestDate: maxD, bytes: info.bytes, sheets: info.sheets, matchedSheets: info.matchedSheets };
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
        if (!minNew[p] || r[7] < minNew[p]) { rows.push(r); mergedBy[p] = (mergedBy[p] || 0) + 1; }
      });
      if (Object.keys(mergedBy).length) {
        log('  history merged from previous publish: ' + JSON.stringify(mergedBy));
        Object.keys(mergedBy).forEach(p => { perSource[p + ' (history)'] = { rows: mergedBy[p] }; });
      }
    }
  } catch (e) { log('  history merge failed: ' + e.message); }
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
  return { records, xlsxBuffer, counts, missing, total: rows.length, perSource };
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

module.exports = { refreshVms, commitVisitorXlsx, cleanSource, parseDateCell, prevRowsAll };
