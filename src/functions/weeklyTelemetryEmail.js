/*
 * VISITOR TELEMETRY — WEEKLY DAYS COMPARISON EMAIL (Mondays only).
 *
 * Sends every MONDAY 9:00 AM Dubai (05:00 UTC), alongside the daily email.
 * The 8:50 AM refresh job (telemetryEmail.js) has already committed a fresh
 * visitor.xlsx containing Sunday's check-ins, so this prefers that published file.
 * If it looks stale (no Sunday data) it refreshes from the OneDrive sources itself
 * (without committing — the daily pipeline stays the only committer), and as a
 * last resort uses the stale published file with a warning banner.
 * S&C is a LIVE Dataverse query at send time (never cached), so bookings entered
 * late for earlier days are always included.
 *
 * Layout (approved by CK 2026-08-16, preview: Visitor_Telemetry_WEEKLY_Email_preview.html):
 *   Per flagship project (6): rows = last 4 ISO weeks (Last Week + 3 prior),
 *   columns = Mon..Sun + Week Total, repeated per service band
 *   (All Visitors / Laundry / Housekeeping / Maintenance / Fitout).
 *   Cell = "S&C / Other" + rectangular share% pill (red <20 <= amber <50 <= green).
 *   Per-project identity colours; black separator line after each project.
 *
 * Counting rules = IDENTICAL to the daily email (telemetryEmail.js), just a 4-week window:
 *   - S&C  = Candoo bookings (Reactive only, no cancelled, no Team Leader/Member,
 *            no laundry/fit-out depts) + laundry msdyn_workorders + our fit-out
 *            check-ins from visitor.xlsx. Same de-dupe (date|unit|title|dept).
 *   - Other = visitor.xlsx Unit Visit rows minus our companies minus Dima.
 *   - Recipients/sender/env = same as daily (MAIL_FROM, MAIL_TO, DV_RESOURCE, ...).
 *
 * HTTP endpoints for testing:
 *   GET /api/weekly-telemetry?code=...               -> JSON summary (no send)
 *   GET /api/weekly-telemetry?code=...&format=html   -> preview the email in the browser
 *   GET /api/weekly-telemetry?code=...&send=1        -> build + SEND now
 */
const { app } = require('@azure/functions');
const { parseVms, canonService } = require('./telemetryEmail.js');
const { refreshVms } = require('../shared/vmsRefresh.js');

const VMS_URL = process.env.VMS_URL || 'https://strive-services-group.github.io/Visitor-Competitor-Dashboard/visitor.xlsx';
const DASHBOARD_URL = process.env.DASHBOARD_URL || 'https://strive-services-group.github.io/Visitor-Competitor-Dashboard/';
const DV_DIVISION_CANDOO = process.env.DV_DIVISION_CANDOO || '93bd43eb-0a44-ef11-a316-6045bd6a8335';
const F = '@OData.Community.Display.V1.FormattedValue';

const TELE_PROJECTS = ['BALQIS RESIDENCE', 'THE8', 'SOUTH RESIDENCE', 'NORTH RESIDENCE', 'MAG 318', 'SHORELINE 7AND8'];
const TELE_PROJ_LBL = { 'BALQIS RESIDENCE': 'Balqis Residence', 'THE8': 'THE8', 'SOUTH RESIDENCE': 'South Residence', 'NORTH RESIDENCE': 'North Residence', 'MAG 318': 'MAG 318', 'SHORELINE 7AND8': 'Shoreline 7AND8' };
const PROJ_UNITS = { 'BALQIS RESIDENCE': 280, 'THE8': 195, 'SOUTH RESIDENCE': 281, 'NORTH RESIDENCE': 281, 'MAG 318': 439, 'SHORELINE 7AND8': 250 };
// per-project identity colours: [accent, light tint]
const PROJ_COL = {
  'BALQIS RESIDENCE': ['#1D4ED8', '#EAF1FF'], 'THE8': ['#7C3AED', '#F3EDFF'], 'SOUTH RESIDENCE': ['#0E9F6E', '#E6F9F1'],
  'NORTH RESIDENCE': ['#D97706', '#FFF3E0'], 'MAG 318': ['#DC2626', '#FDECEC'], 'SHORELINE 7AND8': ['#0E7490', '#E4F6FA']
};
const PROJ_ICON = { 'BALQIS RESIDENCE': '&#127961;&#65039;', 'THE8': '&#127970;', 'SOUTH RESIDENCE': '&#127960;&#65039;', 'NORTH RESIDENCE': '&#127980;', 'MAG 318': '&#127959;&#65039;', 'SHORELINE 7AND8': '&#127754;' };
// service bands: [label, accent, light tint]
const SVCS = [
  ['ALL', '&#128101; All Visitors', '#1D4ED8', '#EFF4FF'],
  ['LAUNDRY', '&#129530; Laundry', '#6D28D9', '#F5F0FF'],
  ['HOUSEKEEP', '&#129529; Housekeeping', '#16A34A', '#EFFBF3'],
  ['MAINT', '&#128295; Maintenance/ Handyman', '#EA580C', '#FFF4EC'],
  ['FITOUT', '&#128736;&#65039; Fitout', '#0891B2', '#EDFAFD']
];
const DAYS = ['MON', 'TUE', 'WED', 'THU', 'FRI', 'SAT', 'SUN'];
const FONT = 'Aptos,Segoe UI,Arial,sans-serif';
const INK = '#1E293B';

function dvIsOurCompany(name) {
  const n = (name || '').trim().toLowerCase();
  return n.indexOf('candoo') === 0 || n.indexOf('s & c') === 0 || n.indexOf('s&c') === 0 || n.indexOf('strive') === 0;
}
function fmt(v) { return v || ''; }
function dubaiDate(iso) { return new Date(new Date(iso).getTime() + 4 * 3600 * 1000).toISOString().slice(0, 10); }
function escHtml(s) { return String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;'); }
function isoWeekNum(d) { // d = Date (UTC)
  const t = new Date(Date.UTC(d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate()));
  const dayNum = t.getUTCDay() || 7; t.setUTCDate(t.getUTCDate() + 4 - dayNum);
  const y0 = new Date(Date.UTC(t.getUTCFullYear(), 0, 1));
  return Math.ceil((((t - y0) / 86400000) + 1) / 7);
}
function fmtDM(d) { return d.toLocaleDateString('en-GB', { day: '2-digit', month: 'short', timeZone: 'UTC' }); }

/* Last 4 full ISO weeks (Mon..Sun), newest first. "Today" = Dubai date of the run. */
function lastFourWeeks() {
  const nowDubai = new Date(Date.now() + 4 * 3600 * 1000);
  const today = new Date(Date.UTC(nowDubai.getUTCFullYear(), nowDubai.getUTCMonth(), nowDubai.getUTCDate()));
  const dow = today.getUTCDay() || 7;                    // Mon=1..Sun=7
  const thisMon = new Date(today); thisMon.setUTCDate(thisMon.getUTCDate() - (dow - 1));
  const weeks = [];
  for (let i = 1; i <= 4; i++) {                          // i=1 -> last full week
    const mon = new Date(thisMon); mon.setUTCDate(mon.getUTCDate() - 7 * i);
    const sun = new Date(mon); sun.setUTCDate(sun.getUTCDate() + 6);
    weeks.push({ mon, sun });
  }
  return weeks;                                           // [last week, w-2, w-3, w-4]
}

/* ================= OAuth + Dataverse (same pattern as telemetryEmail.js) ================= */
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
  if (!r.ok) throw new Error('odata ' + r.status + ' ' + JSON.stringify(j.error || j).slice(0, 300));
  return j;
}

/* S&C rows over the 4-week window — 1:1 filters with the daily email / dashboard. */
async function fetchScRows(minDate, maxDate, context) {
  const resource = (process.env.DV_RESOURCE || '').replace(/\/+$/, '');
  if (!resource) throw new Error('DV_RESOURCE app setting is not set');
  const token = await getToken(resource);
  const API = resource + '/api/data/v9.2';
  const cut = new Date(minDate + 'T00:00:00Z'); cut.setUTCHours(cut.getUTCHours() - 4); // Dubai day start in UTC
  const inWin = d => d >= minDate && d <= maxDate;
  const rows = [];

  let url = API + '/bookableresourcebookings?$select=ssg_plannedstartdate,_resource_value,_bookingstatus_value'
    + '&$expand=msdyn_workorder($select=_ssg_project_value,_ssg_department_value,_ssg_building_value,_ssg_unit_value,ssg_title,_msdyn_workordertype_value)'
    + '&$filter=' + encodeURIComponent('msdyn_workorder/_ssg_division_value eq ' + DV_DIVISION_CANDOO + ' and ssg_plannedstartdate ge ' + cut.toISOString());
  let guard = 0;
  while (url && guard < 120) {
    const j = await dvGet(url, token);
    (j.value || []).forEach(r => {
      const w = r.msdyn_workorder || {};
      const stt = fmt(r['_bookingstatus_value' + F]).toLowerCase(); if (stt.indexOf('cancel') >= 0) return;
      const res = fmt(r['_resource_value' + F]).toLowerCase();
      if (res.indexOf('team leader') >= 0 || res.indexOf('team member') >= 0) return;
      const wtype = fmt(w['_msdyn_workordertype_value' + F]); if (!/reactive/i.test(wtype)) return;
      const dept0 = fmt(w['_ssg_department_value' + F]); if (/laundry/i.test(dept0)) return;
      if (canonService(dept0) === 'FITOUT') return;
      if (!r.ssg_plannedstartdate) return;
      const d = dubaiDate(r.ssg_plannedstartdate);
      if (!inWin(d)) return;
      rows.push({
        date: d, project: fmt(w['_ssg_project_value' + F]).trim().toUpperCase(), service: canonService(dept0),
        unit: (fmt(w['_ssg_building_value' + F]).trim() + '_' + fmt(w['_ssg_unit_value' + F]).trim()),
        title: (w.ssg_title || ''), dept: dept0
      });
    });
    url = j['@odata.nextLink'] || null; guard++;
  }

  let wurl = API + '/msdyn_workorders?$select=msdyn_name,ssg_title,createdon,msdyn_systemstatus,_ssg_building_value,_ssg_unit_value,_ssg_project_value,_ssg_department_value,_msdyn_workordertype_value'
    + '&$filter=' + encodeURIComponent('_ssg_division_value eq ' + DV_DIVISION_CANDOO + ' and createdon ge ' + cut.toISOString());
  let wguard = 0;
  while (wurl && wguard < 120) {
    const wj = await dvGet(wurl, token);
    (wj.value || []).forEach(wo => {
      const wdept = fmt(wo['_ssg_department_value' + F]); if (!/laundry/i.test(wdept)) return;
      const wst = fmt(wo['msdyn_systemstatus' + F]); if (/cancel/i.test(wst)) return;
      if (!wo.createdon) return;
      const wd = dubaiDate(wo.createdon);
      if (!inWin(wd)) return;
      rows.push({
        date: wd, project: fmt(wo['_ssg_project_value' + F]).trim().toUpperCase(), service: 'LAUNDRY',
        unit: (fmt(wo['_ssg_building_value' + F]).trim() + '_' + fmt(wo['_ssg_unit_value' + F]).trim()),
        title: (wo.ssg_title || ''), dept: wdept
      });
    });
    wurl = wj['@odata.nextLink'] || null; wguard++;
  }

  const seen = {}, out = [];
  rows.forEach(rr => {
    const k = rr.date + '|' + rr.unit + '|' + String(rr.title || '').trim().toUpperCase() + '|' + String(rr.dept || rr.service || '').trim().toUpperCase();
    if (seen[k]) return; seen[k] = 1; out.push(rr);
  });
  if (context) context.log('weekly S&C rows after de-dupe:', out.length);
  return out;
}

/* ================= aggregate + build the email HTML ================= */
function buildWeeklyEmail(vmsRecords, scRows, warnNote) {
  const weeks = lastFourWeeks();
  const minDate = weeks[3].mon.toISOString().slice(0, 10);
  const maxDate = weeks[0].sun.toISOString().slice(0, 10);
  const inWin = d => d >= minDate && d <= maxDate;

  // agg[project][service][date] = { sc, ot }
  const agg = {};
  const bump = (p, s, d, side) => {
    ((agg[p] = agg[p] || {})[s] = agg[p][s] || {});
    (agg[p][s][d] = agg[p][s][d] || { sc: 0, ot: 0 })[side]++;
  };
  vmsRecords.forEach(r => {
    if (!r || !r.date || !inWin(r.date)) return;
    const p = (r.project || '').toUpperCase(); if (TELE_PROJECTS.indexOf(p) < 0) return;
    if (dvIsOurCompany(r.company)) {                 // our fit-out check-ins = OURS; other our check-ins not competitor
      if (canonService(r.purpose) === 'FITOUT') { bump(p, 'FITOUT', r.date, 'sc'); bump(p, 'ALL', r.date, 'sc'); }
      return;
    }
    const svc = canonService(r.purpose);             // Dima already excluded by parseVms
    bump(p, svc, r.date, 'ot'); bump(p, 'ALL', r.date, 'ot');
  });
  scRows.forEach(r => {
    if (!inWin(r.date) || TELE_PROJECTS.indexOf(r.project) < 0) return;
    bump(r.project, r.service, r.date, 'sc'); bump(r.project, 'ALL', r.date, 'sc');
  });

  const BD = 'border:1px solid #EDF2F7;';
  function pill(sc, ot) {
    const tot = sc + ot; if (!tot) return '';
    const pn = Math.round(sc / tot * 100);
    let bg = '#FDE8E8', fg = '#C81E1E';
    if (pn >= 50) { bg = '#DEF7EC'; fg = '#03543F'; } else if (pn >= 20) { bg = '#FDF6B2'; fg = '#723B13'; }
    // one-cell table: Outlook ignores span padding but respects td padding
    return '<table cellpadding="0" cellspacing="0" border="0" align="center" style="border-collapse:separate;margin:3px auto 0;"><tr>'
      + '<td bgcolor="' + bg + '" style="background:' + bg + ';border-radius:3px;padding:2px 10px;color:' + fg + ';font-weight:700;font-size:10px;font-family:' + FONT + ';text-align:center;white-space:nowrap;">' + pn + '%</td></tr></table>';
  }
  function cell(sc, ot, total, tint) {
    let base = BD + 'padding:6px 3px 5px;text-align:center;font-size:12px;font-family:' + FONT + ';white-space:nowrap;vertical-align:top;';
    if (total) base += 'background:' + (tint || '#F8FAFC') + ';';
    if (!sc && !ot) return '<td style="' + base + 'color:#D7DEE8;">&ndash;</td>';
    const w = total ? 'font-weight:800;' : 'font-weight:700;';
    return '<td style="' + base + '"><span style="' + w + 'color:' + INK + ';">' + sc + '</span>'
      + '<span style="color:#B6C0CE;"> / </span><span style="color:#64748B;">' + ot + '</span>' + pill(sc, ot) + '</td>';
  }

  let inner = '';
  TELE_PROJECTS.forEach(p => {
    const pc = PROJ_COL[p][0], pt = PROJ_COL[p][1];
    inner += '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="margin-top:20px;">'
      + '<tr><td style="background:' + pt + ';border:1px solid ' + pc + '33;border-left:6px solid ' + pc + ';border-radius:8px;padding:10px 14px;">'
      + '<span style="font-family:' + FONT + ';font-size:15px;font-weight:800;color:' + pc + ';">' + PROJ_ICON[p] + ' ' + escHtml(TELE_PROJ_LBL[p]) + '</span>'
      + '<span style="font-family:' + FONT + ';font-weight:600;color:#71809B;font-size:12px;"> &#183; ' + (PROJ_UNITS[p] || 0) + ' units</span></td></tr></table>';
    inner += '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="border-collapse:collapse;margin-top:8px;table-layout:fixed;">';
    inner += '<tr><th style="' + BD + 'border-left:4px solid ' + pc + ';background:' + pt + ';font-family:' + FONT + ';font-size:10px;color:' + pc + ';letter-spacing:.8px;padding:7px 8px;text-align:left;width:150px;">WEEK</th>';
    DAYS.forEach(d => { inner += '<th style="' + BD + 'background:' + pt + ';font-family:' + FONT + ';font-size:10px;color:' + pc + ';letter-spacing:.8px;padding:7px 3px;width:92px;">' + d + '</th>'; });
    inner += '<th style="' + BD + 'background:' + pt + ';font-family:' + FONT + ';font-size:10px;color:' + pc + ';letter-spacing:.8px;padding:7px 3px;width:96px;">TOTAL</th></tr>';
    const P = agg[p] || {};
    SVCS.forEach(sv => {
      const key = sv[0], lbl = sv[1], accent = sv[2], tint = sv[3];
      inner += '<tr><td colspan="9" style="' + BD + 'border-left:4px solid ' + accent + ';background:' + tint + ';color:' + accent + ';font-family:' + FONT + ';font-size:12px;font-weight:800;padding:6px 10px;">' + lbl + '</td></tr>';
      const S = P[key] || {};
      weeks.forEach((w, i) => {
        const wl = i === 0 ? 'Last Week&nbsp;&#183;&nbsp;W' + isoWeekNum(w.mon) : 'Week ' + isoWeekNum(w.mon);
        const rng = fmtDM(w.mon) + ' &ndash; ' + fmtDM(w.sun);
        inner += '<tr><td style="' + BD + 'padding:6px 10px;font-family:' + FONT + ';font-size:11px;"><b style="color:' + INK + ';">' + wl + '</b><br><span style="color:#9AA7B5;font-size:10px;">' + rng + '</span></td>';
        let tsc = 0, tot = 0;
        for (let di = 0; di < 7; di++) {
          const day = new Date(w.mon); day.setUTCDate(day.getUTCDate() + di);
          const v = S[day.toISOString().slice(0, 10)] || { sc: 0, ot: 0 };
          tsc += v.sc; tot += v.ot;
          inner += cell(v.sc, v.ot, false);
        }
        inner += cell(tsc, tot, true, pt) + '</tr>';
      });
    });
    inner += '</table>';
    // full-width black separator after each project
    inner += '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="margin-top:20px;"><tr><td style="height:3px;background:#111827;font-size:0;line-height:0;border-radius:2px;">&nbsp;</td></tr></table>';
  });

  const nowDubai = new Date(Date.now() + 4 * 3600 * 1000);
  const stamp = nowDubai.toISOString().slice(0, 10) + ' ' + nowDubai.toISOString().slice(11, 16);
  const subject = 'Visitor Telemetry (Flagship Projects) - Weekly Days Comparison (' + fmtDM(weeks[3].mon) + ' - ' + fmtDM(weeks[0].sun) + ')';
  const html =
    '<div style="font-family:' + FONT + ';color:#22303C;background:#EEF2F7;padding:16px;">'
    + '<table role="presentation" width="1000" cellpadding="0" cellspacing="0" align="center">'
    + '<tr><td bgcolor="#14315E" style="background:#14315E;border-radius:10px 10px 0 0;padding:18px 24px;">'
    + '<div style="font-family:' + FONT + ';font-size:17px;font-weight:800;color:#FFFFFF;letter-spacing:.3px;">&#128202; VISITOR TELEMETRY &mdash; WEEKLY DAYS COMPARISON</div>'
    + '<div style="font-family:' + FONT + ';font-size:12px;color:#C9D6EC;margin-top:5px;">Flagship projects &#183; last 4 weeks &#183; each cell = <b style="color:#FFFFFF;">S &amp; C</b> / Other (competitor) &#183; pill = our share % &#183; Dima excluded &#183; snapshot ' + stamp + ' (Dubai) &#183; <a href="' + DASHBOARD_URL + '" style="color:#9CC6FF;font-weight:700;">Open the Live Dashboard</a></div>'
    + '</td></tr>'
    + '<tr><td bgcolor="#FFFFFF" style="background:#FFFFFF;border:1px solid #E2E8F0;border-top:none;border-radius:0 0 10px 10px;padding:6px 24px 20px;">'
    + (warnNote ? '<div style="font-family:' + FONT + ';font-size:12px;font-weight:700;color:#9A3412;background:#FFF7ED;border:1px solid #FDBA74;border-radius:5px;padding:6px 10px;margin:14px 0 0;display:inline-block;">&#9888;&#65039; ' + escHtml(warnNote) + '</div>' : '')
    + inner
    + '<div style="font-family:' + FONT + ';margin-top:14px;color:#9AA7B5;font-size:11px;">Automated Monday report &#183; Strive Services Group &#183; S &amp; C: Candoo bookings &amp; work orders (Dynamics 365) + our fit-out check-ins &#183; Other: building visitor logs</div>'
    + '</td></tr></table></div>';

  return { subject, html, window: { from: minDate, to: maxDate } };
}

/* ================= send (same Graph pattern + recipients as daily) ================= */
async function sendMail(subject, html, context) {
  const from = process.env.MAIL_FROM;
  const toList = (process.env.MAIL_TO || '').split(/[;,]/).map(s => s.trim()).filter(Boolean);
  if (!from) throw new Error('MAIL_FROM app setting is not set');
  if (!toList.length) throw new Error('MAIL_TO app setting is not set');
  const token = await getToken('https://graph.microsoft.com');
  const r = await fetch('https://graph.microsoft.com/v1.0/users/' + encodeURIComponent(from) + '/sendMail', {
    method: 'POST',
    headers: { Authorization: 'Bearer ' + token, 'Content-Type': 'application/json' },
    body: JSON.stringify({
      message: { subject, body: { contentType: 'HTML', content: html }, toRecipients: toList.map(a => ({ emailAddress: { address: a } })) },
      saveToSentItems: true
    })
  });
  if (r.status !== 202) {
    const j = await r.json().catch(() => ({}));
    throw new Error('sendMail ' + r.status + ' ' + JSON.stringify(j.error || j).slice(0, 300));
  }
  if (context) context.log('weekly email sent to', toList.length, 'recipients');
}

/* ================= the full job ================= */
// Prefers the visitor.xlsx committed by the 8:50 refresh job (contains Sunday).
// If stale/unreachable: refreshes from OneDrive itself WITHOUT committing (the
// daily pipeline stays the only committer). Last resort: stale file + banner.
async function runWeekly(context, doSend) {
  const wks = lastFourWeeks();
  const needThrough = wks[0].sun.toISOString().slice(0, 10); // last Sunday must be present
  const newest = recs => { let m = ''; (recs || []).forEach(r => { if (r && r.date && r.date > m) m = r.date; }); return m; };
  let vms = null, warnNote = null;
  // 1) published file (committed minutes ago by the 8:50 job)
  try {
    const vr = await fetch(VMS_URL + '?t=' + Date.now());
    if (vr.ok) {
      const recs = parseVms(Buffer.from(await vr.arrayBuffer()));
      if (newest(recs) >= needThrough) vms = recs;
      else if (context) context.warn('weekly: published visitor.xlsx stale (newest ' + newest(recs) + ', need ' + needThrough + ') — refreshing inline');
    }
  } catch (e) { if (context) context.warn('weekly: published visitor.xlsx fetch failed: ' + e.message); }
  // 2) own refresh, no commit
  if (!vms) {
    try { const rf = await refreshVms(getToken, VMS_URL, context); vms = rf.records; }
    catch (e) { if (context) context.warn('weekly: VMS refresh failed, falling back to last published visitor.xlsx: ' + e.message); }
  }
  // 3) stale published + banner
  if (!vms) {
    const vr = await fetch(VMS_URL + '?t=' + Date.now());
    if (!vr.ok) throw new Error('visitor.xlsx fetch failed: ' + vr.status);
    vms = parseVms(Buffer.from(await vr.arrayBuffer()));
    warnNote = 'Live visitor-log refresh failed this morning - competitor (Other) figures show the last published data (Sunday may be missing). S & C figures are live.';
  }
  const weeks = lastFourWeeks();
  const minDate = weeks[3].mon.toISOString().slice(0, 10);
  const maxDate = weeks[0].sun.toISOString().slice(0, 10);
  const scRows = await fetchScRows(minDate, maxDate, context);
  const out = buildWeeklyEmail(vms, scRows, warnNote);
  out.scRowCount = scRows.length;
  if (doSend) await sendMail(out.subject, out.html, context);
  return out;
}

/* ================= triggers ================= */
// Mondays 05:00 UTC = 9:00 AM Dubai (daily email follows at 9:05; weekly refreshes VMS itself)
app.timer('weekly-telemetry-monday', {
  schedule: '0 0 5 * * 1',
  handler: async (timer, context) => {
    try { await runWeekly(context, true); }
    catch (e) { context.error('weekly telemetry email FAILED:', e.message); throw e; }
  }
});

app.http('weekly-telemetry', {
  methods: ['GET', 'OPTIONS'],
  authLevel: 'function',
  route: 'weekly-telemetry',
  handler: async (request, context) => {
    try {
      const url = new URL(request.url);
      const doSend = url.searchParams.get('send') === '1';
      const out = await runWeekly(context, doSend);
      if (url.searchParams.get('format') === 'html' && !doSend) {
        return { status: 200, headers: { 'Content-Type': 'text/html; charset=utf-8' }, body: out.html };
      }
      return { status: 200, jsonBody: { sent: doSend, subject: out.subject, window: out.window, scRows: out.scRowCount } };
    } catch (e) {
      context.error('weekly-telemetry failed:', e);
      return { status: 500, jsonBody: { error: e.message } };
    }
  }
});

module.exports = { buildWeeklyEmail, lastFourWeeks };
