'use strict';

const crypto = require('crypto');
const PO_STEP_MAP = require('../../poStepMap.json');
const PR_STAGE_MAP = require('../../prStageMap.json');
const PO_APPROVAL_STAGE = {
  'Procurement Manager': 'Procurement',
  'Advance payment request submitted (if applicable)': 'Procurement',
  'Accounting Manager': 'Finance',
  'Finance and Accounts Director': 'Director',
  'Finance': 'Finance',
  'CEO': 'CEO'
};

const CACHE_MS = Number(process.env.PRPO_DATASET_CACHE_MS || 180000);
const WORKBOOK_EXPORT_UTC = '2026-09-07T05:30:00.000Z';
const OBS_PREFIX = 'ifahr-live|PO|';
let cache = null;
let refreshPromise = null;
let foTokenCache = null;
let dvTokenCache = null;

function norm(value) { return String(value == null ? '' : value).trim().toLowerCase().replace(/\s+/g, ' '); }
function doc(value) { return String(value == null ? '' : value).trim().toUpperCase(); }
function num(value) { const n = Number(value); return Number.isFinite(n) ? n : 0; }
function iso(value) { if (!value) return null; const d = new Date(value); return Number.isNaN(d.getTime()) ? null : d.toISOString(); }
function earliest(values) {
  const dates = values.map(iso).filter(Boolean).sort();
  return dates[0] || null;
}
function parseDim(value) {
  if (!value) return { contract: null, department: null, location: null };
  const parts = String(value).split('-').map(x => x.trim());
  return { contract: parts[1] || null, department: parts[3] || null, location: parts[5] || null };
}
function entityKey(company, number) { return `${norm(company)}|${doc(number)}`; }
function observationKey(company, number) { return `${OBS_PREFIX}${norm(company)}|${doc(number)}`; }
function prNumber(value) { const m = String(value || '').match(/C?PR-\d+/i); return m ? m[0].toUpperCase() : null; }
function poNumber(value) {
  const subject = String(value || '');
  const labelled = subject.match(/Purchase Order:\s*([A-Za-z0-9-]+)/i);
  if (labelled) return labelled[1].toUpperCase();
  const bare = subject.match(/[A-Z]+-PO\d+|P?O?\d{6,}/i);
  return bare ? bare[0].toUpperCase() : null;
}
function encodeKey(value) { return String(value).replace(/'/g, "''"); }

async function foToken() {
  if (foTokenCache && foTokenCache.expiresAt > Date.now() + 120000) return foTokenCache.token;
  const resource = String(process.env.FO_RESOURCE || '').replace(/\/+$/, '');
  const body = new URLSearchParams({
    client_id: process.env.CLIENT_ID,
    client_secret: process.env.CLIENT_SECRET,
    grant_type: 'client_credentials',
    scope: `${resource}/.default`
  });
  const response = await fetch(`https://login.microsoftonline.com/${process.env.TENANT_ID}/oauth2/v2.0/token`, {
    method: 'POST', headers: { 'Content-Type': 'application/x-www-form-urlencoded' }, body
  });
  if (!response.ok) throw new Error(`F&O token ${response.status}: ${(await response.text()).slice(0, 240)}`);
  const payload = await response.json();
  foTokenCache = { token: payload.access_token, expiresAt: Date.now() + Number(payload.expires_in || 300) * 1000 };
  return foTokenCache.token;
}

function crossCompanyPath(path) {
  return path + (path.includes('?') ? '&' : '?') + 'cross-company=true';
}

async function odataAll(token, path) {
  const root = String(process.env.FO_RESOURCE || '').replace(/\/+$/, '') + '/data/';
  let url = root + crossCompanyPath(path);
  const rows = [];
  let pages = 0;
  while (url && pages++ < 500) {
    const response = await fetch(url, { headers: { Authorization: `Bearer ${token}`, Accept: 'application/json' } });
    if (!response.ok) throw new Error(`F&O ${response.status} @ ${new URL(url).pathname}: ${(await response.text()).slice(0, 240)}`);
    const payload = await response.json();
    rows.push(...(payload.value || []));
    url = payload['@odata.nextLink'] || null;
  }
  if (pages >= 500) throw new Error('F&O page safety limit reached');
  return rows;
}

async function dvToken() {
  if (process.env.PRPO_DEV_TOKEN) return process.env.PRPO_DEV_TOKEN;
  if (dvTokenCache && dvTokenCache.expiresAt > Date.now() + 120000) return dvTokenCache.token;
  if (!process.env.IDENTITY_ENDPOINT || !process.env.IDENTITY_HEADER) throw new Error('Dataverse managed identity is unavailable');
  const resource = String(process.env.DATAVERSE_RESOURCE || 'https://operations-ifahr-dev.crm15.dynamics.com').replace(/\/+$/, '');
  const url = new URL(process.env.IDENTITY_ENDPOINT);
  url.searchParams.set('api-version', '2019-08-01');
  url.searchParams.set('resource', resource);
  const clientId = process.env.DATAVERSE_MI_CLIENT_ID || process.env.AZURE_CLIENT_ID;
  if (clientId) url.searchParams.set('client_id', clientId);
  const response = await fetch(url, { headers: { 'X-IDENTITY-HEADER': process.env.IDENTITY_HEADER } });
  if (!response.ok) throw new Error(`Dataverse managed identity ${response.status}`);
  const payload = await response.json();
  dvTokenCache = { token: payload.access_token, expiresAt: Number(payload.expires_on || 0) * 1000 || Date.now() + 300000 };
  return dvTokenCache.token;
}

function dvRoot() {
  return String(process.env.DATAVERSE_API_URL || 'https://operations-ifahr-dev.api.crm15.dynamics.com/api/data/v9.2').replace(/\/+$/, '');
}

async function dvRequest(path, options = {}) {
  const token = await dvToken();
  const response = await fetch(/^https?:/i.test(path) ? path : `${dvRoot()}/${path.replace(/^\//, '')}`, {
    ...options,
    headers: {
      Authorization: `Bearer ${token}`,
      Accept: 'application/json',
      'OData-MaxVersion': '4.0',
      'OData-Version': '4.0',
      ...(options.body ? { 'Content-Type': 'application/json; charset=utf-8' } : {}),
      ...(options.headers || {})
    }
  });
  if (!response.ok) throw new Error(`Dataverse ${response.status}: ${(await response.text()).slice(0, 400)}`);
  return response;
}

async function dvAll(setName, select, filter) {
  const params = new URLSearchParams({ '$select': select.join(',') });
  if (filter) params.set('$filter', filter);
  let next = `${setName}?${params}`;
  const rows = [];
  let pages = 0;
  while (next && pages++ < 200) {
    const response = await dvRequest(next);
    const payload = await response.json();
    rows.push(...(payload.value || []));
    next = payload['@odata.nextLink'] || null;
  }
  return rows;
}

async function upsertObservation(key, payload) {
  const path = `ssg_prpodocuments(ssg_documentkey='${encodeKey(key)}')`;
  await dvRequest(path, { method: 'PATCH', headers: { Prefer: 'return=minimal' }, body: JSON.stringify(payload) });
}

async function readObservation(key) {
  const rows = await dvAll('ssg_prpodocuments', [
    'ssg_documentkey','ssg_currentstepname','ssg_observedpendingsince',
    'ssg_dataqualitynotes','ssg_isreportable','ssg_payloadhash'
  ], `ssg_documentkey eq '${encodeKey(key)}'`);
  return rows[0] || null;
}

function parseObservationNotes(value) {
  try { const parsed = JSON.parse(String(value || '')); return parsed && typeof parsed === 'object' ? parsed : {}; }
  catch (_) { return {}; }
}

function observationPayload(po, clock, now) {
  const notes = {
    schema: 'PRPO_STAGE_CLOCK_V1',
    stage: po.liveStage,
    clockProvenance: clock.provenance,
    workbookValue: clock.workbookValue || null,
    workbookExportTimestamp: clock.workbookExportTimestamp || null,
    liveEvent: clock.liveEvent || null
  };
  return {
    ssg_documentkey: observationKey(po.legalEntity, po.number),
    ssg_name: `${po.legalEntity} ${po.number}`.slice(0, 100),
    ssg_documentnumber: po.number,
    ssg_documenttype: 'PO',
    ssg_legalentity: po.legalEntity,
    ssg_sourceenvironment: 'ifahr-live',
    ssg_sourceentity: 'F&O PO lifecycle stage observation',
    ssg_currentstepname: po.liveStage,
    ssg_stepenteredon: clock.provenance === 'LIVE_EVENT_DATE' ? clock.timestamp : null,
    ssg_observedpendingsince: clock.timestamp,
    ssg_sourceupdatedon: now,
    ssg_dataqualitystatus: po.liveStage === 'STAGE_NOT_EVIDENCED' ? 'Failed' : 'Complete',
    ssg_dataqualitynotes: JSON.stringify(notes),
    ssg_isreportable: po.isReportable !== false,
    ssg_payloadhash: crypto.createHash('sha256').update(JSON.stringify(notes)).digest('hex')
  };
}

function deriveClock(po, existing, now) {
  if (po.liveEventTimestamp) {
    return { timestamp: po.liveEventTimestamp, provenance: 'LIVE_EVENT_DATE', label: 'since', liveEvent: po.liveEvent };
  }
  if (existing && norm(existing.ssg_currentstepname) === norm(po.liveStage)) {
    const notes = parseObservationNotes(existing.ssg_dataqualitynotes);
    const provenance = notes.clockProvenance || (existing.ssg_observedpendingsince ? 'PENDING_SINCE_FIRST_OBSERVED' : 'NOT_RECORDED');
    return {
      timestamp: iso(existing.ssg_observedpendingsince), provenance,
      label: provenance === 'SEEDED_FROM_FINAL_WORKBOOK' ? 'since (from last export)' :
        (existing.ssg_observedpendingsince ? 'since' : 'since — not recorded'),
      workbookValue: notes.workbookValue || null,
      workbookExportTimestamp: notes.workbookExportTimestamp || null,
      liveEvent: notes.liveEvent || null
    };
  }
  if (existing) {
    return { timestamp: now, provenance: 'PENDING_SINCE_FIRST_OBSERVED', label: 'since', liveEvent: 'first observed after cutover' };
  }
  return { timestamp: null, provenance: 'NOT_RECORDED', label: 'since — not recorded', liveEvent: null };
}

function openPO(status, approval) {
  return !['invoiced', 'closed', 'cancelled', 'canceled'].includes(norm(status)) && norm(approval) !== 'rejected';
}

function lifecycleStage(header, key, context) {
  const number = doc(header.PurchaseOrderNumber);
  const status = norm(header.PurchaseOrderStatus);
  const approval = norm(header.DocumentApprovalStatus);
  const uniqueNumber = (context.numberKeys.get(number) || []).length === 1;
  const capture = uniqueNumber ? context.capture.get(number) : null;
  if (capture && capture.conflict) return { stage: 'STAGE_NOT_EVIDENCED', reason: 'CONTRADICTORY_APPROVAL_STAGES' };
  if (capture && capture.stage) return { stage: capture.stage, reason: capture.reason || 'APPROVAL_CAPTURE' };
  if (status === 'invoiced') return { stage: 'Invoiced', reason: 'F_AND_O_ORDER_STATUS' };
  if (status === 'received' || context.packing.has(key)) return { stage: 'Receipt posted', reason: status === 'received' ? 'F_AND_O_ORDER_STATUS' : 'POSTED_PACKING_SLIP' };
  if (approval === 'confirmed' && ['backorder', 'open order'].includes(status)) return { stage: 'Sent to supplier', reason: 'F_AND_O_ORDER_AND_APPROVAL_STATUS' };
  if (['backorder', 'open order'].includes(status) && ['draft', 'approved', 'in review', 'inreview'].includes(approval)) return { stage: 'Not yet sent', reason: 'F_AND_O_ORDER_AND_APPROVAL_STATUS' };
  return { stage: 'STAGE_NOT_EVIDENCED', reason: (!status || !approval) ? 'F_AND_O_STATUS_MISSING' : 'F_AND_O_STATUS_CONTRADICTORY' };
}

async function buildDataset() {
  const token = await foToken();
  const now = new Date().toISOString();
  const [prHeaders, prLines, prBi, poHeaders, poLines, vendors, confirmations, packing, invoices,
    snapshots, instances, captureItems, observations] = await Promise.all([
    odataAll(token, 'PurchaseRequisitionHeaders?$select=RequisitionNumber,RequisitionName,RequisitionStatus,DefaultProjectId,IFAHRQuotationReference,PreparerPersonnelNumber,RequisitionPurpose,DefaultRequestedDate,ProjectBuyingLegalEntityId'),
    odataAll(token, 'PurchaseRequisitionLines?$select=RequisitionNumber,LineAmount,PurchasePrice,LineStatus,DefaultLedgerDimensionDisplayValue,BuyingLegalEntityId'),
    odataAll(token, 'PurchReqTableBiEntities?$select=PurchReqId,TransDate,SubmittedDateTime,SysModifiedDateTime,CompanyInfoDefault'),
    odataAll(token, 'PurchaseOrderHeadersV2?$select=PurchaseOrderNumber,OrderVendorAccountNumber,InvoiceVendorAccountNumber,PurchaseOrderName,CurrencyCode,PurchaseOrderStatus,DocumentApprovalStatus,ProjectId,RequestedDeliveryDate,AccountingDate,dataAreaId'),
    odataAll(token, 'PurchaseOrderLinesV2?$select=PurchaseOrderNumber,LineAmount,PurchaseOrderLineStatus,DefaultLedgerDimensionDisplayValue,dataAreaId'),
    odataAll(token, 'VendorsV2?$select=VendorAccountNumber,VendorOrganizationName,dataAreaId'),
    odataAll(token, 'PurchaseOrderConfirmationHeaders?$select=PurchaseOrderNumber,ConfirmationDate,ConfirmationNumber,dataAreaId'),
    odataAll(token, 'VendPackingSlipJourBiEntities?$select=PurchId,PackingSlipId,DocumentDate,dataAreaId'),
    odataAll(token, 'VendInvoiceJourBiEntities?$select=PurchId,InvoiceId,InvoiceDate,SysModifiedDateTime,dataAreaId'),
    dvAll('ssg_prpocurrentapprovalsnapshots', ['ssg_documentnumber','ssg_documenttype','ssg_pendingapprovercount','ssg_pendingstepnames','ssg_oldestpendingsince','ssg_pendingapprovernames','ssg_pendinguserids','ssg_lastreconciledon'], 'ssg_pendingapprovercount gt 0'),
    dvAll('ssg_prpoapprovalinstances', ['ssg_prpoapprovalinstanceid','ssg_documentnumber','ssg_documenttype']),
    dvAll('ssg_prpoapprovalworkitems', ['_ssg_approvalinstance_value','ssg_assignedon','ssg_firstobservedon','ssg_stepelementid','ssg_iscurrent','ssg_approveruserid','ssg_approvername'], 'ssg_iscurrent eq true'),
    dvAll('ssg_prpodocuments', ['ssg_documentkey','ssg_currentstepname','ssg_observedpendingsince','ssg_dataqualitynotes','ssg_payloadhash'], "ssg_documenttype eq 'PO'")
  ]);

  const byKey = (rows, numberField) => {
    const map = new Map();
    for (const row of rows) {
      const key = entityKey(row.dataAreaId, row[numberField]);
      if (!map.has(key)) map.set(key, []);
      map.get(key).push(row);
    }
    return map;
  };
  const confirmationsBy = byKey(confirmations, 'PurchaseOrderNumber');
  const packingBy = byKey(packing, 'PurchId');
  const invoicesBy = byKey(invoices, 'PurchId');
  const poLineBy = byKey(poLines, 'PurchaseOrderNumber');
  const vendorBy = new Map(vendors.map(v => [entityKey(v.dataAreaId, v.VendorAccountNumber), v.VendorOrganizationName]));
  const numberKeys = new Map();
  for (const h of poHeaders) {
    const n = doc(h.PurchaseOrderNumber); const k = entityKey(h.dataAreaId, n);
    if (!numberKeys.has(n)) numberKeys.set(n, []);
    if (!numberKeys.get(n).includes(k)) numberKeys.get(n).push(k);
  }

  const instanceBy = new Map(instances.map(x => [x.ssg_prpoapprovalinstanceid, x]));
  const itemByNumber = new Map();
  const prItemByNumber = new Map();
  for (const item of captureItems) {
    const instance = instanceBy.get(item._ssg_approvalinstance_value);
    if (!instance) continue;
    const type = norm(instance.ssg_documenttype);
    const target = type === 'po' ? itemByNumber : (type === 'pr' ? prItemByNumber : null);
    if (!target) continue;
    const n = doc(instance.ssg_documentnumber);
    if (!target.has(n)) target.set(n, []);
    target.get(n).push(item);
  }
  const snapshotBy = new Map(snapshots.filter(x => norm(x.ssg_documenttype) === 'po').map(x => [doc(x.ssg_documentnumber), x]));
  const prSnapshotBy = new Map(snapshots.filter(x => norm(x.ssg_documenttype) === 'pr').map(x => [doc(x.ssg_documentnumber), x]));
  const capture = new Map();
  for (const [n, items] of itemByNumber) {
    const stages = new Set();
    for (const item of items) {
      const label = PO_STEP_MAP[norm(item.ssg_stepelementid)] || PO_STEP_MAP[item.ssg_stepelementid];
      if (label) stages.add(PO_APPROVAL_STAGE[label] || label);
    }
    const snapshot = snapshotBy.get(n);
    const snapshotIds = String(snapshot && snapshot.ssg_pendingstepnames || '').match(/[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}/ig) || [];
    for (const id of snapshotIds) {
      const label = PO_STEP_MAP[id.toLowerCase()];
      if (label) stages.add(PO_APPROVAL_STAGE[label] || label);
    }
    capture.set(n, {
      conflict: stages.size > 1,
      stage: stages.size === 1 ? [...stages][0] : (snapshot ? 'Approval — unmapped element' : null),
      reason: stages.size === 1 ? 'APPROVAL_CAPTURE' : 'UNMAPPED_APPROVAL_ELEMENT',
      timestamp: earliest([snapshot && snapshot.ssg_oldestpendingsince, ...items.map(x => x.ssg_assignedon || x.ssg_firstobservedon)]),
      approver: items.map(x => x.ssg_approvername || x.ssg_approveruserid).filter(Boolean).join(', ') || null
    });
  }
  for (const [n, snapshot] of snapshotBy) {
    if (!capture.has(n)) capture.set(n, { conflict: false, stage: 'Approval — unmapped element', reason: 'UNMAPPED_APPROVAL_ELEMENT', timestamp: iso(snapshot.ssg_oldestpendingsince), approver: null });
  }
  const existingObs = new Map(observations.filter(x => String(x.ssg_documentkey || '').startsWith(OBS_PREFIX)).map(x => [x.ssg_documentkey, x]));
  const context = { numberKeys, capture, packing: packingBy };

  const poRows = [];
  const pendingWrites = [];
  for (const h of poHeaders) {
    const number = doc(h.PurchaseOrderNumber); const company = norm(h.dataAreaId); const key = entityKey(company, number);
    const lifecycle = lifecycleStage(h, key, context);
    const cap = capture.get(number);
    let liveEvent = null; let liveEventTimestamp = null;
    if (lifecycle.stage === 'Invoiced') { liveEvent = 'posted vendor invoice'; liveEventTimestamp = earliest((invoicesBy.get(key) || []).map(x => x.SysModifiedDateTime || x.InvoiceDate)); }
    else if (lifecycle.stage === 'Receipt posted') { liveEvent = 'posted packing slip'; liveEventTimestamp = earliest((packingBy.get(key) || []).map(x => x.DocumentDate)); }
    else if (lifecycle.stage === 'Sent to supplier') { liveEvent = 'PO confirmation'; liveEventTimestamp = earliest((confirmationsBy.get(key) || []).map(x => x.ConfirmationDate)); }
    else if (lifecycle.stage.includes('Approval') || ['Procurement','Finance','Director','CEO'].includes(lifecycle.stage)) { liveEvent = 'approval capture assignment'; liveEventTimestamp = cap && cap.timestamp; }
    const po = { number, legalEntity: company, liveStage: lifecycle.stage, liveStageReason: lifecycle.reason, liveEvent, liveEventTimestamp };
    const clock = deriveClock(po, existingObs.get(observationKey(company, number)), now);
    const allLines = poLineBy.get(key) || [];
    const lines = allLines.filter(line => !norm(line.PurchaseOrderLineStatus).includes('cancel'));
    const firstLine = lines[0] || allLines[0] || {}; const dim = parseDim(firstLine.DefaultLedgerDimensionDisplayValue);
    const total = Math.round(lines.reduce((sum, line) => sum + num(line.LineAmount), 0) * 100) / 100;
    const row = {
      'Purchase order': number,
      'Legal entity': company,
      'Vendor account': h.OrderVendorAccountNumber || null,
      'Invoice account': h.InvoiceVendorAccountNumber || null,
      'Vendor name': vendorBy.get(entityKey(company, h.OrderVendorAccountNumber)) || h.OrderVendorAccountNumber || null,
      'Purchase type': null,
      'Approval status': h.DocumentApprovalStatus || null,
      'Purchase order status': h.PurchaseOrderStatus || null,
      'Currency': h.CurrencyCode || null,
      'Requested receipt date': h.RequestedDeliveryDate || null,
      'Created date and time': h.AccountingDate || null,
      'Purchase requisition': null,
      'RFQ number': null,
      'Total amount': total,
      'Department': dim.department,
      'Location': dim.location,
      'Contract': dim.contract,
      'Pending Approver/User': cap && cap.approver,
      'Step name': lifecycle.stage,
      'Step date and time': clock.timestamp,
      'Created by': null,
      'Live stage': lifecycle.stage,
      'Stage reason code': lifecycle.reason,
      'Stage event': liveEvent,
      'Stage event date and time': liveEventTimestamp,
      'Clock provenance': clock.provenance,
      'Clock label': clock.label,
      'Workbook seed value': clock.workbookValue || null,
      'Workbook export date and time': clock.workbookExportTimestamp || null,
      'Open pipeline': openPO(h.PurchaseOrderStatus, h.DocumentApprovalStatus)
    };
    poRows.push(row);
    if (row['Open pipeline'] && lifecycle.stage !== 'STAGE_NOT_EVIDENCED') {
      const payload = observationPayload(po, clock, now);
      const old = existingObs.get(observationKey(company, number));
      if (!old || old.ssg_payloadhash !== payload.ssg_payloadhash) pendingWrites.push([payload.ssg_documentkey, payload]);
    }
  }

  if (process.env.PRPO_STAGE_OBSERVATION_WRITE === '1') {
    const concurrency = 5; let cursor = 0;
    async function worker() { while (cursor < pendingWrites.length) { const index = cursor++; await upsertObservation(...pendingWrites[index]); } }
    await Promise.all(Array.from({ length: Math.min(concurrency, pendingWrites.length) }, worker));
  }

  const prLineBy = new Map();
  for (const line of prLines) {
    const number = doc(line.RequisitionNumber);
    if (!prLineBy.has(number)) prLineBy.set(number, []);
    prLineBy.get(number).push(line);
  }
  const prBiBy = new Map(prBi.map(x => [doc(x.PurchReqId), x]));
  const prApproval = Object.fromEntries(Object.entries(PR_STAGE_MAP.approval || {}).map(([key, value]) => [norm(key), value]));
  const procurementElements = new Set((PR_STAGE_MAP.procurementElements || []).map(norm));
  function prCapture(number) {
    const items = prItemByNumber.get(number) || [];
    const snapshot = prSnapshotBy.get(number);
    const snapshotIds = String(snapshot && snapshot.ssg_pendingstepnames || '').match(/[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}/ig) || [];
    const elements = new Set([
      ...snapshotIds.map(norm),
      ...items.map(item => norm(item.ssg_stepelementid)).filter(Boolean)
    ]);
    return {
      present: Boolean(snapshot || items.length),
      elements,
      timestamp: earliest([snapshot && snapshot.ssg_oldestpendingsince, ...items.map(x => x.ssg_assignedon || x.ssg_firstobservedon)]),
      approver: items.map(x => x.ssg_approvername || x.ssg_approveruserid).filter(Boolean).join(', ') ||
        (snapshot && (snapshot.ssg_pendingapprovernames || snapshot.ssg_pendinguserids)) || null
    };
  }
  const prRows = prHeaders.map(h => {
    const n = doc(h.RequisitionNumber); const company = norm(h.ProjectBuyingLegalEntityId);
    const allLines = prLineBy.get(n) || [];
    const lines = allLines.filter(line => !norm(line.LineStatus).includes('cancel'));
    const firstLine = lines[0] || allLines[0] || {}; const dim = parseDim(firstLine.DefaultLedgerDimensionDisplayValue);
    const current = prCapture(n);
    const approvalStages = new Set([...current.elements].map(element => prApproval[element]).filter(Boolean));
    let stage = null; let reason = 'NO_CURRENT_WORK_ITEM';
    if (approvalStages.size) {
      stage = approvalStages.size === 1 ? [...approvalStages][0] : 'Approval — unmapped element';
      reason = approvalStages.size === 1 ? 'APPROVAL_CAPTURE' : 'CONFLICTING_APPROVAL_STAGES';
    } else if ([...current.elements].some(element => procurementElements.has(element))) {
      if (!lines.length) { stage = 'Approval — unmapped element'; reason = 'ZERO_ACTIVE_LINES'; }
      else if (lines.every(line => num(line.PurchasePrice) > 0)) { stage = 'Priced — awaiting approval'; reason = 'ACTIVE_LINES_PRICED'; }
      else { stage = 'Sourcing'; reason = 'ACTIVE_LINES_NOT_FULLY_PRICED'; }
    } else if (current.present) {
      stage = 'Approval — unmapped element'; reason = 'UNMAPPED_ELEMENT';
    }
    const bi = prBiBy.get(n) || {};
    const status = norm(h.RequisitionStatus) === 'inreview' ? 'In review' : h.RequisitionStatus || null;
    return {
      'Purchase requisition': n, 'Legal entity': company,
      'Quotation reference': h.IFAHRQuotationReference || null, 'Name': h.RequisitionName || null,
      'Preparer': h.PreparerPersonnelNumber || null, 'Status': status,
      'Created date': bi.TransDate || h.DefaultRequestedDate || null,
      'Submitted date': bi.SubmittedDateTime || null,
      'Requisition purpose': h.RequisitionPurpose || null, 'Submission Status': null,
      'Accepted By/Assign To': null, 'Department': dim.department,
      'Location': dim.location, 'Contract': dim.contract, 'Request for quotation case': null,
      'Total amount': Math.round(lines.reduce((sum, line) => sum + num(line.LineAmount), 0) * 100) / 100,
      'Pending Approver/User': current.approver,
      'Step name': stage,
      'Step date and time': stage === 'Sourcing' || stage === 'Priced — awaiting approval' ? iso(bi.SysModifiedDateTime) : current.timestamp,
      'Stage reason code': reason
    };
  });

  const revisionMaterial = JSON.stringify({ prRows, poRows });
  const revision = crypto.createHash('sha256').update(revisionMaterial).digest('hex');
  const openRows = poRows.filter(row => row['Open pipeline']);
  const clocks = new Map();
  for (const row of openRows) clocks.set(row['Clock provenance'], (clocks.get(row['Clock provenance']) || 0) + 1);
  const captureReconciledUtc = [...snapshots].map(x => iso(x.ssg_lastreconciledon)).filter(Boolean).sort().at(-1) || null;
  return {
    revision, generatedAt: now, sourceState: 'LIVE', workbookExportTimestamp: WORKBOOK_EXPORT_UTC,
    freshness: {
      datasetGeneratedUtc: now,
      fAndOReadUtc: now,
      approvalCaptureReconciledUtc: captureReconciledUtc,
      effectiveDataTimeUtc: earliest([now, captureReconciledUtc])
    },
    pendingObservationWrites: pendingWrites.length,
    pr: { count: prRows.length, rows: prRows },
    po: {
      count: poRows.length, openCount: openRows.length,
      stageNotEvidenced: openRows.filter(row => row['Live stage'] === 'STAGE_NOT_EVIDENCED').length,
      clockCounts: {
        liveDated: clocks.get('LIVE_EVENT_DATE') || 0,
        seededFromFinalWorkbook: clocks.get('SEEDED_FROM_FINAL_WORKBOOK') || 0,
        pendingSinceFirstObserved: clocks.get('PENDING_SINCE_FIRST_OBSERVED') || 0,
        notRecorded: clocks.get('NOT_RECORDED') || 0
      },
      rows: poRows
    }
  };
}

async function refreshWithFallback(loader, previous) {
  try { return await loader(); }
  catch (error) {
    if (!previous) throw error;
    return { ...previous, cached: true, sourceState: 'STALE', refreshError: error.message };
  }
}

async function getDataset(options = {}) {
  const now = Date.now();
  if (!options.force && cache && now - cache.at < CACHE_MS) return { ...cache.data, cached: true };
  if (refreshPromise) return refreshPromise;
  refreshPromise = refreshWithFallback(buildDataset, cache && cache.data).then(data => {
    if (data.sourceState !== 'STALE') cache = { at: Date.now(), data };
    return data;
  }).finally(() => { refreshPromise = null; });
  return refreshPromise;
}

function clearCache() { cache = null; refreshPromise = null; }

module.exports = {
  WORKBOOK_EXPORT_UTC,
  getDataset,
  buildDataset,
  clearCache,
  deriveClock,
  lifecycleStage,
  observationKey,
  observationPayload,
  readObservation,
  upsertObservation,
  openPO,
  refreshWithFallback
};
