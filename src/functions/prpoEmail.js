/*
 * PR / PO PIPELINE — Waqas-only production test channel, scheduled 10:00 AM Dubai.
 *
 * TIMER 0 0 6 * * 1-5 (06:00 UTC = 10:00 AM Dubai). SEND MODEL:
 *   1. PERSONAL action emails -> one per F&O Pending Approver/User with pending items.
 *      Ops-confirm split "Pending Internal" vs "Pending Client" ('Unit prices updated in PR lines' = with
 *      client): client items in cards + attached Excel ONLY, not the body table.
 *   2. Suppliers & Open Orders team list:
 *      Sent-to-Supplier POs + Confirmed POs with 'Open order' status (the stale-approver ones).
 *   3. Pending Invoicing team list. The old finance/ops_hm/ops_all division emails remain preview-only.
 *
 * Optional: division-scoped emails (produce four separate emails with scoped analysis + Excel attachments).
 * These use the dashboard live-pipeline logic. Any actual Graph message is addressed in code to Waqas only,
 * has no Cc/Bcc, and identifies the intended person/team in a [FOR ...] subject prefix. Environment recipient
 * settings and PRPO_PERSONAL_TEST cannot override this guard.
 *
 * HTTP test endpoints (authLevel: function, add &code=<your default host key>):
 *   ?division=procurement|finance|ops_hm|ops_all & format=html|debug=1|send=1   -> that division (ops_* preview-only)
 *   ?person=<username> & format=html | send=1                                   -> one personal email
 *   ?personal=1 [& send=1]                                                      -> list (or send) ALL personal emails
 *   ?send=1                                                                     -> send everything (divisions + personal)
 *   (no params)                                                                 -> JSON summary (no send)
 *
 * Env: TENANT_ID, CLIENT_ID, CLIENT_SECRET, MAIL_FROM. Optional PRPO_MAIL_FROM and PRPO_DASH_URL.
 */
const { app } = require('@azure/functions');
const XLSX = require('xlsx');
const ExcelJS = require('exceljs');
const WORK_CLASS_RULE = require('../../work-class-rule.json');
const INACTIVE_USERNAMES = new Set(require('../../inactive-usernames.json').inactiveUsernames.map(v=>String(v).trim().toLowerCase().replace(/\s+/g,' ')));

const DASH   = process.env.PRPO_DASH_URL || 'https://strive-services-group.github.io/PR-PO-Pipeline-Dashboard/';
const DATASET_URL = process.env.PRPO_DATASET_URL || 'https://ssg-prpo-proxy-h4cvfegaduftedhz.uaenorth-01.azurewebsites.net/api/dataset';
const WAQAS_ONLY_RECIPIENT = 'w.amjad@striveservicesgroup.com';
const FONT = 'Aptos,Segoe UI,Arial,sans-serif', NAVY = '#14315E', RED = '#dc2626', TEAL = '#0f766e', W = 1000;

const PR_MAP = {"Handyman Services_Manager":"Dep Managers","Building Services_Asst. Facility Managers 1":"Dep Managers","PurchReqReviewTask":"PR In Review","Procurement sends inquiry/RFQ to suppliers":"RFQ to suppliers","Quotation received and logged/attached":"Qt received & Logged","Quotation shared to Operations for confirmation":"Qt Shared to Op","Operations confirms material/scope":"OP confirms material","Unit prices updated in PR lines":"Unit Price Updated","Building Services_Asst. Facility Managers 2":"Dep Managers","Building Services_Facilities Manager":"Dep Managers","PAC Services_Manager":"Dep Managers","Concierge Services_Manager":"Dep Managers","Security Services_Manager":"Dep Managers","Home Services_Operations Manager":"Dep Managers","Landscaping_Manager":"Dep Managers","Finance & Accounts_Accounting Manager":"Finance","Facilities Management_Director":"Director","Commercial_Director":"Director","Executive Management_CEO":"CEO"};
const PO_MAP = {"Advance payment request submitted (if applicable)":"Procurement","Procurement Manager":"Procurement","Accounting Manager":"Finance","Finance and Accounts Director":"Director","CEO":"CEO","LPO sent/shared with supplier":"Sent to Supplier"};
const COLOR = {'Procurement':'#3b82f6','Sourcing':'#0ea5e9','Priced — awaiting approval':'#84cc16','Operations to Confirm':'#14b8a6','Dep Managers':'#8b5cf6','Finance':'#22c55e','Director':'#ec4899','CEO':'#f59e0b','Sent to Supplier':'#a855f7','Pending Invoicing':'#f97316','Re-Assigned/Rejected':'#dc2626','Pending Internal':'#14b8a6','Pending Client':'#6366f1','Confirmed Open Order':'#0891b2'};
const GRAD = {'Procurement':'#eff5ff','Sourcing':'#effbff','Priced — awaiting approval':'#f5fae8','Operations to Confirm':'#ebfbf7','Dep Managers':'#f4f1fe','Finance':'#eefbf3','Director':'#fdeff7','CEO':'#fff9ec','Sent to Supplier':'#f9f2ff','Pending Invoicing':'#fff4e8','Re-Assigned/Rejected':'#fef2f2','Pending Internal':'#ebfbf7','Pending Client':'#eef2ff','Confirmed Open Order':'#ecfeff'};
const TYCOL = {'PR':'#2563eb','CPR':'#7c3aed','PO':'#0891b2'};

/* ---- helpers ---- */
function esc(s){ return String(s==null?'':s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;'); }
function money(v){ const n=Number(v); return (isFinite(n)?n:0).toLocaleString('en-US',{maximumFractionDigits:0}); }
function amt(r){ const v=Number(r['Total amount']); return isFinite(v)?v:0; }
function hasRecordedValue(r){ const raw=r&&r['Total amount']; return raw!==null&&raw!==''&&raw!==undefined&&Number.isFinite(Number(raw))&&Number(raw)>0; }
function xd(v){ if(v instanceof Date) return isNaN(v)?null:v; if(typeof v==='number'&&isFinite(v)){ const o=XLSX.SSF.parse_date_code(v); if(!o||!o.y) return null; return new Date(Date.UTC(o.y,o.m-1,o.d,o.H||0,o.M||0,Math.floor(o.S||0))); } if(typeof v==='string'&&v){ const d=new Date(v); return isNaN(d)?null:d; } return null; }
function ageDays(v){ const d=xd(v); return d? Math.max(0,Math.floor((Date.now()-d.getTime())/86400000)) : null; }
function ymdStr(v){ const d=xd(v); if(!d) return ''; const p=n=>String(n).padStart(2,'0'); return d.getUTCFullYear()+'-'+p(d.getUTCMonth()+1)+'-'+p(d.getUTCDate()); }
function avg(l){ return l.length? l.reduce((a,b)=>a+b,0)/l.length : 0; }
function parseXlsx(buf){
  const wb=XLSX.read(buf,{type:'buffer'}), ws=wb.Sheets[wb.SheetNames[0]];
  const aoa=XLSX.utils.sheet_to_json(ws,{header:1,defval:null,raw:true});
  if(!aoa.length)return [];
  const hdr=aoa[0].map(x=>String(x)), rows=[];
  for(let i=1;i<aoa.length;i++){ const o={}; for(let j=0;j<hdr.length;j++)o[hdr[j]]=aoa[i][j]; rows.push(o); }
  const metaSheet=wb.Sheets['Routing metadata'];
  const metadata=metaSheet?XLSX.utils.sheet_to_json(metaSheet,{defval:null,raw:true}):[];
  Object.defineProperty(rows,'routingMetadata',{value:metadata,enumerable:false});
  return rows;
}
function prLive(r){ const st=_norm(r['Status']); return ['draft','in review','approved'].includes(st); }
function poBucket(r){ const live=String(r['Live stage']||r['Step name']||''); const appr=String(r['Approval status']||''); const pos=String(r['Purchase order status']||''); let b=PO_MAP[r['Step name']]||null; if(['Procurement','Finance','Director','CEO'].includes(live)) b=live; if(live==='Approval — unmapped element'||live==='Not yet sent') b='Procurement'; if(live==='Sent to supplier') b='Sent to Supplier'; if(live==='Receipt posted') b='Pending Invoicing'; if(live==='Invoiced') return null; if(appr==='Rejected'||pos==='Canceled'||pos==='Cancelled'||pos==='Closed'||pos==='Invoiced'||!b) return null; return b; }
function sameMoment(a,b){ const da=xd(a),db=xd(b); return !!(da&&db&&Math.abs(da.getTime()-db.getTime())<1000); }
function prClock(r){
  const created=r['Created date'], step=r['Step date and time'];
  if(!step||sameMoment(step,created)){ const age=ageDays(created); return {age,basis:'raised',label:age==null?'raised date not recorded':'raised '+age+' days ago'}; }
  const age=ageDays(step), date=ymdStr(step);
  return {age,basis:'current step',label:age==null?'current-step date not recorded':'with you since '+date+' ('+age+' days)'};
}
function prAge(r){ return prClock(r).age; }
function poAge(r){ const sd=r['Step date and time']; return ageDays(sd!=null? sd : (r['Created date and time']!=null? r['Created date and time'] : r['Requested receipt date'])); }
function clockNote(r){ const p=String(r&&r['Clock provenance']||''); if(p==='SEEDED_FROM_FINAL_WORKBOOK') return ' — since (from last export)'; if(p==='NOT_RECORDED') return ' — since — not recorded'; return ''; }

/* ---- USER-based routing: pending-with (by status) -> user's department -> division ---- */
const USER_DEPT={"Abdul Basit Raza":"Building Services","Abdul.basit":"IT","Abdul.Muqeet":"Security Services","Admin":"IT","admin.hk":"Housekeeping Services","Adnan.Ullah":"Procurement","Ahamed Noorullah Mohamed":"Accomodation Services","Ahmed.Odeh":"Building Services","Aparna.Pauly":"Procurement","arman.b":"Accounts & Tax","ayman.g":"Accounts & Tax","Ayman.ismail":"Accounts & Tax","Buying Agent Concierge":"Concierge Services","D365CRM ADMIN":"IT","D365CRMADMIN":"IT","Dinesh Laxman Laxman":"Building Services","dinesh.laxman":"Building Services","Gokul Krishna Pillai":"Contracted Cleaning Services","Gokul.Krishna":"Contracted Cleaning Services","IT DEPARTMENT":"IT","Joe Orlain Jamisola":"Concierge Services","Judhin.prabhakar":"Contracted Cleaning Services","Layusha.cleatus":"Procurement","Mohamed.Ashraf":"Procurement","Mohammad.w":"Building Services","Muhammad Shehzad Ahmeduddin":"IT","muhammad.mustajab":"Accounts & Tax","Nathan.Buys":"Building Services","Patrick.Smith":"Accounts & Tax","Pramod Chandrasenan Chandrasenan":"Security Services","pramod.c":"Security Services","Qasim Jahangir":"QHSE","Roderick Red Palma":"Procurement","roderick.red":"Procurement","Shaik.baba":"Housekeeping Services","Shakir Ameer Bakhsh":"FitOut Services","Shijil Choyaprath Chandran":"Home Maintenance Services","shijil.c":"Home Maintenance Services","Sirinikhil":"Housekeeping Services","teena.k":"Concierge Services","Ubaid":"IT","Zaheer Ahmed Ameer":"Accomodation Services","Zaheer.Ahmed":"Accomodation Services"};
const _norm=s=>String(s==null?'':s).trim().toLowerCase().replace(/\s+/g,' ');
const _UN={}; for(const k in USER_DEPT){ _UN[_norm(k)]=USER_DEPT[k]; }
function deptForUser(u){ return _UN[_norm(u)]||''; }
const DEPT_DIV={'Procurement':'procurement','Accounts & Tax':'finance','Home Maintenance Services':'ops_hm','FitOut Services':'ops_hm'};
function divForDept(d){ return DEPT_DIV[d]||'ops_all'; }
// The F&O step name is unreliable for the finance chain, so the Finance/Director/CEO bucket is reconstructed from
// WHO holds the item. roleOf: Accounts & Tax approver -> 'Finance' (ayman.g -> 'Director', Patrick.Smith -> 'CEO'),
// procurement user -> 'Procurement', anyone else (operations) -> '' (falls through to the operations split).
function roleOf(u){ const d=deptForUser(u), ul=_norm(u); if(d==='Accounts & Tax'){ if(ul==='ayman.g') return 'Director'; if(ul==='patrick.smith') return 'CEO'; return 'Finance'; } if(d==='Procurement') return 'Procurement'; return ''; }
function opsDivFor(reqdept){ return (reqdept==='Home Maintenance Services'||reqdept==='FitOut Services')?'ops_hm':'ops_all'; }
function prPendingWith(r){ return String(r['Pending Approver/User']||'').trim(); }
const NO_NAMED_OWNER='No named owner';
function isNoNamedOwner(value){ const v=_norm(value); return !v||v==='(unassigned)'||v==='not recorded'||v.startsWith('no named owner')||v.startsWith('employee number '); }
function fnoOwnerNames(value){
  const raw=String(value==null?'':value).trim();
  if(!raw)return [];
  if(raw.includes(','))throw new Error('F&O export data fault: more than one owner was supplied in one record');
  if(/^\d+$/.test(raw))return [NO_NAMED_OWNER+' — F&O Pending Approver/User is employee number '+raw];
  return [raw];
}
function ageBand(days){ if(days==null)return 'Age not recorded'; if(days<=7)return '0–7'; if(days<=30)return '8–30'; if(days<=60)return '31–60'; if(days<=90)return '61–90'; return 'Over 90'; }
function ageBandLabel(it){ return (it.clockBasis==='raised'?'Raised':'Current step')+' · '+it.ageBand; }
function legacyClassCode(step){
  if(step==='Unit prices updated in PR lines'||step==='Quotation shared to Operations for confirmation')return 'ACTIVE_LINES_PRICED';
  if(['Procurement sends inquiry/RFQ to suppliers','Quotation received and logged/attached','Operations confirms material/scope'].includes(step))return 'ACTIVE_LINES_NOT_FULLY_PRICED';
  return '';
}
function workClassFor(r){
  const supplied=String(r['Stage reason code']||'').trim();
  const code=supplied||legacyClassCode(String(r['Step name']||'').trim())||'NOT_REPORTED';
  const rule=WORK_CLASS_RULE.classes[code];
  if(rule)return {code,label:rule.label,action:rule.action,order:rule.order,rule,fromWorkbook:!!supplied};
  return {code,label:code==='NOT_REPORTED'?'Work class not reported by workbook':WORK_CLASS_RULE.unknownLabel+' — '+code,action:WORK_CLASS_RULE.unknownAction,order:999,rule:{headerBucket:'Step not reported by F&O'},fromWorkbook:!!supplied};
}
function routingMetadataMap(rows){
  const out=new Map();
  for(const row of (rows&&rows.routingMetadata)||[]){
    const ref=String(row['Purchase requisition']||'').trim().toUpperCase(), holder=String(row['Source holder']||'').trim();
    if(!ref||!holder)continue;
    const list=out.get(ref)||[], key=_norm(canonOwner(holder));
    if(!list.some(name=>_norm(canonOwner(name))===key))list.push(holder);
    out.set(ref,list);
  }
  return out;
}
function sharedBuyerState(sourceOwners,owner,workCode){
  return {sourceShared:false,otherLiveBuyers:[],label:''};
}

function prQueueForStep(step){
  if(['Quotation shared to Operations for confirmation','Operations confirms material/scope','Unit prices updated in PR lines'].includes(step))return 'Operations to Confirm';
  if(step==='Finance & Accounts_Accounting Manager')return 'Finance';
  if(['Facilities Management_Director','Commercial_Director'].includes(step))return 'Director';
  if(step==='Executive Management_CEO')return 'CEO';
  if(/_(?:Manager|Asst\. Facility Managers|Facilities Manager|Operations Manager)/.test(step))return 'Dep Managers';
  if(['PurchReqReviewTask','Procurement sends inquiry/RFQ to suppliers','Quotation received and logged/attached'].includes(step))return 'Procurement';
  return 'Step not reported by F&O';
}
function pendingSide(it){
  if(it.doc!=='PR'||it.stage!=='Operations to Confirm')return '';
  return String(it.raw['Step name']||'').trim()==='Unit prices updated in PR lines'?'Pending Client':'Pending Internal';
}
// Routing: PO -> functional home (Sent-to-Supplier/Procurement->procurement, else finance).
// PR "Operations to Confirm" bucket (Unit-price-updated / Quotation-shared-to-Ops steps) -> by the REQUISITION's own
//   department (row Department col), so it lands in the Operations sheets regardless of which procurement user holds it.
// All other PR -> by the pending-with USER's assigned department (falls back to row dept if the user is unmapped).
function itemDivision(it){ return it.div; }  // stage(bucket) + div are reconstructed in buildItems (see below)
function _unused_itemDivision(it){
  if(it.doc==='PO'){
    if(it.stage==='Sent to Supplier') return 'procurement';   // vendor-side -> by bucket
    if(it.stage==='Pending Invoicing') return 'finance';      // vendor-side -> by bucket
    if(it.ppend) return (deptForUser(it.owner)==='Accounts & Tax')?'finance':'procurement'; // pending approval -> by the person holding it
    return (it.stage==='Procurement')?'procurement':'finance'; // settled -> by bucket
  }
  // PR -> by workflow step/bucket (NOT by who holds it)
  if(it.stage==='Procurement') return 'procurement';
  if(it.stage==='Finance'||it.stage==='Director'||it.stage==='CEO') return 'finance';
  // Operations to Confirm + Dep Managers -> Operations, split by REQUISITION department (unmapped dept -> All-Depts)
  return (it.dept==='Home Maintenance Services'||it.dept==='FitOut Services')?'ops_hm':'ops_all';
}
const STAGE_ORDER=[['PR','Re-Assigned/Rejected'],['PR','Procurement'],['PR','Operations to Confirm'],['PR','Step not reported by F&O'],['PR','Dep Managers'],['PR','Finance'],['PR','Director'],['PR','CEO'],['PO','Procurement'],['PO','Finance'],['PO','Director'],['PO','CEO'],['PO','Confirmed Open Order'],['PO','Sent to Supplier'],['PO','Pending Invoicing']];

function buildItems(prRows, poRows){
  const items=[];
  for(const r of prRows){ if(!prLive(r)) continue; const work=workClassFor(r); const rowdept=String(r['Department']||'').trim(); const sourceOwners=fnoOwnerNames(prPendingWith(r)); const owner=sourceOwners[0]||NO_NAMED_OWNER+' — Pending Approver/User not recorded in F&O'; const noNamedOwner=isNoNamedOwner(owner);
    const stage=prQueueForStep(String(r['Step name']||'').trim());
    const div=stage==='Procurement'?'procurement':(['Finance','Director','CEO'].includes(stage)?'finance':opsDivFor(rowdept));
    const clock=prClock(r), shared=sharedBuyerState(sourceOwners,owner,work.code);
    items.push({ref:r['Purchase requisition'],doc:'PR',typ:String(r['Purchase requisition']||'').startsWith('CPR')?'CPR':'PR',stage:stage,div:div,age:clock.age,ageBand:ageBand(clock.age),clockBasis:clock.basis,clockLabel:clock.label,owner:owner,dept:rowdept,value:amt(r),hasRecordedValue:hasRecordedValue(r),vendor:'',ppend:true,noNamedOwner,workClassCode:work.code,workClass:work.label,workAction:work.action,classOrder:work.order,sourceShared:shared.sourceShared,otherLiveBuyers:shared.otherLiveBuyers,sharedLabel:shared.label,raw:r}); }
  // PO: Pending Approver/User is the only person-of-record; vendor lifecycle stages are not personal queues.
  // Vendor stages (Sent-to-Supplier/Pending-Invoicing) route by bucket; every other PO's bucket+division is reconstructed from the holder's role.
  for(const r of poRows){ const bk=poBucket(r); if(!bk) continue; const ven=String(r['Vendor name']||'-').trim(); const poStat=String(r['Approval status']||''); const poPend=fnoOwnerNames(r['Pending Approver/User']); const isVenBk=(bk==='Sent to Supplier'||bk==='Pending Invoicing'); const owners=isVenBk?[ven]:(poPend.length?poPend:[NO_NAMED_OWNER+' — Pending Approver/User not recorded in F&O']); const ppend=!isVenBk;
   for(const own of owners){
    let stage,div;
    if(bk==='Sent to Supplier'){ stage=bk; div='procurement'; }
    else if(bk==='Pending Invoicing'){ stage=bk; div='finance'; }
    else { const rl=roleOf(own); if(rl==='Finance'||rl==='Director'||rl==='CEO'){ stage=rl; div='finance'; } else if(rl==='Procurement'){ stage='Procurement'; div='procurement'; } else { stage='Procurement'; div='procurement'; } }
    const age=poAge(r);
    items.push({ref:r['Purchase order'],doc:'PO',typ:'PO',stage:stage,div:div,age,ageBand:ageBand(age),clockBasis:'current step',clockLabel:'at current step for '+age+' days',owner:own,dept:String(r['Department']||'').trim(),value:amt(r),hasRecordedValue:hasRecordedValue(r),vendor:ven,ppend:ppend,noNamedOwner:isNoNamedOwner(own),workClassCode:String(r['Stage reason code']||''),workClass:'Purchase order action',workAction:'Complete the current purchase-order step.',classOrder:100,sourceShared:false,otherLiveBuyers:[],sharedLabel:'',raw:r}); }}
  return items;
}

/* ---- render helpers ---- */
function chip(doc){ const c=doc==='PR'?'#2563eb':'#0891b2'; return '<span style="display:inline-block;background:'+c+';color:#fff;font-family:'+FONT+';font-size:9px;font-weight:800;padding:1px 5px;margin-right:6px;vertical-align:middle;">'+doc+'</span>'; }
function card2(doc,bk,x,tr){ const a=x.c?(x.sum/x.c):0, price=x.priced?(x.priced+' priced &#183; AED '+money(x.amt)):'not yet priced'; return '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="width:100%;background:'+GRAD[bk]+';border:1px solid #e8ecf2;border-top:3px solid '+COLOR[bk]+';border-radius:10px;box-shadow:0 1px 3px rgba(16,24,40,0.08);"><tr><td style="padding:11px 13px;"><table role="presentation" cellpadding="0" cellspacing="0"><tr><td valign="middle" bgcolor="'+(doc==='PR'?'#2563eb':'#0891b2')+'" style="background:'+(doc==='PR'?'#2563eb':'#0891b2')+';padding:2px 6px;font-family:'+FONT+';font-size:9px;font-weight:800;color:#ffffff;border-radius:3px;">'+doc+'</td><td width="6" style="width:6px;">&#160;</td><td valign="middle" style="font-family:'+FONT+';font-size:9.5px;font-weight:800;color:#5b6b7f;text-transform:uppercase;">'+esc(bk)+'</td></tr></table><div style="margin:6px 0 2px;white-space:nowrap;"><span style="font-family:'+FONT+';font-size:24px;font-weight:800;color:'+COLOR[bk]+';">'+x.n+'</span><span style="font-family:'+FONT+';font-size:12px;font-weight:800;color:'+RED+';"> ('+a.toFixed(1)+'d)</span>'+deltaBadge(tr,x.n)+'</div><div style="font-family:'+FONT+';font-size:11px;font-weight:700;color:'+TEAL+';">'+price+(x.unpriced?' &#183; '+x.unpriced+' not yet priced':'')+'</div>'+histLine(tr)+'</td></tr></table>'; }
// Compact cards in ONE row: PR cards | vertical divider | PO cards. Width auto-shrinks to fit W (capped at 200px).
function cardrow2(pairs,agg,trFn){
  const live=pairs.filter(([d,bk])=>agg[d+'|'+bk]&&agg[d+'|'+bk].n>0); if(!live.length) return '';
  const G=12, pr=live.filter(p=>p[0]==='PR'), po=live.filter(p=>p[0]==='PO');
  const both=(pr.length&&po.length), n=live.length;
  const divW=both?(2+2*G):0, gutW=G*(n-(both?2:1));
  const cw=Math.min(200,Math.floor((W-divW-gutW)/n));
  const cell=p=>'<td width="'+cw+'" valign="top" style="width:'+cw+'px;">'+card2(p[0],p[1],agg[p[0]+'|'+p[1]],trFn?trFn(p[0],p[1]):null)+'</td>';
  // Outlook Word engine ignores font-size:0 — use tiny-but-nonzero font/line-height so spacers keep their declared width.
  const gutter='<td width="'+G+'" style="width:'+G+'px;font-size:6px;line-height:6px;mso-line-height-rule:exactly;">&#160;</td>';
  const seg=arr=>arr.map((p,i)=>cell(p)+(i<arr.length-1?gutter:'')).join('');
  const divider=gutter+'<td width="2" bgcolor="#b9c5d6" style="width:2px;background:#b9c5d6;font-size:2px;line-height:2px;mso-line-height-rule:exactly;">&#160;</td>'+gutter;
  return '<table role="presentation" cellpadding="0" cellspacing="0" style="border-collapse:separate;"><tr>'+seg(pr)+(both?divider:'')+seg(po)+'</tr></table><div style="height:'+G+'px;font-size:'+G+'px;line-height:'+G+'px;">&#160;</div>';
}
function badge(L,col){ return '<span style="display:inline-block;width:24px;height:24px;line-height:24px;text-align:center;background:'+col+';color:#fff;font-family:'+FONT+';font-weight:800;font-size:12px;margin-right:9px;vertical-align:middle;">'+L+'</span>'; }
function finding(L,col,title,right,narr,tbl){ const head='<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="margin:0 0 7px;"><tr><td width="24" height="24" align="center" valign="middle" bgcolor="'+col+'" style="width:24px;height:24px;background:'+col+';font-family:'+FONT+';font-weight:800;font-size:12px;color:#ffffff;">'+L+'</td><td width="9" style="width:9px;">&#160;</td><td valign="middle" style="font-family:'+FONT+';font-weight:800;font-size:14px;color:'+NAVY+';">'+title+'</td>'+(right?'<td align="right" valign="middle" style="font-family:'+FONT+';font-size:11px;font-weight:700;color:#94a3b8;white-space:nowrap;padding-left:10px;">'+right+'</td>':'')+'</tr></table>'; const body='<div style="font-family:'+FONT+';font-size:12.5px;color:#3f4b5b;line-height:1.5;margin:0 0 '+(tbl?'8px':'2px')+';">'+narr+'</div>'+(tbl||''); return '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="margin:0 0 14px;"><tr><td style="border:1px solid #e6ebf1;border-left:4px solid '+col+';background:#fff;padding:14px 16px;border-radius:10px;box-shadow:0 1px 3px rgba(16,24,40,0.07);">'+head+body+'</td></tr></table>'; }
function b(s){ return '<b style="color:'+NAVY+';">'+s+'</b>'; }
function nm(u){ return '<span style="font-weight:700;color:'+NAVY+';">'+esc(u)+'</span>'; }
function sv(t,col){ return '<span style="color:'+col+';font-weight:700;">'+t+'</span>'; }
function agec(a){ a=Math.round(a||0); const col=a>30?'#b91c1c':(a>7?'#c2410c':'#16794a'); return '<span style="font-weight:700;color:'+col+';">'+a+'d</span>'; }
function pricedStats(items){
  const priced=(items||[]).filter(it=>it.hasRecordedValue), unpriced=(items||[]).length-priced.length;
  return {priced:priced.length,unpriced,value:priced.reduce((sum,it)=>sum+it.value,0)};
}
function priceLine(items){
  const s=pricedStats(items);
  if(!s.priced)return b('not yet priced');
  return b(s.priced+' priced')+', worth '+b('AED '+money(s.value))+(s.unpriced?' &#183; '+b(s.unpriced+' not yet priced'):'');
}
function itemPrice(it){ return it.hasRecordedValue?'AED '+money(it.value):'Not yet priced'; }
function agePhrase(it){ return it&&it.clockLabel?it.clockLabel:('age '+Math.round(it&&it.age||0)+' days'); }
function otable(cols,rows){ const th='color:#fff;font-family:'+FONT+';font-weight:800;font-size:10.5px;padding:8px 11px;white-space:nowrap;background:'+NAVY+';'; const head='<tr>'+cols.map(c=>'<th style="'+th+(c[2]==='r'?'text-align:right;':(c[2]==='c'?'text-align:center;':'text-align:left;'))+'width:'+c[1]+'px;">'+c[0]+'</th>').join('')+'</tr>'; let body=''; rows.forEach((cells,i)=>{ const bg=i%2===0?'#ffffff':'#f7f9fc'; const last=(i===rows.length-1); const td='padding:8px 11px;'+(last?'':'border-bottom:1px solid #eef1f6;')+'font-family:'+FONT+';font-size:11.5px;background:'+bg+';white-space:nowrap;'; body+='<tr>'+cells.map((cell,j)=>'<td style="'+td+(j>0?'border-left:1px solid #f1f4f8;':'')+(cols[j][2]==='r'?'text-align:right;':(cols[j][2]==='c'?'text-align:center;':''))+'">'+cell+'</td>').join('')+'</tr>'; }); return '<table width="100%" cellpadding="0" cellspacing="0" border="0" style="width:100%;border-collapse:separate;border-spacing:0;background:#fff;border:1px solid #e2e8f0;border-radius:8px;overflow:hidden;">'+head+body+'</table>'; }
function SH(t,s){ return '<div style="font-family:'+FONT+';font-weight:800;font-size:15px;color:'+NAVY+';margin:0 0 2px;">'+t+'</div>'+(s?'<div style="font-family:'+FONT+';font-size:11.5px;color:#7688a0;margin:0 0 9px;">'+s+'</div>':''); }

/* ---- finding builders ---- */
function grpBy(arr,key){ const g={}; for(const it of arr){ const k=key(it); (g[k]=g[k]||[]).push(it); } return g; }
function dchip(doc){ const c=doc==='PO'?'#0891b2':'#2563eb'; return ' <b style="font-family:'+FONT+';font-size:11px;color:'+c+';">('+doc+')</b>'; }
function f_owners(its,L,col,title,label){
  const persons=its.filter(it=>it.ppend!==false&&it.stage!=='Sent to Supplier'&&it.stage!=='Pending Invoicing');
  const g={}; for(const it of persons){ const key=String(it.owner==null?'':it.owner).trim().toLowerCase()+'|'+it.doc; const e=g[key]||(g[key]={n:0,ages:[],val:0,priced:0,br:0,bk:{},disp:{},doc:it.doc}); e.n++; if(it.hasRecordedValue){e.val+=it.value;e.priced++;} e.bk[it.stage]=(e.bk[it.stage]||0)+1; e.disp[it.owner]=(e.disp[it.owner]||0)+1; if(it.age!=null){e.ages.push(it.age); if(it.age>7)e.br++;} }
  const rows=[]; for(const [u,e] of Object.entries(g).sort((a,b2)=>b2[1].n-a[1].n)){ const okey=u.slice(0,u.lastIndexOf('|')); if(okey==='(unassigned)'||okey==='')continue; const disp=Object.entries(e.disp).sort((a,b2)=>b2[1]-a[1])[0][0]; const stg=Object.entries(e.bk).sort((a,b2)=>b2[1]-a[1])[0][0]; rows.push([nm(disp)+dchip(e.doc),sv(stg,COLOR[stg]),String(e.n),agec(avg(e.ages)),agec(e.ages.length?Math.max(...e.ages):0),sv(String(e.br),'#b91c1c'),e.priced?(e.priced+' priced &#183; AED '+money(e.val)):'Not yet priced']); if(rows.length>=6)break; }
  if(!rows.length) return '';
  const named=persons.filter(it=>{const o=String(it.owner==null?'':it.owner).trim().toLowerCase(); return o!==''&&o!=='(unassigned)';});
  const n=named.length, br=named.filter(it=>(it.age||0)>7).length;
  return finding(L,col,title,n+' items',b(n)+' items are pending with a person; '+b(br)+' are past the 7-day SLA. Top owners &#8212; chase these queues first.',
    otable([[label||'Pending with',172,'l'],['Stage',140,'l'],['Items',52,'c'],['Avg',52,'c'],['Oldest',58,'c'],['Br&gt;7',50,'c'],['Recorded value',180,'r']],rows));
}
function f_vendor(its,bucket,L,col,title,action){
  const vs=its.filter(it=>it.stage===bucket); if(!vs.length) return '';
  const g={}; for(const it of vs){ const e=g[it.vendor||'-']||(g[it.vendor||'-']={n:0,ages:[],val:0}); e.n++; e.val+=it.value; if(it.age!=null)e.ages.push(it.age); }
  const rows=[]; for(const [v,e] of Object.entries(g).sort((a,b2)=>b2[1].val-a[1].val)){ rows.push([esc(v.slice(0,30)),String(e.n),agec(avg(e.ages)),agec(e.ages.length?Math.max(...e.ages):0),'AED '+money(e.val)]); if(rows.length>=6)break; }
  const tot=vs.reduce((a,it)=>a+it.value,0), av=avg(vs.filter(it=>it.age!=null).map(it=>it.age));
  return finding(L,col,title,'AED '+money(tot),b(vs.length)+' POs at &#8220;'+bucket+'&#8221; ('+b('AED '+money(tot))+', avg '+b(Math.round(av)+'d')+'). '+action,
    otable([['Vendor',210,'l'],['POs',52,'c'],['Avg',56,'c'],['Oldest',60,'c'],['Value',130,'r']],rows));
}
function f_oldest(its,L,col){
  const old=its.slice().sort((a,b2)=>(b2.age||0)-(a.age||0)).slice(0,6); if(!old.length) return '';
  const rows=old.map(it=>['<span style="font-weight:700;color:'+NAVY+';">'+esc(it.ref)+'</span>',sv(it.typ,TYCOL[it.typ]),esc(agePhrase(it)),sv(it.stage,COLOR[it.stage]),esc(String(it.owner||'-').slice(0,20)),esc(String(it.dept||'-').slice(0,20))]);
  return finding(L,col,'Escalate &#8212; oldest &amp; most overdue','top '+old.length,'Oldest item was '+b(agePhrase(old[0]))+' &#8212; '+esc(old[0].ref)+' with '+nm(old[0].owner)+'. Escalate the top rows.',
    otable([['Ref',118,'l'],['Type',46,'c'],['Age / source',145,'l'],['Stage',150,'l'],['Owner/Vendor',150,'l'],['Dept',150,'l']],rows));
}
function f_value(its,L,col){
  const tv=its.filter(it=>it.hasRecordedValue).sort((a,b2)=>b2.value-a.value).slice(0,6); if(!tv.length) return '';
  const s=tv.reduce((a,it)=>a+it.value,0);
  const rows=tv.map(it=>['<span style="font-weight:700;color:'+NAVY+';">'+esc(it.ref)+'</span>',sv(it.typ,TYCOL[it.typ]),'<b>AED '+money(it.value)+'</b>',agec(it.age),sv(it.stage,COLOR[it.stage]),esc(String(it.owner||'-').slice(0,20))]);
  return finding(L,col,'High-value at risk','AED '+money(s),'The biggest exposure is '+b('AED '+money(tv[0].value))+' ('+esc(tv[0].ref)+'). Top 6 total '+b('AED '+money(s))+' &#8212; clear these for the biggest cash impact.',
    otable([['Ref',118,'l'],['Type',46,'c'],['Value',120,'l'],['Age',54,'c'],['Stage',150,'l'],['Owner/Vendor',160,'l']],rows));
}
function f_sla(its,L,col){
  const g={}; for(const it of its){ const e=g[it.stage]||(g[it.stage]={n:0,br:0,ages:[]}); e.n++; if(it.age!=null){e.ages.push(it.age); if(it.age>7)e.br++;} }
  const order=['Procurement','Operations to Confirm','Dep Managers','Finance','Director','CEO','Confirmed Open Order','Sent to Supplier','Pending Invoicing'].filter(s=>g[s]);
  const rows=order.map(s=>{ const e=g[s], pct=e.n?Math.round(100*e.br/e.n):0, pcol=pct>=70?'#b91c1c':(pct>=40?'#c2410c':'#16794a'); return [sv(s,COLOR[s]),String(e.n),sv(String(e.br),'#b91c1c'),sv(pct+'%',pcol),agec(avg(e.ages))]; });
  const tn=order.reduce((a,s)=>a+g[s].n,0), tb=order.reduce((a,s)=>a+g[s].br,0);
  return finding(L,col,'Stage SLA performance',(tn?Math.round(100*tb/tn):0)+'% breached',b(tb+' of '+tn)+' items are past the 7-day SLA ('+b((tn?Math.round(100*tb/tn):0)+'%')+'). Breach rate by stage below.',
    otable([['Stage',180,'l'],['Items',56,'c'],['Breach&gt;7d',82,'c'],['Breach%',66,'c'],['Avg',52,'c']],rows));
}
function f_dept(its,L,col){
  const g={}; for(const it of its){ const d=it.dept||'(unspecified)'; const e=g[d]||(g[d]={n:0,br:0,ages:[],val:0,priced:0}); e.n++; if(it.hasRecordedValue){e.val+=it.value;e.priced++;} if(it.age!=null){e.ages.push(it.age); if(it.age>7)e.br++;} }
  const ent=Object.entries(g).sort((a,b2)=>b2[1].n-a[1].n);
  const rows=ent.slice(0,8).map(([d,e])=>[esc(d.slice(0,28)),String(e.n),sv(String(e.br),'#b91c1c'),agec(avg(e.ages)),e.priced?(e.priced+' priced &#183; AED '+money(e.val)):'Not yet priced']);
  return finding(L,col,'By department &#8212; where it sits','',b(ent[0][0])+' has the most ('+b(ent[0][1].n)+' items). Full split below &#8212; route to each department head.',
    otable([['Department',210,'l'],['Items',56,'c'],['Breach&gt;7d',82,'c'],['Avg',52,'c'],['Value',130,'r']],rows));
}
function f_noNamedOwner(its,L,col){
  const rows=its.filter(it=>it.doc==='PR'&&it.noNamedOwner); if(!rows.length)return '';
  const groups=grpBy(rows,it=>(it.originalOwner||it.owner||'Owner not recorded')+'|'+(it.deliveryIssue||'owner not recorded in F&O')+'|'+(it.workClass||'Work class not reported')+'|'+(it.dept||'Department not reported'));
  const body=Object.entries(groups).map(([key,list])=>{
    const [owner,reason,cls,dept]=key.split('|');
    const oldest=Math.max(...list.map(it=>Number(it.age)||0));
    const pricing=pricedStats(list);
    return [esc(owner),esc(reason),esc(cls),esc(dept),String(list.length),agec(oldest),pricing.priced?(pricing.priced+' priced &#183; AED '+money(pricing.value)):'Not yet priced'];
  }).sort((a,b2)=>Number(b2[4])-Number(a[4]));
  return finding(L,col,'No named owner',rows.length+' requisitions',b(rows.length)+' requisitions have no active owner or no usable email address. They remain visible here and are not emailed as personal queues.',
    otable([['Recorded holder',150,'l'],['Reason',150,'l'],['Class of work',220,'l'],['Department',155,'l'],['Items',54,'c'],['Oldest',58,'c'],['Recorded value',180,'r']],body));
}

/* ---- divisions ---- */
const DIVS = [
 {key:'procurement', xlsx:'PRPO_Suppliers_OpenOrders_list.xlsx', title:'Suppliers, Open Orders &amp; Unowned PRs',
  heading:'PR / PO Pipeline &#8212; Suppliers, Open Orders &amp; Unowned PRs', sub:'Supplier-side POs plus requisitions that have no named owner &#183; member queues arrive as individual action emails', accent:'#a855f7',
  pr:[], po:['Confirmed Open Order','Sent to Supplier'], restage:'Confirmed Open Order',
  match:it=>(it.doc==='PR'&&it.noNamedOwner)||(it.doc==='PO'&&(it.stage==='Sent to Supplier'||(it.ppend===false&&String(it.raw['Approval status']||'')==='Confirmed'&&String(it.raw['Purchase order status']||'')==='Open order'))),
  findings:[f=>f_noNamedOwner(f,'A','#dc2626'), f=>f_vendor(f,'Sent to Supplier','B','#a855f7','Sent to Supplier &#8212; awaiting delivery / GRN','Chase the suppliers below for delivery, then move to invoicing.'), f=>f_value(f,'C','#2563eb'), f=>f_oldest(f,'D','#dc2626'), f=>f_sla(f,'E','#e11d48')]},
 {key:'invoicing', xlsx:'PRPO_PendingInvoicing_list.xlsx', title:'Pending Invoicing',
  heading:'PR / PO Pipeline &#8212; Pending Invoicing', sub:'POs confirmed &amp; received &#8212; awaiting supplier invoice posting by Accounts', accent:'#f97316',
  pr:[], po:['Pending Invoicing'], match:it=>it.stage==='Pending Invoicing',
  findings:[f=>f_vendor(f,'Pending Invoicing','A','#f97316','Pending Invoicing &#8212; by vendor','Post the supplier invoices below to clear these from the ledger.'), f=>f_value(f,'B','#2563eb'), f=>f_oldest(f,'C','#dc2626')]},
 {key:'ops_hm', send:false, xlsx:'PRPO_Operations_HomeMaint_FitOut_list.xlsx', title:'Operations &#183; Home Maintenance + FitOut',
  heading:'PR / PO Pipeline &#8212; Operations (Home Maintenance &amp; FitOut)', sub:'Operations-to-confirm &amp; dep-manager PRs whose requisition department is Home Maintenance or FitOut', accent:'#14b8a6',
  pr:['Operations to Confirm','Dep Managers'], po:[], depts:new Set(['Home Maintenance Services','FitOut Services']),
  findings:[f=>f_owners(f,'A','#14b8a6','Pending with &#8212; Home Maintenance &amp; FitOut queue'), f=>f_value(f,'B','#2563eb'), f=>f_oldest(f,'C','#dc2626'), f=>f_sla(f,'D','#e11d48')]},
 {key:'ops_all', send:false, xlsx:'PRPO_Operations_AllDepts_list.xlsx', title:'Operations &#183; All Departments',
  heading:'PR / PO Pipeline &#8212; Operations (All Departments)', sub:'Operations-to-confirm &amp; dep-manager PRs for all other requisition departments (Building Services, Contracted Cleaning, Landscaping, etc.)', accent:'#8b5cf6',
  pr:['Operations to Confirm','Dep Managers'], po:[], xdepts:new Set(['Home Maintenance Services','FitOut Services']),
  findings:[f=>f_owners(f,'A','#8b5cf6','Pending with &#8212; who is holding the queue'), f=>f_dept(f,'B','#4f46e5'), f=>f_value(f,'C','#2563eb'), f=>f_oldest(f,'D','#dc2626'), f=>f_sla(f,'E','#e11d48')]},
];
function filterDiv(items,cfg){ let fil=items.filter(it=> cfg.match? cfg.match(it) : (itemDivision(it)===cfg.key&&(!cfg.keep||cfg.keep(it))) ); if(cfg.restage) fil=fil.map(it=>it.doc==='PO'&&it.stage!=='Sent to Supplier'?Object.assign({},it,{stage:cfg.restage}):it); return fil; }

async function buildXlsxBase64(fil, cfg){
  const wb=new ExcelJS.Workbook(); wb.creator='Strive Services Group'; wb.created=new Date();
  const ws=wb.addWorksheet('Open Items', { views:[{ state:'frozen', ySplit:1 }], properties:{ defaultRowHeight:16 } });
  ws.columns=[
    {header:'Ref',key:'ref',width:16},{header:'Doc',key:'doc',width:7},{header:'Quote Ref',key:'qref',width:12},{header:'Stage / Bucket',key:'stage',width:22},
    {header:'Stage reason code',key:'classcode',width:32},{header:'Class of work',key:'workclass',width:58},{header:'What to do',key:'workaction',width:48},
    {header:'Pending Internal / Client',key:'pendingside',width:25},{header:'Step name',key:'step',width:34},{header:'Status',key:'status',width:22},{header:'Department',key:'dept',width:26},
    {header:'Location',key:'loc',width:22},{header:'Pending With',key:'pend',width:20},{header:'Delivery issue',key:'deliveryissue',width:28},{header:'Vendor',key:'vendor',width:30},
    {header:'F&O export total (AED)',key:'value',width:19},{header:'Pricing state',key:'pricing',width:18},{header:'Age (days)',key:'age',width:11},{header:'Age basis',key:'agebasis',width:18},{header:'Age band',key:'ageband',width:24},{header:'Owner data quality',key:'shared',width:42},{header:'Created',key:'created',width:13},
    {header:'Step date',key:'stepd',width:13},{header:'Clock source',key:'clocksrc',width:30},{header:'Title / Name',key:'title',width:34},{header:'Preparer / Linked PR',key:'prep',width:20}
  ];
  fil.slice().sort((a,b2)=>(a.classOrder||999)-(b2.classOrder||999)||String(a.dept||'').localeCompare(String(b2.dept||''))||(b2.age||0)-(a.age||0)||b2.value-a.value).forEach(it=>{ const r=it.raw; let status,loc,ven,created,stepd,title,prep;
    if(it.doc!=='PO'){ status=String(r['Status']||''); loc=r['Location']; ven=''; created=ymdStr(r['Created date']); stepd=ymdStr(r['Step date and time']); title=r['Name']; prep=r['Preparer']; }
    else { status=(String(r['Approval status']||'')+' / '+String(r['Purchase order status']||'')).replace(/^ \/ | \/ $/g,''); loc=r['Location']; ven=r['Vendor name']; created=ymdStr(r['Created date and time']); stepd=ymdStr(r['Step date and time']); title=''; prep=r['Purchase requisition']; }
    const pend=it.owner;  // computed owner: ops-user for ops-confirm, Created-by for Draft POs, approver otherwise
    const qref=(it.doc!=='PO'? String(r['Quotation reference']||'') : '');
    ws.addRow({ref:it.ref,doc:it.typ,qref,stage:it.stage,classcode:it.workClassCode,workclass:it.workClass,workaction:it.workAction,pendingside:it.pendingSide||pendingSide(it),step:r['Step name'],status,dept:it.dept,loc,pend,deliveryissue:it.deliveryIssue||'',vendor:ven,value:it.hasRecordedValue?Math.round(it.value*100)/100:null,pricing:it.hasRecordedValue?'Price recorded':'Not yet priced',age:(it.age==null?null:it.age),agebasis:it.clockBasis==='raised'?'Raised date':'Current step date',ageband:ageBandLabel(it),shared:it.sharedLabel||'Not shared',created,stepd,clocksrc:it.clockLabel||'',title,prep}); });
  const DIVCOL={procurement:'FF1D4ED8',finance:'FF16A34A',ops_hm:'FF0F766E',ops_all:'FF7C3AED'};
  const HEAD=DIVCOL[cfg&&cfg.key]||'FF14315E';
  const h=ws.getRow(1); h.height=26;
  h.eachCell(c=>{ c.fill={type:'pattern',pattern:'solid',fgColor:{argb:HEAD}}; c.font={bold:true,color:{argb:'FFFFFFFF'},size:11}; c.alignment={vertical:'middle',horizontal:'center',wrapText:true}; c.border={bottom:{style:'thin',color:{argb:'FF0B2350'}}}; });
  const N=ws.rowCount;
  for(let i=2;i<=N;i++){ const row=ws.getRow(i); row.height=15; const zeb=(i%2===1);
    row.eachCell({includeEmpty:true},c=>{ c.alignment={vertical:'middle'}; c.border={bottom:{style:'hair',color:{argb:'FFE1E7F0'}}}; c.font={size:10.5}; if(zeb){ c.fill={type:'pattern',pattern:'solid',fgColor:{argb:'FFEFF3FA'}}; } });
    row.getCell('ref').font={bold:true,size:10.5,color:{argb:'FF14315E'}};
    const vc=row.getCell('value'); vc.numFmt='#,##0'; vc.alignment={vertical:'middle',horizontal:'right'}; vc.font={size:10.5,color:{argb:'FF0F766E'},bold:true};
    const ac=row.getCell('age'); ac.numFmt='0'; ac.alignment={vertical:'middle',horizontal:'center'};
    const av=ac.value; if(typeof av==='number'){ ac.font={bold:true,size:10.5,color:{argb: av>30?'FFB42318': av>7?'FF9A6700':'FF1F7A33'}}; }
    row.getCell('doc').alignment={vertical:'middle',horizontal:'center'};
  }
  ws.autoFilter={ from:{row:1,column:1}, to:{row:1,column:ws.columns.length} };
  const buf=await wb.xlsx.writeBuffer();
  return Buffer.from(buf).toString('base64');
}

/* ---- HS-D08 style: navy shell header/footer + PR/PO detail line-item list ---- */
const HF="'Segoe UI', Aptos, Verdana, Arial, sans-serif", HNAVY='#0F2A6B', HGOLD='#FAC775', HMUT='#C9D3EA', HBORD='#D8DEE8';
function f_details(fil, cfg){
  const rows=fil.slice().sort((a,b2)=>(b2.age||0)-(a.age||0)).slice(0,15);
  const desc=it=> it.doc==='PO' ? (it.raw['Vendor name']||'-') : (it.raw['Name']||'-');
  const th='padding:7px 9px;font:700 10px '+HF+';color:#5A6578;text-transform:uppercase;background:#F4F6FB;';
  const head='<tr>'+[['Document','l'],['Description','l'],['Waiting on step','l'],['Pending with','l'],['Value','r'],['Age','r']].map(c=>'<td '+(c[1]==='r'?'align="right" ':'')+'style="'+th+'">'+c[0]+'</td>').join('')+'</tr>';
  const body=rows.map(it=>{ const a=Math.round(it.age||0); const acol=a>30?'#B42318':(a>7?'#9A6700':'#1F7A33'); const td='padding:6px 9px;font:400 12px '+HF+';color:#1A2233;border-bottom:1px solid '+HBORD+';'; const tdl=td+'border-left:1px solid #f1f4f8;';
    return '<tr><td style="'+td+'font-weight:600;white-space:nowrap;">'+esc(it.ref)+'</td>'
      +'<td style="'+tdl+'">'+esc(String(desc(it)).slice(0,46))+'</td>'
      +'<td style="'+tdl+'color:'+HNAVY+';white-space:nowrap;">'+esc(String(it.raw['Step name']||'-').slice(0,28)+clockNote(it.raw))+'</td>'
      +'<td style="'+tdl+'font-weight:600;white-space:nowrap;">'+esc(String(it.owner||'-').slice(0,18))+'</td>'
      +'<td align="right" style="'+tdl+'font-weight:600;white-space:nowrap;">'+esc(itemPrice(it))+'</td>'
      +'<td align="right" style="'+tdl+'font-weight:700;color:'+acol+';white-space:nowrap;">'+esc(agePhrase(it))+'</td></tr>'; }).join('');
  const more=fil.length>15?'<tr><td colspan="6" style="padding:7px 9px;font:400 11px '+HF+';color:#5A6578;background:#F9FAFC;">&#8230;and '+(fil.length-15)+' more &#8212; full list in the attached '+cfg.xlsx+'.</td></tr>':'';
  return '<div style="font:700 14px '+HF+';color:'+HNAVY+';margin:16px 0 2px;">Details &#8212; PR / PO list</div>'
    +'<div style="font:400 11px '+HF+';color:#5A6578;margin:0 0 10px;">Oldest first &#183; every age says whether it starts at the raised date or a distinct current-step date.</div>'
    +'<table role="presentation" width="'+W+'" cellpadding="0" cellspacing="0" style="width:'+W+'px;border:1px solid '+HBORD+';border-collapse:separate;border-spacing:0;border-radius:8px;overflow:hidden;box-shadow:0 1px 3px rgba(16,24,40,0.07);">'+head+body+more+'</table>';
}

function buildDivision(cfg, items, hist){
  const fil=filterDiv(items,cfg);
  const agg={}; for(const it of fil){ const k=it.doc+'|'+it.stage; const x=agg[k]||(agg[k]={n:0,sum:0,c:0,amt:0,priced:0,unpriced:0}); x.n++; if(it.hasRecordedValue){x.amt+=it.value;x.priced++;}else x.unpriced++; if(it.age!=null){x.sum+=it.age;x.c++;} }
  const pairs=STAGE_ORDER.filter(p=>agg[p[0]+'|'+p[1]]);
  // Per-stage 3-day history (from the git-snapshot data), shown vertically inside each card.
  const hdays=histDayList(); const histAgg={};
  for(const hd of hdays){ if(hist&&hist[hd.ymd]){ const m={}; for(const it2 of filterDiv(hist[hd.ymd],cfg)){ const k2=it2.doc+'|'+it2.stage; m[k2]=(m[k2]||0)+1; } histAgg[hd.ymd]=m; } }
  const trFn=(doc,bk)=>hdays.map(hd=>({lab:hd.lab,n:histAgg[hd.ymd]?String(histAgg[hd.ymd][doc+'|'+bk]||0):'&#8211;'}));
  const cards=cardrow2(pairs,agg,trFn);
  const analysis=cfg.findings.map(fn=>fn(fil)).join('');
  const stamp=new Date(Date.now()+4*3600*1000).toLocaleDateString('en-GB',{day:'2-digit',month:'short',year:'numeric',timeZone:'UTC'});
  const tot=fil.length, pricing=pricedStats(fil), totv=pricing.value;
  const att='<div style="border:1px solid #cbd9ec;background:#f2f7ff;padding:11px 14px;margin:0 0 2px;border-radius:8px;font:400 12px '+HF+';color:#334867;">&#8505;&#65039; <b style="color:'+HNAVY+';">Source:</b> '+esc(provenanceSentence(fil.length?fil:items))+'</div>';
  const titleTxt=cfg.title.replace(/&#183;/g,'·').replace(/&amp;/g,'&');
  const trendBlock='<table role="presentation" cellpadding="0" cellspacing="0" style="border-collapse:separate;"><tr><td width="300" valign="top" style="width:300px;">'
    +tcard('Open items &#8212; 3 days',trendCols(hist,s=>filterDiv(s,cfg).length,tot),cfg.accent||'#3b82f6',pricing.priced?pricing.priced+' priced &#183; AED '+money(totv):'not yet priced')
    +'</td></tr></table><div style="height:12px;font-size:12px;line-height:12px;">&#160;</div>';
  const warning=(fil[0]&&fil[0].freshnessWarning)||items.freshnessWarning||'';
  const inner='<div style="width:'+W+'px;font-family:'+FONT+';color:#22303c;">'
    +freshnessBanner(warning)
    +'<div style="font:400 12px '+HF+';color:#607083;margin:0 0 10px;">This queue: '+b(tot)+' open items &#183; '+priceLine(fil)+' &#183; the same export values shown on the dashboard.</div>'
    +trendBlock+cards+att
    +'<div style="font-family:'+FONT+';font-weight:800;font-size:15px;color:'+NAVY+';margin:16px 0 2px;">&#128269; Analysis &#8212; who to chase today</div>'
    +'<div style="font-family:'+FONT+';font-size:11.5px;color:#7688a0;margin:0 0 11px;">Auto-generated daily from the latest data, scoped to '+cfg.title+'.</div>'
    +'<div style="width:'+W+'px;">'+analysis+'</div>'
    +f_details(fil,cfg)
    +'</div>';
  const shell='<table role="presentation" width="1040" cellpadding="0" cellspacing="0" style="width:1040px;max-width:1040px;">'
    +'<tr><td style="background:'+HNAVY+';padding:18px 22px;border-bottom:3px solid '+HGOLD+';border-radius:14px 14px 0 0;"><table role="presentation" width="100%" cellpadding="0" cellspacing="0"><tr>'
    +'<td style="font:700 20px '+HF+';color:#fff;">'+cfg.heading+'</td>'
    +'<td align="right" valign="top" style="font:600 12px '+HF+';color:'+HGOLD+';">'+stamp+'</td></tr></table>'
    +'<div style="font:400 12px '+HF+';color:'+HMUT+';margin-top:4px;">'+cfg.sub+' &#183; <a href="'+DASH+'" style="color:'+HGOLD+';font-weight:800;text-decoration:underline;">Open Live Dashboard &#8599;</a></div></td></tr>'
    +'<tr><td style="background:#fff;border-left:1px solid '+HBORD+';border-right:1px solid '+HBORD+';padding:18px 20px;">'+inner+'</td></tr>'
    +'<tr><td style="background:'+HNAVY+';padding:14px 20px;border-top:3px solid '+HGOLD+';border-radius:0 0 14px 14px;font:400 11px '+HF+';color:'+HMUT+';">'
    +'<div style="color:'+HGOLD+';font-weight:700;">FOR EXCELLENCE WE STRIVE</div>'
    +'<div style="margin-top:6px;">PR / PO Pipeline &#183; '+titleTxt+' &#183; '+stamp+' &#183; automated daily 10:00 AM Dubai. '+esc(provenanceSentence(fil.length?fil:items))+' <b>Each age is labelled by its source event.</b> Full line-item list attached ('+cfg.xlsx+').</div></td></tr></table>';
  const wrap='<!doctype html><html xmlns:o="urn:schemas-microsoft-com:office:office"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><meta http-equiv="X-UA-Compatible" content="IE=edge"><title>'+cfg.heading+'</title>'
    +'<!--[if mso]><xml><o:OfficeDocumentSettings><o:PixelsPerInch>96</o:PixelsPerInch></o:OfficeDocumentSettings></xml><![endif]-->'
    +'<style>table{mso-table-lspace:0pt;mso-table-rspace:0pt;border-collapse:collapse;} td{mso-line-height-rule:exactly;} img{-ms-interpolation-mode:bicubic;} body{-webkit-text-size-adjust:100%;-ms-text-size-adjust:100%;}</style></head>'
    +'<body style="margin:0;padding:0;background:#EEF1F6;"><table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background:#EEF1F6;padding:20px 0;"><tr><td align="center">'+shell+'</td></tr></table></body></html>';
  const subject='PR / PO Pipeline — '+cfg.title.replace(/&#183;/g,'·').replace(/&amp;/g,'&')+' ('+stamp+')';
  return { subject, html:wrap, fil, count:tot, value:totv, cfg };
}

/* ---- PERSONAL action emails: one per named F&O Pending Approver/User ---- */
function personalPool(items){ return items.filter(it=>it.ppend!==false && !it.noNamedOwner && !it.deliveryIssue && it.stage!=='Sent to Supplier'&&it.stage!=='Pending Invoicing' && String(it.owner==null?'':it.owner).trim()!=='' && it.owner!=='(unassigned)'); }
// Same human appears under both full-name and username F&O accounts — fold them into one personal email.
function canonOwner(u){ return String(u==null?'':u).trim(); }
function applyDeliveryPolicy(items){
  for(const it of items){
    if(it.doc!=='PR')continue;
    const original=String(it.owner==null?'':it.owner).trim();
    const issue=it.noNamedOwner?'owner not recorded in F&O':'';
    if(issue){ it.originalOwner=original||'Owner not recorded'; it.deliveryIssue=issue; it.noNamedOwner=true; it.div='procurement'; }
  }
  return items;
}
function groupByOwner(pool){ const g={}; for(const it of pool){ const cu=canonOwner(it.owner); const k=_norm(cu); const e=g[k]||(g[k]={key:k,items:[],disp:{}}); e.items.push(it); e.disp[cu]=(e.disp[cu]||0)+1; } return Object.values(g).map(e=>({key:e.key,user:Object.entries(e.disp).sort((a,b2)=>b2[1]-a[1])[0][0],items:e.items})).sort((a,b2)=>b2.items.length-a.items.length); }
function firstName(u){ const t=String(u==null?'':u).trim().split(/[.\s]+/)[0]||''; return t? t.charAt(0).toUpperCase()+t.slice(1) : 'there'; }
const PW=1000;
function pcard(label,val,sub,col,bg,vs){ return '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="width:100%;background:'+bg+';border:1px solid #e8ecf2;border-top:3px solid '+col+';border-radius:10px;box-shadow:0 1px 3px rgba(16,24,40,0.08);"><tr><td height="92" valign="middle" style="padding:11px 13px;height:92px;"><div style="font-family:'+FONT+';font-size:9.5px;font-weight:800;color:#5b6b7f;text-transform:uppercase;">'+label+'</div><div style="margin:6px 0 2px;white-space:nowrap;font-family:'+FONT+';font-size:'+(vs||24)+'px;font-weight:800;color:'+col+';">'+val+'</div><div style="font-family:'+FONT+';font-size:11px;font-weight:700;color:'+TEAL+';white-space:nowrap;">'+sub+'</div></td></tr></table>'; }
function buildPersonal(p,hist){
  const fil=p.items.slice().sort((a,b2)=>(a.classOrder||999)-(b2.classOrder||999)||String(a.dept||'').localeCompare(String(b2.dept||''))||(b2.age||0)-(a.age||0)||b2.value-a.value);
  const n=fil.length, pricing=pricedStats(fil), totv=pricing.value;
  const pricingQueue=fil.filter(it=>it.workClassCode==='ACTIVE_LINES_NOT_FULLY_PRICED').length;
  const unvaluedOutsidePricing=fil.filter(it=>!it.hasRecordedValue&&it.workClassCode!=='ACTIVE_LINES_NOT_FULLY_PRICED').length;
  const sourceShared=fil.filter(it=>it.sourceShared).length;
  const sharedWithOthers=fil.filter(it=>it.otherLiveBuyers&&it.otherLiveBuyers.length).length;
  const prn=fil.filter(it=>it.doc!=='PO').length, pon=n-prn;
  const old=fil.slice().sort((a,b2)=>(b2.age||0)-(a.age||0))[0], oldAge=Math.round((old&&old.age)||0);
  const priceSummaryText=n+' item'+(n===1?'':'s')+' pending your action · '+pricingQueue+' still being priced · '+(pricing.priced?pricing.priced+' priced, worth AED '+money(totv):'not yet priced')+(unvaluedOutsidePricing?' · '+unvaluedOutsidePricing+' other item'+(unvaluedOutsidePricing===1?'':'s')+' without a recorded price':'');
  const oldestSummaryText=oldAge>0?'Oldest item was '+agePhrase(old)+'.':'';
  const sharedSummaryText=sourceShared?sourceShared+' source-shared item'+(sourceShared===1?'':'s')+' · '+sharedWithOthers+' of these are shared with other active buyers. Inactive usernames are excluded from the names shown on each line.':'';
  const br=fil.filter(it=>(it.age||0)>7).length;
  const stamp=new Date(Date.now()+4*3600*1000).toLocaleDateString('en-GB',{day:'2-digit',month:'short',year:'numeric',timeZone:'UTC'});
  // Operations-to-Confirm split (per CK): "Unit prices updated" = shared with the CLIENT for confirmation ->
  // "Pending Client" (cards + Excel only, NOT the body table); every other ops-confirm step -> "Pending Internal".
  const CLIENT_STEP='Unit prices updated in PR lines';
  const pst=it=>it.stage==='Operations to Confirm'?(String(it.raw['Step name']||'').trim()===CLIENT_STEP?'Pending Client':'Pending Internal'):it.stage;
  const xfil=fil.map(it=>Object.assign({},it,{pendingSide:pst(it),stage:pst(it)}));
  const G=12, TW=268, cw=Math.floor((PW-G*3-TW)/3);
  const gutter='<td width="'+G+'" style="width:'+G+'px;font-size:6px;line-height:6px;mso-line-height-rule:exactly;">&#160;</td>';
  // First card = 3-day pending trend (11th / 12th / Today); Total value, Oldest, SLA cards stay.
  const myCount=snap=>{ const e=groupByOwner(personalPool(snap)).find(x=>x.key===p.key); return e?e.items.length:0; };
  const trendCard=tcard('Items pending &#8212; 3 days',trendCols(hist,myCount,n),'#3b82f6','PR '+prn+' &#183; PO '+pon);
  const defs=[['Priced value',pricing.priced?'AED '+money(totv):'Not yet priced',pricing.priced+' priced &#183; '+pricing.unpriced+' without a recorded price','#0f766e','#ebfbf7',17],
              ['Oldest item',oldAge+'d',esc(old?String(old.ref):'-'),'#dc2626','#fef2f2',22],
              ['Past 7-day SLA',String(br),'of '+n+' item'+(n===1?'':'s'),'#f59e0b','#fff9ec',22]];
  const cardsHtml='<table class="summary-grid" role="presentation" cellpadding="0" cellspacing="0" style="border-collapse:separate;"><tr><td width="'+TW+'" valign="top" style="width:'+TW+'px;">'+trendCard+'</td>'+gutter+defs.map((c,i)=>'<td width="'+cw+'" valign="top" style="width:'+cw+'px;">'+pcard(c[0],c[1],c[2],c[3],c[4],c[5])+'</td>'+(i<defs.length-1?gutter:'')).join('')+'</tr></table><div style="height:'+G+'px;font-size:'+G+'px;line-height:'+G+'px;">&#160;</div>';
  // Stage cards row — merged per STAGE (no PR/PO duplicates; doc split already shown on "Items pending").
  // Rendered only when the person has MORE than one stage; a single redundant stage card is dropped.
  const sagg={}; for(const it of xfil){ const x=sagg[it.stage]||(sagg[it.stage]={n:0,sum:0,c:0,amt:0,priced:0,unpriced:0}); x.n++; if(it.hasRecordedValue){x.amt+=it.value;x.priced++;}else x.unpriced++; if(it.age!=null){x.sum+=it.age;x.c++;} }
  const SORD=['Re-Assigned/Rejected','Pending Internal','Pending Client','Procurement','Dep Managers','Finance'];
  const sl=SORD.filter(s=>sagg[s]).concat(Object.keys(sagg).filter(s=>!SORD.includes(s)));
  // Per-stage 3-day history for THIS person (vertical rows inside each stage card).
  const hdays=histDayList(); const histStage={};
  for(const hd of hdays){ if(hist&&hist[hd.ymd]){ const e2=groupByOwner(personalPool(hist[hd.ymd])).find(x=>x.key===p.key); const m={}; if(e2){ for(const it2 of e2.items){ const s2=pst(it2); m[s2]=(m[s2]||0)+1; } } histStage[hd.ymd]=m; } }
  const strFn=s=>hdays.map(hd=>({lab:hd.lab,n:histStage[hd.ymd]?String(histStage[hd.ymd][s]||0):'&#8211;'}));
  const scard2=(bk,x,tr)=>{ const a=x.c?(x.sum/x.c):0, price=x.priced?(x.priced+' priced &#183; AED '+money(x.amt)):'not yet priced'; return '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="width:100%;background:'+(GRAD[bk]||'#f6f8fb')+';border:1px solid #e8ecf2;border-top:3px solid '+(COLOR[bk]||NAVY)+';border-radius:10px;box-shadow:0 1px 3px rgba(16,24,40,0.08);"><tr><td style="padding:11px 13px;"><div style="font-family:'+FONT+';font-size:9.5px;font-weight:800;color:#5b6b7f;text-transform:uppercase;">'+esc(bk)+'</div><div style="margin:6px 0 2px;white-space:nowrap;"><span style="font-family:'+FONT+';font-size:24px;font-weight:800;color:'+(COLOR[bk]||NAVY)+';">'+x.n+'</span><span style="font-family:'+FONT+';font-size:12px;font-weight:800;color:'+RED+';"> ('+a.toFixed(1)+'d)</span>'+deltaBadge(tr,x.n)+'</div><div style="font-family:'+FONT+';font-size:11px;font-weight:700;color:'+TEAL+';">'+price+(x.unpriced?' &#183; '+x.unpriced+' not yet priced':'')+'</div>'+histLine(tr)+'</td></tr></table>'; };
  const cw2=Math.min(200,Math.floor((PW-G*Math.max(0,sl.length-1))/Math.max(1,sl.length)));
  const stageCards=sl.length>1?'<table class="summary-grid" role="presentation" cellpadding="0" cellspacing="0" style="border-collapse:separate;"><tr>'+sl.map((s,i)=>'<td width="'+cw2+'" valign="top" style="width:'+cw2+'px;">'+scard2(s,sagg[s],strFn(s))+'</td>'+(i<sl.length-1?gutter:'')).join('')+'</tr></table><div style="height:'+G+'px;font-size:'+G+'px;line-height:'+G+'px;">&#160;</div>':'';
  const prItems=xfil.filter(it=>it.doc==='PR');
  const classGroups=grpBy(prItems,it=>it.workClassCode||'NOT_REPORTED');
  const classKeys=Object.keys(classGroups).sort((a,b2)=>(classGroups[a][0].classOrder||999)-(classGroups[b2][0].classOrder||999));
  const classSummary=classKeys.length?'<div style="font-family:'+FONT+';font-size:11.5px;color:#4b5c74;line-height:1.65;margin:0 0 12px;"><b style="color:'+NAVY+';">Shape of your PR queue:</b> '+classKeys.map(k=>esc(classGroups[k][0].workClass)+' <b style="color:'+NAVY+';">'+classGroups[k].length+'</b>').join(' &#183; ')+'</div>':'';
  const sectionMeta=[];
  const detailTable=list=>'<div class="detail-scroll">'+otable([['PR #',80,'l'],['Site',86,'l'],['Description',125,'l'],['Class of work',160,'l'],['Department',105,'l'],['Price',92,'r'],['Age / source',135,'l'],['Band',105,'l'],['Shared assignment',150,'l'],['Queue',70,'l']],list.map(it=>[
    '<span style="font-weight:700;color:'+NAVY+';">'+esc(it.ref)+'</span>',esc(String(it.raw['Location']||'-').slice(0,22)),esc(String(it.raw['Name']||'-').slice(0,34)),esc(it.workClass),esc(it.dept||'-'),esc(itemPrice(it)),esc(agePhrase(it)),esc(ageBandLabel(it)),esc(it.sharedLabel||'Not shared'),esc(it.pendingSide==='Pending Client'||it.pendingSide==='Pending Internal'?it.pendingSide:'—')
  ]))+'</div>';
  const sectionHtml=classKeys.map(code=>{
    const items=classGroups[code], first=items[0], deptGroups=grpBy(items,it=>it.dept||'Department not reported');
    const depts=Object.keys(deptGroups).sort((a,b2)=>deptGroups[b2].length-deptGroups[a].length||a.localeCompare(b2));
    sectionMeta.push({code,label:first.workClass,count:items.length,departments:depts.map(dept=>({department:dept,count:deptGroups[dept].length}))});
    const deptHtml=depts.map(dept=>{
      const rows=deptGroups[dept].slice().sort((a,b2)=>(b2.age||0)-(a.age||0)||b2.value-a.value);
      return '<div style="font-family:'+FONT+';font-size:12px;font-weight:800;color:'+NAVY+';margin:10px 0 6px;">'+esc(dept)+' &#183; '+rows.length+'</div>'
        +detailTable(rows);
    }).join('');
    const sharedN=items.filter(it=>it.otherLiveBuyers&&it.otherLiveBuyers.length).length;
    return '<div style="border-left:4px solid '+(COLOR[first.stage]||'#145A95')+';padding-left:12px;margin:18px 0 8px;"><div style="font-family:'+FONT+';font-size:15px;font-weight:800;color:'+NAVY+';">'+esc(first.workClass)+' &#183; '+items.length+'</div><div style="font-family:'+FONT+';font-size:11.5px;color:#607083;margin-top:2px;">'+esc(first.workAction)+(sharedN?' &#183; '+sharedN+' of these are shared with other buyers':'')+'</div></div>'+deptHtml;
  }).join('');
  const poItems=xfil.filter(it=>it.doc==='PO').sort((a,b2)=>(b2.age||0)-(a.age||0)||b2.value-a.value);
  if(poItems.length)sectionMeta.push({code:'PO_ACTION',label:'Purchase order action',count:poItems.length,departments:[]});
  const poHtml=poItems.length?'<div style="border-left:4px solid #0891b2;padding-left:12px;margin:18px 0 8px;"><div style="font-family:'+FONT+';font-size:15px;font-weight:800;color:'+NAVY+';">Purchase order action &#183; '+poItems.length+'</div></div><div class="detail-scroll">'+otable([['PO #',95,'l'],['Vendor',190,'l'],['Department',145,'l'],['Step',170,'l'],['Value',105,'r'],['Age',50,'c'],['Band',70,'c']],poItems.map(it=>['<span style="font-weight:700;color:'+NAVY+';">'+esc(it.ref)+'</span>',esc(it.vendor||'-'),esc(it.dept||'-'),esc(String(it.raw['Step name']||'-')+clockNote(it.raw)),'AED '+money(it.value),agec(it.age),esc(it.ageBand)]))+'</div>':'';
  const clientN=xfil.filter(it=>it.pendingSide==='Pending Client').length;
  const att='<div style="border:1px solid #cbd9ec;background:#f2f7ff;padding:10px 13px;margin:0 0 2px;border-radius:8px;font:400 12px '+HF+';color:#334867;">&#8505;&#65039; <b style="color:'+HNAVY+';">Source:</b> '+esc(provenanceSentence(fil))+'</div>';
  const warning=(fil[0]&&fil[0].freshnessWarning)||'';
  const inner='<div class="mail-inner" style="width:'+PW+'px;font-family:'+FONT+';color:#22303c;">'
    +freshnessBanner(warning)
    +'<div style="font:400 13px '+HF+';color:#334155;margin:0 0 8px;">Hi <b style="color:'+HNAVY+';">'+esc(firstName(p.user))+'</b> &#8212; '+esc(priceSummaryText)+'.'+(oldestSummaryText?' '+esc(oldestSummaryText):'')+'</div>'
    +(sharedSummaryText?'<div style="font:700 12px '+HF+';color:#334155;margin:0 0 10px;">'+esc(sharedSummaryText)+'</div>':'')
    +cardsHtml+stageCards+att
    +'<div style="font-family:'+FONT+';font-weight:800;font-size:15px;color:'+NAVY+';margin:16px 0 2px;">&#128203; Your pending items</div>'
    +'<div style="font-family:'+FONT+';font-size:11.5px;color:#7688a0;margin:0 0 9px;">Class first &#183; department next &#183; oldest then largest &#183; each age names its source event'+(clientN>0?' &#183; Queue shows Pending Internal or Pending Client':'')+'.</div>'
    +classSummary+sectionHtml+poHtml+'</div>';
  const shell='<table class="mail-shell" role="presentation" width="'+(PW+40)+'" cellpadding="0" cellspacing="0" style="width:'+(PW+40)+'px;max-width:'+(PW+40)+'px;">'
    +'<tr><td style="background:'+HNAVY+';padding:16px 20px;border-bottom:3px solid '+HGOLD+';border-radius:14px 14px 0 0;"><table role="presentation" width="100%" cellpadding="0" cellspacing="0"><tr>'
    +'<td style="font:700 19px '+HF+';color:#fff;">Your PR / PO Action List</td>'
    +'<td align="right" valign="top" style="font:600 12px '+HF+';color:'+HGOLD+';">'+stamp+'</td></tr></table>'
    +'<div style="font:400 12px '+HF+';color:'+HMUT+';margin-top:4px;">Personal daily digest of items waiting on <b style="color:#fff;">'+esc(p.user)+'</b> &#183; <a href="'+DASH+'" style="color:'+HGOLD+';font-weight:800;text-decoration:underline;">Open Live Dashboard &#8599;</a></div></td></tr>'
    +'<tr><td style="background:#fff;border-left:1px solid '+HBORD+';border-right:1px solid '+HBORD+';padding:16px 20px;">'+inner+'</td></tr>'
    +'<tr><td style="background:'+HNAVY+';padding:12px 20px;border-top:3px solid '+HGOLD+';border-radius:0 0 14px 14px;font:400 11px '+HF+';color:'+HMUT+';">'
    +'<div style="color:'+HGOLD+';font-weight:700;">FOR EXCELLENCE WE STRIVE</div>'
    +'<div style="margin-top:5px;">Sent automatically every day at 10:00 AM Dubai while items are pending with you. <b>Each age is labelled by its source event.</b> Your full line-item list is attached.</div></td></tr></table>';
  const heading='Your PR / PO Action List &#8212; '+esc(p.user);
  const wrap='<!doctype html><html xmlns:o="urn:schemas-microsoft-com:office:office"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><meta http-equiv="X-UA-Compatible" content="IE=edge"><title>'+heading+'</title>'
    +'<!--[if mso]><xml><o:OfficeDocumentSettings><o:PixelsPerInch>96</o:PixelsPerInch></o:OfficeDocumentSettings></xml><![endif]-->'
    +'<style>table{mso-table-lspace:0pt;mso-table-rspace:0pt;border-collapse:collapse;} td{mso-line-height-rule:exactly;} img{-ms-interpolation-mode:bicubic;} body{-webkit-text-size-adjust:100%;-ms-text-size-adjust:100%;}@media only screen and (max-width:640px){html,body{width:100%!important;margin:0!important;overflow-x:hidden!important}body>table,body>table>tbody,body>table>tbody>tr,body>table>tbody>tr>td,.mail-shell,.mail-shell>tbody,.mail-shell>tbody>tr,.mail-shell>tbody>tr>td{display:block!important;width:100%!important;max-width:100%!important;box-sizing:border-box!important}.mail-inner{display:block!important;width:100%!important;max-width:100%!important;min-width:0!important;box-sizing:border-box!important}.summary-grid,.summary-grid>tbody,.summary-grid>tbody>tr{display:block!important;width:100%!important;max-width:100%!important}.summary-grid>tbody>tr>td{display:block!important;width:auto!important;max-width:100%!important;margin-bottom:8px!important}.summary-grid>tbody>tr>td[width="12"]{display:none!important}.detail-scroll{display:block!important;width:100%!important;max-width:100%!important;overflow-x:auto!important;-webkit-overflow-scrolling:touch}}</style></head>'
    +'<body style="margin:0;padding:0;background:#EEF1F6;"><table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background:#EEF1F6;padding:20px 0;"><tr><td align="center">'+shell+'</td></tr></table></body></html>';
  const subject='Action needed — '+n+' PR/PO item'+(n===1?'':'s')+' pending with you ('+stamp+')';
  const xlsx='PRPO_'+p.user.replace(/[^\w.-]+/g,'_')+'_pending.xlsx';
  return { subject, html:wrap, fil:xfil, count:n, value:totv, pricedCount:pricing.priced,pricingQueueCount:pricingQueue,unvaluedOutsidePricing,sourceSharedCount:sourceShared,sharedWithOtherActiveBuyers:sharedWithOthers,priceSummaryText,oldestSummaryText,sharedSummaryText,user:p.user,key:p.key,xlsx,sections:sectionMeta,classCounts:Object.fromEntries(classKeys.map(k=>[k,classGroups[k].length])) };
}
async function sendPersonal(out, context){
  const from=process.env.PRPO_MAIL_FROM||process.env.MAIL_FROM;
  if(!from) throw new Error('MAIL_FROM / PRPO_MAIL_FROM not set');
  const msg=await buildPersonalMessage(out);
  const token=await getToken('https://graph.microsoft.com');
  const r=await fetch('https://graph.microsoft.com/v1.0/users/'+encodeURIComponent(from)+'/sendMail',{method:'POST',headers:{Authorization:'Bearer '+token,'Content-Type':'application/json'},body:JSON.stringify({message:msg,saveToSentItems:true})});
  if(r.status!==202){ const j=await r.json().catch(()=>({})); throw new Error('personal sendMail '+r.status+' '+JSON.stringify(j.error||j).slice(0,300)); }
  if(context) context.log('personal test-channel message for '+out.user+' -> '+WAQAS_ONLY_RECIPIENT);
  return {user:out.user,sent:true,to:WAQAS_ONLY_RECIPIENT,waqasOnly:true};
}

/* ---- auth + send ---- */
function guardedSubject(intended,subject){ return '[FOR '+String(intended||'PR / PO team')+'] '+subject; }
async function buildPersonalMessage(out){
  return {subject:guardedSubject(out.user,out.subject),body:{contentType:'HTML',content:out.html},toRecipients:[{emailAddress:{address:WAQAS_ONLY_RECIPIENT}}],attachments:[{'@odata.type':'#microsoft.graph.fileAttachment',name:out.xlsx,contentType:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',contentBytes:await buildXlsxBase64(out.fil,{key:'personal'})}]};
}
function buildDivisionMessage(out,xlsxB64){
  return {subject:guardedSubject(out.cfg.key+' team',out.subject),body:{contentType:'HTML',content:out.html},toRecipients:[{emailAddress:{address:WAQAS_ONLY_RECIPIENT}}],attachments:[{'@odata.type':'#microsoft.graph.fileAttachment',name:out.cfg.xlsx,contentType:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',contentBytes:xlsxB64}]};
}
async function getToken(scopeBase){ const body=new URLSearchParams({client_id:process.env.CLIENT_ID,client_secret:process.env.CLIENT_SECRET,grant_type:'client_credentials',scope:scopeBase.replace(/\/+$/,'')+'/.default'}); const r=await fetch(`https://login.microsoftonline.com/${process.env.TENANT_ID}/oauth2/v2.0/token`,{method:'POST',headers:{'Content-Type':'application/x-www-form-urlencoded'},body}); const j=await r.json(); if(!r.ok||!j.access_token) throw new Error('token '+r.status+' '+(j.error_description||j.error||'')); return j.access_token; }
async function sendDivision(out, context){
  const from=process.env.PRPO_MAIL_FROM||process.env.MAIL_FROM;
  if(!from) throw new Error('MAIL_FROM / PRPO_MAIL_FROM not set');
  const xlsxB64=await buildXlsxBase64(out.fil, out.cfg);
  const msg=buildDivisionMessage(out,xlsxB64);
  const token=await getToken('https://graph.microsoft.com');
  const r=await fetch('https://graph.microsoft.com/v1.0/users/'+encodeURIComponent(from)+'/sendMail',{method:'POST',headers:{Authorization:'Bearer '+token,'Content-Type':'application/json'},
    body:JSON.stringify({message:msg,saveToSentItems:true})});
  if(r.status!==202){ const j=await r.json().catch(()=>({})); throw new Error('sendMail '+r.status+' '+JSON.stringify(j.error||j).slice(0,300)); }
  if(context) context.log('division test-channel message for '+out.cfg.key+' -> '+WAQAS_ONLY_RECIPIENT);
  return {sent:true,to:WAQAS_ONLY_RECIPIENT,waqasOnly:true};
}

async function fetchXlsx(url){ const r=await fetch(url+(url.includes('?')?'&':'?')+'t='+Date.now()); if(!r.ok) throw new Error('fetch '+r.status+' '+url); return parseXlsx(Buffer.from(await r.arrayBuffer())); }
async function fetchJson(url){ const r=await fetch(url+(url.includes('?')?'&':'?')+'t='+Date.now(),{headers:{Accept:'application/json'}}); if(!r.ok) throw new Error('fetch '+r.status+' '+url); return r.json(); }
async function safeFetchJson(url){ try{return await fetchJson(url);}catch(e){return {}; } }
function dubaiPreparedAt(value){
  const d=new Date(value); if(isNaN(d))return '';
  const parts=new Intl.DateTimeFormat('en-GB',{timeZone:'Asia/Dubai',hour:'2-digit',minute:'2-digit',hour12:false,day:'numeric',month:'long'}).formatToParts(d);
  const get=t=>(parts.find(p=>p.type===t)||{}).value||'';
  return get('hour')+':'+get('minute')+' on '+get('day')+' '+get('month');
}
function freshnessWarning(workbookGeneratedAt, liveGeneratedAt){
  const workbook=new Date(workbookGeneratedAt), live=new Date(liveGeneratedAt);
  if(isNaN(workbook)||isNaN(live))return 'The workbook freshness could not be confirmed against the live feed.';
  if(live.getTime()-workbook.getTime()<=6*3600000)return '';
  return 'These figures were prepared at '+dubaiPreparedAt(workbookGeneratedAt)+' and may not include work raised since.';
}
function freshnessBanner(message){ return message?'<div style="border:1px solid #f59e0b;background:#fff7ed;color:#9a3412;padding:10px 13px;margin:0 0 12px;border-radius:8px;font:700 12px '+HF+';">'+esc(message)+'</div>':''; }
function provenanceSentence(items){ return (items&&items[0]&&items[0].provenanceSentence)||items.provenanceSentence||'These figures come from the Dynamics 365 F&O export supplied by IT.'; }
async function loadItems(){
  const live=await fetchJson(DATASET_URL);
  if(live.sourceState!=='LIVE'||!live.pr||!Array.isArray(live.pr.rows)||!live.po||!Array.isArray(live.po.rows))throw new Error('live F&O dataset is unavailable');
  const items=applyDeliveryPolicy(buildItems(live.pr.rows,live.po.rows));
  const authority=live.exportAuthority||{};
  items.freshnessWarning=authority.staleWarning||'';
  items.provenanceSentence=authority.provenanceSentence||'';
  items.exportDateUtc=authority.exportDateUtc||null;
  for(const item of items){
    item.freshnessWarning=items.freshnessWarning;
    item.provenanceSentence=items.provenanceSentence;
    item.exportDateUtc=items.exportDateUtc;
  }
  items.datasetRevision=live.revision;
  items.datasetGeneratedAt=live.generatedAt||null;
  items.liveDatasetGeneratedAt=live.generatedAt||null;
  items.sourceState='LIVE';
  return items;
}

/* ---- 3-day trend: the dashboard repo commits pr/po.xlsx daily, so git history IS the snapshot archive.
 * Fetch the latest commit on/before each of the last 2 Dubai days and rebuild items with the same logic. ---- */
const GH_HIST_REPO='Strive-Services-Group/PR-PO-Pipeline-Dashboard';
function dubaiYmd(d){ const x=new Date(d.getTime()+4*3600*1000); return x.getUTCFullYear()+'-'+String(x.getUTCMonth()+1).padStart(2,'0')+'-'+String(x.getUTCDate()).padStart(2,'0'); }
async function historyItems(){ return {}; }
const MON3=['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
function histDayList(){ const l=[]; for(let i=2;i>=1;i--){ const ymd=dubaiYmd(new Date(Date.now()-i*86400000)); const d=new Date(ymd+'T00:00:00Z'); l.push({ymd,lab:d.getUTCDate()+' '+MON3[d.getUTCMonth()]}); } return l; }
// History footer line inside a stage card: thin divider + "11 Aug 19 · 12 Aug 19" (muted, one line).
function histLine(tr){ if(!tr||!tr.length) return ''; return '<div style="border-top:1px solid rgba(20,49,94,0.10);margin:8px 0 0;padding:5px 0 0;font-family:'+FONT+';font-size:9.5px;font-weight:700;color:#8795a9;white-space:nowrap;">'+tr.map(r=>r.lab+' <b style="font-size:14px;color:#33415c;">'+r.n+'</b>').join(' &#160;&#183;&#160; ')+'</div>'; }
// ▲/▼ badge on today's number vs yesterday (green when the queue went DOWN).
function deltaBadge(tr,todayN){ if(!tr||!tr.length) return ''; const y=Number(tr[tr.length-1].n); if(!isFinite(y)) return ''; const d=todayN-y; if(!d) return ''; const up=d>0; return ' <span style="font-family:'+FONT+';font-size:11px;font-weight:800;color:'+(up?'#b91c1c':'#16794a')+';">'+(up?'&#9650;':'&#9660;')+Math.abs(d)+'</span>'; }
function trendCols(hist,countFn,todayN){
  const cols=[];
  for(let i=2;i>=1;i--){
    const ymd=dubaiYmd(new Date(Date.now()-i*86400000));
    const d=new Date(ymd+'T00:00:00Z'); const lab=d.getUTCDate()+' '+MON3[d.getUTCMonth()];
    let n='&#8211;'; if(hist&&hist[ymd]){ try{ n=String(countFn(hist[ymd])); }catch(e){} }
    cols.push({lab,n});
  }
  cols.push({lab:'Today',n:String(todayN),em:true});
  return cols;
}
function tcard(label,cols,accent,foot){
  const cells=cols.map((c,i)=>'<td width="'+Math.floor(100/cols.length)+'%" align="center" valign="middle" style="padding:3px 4px;'+(i>0?'border-left:1px solid #e8edf4;':'')+'">'
    +'<div style="font-family:'+FONT+';font-size:8.5px;font-weight:800;color:#8795a9;text-transform:uppercase;white-space:nowrap;">'+c.lab+'</div>'
    +'<div style="font-family:'+FONT+';font-size:'+(c.em?'24':'22')+'px;font-weight:800;color:'+(c.em?accent:'#33415c')+';white-space:nowrap;margin-top:2px;">'+c.n+(c.em?deltaBadge(cols.slice(0,-1),Number(c.n)):'')+'</div>'
    +'</td>').join('');
  return '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="width:100%;background:#ffffff;border:1px solid #e8ecf2;border-top:3px solid '+accent+';border-radius:10px;box-shadow:0 1px 3px rgba(16,24,40,0.08);"><tr><td height="92" valign="middle" style="padding:10px 12px;height:92px;"><div style="font-family:'+FONT+';font-size:9.5px;font-weight:800;color:#5b6b7f;text-transform:uppercase;">'+label+'</div><table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="margin-top:4px;"><tr>'+cells+'</tr></table>'+(foot?'<div style="border-top:1px solid #eef2f7;margin-top:6px;padding-top:5px;text-align:center;font-family:'+FONT+';font-size:9.5px;font-weight:700;color:'+TEAL+';white-space:nowrap;">'+foot+'</div>':'')+'</td></tr></table>';
}

/* ---- triggers ---- */
// NCRONTAB runs in UTC: 06:00 UTC = 10:00 AM Dubai. 1-5 = Monday-Friday (no weekend sends). Do NOT put the Dubai hour here.
app.timer('prpo-email-daily', { schedule:'0 0 6 * * 1-5', handler:async(timer,context)=>{
  // Guard: Azure re-fires a "missed" timer after deploys/restarts. Only send 10:00–10:30 Dubai, Mon–Fri.
  const dxb=new Date(Date.now()+4*3600*1000); const hh=dxb.getUTCHours(), mm=dxb.getUTCMinutes(), dow=dxb.getUTCDay();
  if(dow===0||dow===6){ context.log('prpo-email-daily: weekend ('+['Sun','Mon','Tue','Wed','Thu','Fri','Sat'][dow]+') — skipped'); return; }
  if(hh!==10||mm>30){ context.log('prpo-email-daily: off-schedule fire at '+hh+':'+String(mm).padStart(2,'0')+' Dubai (restart catch-up) — skipped'); return; }
  const items=await loadItems();
  const hist=await historyItems();
  for(const cfg of DIVS){ if(cfg.send===false) continue; try{ await sendDivision(buildDivision(cfg,items,hist),context); }catch(e){ context.error('prpo '+cfg.key+' FAILED: '+e.message); } }
  for(const p of groupByOwner(personalPool(items))){ try{ await sendPersonal(buildPersonal(p,hist),context); }catch(e){ context.error('prpo personal '+p.user+' FAILED: '+e.message); } }
}});
app.http('prpo-email', { methods:['GET','OPTIONS'], authLevel:'function', route:'prpo-email', handler:async(request,context)=>{
  try{
    const url=new URL(request.url); const dk=url.searchParams.get('division'); const pk=url.searchParams.get('person');
    const wantSend=url.searchParams.get('send')==='1'; const sendAll=wantSend&&!dk&&!pk&&url.searchParams.get('personal')!=='1';
    const items=await loadItems();
    const datasetMeta={datasetRevision:items.datasetRevision,datasetGeneratedAt:items.datasetGeneratedAt,sourceState:items.sourceState};
    const hist=await historyItems();
    if(dk){ const cfg=DIVS.find(d=>d.key===dk); if(!cfg) return {status:400,jsonBody:{error:'unknown division; use procurement|invoicing|ops_hm|ops_all'}};
      const out=buildDivision(cfg,items,hist);
      if(url.searchParams.get('format')==='html') return {status:200,headers:{'Content-Type':'text/html; charset=utf-8'},body:out.html};
      if(url.searchParams.get('debug')==='1') return {status:200,jsonBody:{...datasetMeta,division:dk,count:out.count,value:Math.round(out.value),retired:cfg.send===false}};
      if(wantSend){ if(cfg.send===false) return {status:400,jsonBody:{division:dk,sent:false,reason:'retired — replaced by personal emails'}}; const s=await sendDivision(out,context); return {status:200,jsonBody:{division:dk,...s}}; }
      return {status:200,jsonBody:{...datasetMeta,division:dk,count:out.count,value:Math.round(out.value),retired:cfg.send===false}};
    }
    if(pk){ const p=groupByOwner(personalPool(items)).find(e=>e.key===_norm(pk));
      if(!p) return {status:404,jsonBody:{error:'no pending items for this user',user:pk,people:groupByOwner(personalPool(items)).map(e=>e.user)}};
      const out=buildPersonal(p,hist);
      if(url.searchParams.get('format')==='html') return {status:200,headers:{'Content-Type':'text/html; charset=utf-8'},body:out.html};
      if(wantSend){ const s=await sendPersonal(out,context); return {status:200,jsonBody:s}; }
      return {status:200,jsonBody:{...datasetMeta,user:out.user,deliveryAddress:WAQAS_ONLY_RECIPIENT,count:out.count,value:Math.round(out.value)}};
    }
    if(url.searchParams.get('personal')==='1'){ const people=[];
      for(const p of groupByOwner(personalPool(items))){ const out=buildPersonal(p,hist); let s={sent:false}; if(wantSend) s=await sendPersonal(out,context); people.push({user:p.user,deliveryAddress:WAQAS_ONLY_RECIPIENT,count:out.count,value:Math.round(out.value),...s}); }
      return {status:200,jsonBody:{...datasetMeta,waqasOnly:true,sent:wantSend,people}};
    }
    const summary=[]; for(const cfg of DIVS){ if(cfg.send===false) continue; const out=buildDivision(cfg,items,hist); let s={sent:false}; if(sendAll) s=await sendDivision(out,context); summary.push({division:cfg.key,count:out.count,value:Math.round(out.value),...s}); }
    const people=[]; for(const p of groupByOwner(personalPool(items))){ const out=buildPersonal(p,hist); let s={sent:false}; if(sendAll) s=await sendPersonal(out,context); people.push({user:p.user,count:out.count,value:Math.round(out.value),...s}); }
    return {status:200,jsonBody:{...datasetMeta,sentAll:sendAll,waqasOnly:true,divisions:summary,people}};
  }catch(e){ context.error('prpo-email failed:',e); return {status:500,jsonBody:{error:e.message}}; }
}});

module.exports = { buildItems, buildDivision, buildXlsxBase64, parseXlsx, DIVS, personalPool, groupByOwner, buildPersonal, historyItems, loadItems, applyDeliveryPolicy, freshnessWarning, fnoOwnerNames, buildPersonalMessage, buildDivisionMessage, WAQAS_ONLY_RECIPIENT };
