/*
 * PR / PO PIPELINE EMAIL — self-contained daily report, sent 10:00 AM Dubai.
 *
 * TIMER: 0 0 6 * * *  (06:00 UTC = 10:00 AM Asia/Dubai), every day.
 *   1. Fetches pr.xlsx + po.xlsx from the published PR-PO dashboard (GitHub Pages).
 *   2. Computes the pipeline buckets, per-department breakdown and 7-day inflow using
 *      the SAME live-pipeline logic as the dashboard (buildPRRecords / livePipelineFilter /
 *      renderBucketKPIs), so the numbers reconcile 1:1 with the dashboard header.
 *   3. Builds an Outlook-safe HTML report and sends it via Microsoft Graph from CK's mailbox.
 *
 * HTTP test endpoints (authLevel: function, call with ?code=...):
 *   GET /api/prpo-email?code=...&format=html  -> preview the email in the browser (no send)
 *   GET /api/prpo-email?code=...&debug=1       -> JSON of the reconciliation totals (no send)
 *   GET /api/prpo-email?code=...&send=1        -> build AND send now
 *
 * Required Function App settings (portal -> Environment variables) — mostly already set for the telemetry email:
 *   TENANT_ID, CLIENT_ID, CLIENT_SECRET     (existing)
 *   MAIL_FROM                                (existing; e.g. Chandan.kumar@striveservicesgroup.com)
 *   PRPO_MAIL_TO   = semicolon-separated recipients   (NEW — falls back to MAIL_TO if unset)
 * Optional overrides:
 *   PRPO_PR_URL / PRPO_PO_URL  (defaults to the GitHub Pages pr.xlsx / po.xlsx)
 *   PRPO_MAIL_FROM             (defaults to MAIL_FROM)
 *   PRPO_DASH_URL              (dashboard link in the email)
 */
const { app } = require('@azure/functions');
const XLSX = require('xlsx');

const PR_URL = process.env.PRPO_PR_URL || 'https://strive-services-group.github.io/PR-PO-Pipeline-Dashboard/pr.xlsx';
const PO_URL = process.env.PRPO_PO_URL || 'https://strive-services-group.github.io/PR-PO-Pipeline-Dashboard/po.xlsx';
const DASH   = process.env.PRPO_DASH_URL || 'https://strive-services-group.github.io/PR-PO-Pipeline-Dashboard/';
const FONT = 'Aptos,Segoe UI,Arial,sans-serif', NAVY = '#14315E', RED = '#dc2626', TEAL = '#0f766e', W = 1080;

/* ===== step -> bucket maps (copied 1:1 from the dashboard index.html) ===== */
const PR_MAP = {
  "Handyman Services_Manager":"Dep Managers","Building Services_Asst. Facility Managers 1":"Dep Managers",
  "PurchReqReviewTask":"PR In Review","Procurement sends inquiry/RFQ to suppliers":"RFQ to suppliers",
  "Quotation received and logged/attached":"Qt received & Logged","Quotation shared to Operations for confirmation":"Qt Shared to Op",
  "Operations confirms material/scope":"OP confirms material","Unit prices updated in PR lines":"Unit Price Updated",
  "Building Services_Asst. Facility Managers 2":"Dep Managers","Building Services_Facilities Manager":"Dep Managers",
  "PAC Services_Manager":"Dep Managers","Concierge Services_Manager":"Dep Managers","Security Services_Manager":"Dep Managers",
  "Home Services_Operations Manager":"Dep Managers","Landscaping_Manager":"Dep Managers",
  "Finance & Accounts_Accounting Manager":"Finance","Facilities Management_Director":"Director",
  "Commercial_Director":"Director","Executive Management_CEO":"CEO"
};
const PROC = new Set(["PR In Review","RFQ to suppliers","Qt received & Logged","OP confirms material","Procurement (in process)"]);
const OPS  = new Set(["Qt Shared to Op","Unit Price Updated"]);
const PO_MAP = {
  "Advance payment request submitted (if applicable)":"Procurement","Procurement Manager":"Procurement",
  "Accounting Manager":"Finance","Finance and Accounts Director":"Director","CEO":"CEO","LPO sent/shared with supplier":"Sent to Supplier"
};
const PR_ORDER = ['Procurement','Operations to Confirm','Dep Managers','Finance','Director','CEO'];
const PO_ORDER = ['Procurement','Finance','Director','CEO','Sent to Supplier','Pending Invoicing'];
const COLOR = {'Procurement':'#3b82f6','Operations to Confirm':'#14b8a6','Dep Managers':'#8b5cf6','Finance':'#22c55e','Director':'#ec4899','CEO':'#f59e0b','Sent to Supplier':'#a855f7','Pending Invoicing':'#f97316'};
const GRAD  = {'Procurement':'#eff5ff','Operations to Confirm':'#ebfbf7','Dep Managers':'#f4f1fe','Finance':'#eefbf3','Director':'#fdeff7','CEO':'#fff9ec','Sent to Supplier':'#f9f2ff','Pending Invoicing':'#fff4e8'};
const DEPARTMENTS = ["Building Services","Contracted Cleaning Services","Security Services","Landscaping Services","Concierge Services","FitOut Services","Home Maintenance Services","Others"];
const DICON = {'Building Services':'&#127970;','Contracted Cleaning Services':'&#129529;','Security Services':'&#128737;&#65039;','Landscaping Services':'&#127795;','Concierge Services':'&#128717;&#65039;','FitOut Services':'&#127959;&#65039;','Home Maintenance Services':'&#128736;&#65039;','Others':'&#128230;'};
const DCOL  = {'Building Services':'#135573','Contracted Cleaning Services':'#4FB0E5','Security Services':'#ED8624','Landscaping Services':'#3CC24A','Concierge Services':'#B13BB4','FitOut Services':'#DC2626','Home Maintenance Services':'#E63B2E','Others':'#1F6FB5'};
const TC = {'PR':['#2563eb','#3b82f6'],'CPR':['#7c3aed','#9333ea'],'PO':['#0891b2','#06b6d4']};

/* ===== helpers ===== */
function esc(s){ return String(s==null?'':s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;'); }
function money(v){ const n=Number(v); return (isFinite(n)?n:0).toLocaleString('en-US',{maximumFractionDigits:0}); }
function amt(r){ const v=Number(r['Total amount']); return isFinite(v)?v:0; }
// Excel serial (or Date/ISO string) -> UTC Date at the true calendar date/time (timezone-independent).
function xd(v){ if(v instanceof Date) return isNaN(v)?null:v; if(typeof v==='number' && isFinite(v)){ const o=XLSX.SSF.parse_date_code(v); if(!o||!o.y) return null; return new Date(Date.UTC(o.y,o.m-1,o.d,o.H||0,o.M||0,Math.floor(o.S||0))); } if(typeof v==='string' && v){ const d=new Date(v); return isNaN(d)?null:d; } return null; }
function ageDays(v){ const d=xd(v); return d? Math.max(0, Math.floor((Date.now()-d.getTime())/86400000)) : null; }
function ymd(v){ const d=xd(v); if(!d) return null; const p=n=>String(n).padStart(2,'0'); return d.getUTCFullYear()+'-'+p(d.getUTCMonth()+1)+'-'+p(d.getUTCDate()); }
function pref(r){ return String(r['Purchase requisition']||'').startsWith('PR-')?'PR':'CPR'; }

function parseXlsx(buf){
  const wb=XLSX.read(buf,{type:'buffer'}); // raw serials for dates -> converted via xd() (tz-independent)
  const ws=wb.Sheets[wb.SheetNames[0]];
  const aoa=XLSX.utils.sheet_to_json(ws,{header:1,defval:null,raw:true});
  if(!aoa.length) return [];
  const hdr=aoa[0].map(x=>String(x));
  const rows=[];
  for(let i=1;i<aoa.length;i++){ const o={}; for(let j=0;j<hdr.length;j++) o[hdr[j]]=aoa[i][j]; rows.push(o); }
  return rows;
}

/* live-pipeline predicates — identical to dashboard livePipelineFilter */
function prLive(r){ const st=String(r['Status']||''); return !!PR_MAP[r['Step name']] && st.toLowerCase()!=='closed' && st!=='Rejected' && st!=='Cancelled'; }
function poBucket(r){
  const step=r['Step name'], appr=String(r['Approval status']||''), pos=String(r['Purchase order status']||'');
  let b = PO_MAP[step] || null;
  if(appr==='Confirmed' && pos==='Received') b='Pending Invoicing';
  if(appr==='Rejected' || pos==='Canceled' || pos==='Invoiced' || !b) return null;
  return b;
}
function prAge(r){ const sd=r['Step date and time']; return ageDays(sd!=null? sd : r['Created date']); }
function poAge(r){ const sd=r['Step date and time']; return ageDays(sd!=null? sd : (r['Created date and time']!=null? r['Created date and time'] : r['Requested receipt date'])); }
function prHb(b){ return PROC.has(b)?'Procurement':(OPS.has(b)?'Operations to Confirm':b); }
function nd(d){ if(d==null||String(d).trim()==='') return 'Others'; d=String(d).trim(); return DEPARTMENTS.includes(d)?d:'Others'; }

/* ===== build the email ===== */
function buildEmail(prRows, poRows){
  const dubai=new Date(Date.now()+4*3600*1000);
  const WIN=[]; for(let i=0;i<7;i++){ const d=new Date(dubai); d.setUTCDate(d.getUTCDate()-i); const p=n=>String(n).padStart(2,'0'); WIN.push(d.getUTCFullYear()+'-'+p(d.getUTCMonth()+1)+'-'+p(d.getUTCDate())); }
  const WINSET=new Set(WIN);
  const stamp=dubai.toLocaleDateString('en-GB',{day:'2-digit',month:'short',year:'numeric',timeZone:'UTC'});

  const mk=order=>{ const o={}; order.forEach(b=>o[b]={n:0,sum:0,c:0,amt:0}); return o; };
  const pragg=mk(PR_ORDER), poagg=mk(PO_ORDER);
  for(const r of prRows){ if(!prLive(r)) continue; const g=prHb(PR_MAP[r['Step name']]); if(!pragg[g]) continue; const a=prAge(r); pragg[g].n++; if(a!=null){pragg[g].sum+=a;pragg[g].c++;} pragg[g].amt+=amt(r); }
  for(const r of poRows){ const b=poBucket(r); if(!b||!poagg[b]) continue; const a=poAge(r); poagg[b].n++; if(a!=null){poagg[b].sum+=a;poagg[b].c++;} poagg[b].amt+=amt(r); }

  const card=(b,agg)=>{ const x=agg[b], avg=x.c?(x.sum/x.c):0;
    return '<td width="16.6%" valign="top" style="padding:0 5px;">'
      +'<div style="background:'+GRAD[b]+';background-image:linear-gradient(158deg,#ffffff 0%,'+GRAD[b]+' 100%);border:1px solid #e8ecf2;border-top:3px solid '+COLOR[b]+';border-radius:12px;padding:13px 14px;box-shadow:0 1px 3px rgba(15,23,42,.05);">'
      +'<div style="font-family:'+FONT+';font-size:10px;font-weight:800;letter-spacing:.3px;color:#5b6b7f;text-transform:uppercase;line-height:1.2;min-height:25px;">'+esc(b)+'</div>'
      +'<div style="margin:6px 0 3px;white-space:nowrap;"><span style="font-family:'+FONT+';font-size:28px;font-weight:800;color:'+COLOR[b]+';">'+x.n+'</span>'
      +'<span style="font-family:'+FONT+';font-size:13px;font-weight:800;color:'+RED+';"> ('+avg.toFixed(1)+'d)</span></div>'
      +'<div style="font-family:'+FONT+';font-size:11.5px;font-weight:700;color:'+TEAL+';">AED '+money(x.amt)+'</div></div></td>'; };
  const cardrow=(order,agg)=>'<table cellpadding="0" cellspacing="0" border="0" width="'+W+'" style="width:'+W+'px;border-collapse:separate;margin:2px 0 4px;"><tr>'+order.map(b=>card(b,agg)).join('')+'</tr></table>';
  const pill=(t,c1,c2)=>'<span style="display:inline-block;background:'+c1+';background-image:linear-gradient(90deg,'+c1+','+c2+');color:#fff;font-family:'+FONT+';font-size:11px;font-weight:800;letter-spacing:.5px;text-transform:uppercase;padding:5px 14px;border-radius:20px;box-shadow:0 1px 3px rgba(15,23,42,.15);">'+t+'</span>';
  const hcell=(w,base,light,txt,al,pad)=>'<th style="width:'+w+'px;color:#fff;font-family:'+FONT+';font-weight:800;font-size:11.5px;padding:'+(pad||'11px 8px')+';text-align:'+(al||'center')+';border:0;background:'+base+';background-image:linear-gradient(180deg,'+light+','+base+');">'+txt+'</th>';
  const wrapTbl=(w,inner)=>'<div style="border:1px solid #d9e1ea;border-radius:13px;overflow:hidden;box-shadow:0 2px 6px rgba(15,23,42,.07);"><table cellpadding="0" cellspacing="0" border="0" width="'+w+'" style="border-collapse:collapse;table-layout:fixed;width:'+w+'px;background:#fff;">'+inner+'</table></div>';
  const nc=(td,n,col)=> n?('<td style="'+td+'font-weight:700;color:'+col+';">'+n+'</td>'):('<td style="'+td+'color:#cbd5e1;">0</td>');
  const SH=(t,s)=>'<div style="font-family:'+FONT+';font-weight:800;font-size:15px;color:'+NAVY+';margin:0 0 2px;">'+t+'</div>'+(s?'<div style="font-family:'+FONT+';font-size:11.5px;color:#7688a0;margin:0 0 9px;">'+s+'</div>':'');

  /* department table (live pipeline) */
  const byd={}; DEPARTMENTS.forEach(d=>byd[d]={PR:0,CPR:0,PO:0,val:0,ages:[]});
  for(const r of prRows){ if(!prLive(r)) continue; const e=byd[nd(r['Department'])]; e[pref(r)]++; e.val+=amt(r); const a=prAge(r); if(a!=null) e.ages.push(a); }
  for(const r of poRows){ if(poBucket(r)==null) continue; const e=byd[nd(r['Department'])]; e.PO++; e.val+=amt(r); const a=poAge(r); if(a!=null) e.ages.push(a); }
  const Wd=[210,52,56,52,90,140];
  let dHead='<tr>'+hcell(Wd[0],NAVY,'#274a72','Department','left','11px 12px')+hcell(Wd[1],TC.PR[0],TC.PR[1],'PR')+hcell(Wd[2],TC.CPR[0],TC.CPR[1],'CPR')+hcell(Wd[3],TC.PO[0],TC.PO[1],'PO')+hcell(Wd[4],'#334155','#475569','Total')+hcell(Wd[5],'#334155','#475569','Value (AED)','right','11px 12px')+'</tr>';
  let dBody='', tPR=0,tCPR=0,tPO=0,tV=0,tN=0,allages=[], i=0;
  for(const d of DEPARTMENTS){ const e=byd[d], t=e.PR+e.CPR+e.PO; if(!t) continue;
    const avg=e.ages.length?e.ages.reduce((a,b)=>a+b,0)/e.ages.length:0;
    tPR+=e.PR;tCPR+=e.CPR;tPO+=e.PO;tV+=e.val;tN+=t;allages=allages.concat(e.ages);
    const bg=(i%2===0)?'#ffffff':'#f6f9fc'; i++;
    const tdc='padding:9px 6px;border-bottom:1px solid #eef1f6;font-family:'+FONT+';font-size:12px;text-align:center;vertical-align:middle;background:'+bg+';';
    const dn='padding:8px 10px 8px 12px;border-bottom:1px solid #eef1f6;border-left:4px solid '+DCOL[d]+';font-family:'+FONT+';font-size:12px;font-weight:700;color:'+NAVY+';text-align:left;vertical-align:middle;line-height:1.2;background:'+bg+';';
    const logo='<span style="display:inline-block;width:22px;height:22px;line-height:22px;text-align:center;background:'+DCOL[d]+'22;border-radius:6px;font-size:13px;margin-right:7px;vertical-align:middle;">'+DICON[d]+'</span>';
    dBody+='<tr><td style="'+dn+'">'+logo+'<span style="vertical-align:middle;">'+esc(d)+'</span></td>'
      +nc(tdc,e.PR,TC.PR[0])+nc(tdc,e.CPR,TC.CPR[0])+nc(tdc,e.PO,TC.PO[0])
      +'<td style="'+tdc+'"><span style="font-weight:800;color:'+NAVY+';font-size:13px;">'+t+'</span> <span style="font-size:9.5px;color:#93a1b3;font-weight:700;">'+avg.toFixed(0)+'d</span></td>'
      +'<td style="'+tdc+'text-align:right;padding-right:12px;font-weight:700;color:'+TEAL+';">'+money(e.val)+'</td></tr>';
  }
  const tavg=allages.length?allages.reduce((a,b)=>a+b,0)/allages.length:0;
  const tf='padding:10px 6px;border-top:2px solid #cdd7e2;font-family:'+FONT+';font-size:12px;text-align:center;background:#eef2f7;font-weight:800;color:'+NAVY+';';
  dBody+='<tr><td style="'+tf+'text-align:left;padding-left:12px;">TOTAL</td><td style="'+tf+'">'+tPR+'</td><td style="'+tf+'">'+tCPR+'</td><td style="'+tf+'">'+tPO+'</td>'
    +'<td style="'+tf+'">'+tN+' <span style="font-size:9.5px;color:#64748b;">'+tavg.toFixed(0)+'d</span></td><td style="'+tf+'text-align:right;padding-right:12px;color:'+TEAL+';">'+money(tV)+'</td></tr>';
  const deptTable=wrapTbl(600,dHead+dBody);

  /* daily inflow (live pipeline, last 7 days created) */
  const d7=[];
  for(const r of prRows){ if(prLive(r) && WINSET.has(ymd(r['Created date']))) d7.push([pref(r), ymd(r['Created date']), amt(r)]); }
  for(const r of poRows){ if(poBucket(r)!=null && WINSET.has(ymd(r['Created date and time']))) d7.push(['PO', ymd(r['Created date and time']), amt(r)]); }
  const Wy=[130,52,56,52,66,104];
  let yHead='<tr>'+hcell(Wy[0],NAVY,'#274a72','Date','left','11px 12px')+hcell(Wy[1],TC.PR[0],TC.PR[1],'PR')+hcell(Wy[2],TC.CPR[0],TC.CPR[1],'CPR')+hcell(Wy[3],TC.PO[0],TC.PO[1],'PO')+hcell(Wy[4],'#334155','#475569','Total')+hcell(Wy[5],'#334155','#475569','Value (AED)','right','11px 12px')+'</tr>';
  let yBody='', yPR=0,yCPR=0,yPO=0,yV=0,yT=0;
  const asc=WIN.slice().sort(); const todayStr=WIN[0];
  for(let k=0;k<asc.length;k++){ const day=asc[k]; const rr=d7.filter(x=>x[1]===day);
    const p=rr.filter(x=>x[0]==='PR').length, c=rr.filter(x=>x[0]==='CPR').length, o=rr.filter(x=>x[0]==='PO').length;
    const v=rr.reduce((a,b)=>a+b[2],0), t=p+c+o; yPR+=p;yCPR+=c;yPO+=o;yV+=v;yT+=t;
    const isToday=day===todayStr, bg=isToday?'#fffbeb':((k%2===0)?'#ffffff':'#f6f9fc');
    const tdc='padding:9px 6px;border-bottom:1px solid #eef1f6;font-family:'+FONT+';font-size:12px;text-align:center;background:'+bg+';';
    const dn='padding:9px 12px;border-bottom:1px solid #eef1f6;font-family:'+FONT+';font-size:12px;font-weight:700;color:'+(isToday?'#92400e':'#475569')+';text-align:left;background:'+bg+';';
    const lbl=new Date(day+'T00:00:00Z').toLocaleDateString('en-GB',{weekday:'short',day:'2-digit',month:'short',timeZone:'UTC'});
    yBody+='<tr><td style="'+dn+'">'+lbl+'</td>'+nc(tdc,p,TC.PR[0])+nc(tdc,c,TC.CPR[0])+nc(tdc,o,TC.PO[0])
      +'<td style="'+tdc+'font-weight:800;color:'+NAVY+';">'+t+'</td><td style="'+tdc+'text-align:right;padding-right:12px;font-weight:700;color:'+TEAL+';">'+money(v)+'</td></tr>';
  }
  const yf='padding:10px 6px;border-top:2px solid #cdd7e2;font-family:'+FONT+';font-size:12px;text-align:center;background:#eef2f7;font-weight:800;color:'+NAVY+';';
  yBody+='<tr><td style="'+yf+'text-align:left;padding-left:12px;">TOTAL</td><td style="'+yf+'">'+yPR+'</td><td style="'+yf+'">'+yCPR+'</td><td style="'+yf+'">'+yPO+'</td><td style="'+yf+'">'+yT+'</td><td style="'+yf+'text-align:right;padding-right:12px;color:'+TEAL+';">'+money(yV)+'</td></tr>';
  const dailyTable=wrapTbl(460,yHead+yBody);

  const seg=['#3b82f6','#14b8a6','#8b5cf6','#22c55e','#ec4899','#f59e0b','#a855f7','#f97316'].map(c=>'<td width="135" style="height:5px;background:'+c+';font-size:0;line-height:0;">&nbsp;</td>').join('');
  const accent='<table cellpadding="0" cellspacing="0" border="0" width="'+W+'" style="width:'+W+'px;border-collapse:collapse;margin:0 0 16px;border-radius:3px;overflow:hidden;"><tr>'+seg+'</tr></table>';
  const snapshot='<div style="font-family:'+FONT+';font-size:12.5px;color:#607083;margin:2px 0 14px;">Ageing snapshot &#183; '+stamp+' &#183; open = PR/CPR in Draft or In review, PO in Open order &#183; <a href="'+DASH+'" style="color:#145A95;font-weight:700;text-decoration:none;">Open the Live Dashboard</a> for drill-through</div>';
  const twotables='<table cellpadding="0" cellspacing="0" border="0" width="'+W+'" style="width:'+W+'px;table-layout:fixed;border-collapse:collapse;"><tr>'
    +'<td width="620" valign="top" style="padding-right:20px;">'+SH('&#127970; By Department &#8212; All Open','Live-pipeline PR/CPR &amp; PO per department. Grey = avg days at stage. Reconciles to the dashboard.')+deptTable+'</td>'
    +'<td width="460" valign="top">'+SH('&#128197; Daily Inflow &#8212; Last 7 Days Created (Open)','New live-pipeline items by the day they were created.')+dailyTable+'</td></tr></table>';

  const subject='PR / PO Pipeline — Open & Pending ('+stamp+')';
  const html='<div style="font-family:'+FONT+';color:#22303c;width:'+W+'px;">'
    +'<table cellpadding="0" cellspacing="0" border="0" width="'+W+'" style="width:'+W+'px;"><tr>'
    +'<td style="font-family:'+FONT+';font-weight:800;font-size:20px;color:'+NAVY+';">PR / PO Pipeline &#8212; Open &amp; Pending</td>'
    +'<td align="right" style="font-family:'+FONT+';font-size:12px;color:#607083;">'+stamp+' &#183; <a href="'+DASH+'" style="color:#145A95;font-weight:700;text-decoration:none;">Open Dashboard &#8599;</a></td></tr></table>'+accent
    +SH('&#128202; Bucket-wise Counts &#8212; Pipeline Stages','')
    +'<div style="margin:2px 0 6px;">'+pill('PR / CPR Pipeline','#1D4ED8','#3b82f6')+'</div>'+cardrow(PR_ORDER,pragg)
    +'<div style="margin:10px 0 6px;">'+pill('PO Pipeline','#0891b2','#06b6d4')+'</div>'+cardrow(PO_ORDER,poagg)
    +snapshot+'<div style="height:6px;line-height:6px;font-size:0;">&nbsp;</div>'+twotables
    +'<div style="font-family:'+FONT+';font-size:10px;color:#8b98a5;margin-top:14px;">Automated daily report &#183; Strive Services Group &#183; Dynamics 365 PR/PO exports &#183; counts use the dashboard live-pipeline logic</div></div>';

  const debug={ pr:Object.fromEntries(PR_ORDER.map(b=>[b,pragg[b].n])), po:Object.fromEntries(PO_ORDER.map(b=>[b,poagg[b].n])),
    prSum:PR_ORDER.reduce((a,b)=>a+pragg[b].n,0), poSum:PO_ORDER.reduce((a,b)=>a+poagg[b].n,0),
    dept:{PR:tPR,CPR:tCPR,PO:tPO,total:tN}, daily7:yT };
  return { subject, html, debug };
}

/* ===== auth + send (mirrors telemetryEmail.js) ===== */
async function getToken(scopeBase){
  const body=new URLSearchParams({ client_id:process.env.CLIENT_ID, client_secret:process.env.CLIENT_SECRET, grant_type:'client_credentials', scope:scopeBase.replace(/\/+$/,'')+'/.default' });
  const r=await fetch(`https://login.microsoftonline.com/${process.env.TENANT_ID}/oauth2/v2.0/token`,{ method:'POST', headers:{'Content-Type':'application/x-www-form-urlencoded'}, body });
  const j=await r.json();
  if(!r.ok||!j.access_token) throw new Error('token '+r.status+' '+(j.error_description||j.error||''));
  return j.access_token;
}
async function sendMail(subject, html, context){
  const from=process.env.PRPO_MAIL_FROM||process.env.MAIL_FROM;
  const toList=(process.env.PRPO_MAIL_TO||process.env.MAIL_TO||'').split(/[;,]/).map(s=>s.trim()).filter(Boolean);
  if(!from) throw new Error('MAIL_FROM / PRPO_MAIL_FROM not set');
  if(!toList.length) throw new Error('PRPO_MAIL_TO / MAIL_TO not set');
  const token=await getToken('https://graph.microsoft.com');
  const r=await fetch('https://graph.microsoft.com/v1.0/users/'+encodeURIComponent(from)+'/sendMail',{ method:'POST', headers:{ Authorization:'Bearer '+token, 'Content-Type':'application/json' },
    body:JSON.stringify({ message:{ subject, body:{contentType:'HTML',content:html}, toRecipients:toList.map(a=>({emailAddress:{address:a}})) }, saveToSentItems:true }) });
  if(r.status!==202){ const j=await r.json().catch(()=>({})); throw new Error('sendMail '+r.status+' '+JSON.stringify(j.error||j).slice(0,300)); }
  if(context) context.log('PR/PO email sent to', toList.length, 'recipients');
}

async function fetchXlsx(url){
  const r=await fetch(url+(url.includes('?')?'&':'?')+'t='+Date.now());
  if(!r.ok) throw new Error('fetch '+r.status+' '+url);
  return parseXlsx(Buffer.from(await r.arrayBuffer()));
}
async function run(context, doSend){
  const [prRows, poRows]=await Promise.all([fetchXlsx(PR_URL), fetchXlsx(PO_URL)]);
  const out=buildEmail(prRows, poRows);
  if(doSend) await sendMail(out.subject, out.html, context);
  return out;
}

/* Daily at 06:00 UTC = 10:00 AM Dubai */
app.timer('prpo-email-daily', { schedule:'0 0 6 * * *', handler:async(timer,context)=>{
  try{ await run(context,true); } catch(e){ context.error('prpo email FAILED:', e.message); throw e; }
}});
app.http('prpo-email', { methods:['GET','OPTIONS'], authLevel:'function', route:'prpo-email', handler:async(request,context)=>{
  try{
    const url=new URL(request.url);
    const doSend=url.searchParams.get('send')==='1';
    const out=await run(context, doSend);
    if(url.searchParams.get('format')==='html' && !doSend) return { status:200, headers:{'Content-Type':'text/html; charset=utf-8'}, body:out.html };
    if(url.searchParams.get('debug')==='1') return { status:200, jsonBody:out.debug };
    return { status:200, jsonBody:{ sent:doSend, subject:out.subject, ...out.debug } };
  }catch(e){ context.error('prpo-email failed:', e); return { status:500, jsonBody:{ error:e.message } }; }
}});

module.exports = { buildEmail, parseXlsx };
