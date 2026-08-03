/*
 * PR / PO PIPELINE — FOUR DIVISION EMAILS, sent daily 10:00 AM Dubai.
 *
 * TIMER 0 0 6 * * * (06:00 UTC = 10:00 AM Dubai) sends four separate emails, each with a
 * scoped analysis and an Excel line-item list attached:
 *   1. Procurement          -> PRPO_PROC_MAIL_TO     (PR Procurement + PO Procurement + PO Sent to Supplier)
 *   2. Finance              -> PRPO_FIN_MAIL_TO      (PR Finance/Director + PO Finance/Pending Invoicing)
 *   3. Operations HM+FitOut -> PRPO_OPSHM_MAIL_TO    (Ops to Confirm + Dep Managers, Home Maintenance + FitOut)
 *   4. Operations All       -> PRPO_OPSALL_MAIL_TO   (Ops to Confirm + Dep Managers, all other departments)
 * A division with no recipient env set is skipped. All counts use the dashboard live-pipeline logic.
 *
 * HTTP test endpoints (authLevel: function, add &code=<your default host key>):
 *   ?division=procurement|finance|ops_hm|ops_all & format=html  -> preview that email (no send)
 *   ?division=... & debug=1                                      -> JSON counts for that division
 *   ?division=... & send=1                                       -> send just that division now
 *   ?send=1                                                      -> send ALL four now
 *   (no params)                                                  -> JSON summary of all four (no send)
 *
 * Env: TENANT_ID, CLIENT_ID, CLIENT_SECRET, MAIL_FROM (existing) + the four PRPO_*_MAIL_TO above.
 * Optional: PRPO_MAIL_FROM, PRPO_PR_URL, PRPO_PO_URL, PRPO_DASH_URL.
 */
const { app } = require('@azure/functions');
const XLSX = require('xlsx');

const PR_URL = process.env.PRPO_PR_URL || 'https://strive-services-group.github.io/PR-PO-Pipeline-Dashboard/pr.xlsx';
const PO_URL = process.env.PRPO_PO_URL || 'https://strive-services-group.github.io/PR-PO-Pipeline-Dashboard/po.xlsx';
const DASH   = process.env.PRPO_DASH_URL || 'https://strive-services-group.github.io/PR-PO-Pipeline-Dashboard/';
const FONT = 'Aptos,Segoe UI,Arial,sans-serif', NAVY = '#14315E', RED = '#dc2626', TEAL = '#0f766e', W = 1000;

const PR_MAP = {"Handyman Services_Manager":"Dep Managers","Building Services_Asst. Facility Managers 1":"Dep Managers","PurchReqReviewTask":"PR In Review","Procurement sends inquiry/RFQ to suppliers":"RFQ to suppliers","Quotation received and logged/attached":"Qt received & Logged","Quotation shared to Operations for confirmation":"Qt Shared to Op","Operations confirms material/scope":"OP confirms material","Unit prices updated in PR lines":"Unit Price Updated","Building Services_Asst. Facility Managers 2":"Dep Managers","Building Services_Facilities Manager":"Dep Managers","PAC Services_Manager":"Dep Managers","Concierge Services_Manager":"Dep Managers","Security Services_Manager":"Dep Managers","Home Services_Operations Manager":"Dep Managers","Landscaping_Manager":"Dep Managers","Finance & Accounts_Accounting Manager":"Finance","Facilities Management_Director":"Director","Commercial_Director":"Director","Executive Management_CEO":"CEO"};
const PROC = new Set(["PR In Review","RFQ to suppliers","Qt received & Logged","OP confirms material","Procurement (in process)"]);
const OPS = new Set(["Qt Shared to Op","Unit Price Updated"]);
const PO_MAP = {"Advance payment request submitted (if applicable)":"Procurement","Procurement Manager":"Procurement","Accounting Manager":"Finance","Finance and Accounts Director":"Director","CEO":"CEO","LPO sent/shared with supplier":"Sent to Supplier"};
const COLOR = {'Procurement':'#3b82f6','Operations to Confirm':'#14b8a6','Dep Managers':'#8b5cf6','Finance':'#22c55e','Director':'#ec4899','CEO':'#f59e0b','Sent to Supplier':'#a855f7','Pending Invoicing':'#f97316'};
const GRAD = {'Procurement':'#eff5ff','Operations to Confirm':'#ebfbf7','Dep Managers':'#f4f1fe','Finance':'#eefbf3','Director':'#fdeff7','CEO':'#fff9ec','Sent to Supplier':'#f9f2ff','Pending Invoicing':'#fff4e8'};
const TYCOL = {'PR':'#2563eb','CPR':'#7c3aed','PO':'#0891b2'};

/* ---- helpers ---- */
function esc(s){ return String(s==null?'':s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;'); }
function money(v){ const n=Number(v); return (isFinite(n)?n:0).toLocaleString('en-US',{maximumFractionDigits:0}); }
function amt(r){ const v=Number(r['Total amount']); return isFinite(v)?v:0; }
function xd(v){ if(v instanceof Date) return isNaN(v)?null:v; if(typeof v==='number'&&isFinite(v)){ const o=XLSX.SSF.parse_date_code(v); if(!o||!o.y) return null; return new Date(Date.UTC(o.y,o.m-1,o.d,o.H||0,o.M||0,Math.floor(o.S||0))); } if(typeof v==='string'&&v){ const d=new Date(v); return isNaN(d)?null:d; } return null; }
function ageDays(v){ const d=xd(v); return d? Math.max(0,Math.floor((Date.now()-d.getTime())/86400000)) : null; }
function ymdStr(v){ const d=xd(v); if(!d) return ''; const p=n=>String(n).padStart(2,'0'); return d.getUTCFullYear()+'-'+p(d.getUTCMonth()+1)+'-'+p(d.getUTCDate()); }
function avg(l){ return l.length? l.reduce((a,b)=>a+b,0)/l.length : 0; }
function parseXlsx(buf){ const wb=XLSX.read(buf,{type:'buffer'}); const ws=wb.Sheets[wb.SheetNames[0]]; const aoa=XLSX.utils.sheet_to_json(ws,{header:1,defval:null,raw:true}); if(!aoa.length) return []; const hdr=aoa[0].map(x=>String(x)); const rows=[]; for(let i=1;i<aoa.length;i++){ const o={}; for(let j=0;j<hdr.length;j++) o[hdr[j]]=aoa[i][j]; rows.push(o);} return rows; }
function prLive(r){ const st=String(r['Status']||''); return !!PR_MAP[r['Step name']] && st.toLowerCase()!=='closed' && st!=='Rejected' && st!=='Cancelled'; }
function poBucket(r){ const step=r['Step name'],appr=String(r['Approval status']||''),pos=String(r['Purchase order status']||''); let b=PO_MAP[step]||null; if(appr==='Confirmed'&&pos==='Received') b='Pending Invoicing'; if(appr==='Rejected'||pos==='Canceled'||pos==='Invoiced'||!b) return null; return b; }
function prHb(b){ return PROC.has(b)?'Procurement':(OPS.has(b)?'Operations to Confirm':b); }
function prAge(r){ const sd=r['Step date and time']; return ageDays(sd!=null? sd : r['Created date']); }
function poAge(r){ const sd=r['Step date and time']; return ageDays(sd!=null? sd : (r['Created date and time']!=null? r['Created date and time'] : r['Requested receipt date'])); }

function buildItems(prRows, poRows){
  const items=[];
  for(const r of prRows){ if(!prLive(r)) continue; const st=prHb(PR_MAP[r['Step name']]); items.push({ref:r['Purchase requisition'],doc:'PR',typ:String(r['Purchase requisition']||'').startsWith('CPR')?'CPR':'PR',stage:st,age:prAge(r),owner:(String(r['Pending Approver/User']||'').trim()||'(unassigned)'),dept:String(r['Department']||'').trim(),value:amt(r),vendor:'',raw:r}); }
  for(const r of poRows){ const bk=poBucket(r); if(!bk) continue; const ven=String(r['Vendor name']||'-').trim(); const own=(bk==='Sent to Supplier'||bk==='Pending Invoicing')?ven:(String(r['Pending Approver/User']||'').trim()||'(unassigned)'); items.push({ref:r['Purchase order'],doc:'PO',typ:'PO',stage:bk,age:poAge(r),owner:own,dept:String(r['Department']||'').trim(),value:amt(r),vendor:ven,raw:r}); }
  return items;
}

/* ---- render helpers ---- */
function chip(doc){ const c=doc==='PR'?'#2563eb':'#0891b2'; return '<span style="display:inline-block;background:'+c+';color:#fff;font-family:'+FONT+';font-size:9px;font-weight:800;letter-spacing:.3px;padding:1px 5px;border-radius:4px;margin-right:6px;vertical-align:middle;">'+doc+'</span>'; }
function card2(doc,bk,x){ const a=x.c?(x.sum/x.c):0; return '<td width="192" valign="top" style="width:192px;padding:0 5px;"><table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background:'+GRAD[bk]+';border:1px solid #e8ecf2;border-top:3px solid '+COLOR[bk]+';border-radius:12px;"><tr><td style="padding:12px 13px;"><table role="presentation" cellpadding="0" cellspacing="0"><tr><td valign="middle" bgcolor="'+(doc==='PR'?'#2563eb':'#0891b2')+'" style="background:'+(doc==='PR'?'#2563eb':'#0891b2')+';border-radius:4px;padding:2px 6px;font-family:'+FONT+';font-size:9px;font-weight:800;letter-spacing:.3px;color:#ffffff;">'+doc+'</td><td width="6" style="width:6px;">&#160;</td><td valign="middle" style="font-family:'+FONT+';font-size:10px;font-weight:800;letter-spacing:.3px;color:#5b6b7f;text-transform:uppercase;">'+esc(bk)+'</td></tr></table><div style="margin:6px 0 3px;white-space:nowrap;"><span style="font-family:'+FONT+';font-size:28px;font-weight:800;color:'+COLOR[bk]+';">'+x.n+'</span><span style="font-family:'+FONT+';font-size:13px;font-weight:800;color:'+RED+';"> ('+a.toFixed(1)+'d)</span></div><div style="font-family:'+FONT+';font-size:11.5px;font-weight:700;color:'+TEAL+';">AED '+money(x.amt)+'</div></td></tr></table></td>'; }
function cardrow2(pairs,agg){ const cells=pairs.filter(([d,bk])=>agg[d+'|'+bk]&&agg[d+'|'+bk].n>0).map(([d,bk])=>card2(d,bk,agg[d+'|'+bk])).join(''); return '<table cellpadding="0" cellspacing="0" border="0" style="border-collapse:separate;margin:2px 0 4px;"><tr>'+cells+'</tr></table>'; }
function badge(L,col){ return '<span style="display:inline-block;width:24px;height:24px;line-height:24px;text-align:center;border-radius:50%;background:'+col+';color:#fff;font-family:'+FONT+';font-weight:800;font-size:12px;margin-right:9px;vertical-align:middle;">'+L+'</span>'; }
function finding(L,col,title,right,narr,tbl){ const head='<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="margin:0 0 7px;"><tr><td width="24" height="24" align="center" valign="middle" bgcolor="'+col+'" style="width:24px;height:24px;background:'+col+';border-radius:50%;font-family:'+FONT+';font-weight:800;font-size:12px;color:#ffffff;">'+L+'</td><td width="9" style="width:9px;">&#160;</td><td valign="middle" style="font-family:'+FONT+';font-weight:800;font-size:14px;color:'+NAVY+';">'+title+'</td>'+(right?'<td align="right" valign="middle" style="font-family:'+FONT+';font-size:11px;font-weight:700;color:#94a3b8;white-space:nowrap;padding-left:10px;">'+right+'</td>':'')+'</tr></table>'; const body='<div style="font-family:'+FONT+';font-size:12.5px;color:#3f4b5b;line-height:1.5;margin:0 0 '+(tbl?'8px':'2px')+';">'+narr+'</div>'+(tbl||''); return '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="margin:0 0 12px;"><tr><td style="border:1px solid #e6ebf1;border-left:4px solid '+col+';border-radius:10px;background:#fff;padding:13px 15px;">'+head+body+'</td></tr></table>'; }
function b(s){ return '<b style="color:'+NAVY+';">'+s+'</b>'; }
function nm(u){ return '<span style="font-weight:700;color:'+NAVY+';">'+esc(u)+'</span>'; }
function sv(t,col){ return '<span style="color:'+col+';font-weight:700;">'+t+'</span>'; }
function agec(a){ a=Math.round(a||0); const col=a>30?'#b91c1c':(a>7?'#c2410c':'#16794a'); return '<span style="font-weight:700;color:'+col+';">'+a+'d</span>'; }
function otable(cols,rows){ const th='color:#fff;font-family:'+FONT+';font-weight:800;font-size:10.5px;padding:6px 9px;white-space:nowrap;background:'+NAVY+';'; const head='<tr>'+cols.map(c=>'<th style="'+th+(c[2]==='r'?'text-align:right;':(c[2]==='c'?'text-align:center;':'text-align:left;'))+'width:'+c[1]+'px;">'+c[0]+'</th>').join('')+'</tr>'; let body=''; rows.forEach((cells,i)=>{ const bg=i%2===0?'#ffffff':'#f7f9fc'; const td='padding:6px 9px;border-bottom:1px solid #eef1f6;font-family:'+FONT+';font-size:11.5px;background:'+bg+';white-space:nowrap;'; body+='<tr>'+cells.map((cell,j)=>'<td style="'+td+(cols[j][2]==='r'?'text-align:right;':(cols[j][2]==='c'?'text-align:center;':''))+'">'+cell+'</td>').join('')+'</tr>'; }); return '<div style="border:1px solid #e2e8f0;border-radius:8px;overflow:hidden;display:inline-block;"><table cellpadding="0" cellspacing="0" border="0" style="border-collapse:collapse;background:#fff;">'+head+body+'</table></div>'; }
function SH(t,s){ return '<div style="font-family:'+FONT+';font-weight:800;font-size:15px;color:'+NAVY+';margin:0 0 2px;">'+t+'</div>'+(s?'<div style="font-family:'+FONT+';font-size:11.5px;color:#7688a0;margin:0 0 9px;">'+s+'</div>':''); }

/* ---- finding builders ---- */
function grpBy(arr,key){ const g={}; for(const it of arr){ const k=key(it); (g[k]=g[k]||[]).push(it); } return g; }
function f_owners(its,L,col,title,label){
  const persons=its.filter(it=>it.stage!=='Sent to Supplier'&&it.stage!=='Pending Invoicing');
  const g={}; for(const it of persons){ const e=g[it.owner]||(g[it.owner]={n:0,ages:[],val:0,br:0,bk:{}}); e.n++; e.val+=it.value; e.bk[it.stage]=(e.bk[it.stage]||0)+1; if(it.age!=null){e.ages.push(it.age); if(it.age>7)e.br++;} }
  const rows=[]; for(const [u,e] of Object.entries(g).sort((a,b2)=>b2[1].n-a[1].n)){ if(u==='(unassigned)')continue; const stg=Object.entries(e.bk).sort((a,b2)=>b2[1]-a[1])[0][0]; rows.push([nm(u),sv(stg,COLOR[stg]),String(e.n),agec(avg(e.ages)),agec(e.ages.length?Math.max(...e.ages):0),sv(String(e.br),'#b91c1c'),'AED '+money(e.val)]); if(rows.length>=6)break; }
  if(!rows.length) return '';
  const n=persons.length, br=persons.filter(it=>(it.age||0)>7).length;
  return finding(L,col,title,n+' items',b(n)+' items are pending with a person; '+b(br)+' are past the 7-day SLA. Top owners &#8212; chase these queues first.',
    otable([[label||'Pending with',150,'l'],['Stage',150,'l'],['Items',52,'c'],['Avg',52,'c'],['Oldest',58,'c'],['Br&gt;7',50,'c'],['Value',120,'r']],rows));
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
  const rows=old.map(it=>['<span style="font-weight:700;color:'+NAVY+';">'+esc(it.ref)+'</span>',sv(it.typ,TYCOL[it.typ]),agec(it.age),sv(it.stage,COLOR[it.stage]),esc(String(it.owner||'-').slice(0,20)),esc(String(it.dept||'-').slice(0,20))]);
  return finding(L,col,'Escalate &#8212; oldest &amp; most overdue','top '+old.length,'Oldest items are stuck up to '+b(old[0].age+' days')+' &#8212; '+esc(old[0].ref)+' with '+nm(old[0].owner)+'. Escalate the top rows.',
    otable([['Ref',118,'l'],['Type',46,'c'],['Age',54,'c'],['Stage',150,'l'],['Owner/Vendor',150,'l'],['Dept',150,'l']],rows));
}
function f_value(its,L,col){
  const tv=its.slice().sort((a,b2)=>b2.value-a.value).slice(0,6); if(!tv.length) return '';
  const s=tv.reduce((a,it)=>a+it.value,0);
  const rows=tv.map(it=>['<span style="font-weight:700;color:'+NAVY+';">'+esc(it.ref)+'</span>',sv(it.typ,TYCOL[it.typ]),'<b>AED '+money(it.value)+'</b>',agec(it.age),sv(it.stage,COLOR[it.stage]),esc(String(it.owner||'-').slice(0,20))]);
  return finding(L,col,'High-value at risk','AED '+money(s),'The biggest exposure is '+b('AED '+money(tv[0].value))+' ('+esc(tv[0].ref)+'). Top 6 total '+b('AED '+money(s))+' &#8212; clear these for the biggest cash impact.',
    otable([['Ref',118,'l'],['Type',46,'c'],['Value',120,'l'],['Age',54,'c'],['Stage',150,'l'],['Owner/Vendor',160,'l']],rows));
}
function f_sla(its,L,col){
  const g={}; for(const it of its){ const e=g[it.stage]||(g[it.stage]={n:0,br:0,ages:[]}); e.n++; if(it.age!=null){e.ages.push(it.age); if(it.age>7)e.br++;} }
  const order=['Procurement','Operations to Confirm','Dep Managers','Finance','Director','CEO','Sent to Supplier','Pending Invoicing'].filter(s=>g[s]);
  const rows=order.map(s=>{ const e=g[s], pct=e.n?Math.round(100*e.br/e.n):0, pcol=pct>=70?'#b91c1c':(pct>=40?'#c2410c':'#16794a'); return [sv(s,COLOR[s]),String(e.n),sv(String(e.br),'#b91c1c'),sv(pct+'%',pcol),agec(avg(e.ages))]; });
  const tn=order.reduce((a,s)=>a+g[s].n,0), tb=order.reduce((a,s)=>a+g[s].br,0);
  return finding(L,col,'Stage SLA performance',(tn?Math.round(100*tb/tn):0)+'% breached',b(tb+' of '+tn)+' items are past the 7-day SLA ('+b((tn?Math.round(100*tb/tn):0)+'%')+'). Breach rate by stage below.',
    otable([['Stage',180,'l'],['Items',56,'c'],['Breach&gt;7d',82,'c'],['Breach%',66,'c'],['Avg',52,'c']],rows));
}
function f_dept(its,L,col){
  const g={}; for(const it of its){ const d=it.dept||'(unspecified)'; const e=g[d]||(g[d]={n:0,br:0,ages:[],val:0}); e.n++; e.val+=it.value; if(it.age!=null){e.ages.push(it.age); if(it.age>7)e.br++;} }
  const ent=Object.entries(g).sort((a,b2)=>b2[1].n-a[1].n);
  const rows=ent.slice(0,8).map(([d,e])=>[esc(d.slice(0,28)),String(e.n),sv(String(e.br),'#b91c1c'),agec(avg(e.ages)),'AED '+money(e.val)]);
  return finding(L,col,'By department &#8212; where it sits','',b(ent[0][0])+' has the most ('+b(ent[0][1].n)+' items). Full split below &#8212; route to each department head.',
    otable([['Department',210,'l'],['Items',56,'c'],['Breach&gt;7d',82,'c'],['Avg',52,'c'],['Value',130,'r']],rows));
}

/* ---- divisions ---- */
const DIVS = [
 {key:'procurement', mail:'PRPO_PROC_MAIL_TO', xlsx:'PRPO_Procurement_list.xlsx', title:'Procurement',
  heading:'PR / PO Pipeline &#8212; Procurement', sub:'Procurement-stage PR &amp; PO, plus POs sent to suppliers', accent:'#3b82f6',
  pr:['Procurement'], po:['Procurement','Sent to Supplier'],
  findings:[f=>f_owners(f,'A','#3b82f6','Pending with &#8212; procurement queue'), f=>f_vendor(f,'Sent to Supplier','B','#a855f7','Sent to Supplier &#8212; awaiting delivery / GRN','Chase the suppliers below for delivery, then move to invoicing.'), f=>f_value(f,'C','#2563eb'), f=>f_oldest(f,'D','#dc2626'), f=>f_sla(f,'E','#e11d48')]},
 {key:'finance', mail:'PRPO_FIN_MAIL_TO', xlsx:'PRPO_Finance_list.xlsx', title:'Finance &amp; Approvals',
  heading:'PR / PO Pipeline &#8212; Finance &amp; Approvals', sub:'Finance &amp; Director approvals (PR) and POs pending invoicing', accent:'#22c55e',
  pr:['Finance','Director','CEO'], po:['Finance','Pending Invoicing','Director','CEO'],
  findings:[f=>f_owners(f,'A','#22c55e','Pending approvals &#8212; waiting on you'), f=>f_vendor(f,'Pending Invoicing','B','#f97316','Pending Invoicing &#8212; Accounts to post','Post the supplier invoices below to clear these from the ledger.'), f=>f_value(f,'C','#2563eb'), f=>f_oldest(f,'D','#dc2626'), f=>f_sla(f,'E','#e11d48')]},
 {key:'ops_hm', mail:'PRPO_OPSHM_MAIL_TO', xlsx:'PRPO_Operations_HomeMaint_FitOut_list.xlsx', title:'Operations &#183; Home Maintenance + FitOut',
  heading:'PR / PO Pipeline &#8212; Operations (Home Maintenance &amp; FitOut)', sub:'Operations to Confirm &amp; Dep Managers &#8212; Home Maintenance + FitOut Services', accent:'#14b8a6',
  pr:['Operations to Confirm','Dep Managers'], po:[], depts:new Set(['Home Maintenance Services','FitOut Services']),
  findings:[f=>f_owners(f,'A','#14b8a6','Pending with &#8212; Home Maintenance &amp; FitOut queue'), f=>f_value(f,'B','#2563eb'), f=>f_oldest(f,'C','#dc2626'), f=>f_sla(f,'D','#e11d48')]},
 {key:'ops_all', mail:'PRPO_OPSALL_MAIL_TO', xlsx:'PRPO_Operations_AllDepts_list.xlsx', title:'Operations &#183; All Departments',
  heading:'PR / PO Pipeline &#8212; Operations (All Departments)', sub:'Operations to Confirm &amp; Dep Managers &#8212; all departments except Home Maintenance &amp; FitOut', accent:'#8b5cf6',
  pr:['Operations to Confirm','Dep Managers'], po:[], xdepts:new Set(['Home Maintenance Services','FitOut Services']),
  findings:[f=>f_owners(f,'A','#8b5cf6','Pending with &#8212; who is holding the queue'), f=>f_dept(f,'B','#4f46e5'), f=>f_value(f,'C','#2563eb'), f=>f_oldest(f,'D','#dc2626'), f=>f_sla(f,'E','#e11d48')]},
];
function filterDiv(items,cfg){ return items.filter(it=> ((it.doc==='PR'&&cfg.pr.includes(it.stage))||(it.doc==='PO'&&cfg.po.includes(it.stage))) && (!cfg.depts||cfg.depts.has(it.dept)) && (!cfg.xdepts||!cfg.xdepts.has(it.dept)) ); }

function buildXlsxBase64(fil){
  const cols=['Ref','Doc','Stage / Bucket','Step name','Status','Department','Location','Pending With','Vendor','Value (AED)','Age (days)','Created','Step date','Title / Name','Preparer / Linked PR'];
  const aoa=[cols];
  fil.slice().sort((a,b2)=>(b2.age||0)-(a.age||0)).forEach(it=>{ const r=it.raw; let status,loc,pend,ven,created,stepd,title,prep;
    if(it.doc!=='PO'){ status=String(r['Status']||''); loc=r['Location']; pend=r['Pending Approver/User']; ven=''; created=ymdStr(r['Created date']); stepd=ymdStr(r['Step date and time']); title=r['Name']; prep=r['Preparer']; }
    else { status=(String(r['Approval status']||'')+' / '+String(r['Purchase order status']||'')).replace(/^ \/ | \/ $/g,''); loc=r['Location']; pend=r['Pending Approver/User']; ven=r['Vendor name']; created=ymdStr(r['Created date and time']); stepd=ymdStr(r['Step date and time']); title=''; prep=r['Purchase requisition']; }
    aoa.push([it.ref,it.typ,it.stage,r['Step name'],status,it.dept,loc,pend,ven,Math.round(it.value*100)/100,it.age,created,stepd,title,prep]); });
  const ws=XLSX.utils.aoa_to_sheet(aoa);
  ws['!cols']=[15,6,20,30,20,26,20,18,26,14,9,12,12,28,16].map(w=>({wch:w}));
  ws['!autofilter']={ref:'A1:O'+aoa.length};
  const wb=XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb,ws,'Open Items');
  return XLSX.write(wb,{type:'base64',bookType:'xlsx'});
}

/* ---- HS-D08 style: navy shell header/footer + PR/PO detail line-item list ---- */
const HF="'Segoe UI', Aptos, Verdana, Arial, sans-serif", HNAVY='#0F2A6B', HGOLD='#FAC775', HMUT='#C9D3EA', HBORD='#D8DEE8';
function f_details(fil, cfg){
  const rows=fil.slice().sort((a,b2)=>(b2.age||0)-(a.age||0)).slice(0,15);
  const desc=it=> it.doc==='PO' ? (it.raw['Vendor name']||'-') : (it.raw['Name']||'-');
  const th='padding:7px 9px;font:700 10px '+HF+';color:#5A6578;text-transform:uppercase;letter-spacing:.3px;background:#F4F6FB;';
  const head='<tr>'+[['Document','l'],['Description','l'],['Waiting on step','l'],['Pending with','l'],['Value','r'],['Age','r']].map(c=>'<td '+(c[1]==='r'?'align="right" ':'')+'style="'+th+'">'+c[0]+'</td>').join('')+'</tr>';
  const body=rows.map(it=>{ const a=Math.round(it.age||0); const acol=a>30?'#B42318':(a>7?'#9A6700':'#1F7A33'); const td='padding:6px 9px;font:400 12px '+HF+';color:#1A2233;border-bottom:1px solid '+HBORD+';';
    return '<tr><td style="'+td+'font-weight:600;white-space:nowrap;">'+esc(it.ref)+'</td>'
      +'<td style="'+td+'">'+esc(String(desc(it)).slice(0,46))+'</td>'
      +'<td style="'+td+'color:'+HNAVY+';white-space:nowrap;">'+esc(String(it.raw['Step name']||'-').slice(0,28))+'</td>'
      +'<td style="'+td+'font-weight:600;white-space:nowrap;">'+esc(String(it.owner||'-').slice(0,18))+'</td>'
      +'<td align="right" style="'+td+'font-weight:600;white-space:nowrap;">AED '+money(it.value)+'</td>'
      +'<td align="right" style="'+td+'font-weight:700;color:'+acol+';white-space:nowrap;">'+a+'d</td></tr>'; }).join('');
  const more=fil.length>15?'<tr><td colspan="6" style="padding:7px 9px;font:400 11px '+HF+';color:#5A6578;background:#F9FAFC;">&#8230;and '+(fil.length-15)+' more &#8212; full list in the attached '+cfg.xlsx+'.</td></tr>':'';
  return '<div style="font:700 14px '+HF+';color:'+HNAVY+';margin:16px 0 2px;">Details &#8212; PR / PO list</div>'
    +'<div style="font:400 11px '+HF+';color:#5A6578;margin:0 0 10px;">Oldest first &#183; age = days at the current step.</div>'
    +'<table role="presentation" width="'+W+'" cellpadding="0" cellspacing="0" style="width:'+W+'px;border:1px solid '+HBORD+';border-radius:8px;border-collapse:separate;overflow:hidden;">'+head+body+more+'</table>';
}

function buildDivision(cfg, items){
  const fil=filterDiv(items,cfg);
  const pairs=cfg.pr.map(bk=>['PR',bk]).concat(cfg.po.map(bk=>['PO',bk]));
  const agg={}; for(const it of fil){ const k=it.doc+'|'+it.stage; const set=pairs.some(([d,bk])=>d===it.doc&&bk===it.stage); if(!set) continue; const x=agg[k]||(agg[k]={n:0,sum:0,c:0,amt:0}); x.n++; x.amt+=it.value; if(it.age!=null){x.sum+=it.age;x.c++;} }
  const cards=cardrow2(pairs,agg);
  const analysis=cfg.findings.map(fn=>fn(fil)).join('');
  const stamp=new Date(Date.now()+4*3600*1000).toLocaleDateString('en-GB',{day:'2-digit',month:'short',year:'numeric',timeZone:'UTC'});
  const tot=fil.length, totv=fil.reduce((a,it)=>a+it.value,0);
  const att='<div style="border:1px dashed #b6c2d4;border-radius:10px;background:#f8fafc;padding:11px 14px;margin:12px 0 2px;font:400 12px '+HF+';color:#475569;">&#128206; <b style="color:'+HNAVY+';">Attachment:</b> <span style="color:'+HNAVY+';font-weight:700;">'+cfg.xlsx+'</span> &#8212; full line-item list of all '+b(tot)+' items (Ref, Doc, Stage, Status, Department, Pending With / Vendor, Value, Age, dates), sorted oldest first.</div>';
  const titleTxt=cfg.title.replace(/&#183;/g,'·').replace(/&amp;/g,'&');
  const inner='<div style="width:'+W+'px;font-family:'+FONT+';color:#22303c;">'
    +'<div style="font:400 12px '+HF+';color:#607083;margin:0 0 10px;">This queue: '+b(tot)+' open items &#183; '+b('AED '+money(totv))+' &#183; live-pipeline logic, reconciles to the dashboard.</div>'
    +cards+att
    +'<div style="font-family:'+FONT+';font-weight:800;font-size:15px;color:'+NAVY+';margin:16px 0 2px;">&#128269; Analysis &#8212; who to chase today</div>'
    +'<div style="font-family:'+FONT+';font-size:11.5px;color:#7688a0;margin:0 0 11px;">Auto-generated daily from the latest data, scoped to '+cfg.title+'.</div>'
    +'<div style="width:'+W+'px;">'+analysis+'</div>'
    +f_details(fil,cfg)
    +'</div>';
  const shell='<table role="presentation" width="1040" cellpadding="0" cellspacing="0" style="width:1040px;max-width:1040px;">'
    +'<tr><td style="background:'+HNAVY+';border-radius:10px 10px 0 0;padding:18px 22px;"><table role="presentation" width="100%" cellpadding="0" cellspacing="0"><tr>'
    +'<td style="font:700 20px '+HF+';color:#fff;">'+cfg.heading+'</td>'
    +'<td align="right" valign="top" style="font:600 12px '+HF+';color:'+HGOLD+';">'+stamp+'</td></tr></table>'
    +'<div style="font:400 12px '+HF+';color:'+HMUT+';margin-top:4px;">'+cfg.sub+' &#183; <a href="'+DASH+'" style="color:#9DB2E6;text-decoration:none;">Live dashboard</a></div></td></tr>'
    +'<tr><td style="background:#fff;border-left:1px solid '+HBORD+';border-right:1px solid '+HBORD+';padding:18px 20px;">'+inner+'</td></tr>'
    +'<tr><td style="background:'+HNAVY+';border-radius:0 0 10px 10px;padding:14px 20px;font:400 11px '+HF+';color:'+HMUT+';">'
    +'<div style="color:'+HGOLD+';font-weight:700;letter-spacing:.5px;">FOR EXCELLENCE WE STRIVE</div>'
    +'<div style="margin-top:6px;">PR / PO Pipeline &#183; '+titleTxt+' &#183; '+stamp+' &#183; automated daily 10:00 AM Dubai. Source: live-pipeline PR/PO in D365 F&amp;O. <b>Age = days at the current workflow step.</b> Full line-item list attached ('+cfg.xlsx+').</div></td></tr></table>';
  const wrap='<!doctype html><html><head><meta charset="utf-8"><title>'+cfg.heading+'</title></head><body style="margin:0;background:#EEF1F6;"><table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background:#EEF1F6;padding:20px 0;"><tr><td align="center">'+shell+'</td></tr></table></body></html>';
  const subject='PR / PO Pipeline — '+cfg.title.replace(/&#183;/g,'·').replace(/&amp;/g,'&')+' ('+stamp+')';
  return { subject, html:wrap, xlsxB64:buildXlsxBase64(fil), count:tot, value:totv, cfg };
}

/* ---- auth + send ---- */
async function getToken(scopeBase){ const body=new URLSearchParams({client_id:process.env.CLIENT_ID,client_secret:process.env.CLIENT_SECRET,grant_type:'client_credentials',scope:scopeBase.replace(/\/+$/,'')+'/.default'}); const r=await fetch(`https://login.microsoftonline.com/${process.env.TENANT_ID}/oauth2/v2.0/token`,{method:'POST',headers:{'Content-Type':'application/x-www-form-urlencoded'},body}); const j=await r.json(); if(!r.ok||!j.access_token) throw new Error('token '+r.status+' '+(j.error_description||j.error||'')); return j.access_token; }
async function sendDivision(out, context){
  const from=process.env.PRPO_MAIL_FROM||process.env.MAIL_FROM;
  const toList=(process.env[out.cfg.mail]||'').split(/[;,]/).map(s=>s.trim()).filter(Boolean);
  if(!from) throw new Error('MAIL_FROM / PRPO_MAIL_FROM not set');
  if(!toList.length){ if(context) context.log('skip '+out.cfg.key+': '+out.cfg.mail+' not set'); return {sent:false,reason:'no recipients'}; }
  const token=await getToken('https://graph.microsoft.com');
  const r=await fetch('https://graph.microsoft.com/v1.0/users/'+encodeURIComponent(from)+'/sendMail',{method:'POST',headers:{Authorization:'Bearer '+token,'Content-Type':'application/json'},
    body:JSON.stringify({message:{subject:out.subject,body:{contentType:'HTML',content:out.html},toRecipients:toList.map(a=>({emailAddress:{address:a}})),attachments:[{'@odata.type':'#microsoft.graph.fileAttachment',name:out.cfg.xlsx,contentType:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',contentBytes:out.xlsxB64}]},saveToSentItems:true})});
  if(r.status!==202){ const j=await r.json().catch(()=>({})); throw new Error('sendMail '+r.status+' '+JSON.stringify(j.error||j).slice(0,300)); }
  if(context) context.log('sent '+out.cfg.key+' to '+toList.length+' recipients');
  return {sent:true,to:toList.length};
}

async function fetchXlsx(url){ const r=await fetch(url+(url.includes('?')?'&':'?')+'t='+Date.now()); if(!r.ok) throw new Error('fetch '+r.status+' '+url); return parseXlsx(Buffer.from(await r.arrayBuffer())); }
async function loadItems(){ const [prRows,poRows]=await Promise.all([fetchXlsx(PR_URL),fetchXlsx(PO_URL)]); return buildItems(prRows,poRows); }

/* ---- triggers ---- */
app.timer('prpo-email-daily', { schedule:'0 0 6 * * *', handler:async(timer,context)=>{
  const items=await loadItems();
  for(const cfg of DIVS){ try{ await sendDivision(buildDivision(cfg,items),context); }catch(e){ context.error('prpo '+cfg.key+' FAILED: '+e.message); } }
}});
app.http('prpo-email', { methods:['GET','OPTIONS'], authLevel:'function', route:'prpo-email', handler:async(request,context)=>{
  try{
    const url=new URL(request.url); const dk=url.searchParams.get('division'); const sendAll=url.searchParams.get('send')==='1'&&!dk;
    const items=await loadItems();
    if(dk){ const cfg=DIVS.find(d=>d.key===dk); if(!cfg) return {status:400,jsonBody:{error:'unknown division; use procurement|finance|ops_hm|ops_all'}};
      const out=buildDivision(cfg,items);
      if(url.searchParams.get('format')==='html') return {status:200,headers:{'Content-Type':'text/html; charset=utf-8'},body:out.html};
      if(url.searchParams.get('debug')==='1') return {status:200,jsonBody:{division:dk,count:out.count,value:Math.round(out.value)}};
      if(url.searchParams.get('send')==='1'){ const s=await sendDivision(out,context); return {status:200,jsonBody:{division:dk,...s}}; }
      return {status:200,jsonBody:{division:dk,count:out.count,value:Math.round(out.value)}};
    }
    const summary=[]; for(const cfg of DIVS){ const out=buildDivision(cfg,items); let s={sent:false}; if(sendAll) s=await sendDivision(out,context); summary.push({division:cfg.key,count:out.count,value:Math.round(out.value),...s}); }
    return {status:200,jsonBody:{sentAll:sendAll,divisions:summary}};
  }catch(e){ context.error('prpo-email failed:',e); return {status:500,jsonBody:{error:e.message}}; }
}});

module.exports = { buildItems, buildDivision, parseXlsx, DIVS };
