// ╔══════════════════════════════════════════════════════╗
//  ANTI FRAUD PORTAL v3.0 — Code.gs
//  Deploy: Extensions → Apps Script → Deploy → Web App
//  Execute as: Me | Access: Anyone with Google account
// ╚══════════════════════════════════════════════════════╝

// ── Sheet names
const SH_CASES    = 'CASES';
const SH_HISTORY  = 'CASE_HISTORY';
const SH_NOTES    = 'CASE_NOTES';
const SH_EVIDENCE = 'CASE_EVIDENCE';
const SH_USERS    = 'USERS';
const SH_PRIORITY = 'PRIORITY_CONFIG';
const SH_CATS     = 'CATEGORIES';
const SH_CLOSE_R  = 'CLOSE_REASONS';

// ── Status flows
const STATUS_FLOW    = ['Open','On Investigation','Case Close'];
const OPEN_SUBS      = ['ASSIGN TO PIC','ASSIGN TO SL','HO CALL'];
const INVEST_SUBS    = ['SPP','JANJI BAYAR (PTP)','UPDATE TGL JANJI BAYAR (BROKEN PTP)','LUNAS (PAID OFF)','CICIL (TERMIN)','HOLD','PENYITAAN ASET'];
const DEFAULT_CATS   = ['COD Fraud','Driver Fraud','Data Manipulation','Unauthorized Transaction','Double Collection','LND Fraud','COD & LND Fraud','Other Fraud','Lainnya'];
const DEFAULT_CLOSE  = ['Setor NEK','Lunas Cicilan','Setor Penampung','Refund','Deduction Invoice','Penyitaan dan Penjualan Aset','Other'];

// ── CASES columns (1-based)
const C = {
  NO:1, CASE_ID:2, DATE_IN:3,
  // Courier
  COURIER_ID:4, COURIER_NM:5, HUB:6,
  // Financial
  AMOUNT:7, REK_AMT:8,
  // === READ-ONLY LOOKUP FIELDS (populated via VLOOKUP formula in sheet) ===
  STAT_PEND:9, FDS_HUB:10, NIK:11,
  EMAIL_K:12, TELEPON:13, ALAMAT:14, KTP:15,
  NAMA_KONDAR:16, NO_KONDAR:17, HUB_KONDAR:18,
  LEAD_REGION:19, PIC:20, KORLAP:21, CITY:22, PROVINCE:23,
  // === CASE MANAGEMENT ===
  STATUS:24, SUB_STATUS:25, CLOSE_REASON:26,
  FIRST_UPD:27, LAST_UPD:28, CLOSE_DATE:29, RUN_DAYS:30,
  PRIORITY:31, CATEGORY:32, AREA:33, ASSIGNED:34,
  LAST_NOTE:35, CREATED_BY:36, UPDATED_BY:37,
  _TOTAL:37
};

// ── CASE_HISTORY columns
const H = { TIMESTAMP:1,CASE_ID:2,ACTION:3,FIELD:4,OLD_VAL:5,NEW_VAL:6,EXTRA_JSON:7,BY:8,_TOTAL:8 };
// ── CASE_NOTES columns
const N = { TIMESTAMP:1,CASE_ID:2,NOTE:3,BY:4,_TOTAL:4 };
// ── CASE_EVIDENCE columns
const EV = { CASE_ID:1,TYPE:2,URL:3,FILENAME:4,DESCRIPTION:5,ADDED_BY:6,ADDED_AT:7,_TOTAL:7 };
// ── USERS columns
const U = { EMAIL:1,ROLE:2,AREA:3,ADDED_BY:4,ADDED_AT:5,_TOTAL:5 };
// ── PRIORITY columns
const P = { AMOUNT_MIN:1,LABEL:2,SORT:3,_TOTAL:3 };
// ── CONFIG list columns (for CATS and CLOSE_REASONS)
const CF = { LABEL:1,ACTIVE:2,SORT:3,_TOTAL:3 };

// ═══════════════════════════════
//  SPREADSHEET RESOLVER
// ═══════════════════════════════
function _ss() {
  try { const s=SpreadsheetApp.getActiveSpreadsheet(); if(s)return s; } catch(e){}
  const id=PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  if(!id) throw new Error('Jalankan Setup dari Google Sheets terlebih dahulu.');
  return SpreadsheetApp.openById(id);
}

// ═══════════════════════════════
//  WEB APP ENTRY
// ═══════════════════════════════
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('Portal')
    .setTitle('Anti Fraud Portal')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport','width=device-width,initial-scale=1.0');
}

// ═══════════════════════════════
//  MENU
// ═══════════════════════════════
function onOpen() {
  SpreadsheetApp.getUi().createMenu('🛡️ Anti Fraud Portal')
    .addItem('🌐 Buka Portal','openPortalInBrowser')
    .addItem('🔗 Lihat URL','showWebAppUrl')
    .addSeparator()
    .addItem('🔧 Setup','setupSheets')
    .addToUi();
}
function openPortalInBrowser() {
  const url=_getWebAppUrl();
  if(!url){SpreadsheetApp.getUi().alert('Deploy Web App terlebih dahulu.');return;}
  const s=url.replace(/'/g,"\\'");
  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutput(`<html><head><style>body{background:#0c0c0c;color:#fff;font-family:sans-serif;display:flex;flex-direction:column;align-items:center;justify-content:center;height:100vh;gap:10px;}a{color:#3b82f6;}</style></head>
    <body><p style="color:#888;font-size:13px">Membuka Anti Fraud Portal…</p>
    <a href="${s}" target="_blank">Klik jika tidak terbuka otomatis</a>
    <script>window.open('${s}','_blank');setTimeout(()=>google.script.host.close(),1500);</script></body></html>`)
    .setWidth(340).setHeight(160),'🛡️ Membuka Portal…');
}
function showWebAppUrl(){const url=_getWebAppUrl();SpreadsheetApp.getUi().alert('URL Portal',url||'Belum di-deploy.',SpreadsheetApp.getUi().ButtonSet.OK);}
function _getWebAppUrl(){try{const u=ScriptApp.getService().getUrl();return(u&&u.length>20)?u:null;}catch(e){return null;}}

// ═══════════════════════════════
//  AUTH
// ═══════════════════════════════
function getCurrentUser() {
  const email=Session.getActiveUser().getEmail()||'';
  try {
    const sheet=_ss().getSheetByName(SH_USERS);
    if(!sheet||sheet.getLastRow()<2) return{email,role:'superadmin',area:'ALL',authorized:true};
    const rows=sheet.getRange(2,1,sheet.getLastRow()-1,U._TOTAL).getValues();
    const f=rows.find(r=>String(r[U.EMAIL-1]).trim().toLowerCase()===email.toLowerCase());
    if(!f) return{email,role:null,area:null,authorized:false};
    return{email,role:f[U.ROLE-1],area:f[U.AREA-1],authorized:true};
  } catch(e){return{email,role:'superadmin',area:'ALL',authorized:true};}
}

// ═══════════════════════════════
//  SETUP
// ═══════════════════════════════
function setupSheets() {
  const ss=SpreadsheetApp.getActiveSpreadsheet(),ui=SpreadsheetApp.getUi();
  PropertiesService.getScriptProperties().setProperty('SPREADSHEET_ID',ss.getId());

  let cs=ss.getSheetByName(SH_CASES);
  if(cs){
    const r=ui.alert('Sheet CASES sudah ada.','Reset? (Data hilang!) No = hanya update config.',ui.ButtonSet.YES_NO);
    if(r===ui.Button.YES){cs.clearContents();cs.clearFormats();cs.clearConditionalFormatRules();}
    else{_ensureExtras(ss);ui.alert('✅ Config disimpan.\n'+(_getWebAppUrl()||'⚠ Belum deploy.'));return;}
  } else {cs=ss.insertSheet(SH_CASES);}
  _ensureExtras(ss);

  // Headers
  const hdrs=[
    'No','Case ID','Tanggal Input',
    'ID Kurir','Nama Kurir','Hub',
    'Amount COD','Nominal Rek Penampung',
    '★ Status Pending','★ FDS (HubScore)','★ NIK',
    '★ Email','★ Telepon','★ Alamat','★ KTP Link',
    '★ Nama Kontak Darurat','★ No Kontak Darurat','★ Hubungan Kondar',
    '★ Lead Region','★ PIC','★ Korlap','★ City','★ Province',
    'Status','Sub Status','Close Reason',
    'First Update','Last Update','Close Date','Running Days',
    'Priority','Kategori','Area','Assigned To',
    'Last Note','Created By','Updated By'
  ];
  cs.getRange(1,1,1,hdrs.length).setValues([hdrs]);
  _fmtHdr(cs,hdrs.length,'#0d1117');

  // Column widths
  const widths=[40,160,100,90,160,120,110,140,100,100,100,120,100,160,120,160,120,120,100,100,100,100,100,120,140,140,100,110,100,80,80,130,100,130,200,130,130];
  widths.forEach((w,i)=>cs.setColumnWidth(i+1,Math.min(w,200)));

  // Highlight lookup columns (light blue bg)
  cs.getRange(1,9,1000,15).setBackground('#e8f4f8');
  cs.getRange(1,9,1,15).setBackground('#1a3a4a').setFontColor('#7dd3fc');

  // Dropdowns (editable columns only)
  _dropdown(cs,C.STATUS,1000,['Open','On Investigation','Case Close','Escalated']);
  _dropdown(cs,C.PRIORITY,1000,['High','Medium','Low']);

  // Conditional formatting STATUS
  const sc={'Open':{bg:'#fef9c3',fc:'#854d0e'},'On Investigation':{bg:'#ffedd5',fc:'#7c2d12'},'Case Close':{bg:'#dcfce7',fc:'#14532d'},'Escalated':{bg:'#fee2e2',fc:'#7f1d1d'}};
  const existRules=[];
  Object.entries(sc).forEach(([v,c])=>existRules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(v).setBackground(c.bg).setFontColor(c.fc).setRanges([cs.getRange(2,C.STATUS,1000)]).build()));
  const pc={'High':{bg:'#fee2e2',fc:'#7f1d1d'},'Medium':{bg:'#fef3c7',fc:'#78350f'},'Low':{bg:'#dcfce7',fc:'#14532d'}};
  Object.entries(pc).forEach(([v,c])=>existRules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(v).setBackground(c.bg).setFontColor(c.fc).setRanges([cs.getRange(2,C.PRIORITY,1000)]).build()));
  cs.setConditionalFormatRules(existRules);

  cs.getRange(2,C.AMOUNT,1000).setNumberFormat('#,##0');
  cs.getRange(2,C.REK_AMT,1000).setNumberFormat('#,##0');
  [C.DATE_IN,C.FIRST_UPD,C.LAST_UPD,C.CLOSE_DATE].forEach(col=>cs.getRange(2,col,1000).setNumberFormat('dd mmm yyyy'));

  ui.alert('✅ Setup Selesai!\n\nCatatan: Kolom bertanda ★ adalah lookup—isi via VLOOKUP/formula dari sheet master kurir.\n\n'+(_getWebAppUrl()?'🌐 '+_getWebAppUrl():'⚠ Belum deploy Web App.'));
}

function _ensureExtras(ss) {
  const ensureSheet=(name,hdrs,bgColor)=>{
    let sh=ss.getSheetByName(name);
    if(!sh){sh=ss.insertSheet(name);sh.getRange(1,1,1,hdrs.length).setValues([hdrs]);_fmtHdr(sh,hdrs.length,bgColor||'#0d1117');}
    return sh;
  };
  ensureSheet(SH_HISTORY,['Timestamp','Case ID','Action','Field','Old Value','New Value','Extra JSON','Changed By']);
  ensureSheet(SH_NOTES,['Timestamp','Case ID','Note','By']);
  ensureSheet(SH_EVIDENCE,['Case ID','Type','URL','Filename','Description','Added By','Added At']);
  const su=ensureSheet(SH_USERS,['Email','Role','Area','Added By','Added At']);
  if(su.getLastRow()===1){_dropdown(su,U.ROLE,500,['superadmin','user']);su.setColumnWidth(1,220);}
  const sp=ensureSheet(SH_PRIORITY,['Amount Min (>=)','Priority Label','Sort Order']);
  if(sp.getLastRow()===1){sp.getRange(2,1,3,3).setValues([[10000000,'High',1],[3000000,'Medium',2],[0,'Low',3]]);sp.getRange(2,1,100).setNumberFormat('#,##0');}
  const sc=ensureSheet(SH_CATS,['Label','Active','Sort']);
  if(sc.getLastRow()===1){DEFAULT_CATS.forEach((cat,i)=>sc.appendRow([cat,true,i+1]));}
  const sr=ensureSheet(SH_CLOSE_R,['Label','Active','Sort']);
  if(sr.getLastRow()===1){DEFAULT_CLOSE.forEach((r,i)=>sr.appendRow([r,true,i+1]));}
}

function _fmtHdr(sheet,n,bg){
  sheet.getRange(1,1,1,n).setBackground(bg||'#0d1117').setFontColor('#fff').setFontWeight('bold').setFontSize(10).setHorizontalAlignment('center').setVerticalAlignment('middle').setWrap(false);
  sheet.setRowHeight(1,34);sheet.setFrozenRows(1);
}
function _dropdown(sheet,col,n,vals){
  sheet.getRange(2,col,n).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(vals,true).setAllowInvalid(false).build());
}

// ═══════════════════════════════
//  CONFIG: CATEGORIES & CLOSE REASONS
// ═══════════════════════════════
function getCategories() {
  try {
    const sh=_ss().getSheetByName(SH_CATS);
    if(!sh||sh.getLastRow()<2) return{ok:true,data:DEFAULT_CATS};
    return{ok:true,data:sh.getRange(2,1,sh.getLastRow()-1,CF._TOTAL).getValues().filter(r=>r[CF.ACTIVE-1]!==false&&r[CF.LABEL-1]!=='').map(r=>r[CF.LABEL-1])};
  } catch(e){return{ok:false,message:e.message};}
}
function saveCategories(cats) {
  try {
    const u=getCurrentUser();if(u.role!=='superadmin')return{ok:false,message:'Unauthorized.'};
    const sh=_ss().getSheetByName(SH_CATS);
    if(sh.getLastRow()>1)sh.getRange(2,1,sh.getLastRow()-1,CF._TOTAL).clearContent();
    cats.forEach((c,i)=>sh.appendRow([c,true,i+1]));
    return{ok:true,message:'Kategori disimpan.'};
  } catch(e){return{ok:false,message:e.message};}
}
function getCloseReasons() {
  try {
    const sh=_ss().getSheetByName(SH_CLOSE_R);
    if(!sh||sh.getLastRow()<2) return{ok:true,data:DEFAULT_CLOSE};
    return{ok:true,data:sh.getRange(2,1,sh.getLastRow()-1,CF._TOTAL).getValues().filter(r=>r[CF.ACTIVE-1]!==false&&r[CF.LABEL-1]!=='').map(r=>r[CF.LABEL-1])};
  } catch(e){return{ok:false,message:e.message};}
}
function saveCloseReasons(reasons) {
  try {
    const u=getCurrentUser();if(u.role!=='superadmin')return{ok:false,message:'Unauthorized.'};
    const sh=_ss().getSheetByName(SH_CLOSE_R);
    if(sh.getLastRow()>1)sh.getRange(2,1,sh.getLastRow()-1,CF._TOTAL).clearContent();
    reasons.forEach((r,i)=>sh.appendRow([r,true,i+1]));
    return{ok:true,message:'Close Reasons disimpan.'};
  } catch(e){return{ok:false,message:e.message};}
}
function getThresholds() {
  try {
    const sh=_ss().getSheetByName(SH_PRIORITY);
    if(!sh||sh.getLastRow()<2)return{ok:true,data:[{amountMin:10000000,label:'High',sortOrder:1},{amountMin:3000000,label:'Medium',sortOrder:2},{amountMin:0,label:'Low',sortOrder:3}]};
    return{ok:true,data:sh.getRange(2,1,sh.getLastRow()-1,P._TOTAL).getValues().filter(r=>r[P.AMOUNT_MIN-1]!=='').map(r=>({amountMin:r[P.AMOUNT_MIN-1],label:r[P.LABEL-1],sortOrder:r[P.SORT-1]})).sort((a,b)=>b.amountMin-a.amountMin)};
  } catch(e){return{ok:false,message:e.message};}
}
function saveThresholds(thresholds) {
  try {
    const u=getCurrentUser();if(u.role!=='superadmin')return{ok:false,message:'Unauthorized.'};
    const sh=_ss().getSheetByName(SH_PRIORITY);
    if(sh.getLastRow()>1)sh.getRange(2,1,sh.getLastRow()-1,P._TOTAL).clearContent();
    thresholds.forEach((t,i)=>sh.appendRow([t.amountMin,t.label,i+1]));
    sh.getRange(2,1,Math.max(sh.getLastRow()-1,1)).setNumberFormat('#,##0');
    return{ok:true,message:'Threshold disimpan.'};
  } catch(e){return{ok:false,message:e.message};}
}

// ═══════════════════════════════
//  AUTO PRIORITY
// ═══════════════════════════════
function _autoPriority(amount) {
  try {
    const sh=_ss().getSheetByName(SH_PRIORITY);
    const rows=sh&&sh.getLastRow()>1?sh.getRange(2,1,sh.getLastRow()-1,P._TOTAL).getValues().filter(r=>r[P.AMOUNT_MIN-1]!=='').sort((a,b)=>b[P.AMOUNT_MIN-1]-a[P.AMOUNT_MIN-1]):[[10000000,'High'],[3000000,'Medium'],[0,'Low']];
    for(const r of rows){if(amount>=r[P.AMOUNT_MIN-1])return r[P.LABEL-1]||'Low';}
    return 'Low';
  } catch(e){return 'Medium';}
}

// ═══════════════════════════════
//  READ
// ═══════════════════════════════
function getCases() {
  try {
    const user=getCurrentUser();
    const sheet=_ss().getSheetByName(SH_CASES);
    if(!sheet)return{ok:false,message:'Sheet CASES tidak ditemukan. Jalankan Setup.'};
    const last=sheet.getLastRow();
    if(last<2)return{ok:true,data:[],user};
    let rows=sheet.getRange(2,1,last-1,C._TOTAL).getValues().filter(r=>r[C.CASE_ID-1]!==''&&r[C.CASE_ID-1]!=null).map(_rowToCase);
    if(user.role!=='superadmin'&&user.area&&user.area!=='ALL'){
      rows=rows.filter(c=>String(c.area||'').trim().toLowerCase()===String(user.area||'').trim().toLowerCase());
    }
    return{ok:true,data:rows,user};
  } catch(e){return{ok:false,message:e.message};}
}

function getDashboardStats() {
  try {
    const r=getCases();if(!r.ok)return r;
    const d=r.data;
    return{ok:true,user:r.user,data:{
      total:d.length,open:d.filter(c=>c.status==='Open').length,
      onInvestigation:d.filter(c=>c.status==='On Investigation').length,
      caseClosed:d.filter(c=>c.status==='Case Close').length,
      escalated:d.filter(c=>c.status==='Escalated').length,
      highPriority:d.filter(c=>c.priority==='High'&&c.status!=='Case Close').length,
      totalAmount:d.filter(c=>c.status!=='Case Close').reduce((s,c)=>s+(parseFloat(c.amount)||0),0),
      currentUser:r.user
    }};
  } catch(e){return{ok:false,message:e.message};}
}

function getCaseHistory(caseId) {
  try {
    const sh=_ss().getSheetByName(SH_HISTORY);
    if(!sh||sh.getLastRow()<2)return{ok:true,data:[]};
    return{ok:true,data:sh.getRange(2,1,sh.getLastRow()-1,H._TOTAL).getValues()
      .filter(r=>String(r[H.CASE_ID-1])===String(caseId))
      .map(r=>({timestamp:_fmtDT(r[H.TIMESTAMP-1]),action:r[H.ACTION-1],field:r[H.FIELD-1],oldVal:r[H.OLD_VAL-1],newVal:r[H.NEW_VAL-1],extraJson:r[H.EXTRA_JSON-1],by:r[H.BY-1]}))
      .reverse()};
  } catch(e){return{ok:false,message:e.message};}
}

function getCaseNotes(caseId) {
  try {
    const sh=_ss().getSheetByName(SH_NOTES);
    if(!sh||sh.getLastRow()<2)return{ok:true,data:[]};
    return{ok:true,data:sh.getRange(2,1,sh.getLastRow()-1,N._TOTAL).getValues()
      .filter(r=>String(r[N.CASE_ID-1])===String(caseId))
      .map(r=>({timestamp:_fmtDT(r[N.TIMESTAMP-1]),note:r[N.NOTE-1],by:r[N.BY-1]}))
      .reverse()};
  } catch(e){return{ok:false,message:e.message};}
}

function getCaseEvidence(caseId) {
  try {
    const sh=_ss().getSheetByName(SH_EVIDENCE);
    if(!sh||sh.getLastRow()<2)return{ok:true,data:[]};
    return{ok:true,data:sh.getRange(2,1,sh.getLastRow()-1,EV._TOTAL).getValues()
      .filter(r=>String(r[EV.CASE_ID-1])===String(caseId))
      .map(r=>({type:r[EV.TYPE-1],url:r[EV.URL-1],filename:r[EV.FILENAME-1],description:r[EV.DESCRIPTION-1],addedBy:r[EV.ADDED_BY-1],addedAt:_fmtDT(r[EV.ADDED_AT-1])}))};
  } catch(e){return{ok:false,message:e.message};}
}

// ═══════════════════════════════
//  WRITE — CASE
// ═══════════════════════════════
function addCase(data) {
  try {
    const user=getCurrentUser();
    const sheet=_ss().getSheetByName(SH_CASES);
    if(!sheet)return{ok:false,message:'Sheet CASES tidak ditemukan.'};
    const caseId=_genId(sheet);
    const now=new Date();
    const amt=parseFloat(data.amount)||0;
    const priority=_autoPriority(amt);
    const area=user.role==='superadmin'?(data.area||''):(user.area||'');
    const no=Math.max(sheet.getLastRow(),1);
    sheet.appendRow([
      no,caseId,now,
      data.courierId||'',data.courierName||'',data.hub||'',
      amt,0, // REK_AMT = 0 (lookup, will be filled by formula)
      '','','','','','','','','','','','','','','', // lookup columns 9-23
      data.status||'Open','','', // STATUS, SUB_STATUS, CLOSE_REASON
      data.firstUpdate?new Date(data.firstUpdate):now, // FIRST_UPD
      now,'', // LAST_UPD, CLOSE_DATE
      0, // RUN_DAYS
      priority,data.category||'',area,data.assignedTo||'',
      data.notes||'',user.email,user.email
    ]);
    const nr=sheet.getLastRow();
    sheet.getRange(nr,C.AMOUNT).setNumberFormat('#,##0');
    sheet.getRange(nr,C.REK_AMT).setNumberFormat('#,##0');
    [C.DATE_IN,C.FIRST_UPD,C.LAST_UPD].forEach(col=>sheet.getRange(nr,col).setNumberFormat('dd mmm yyyy'));
    if(data.notes){
      const ns=_ss().getSheetByName(SH_NOTES);
      if(ns)ns.appendRow([now,caseId,data.notes,user.email]);
    }
    _log(caseId,'CREATE','','','','',user.email);
    return{ok:true,caseId,priority,message:'Case '+caseId+' dibuat. Priority: '+priority};
  } catch(e){return{ok:false,message:e.message};}
}

function updateCase(caseId,newData) {
  try {
    const user=getCurrentUser();
    const sheet=_ss().getSheetByName(SH_CASES);
    if(!sheet)return{ok:false,message:'Sheet tidak ditemukan.'};
    const all=sheet.getDataRange().getValues();
    let row=-1,old=null;
    for(let i=1;i<all.length;i++){if(String(all[i][C.CASE_ID-1])===String(caseId)){row=i+1;old=all[i];break;}}
    if(row<0)return{ok:false,message:'Case ID tidak ditemukan.'};
    const isSA=user.role==='superadmin';
    const changes=[];
    // Field map: key → C.XXX (only EDITABLE fields)
    const fm={courierId:C.COURIER_ID,courierName:C.COURIER_NM,hub:C.HUB,amount:C.AMOUNT,
      status:C.STATUS,subStatus:C.SUB_STATUS,closeReason:C.CLOSE_REASON,
      firstUpdate:C.FIRST_UPD,category:C.CATEGORY,area:C.AREA,assignedTo:C.ASSIGNED};
    const df=new Set(['firstUpdate']);
    const saOnly=new Set(['courierId','courierName','area']);
    Object.keys(newData).forEach(k=>{
      if(!(k in fm))return;
      if(saOnly.has(k)&&!isSA)return;
      const col=fm[k];let nv=newData[k];const ov=old[col-1];
      if(nv===undefined||nv===null)return;
      if(df.has(k)&&nv)nv=new Date(nv);
      const os=ov instanceof Date?_fmtD(ov):String(ov??'');
      const ns=nv instanceof Date?_fmtD(nv):String(nv??'');
      if(os!==ns){sheet.getRange(row,col).setValue(nv!==''?nv:'');changes.push({field:k,oldVal:os,newVal:ns});}
    });
    // Auto-recalculate priority if amount changed
    if(newData.amount!==undefined){
      const np=_autoPriority(parseFloat(newData.amount)||0);
      if(String(old[C.PRIORITY-1])!==np){sheet.getRange(row,C.PRIORITY).setValue(np);changes.push({field:'priority',oldVal:String(old[C.PRIORITY-1]||''),newVal:np});}
    }
    const now=new Date();
    sheet.getRange(row,C.LAST_UPD).setValue(now);sheet.getRange(row,C.LAST_UPD).setNumberFormat('dd mmm yyyy');
    sheet.getRange(row,C.UPDATED_BY).setValue(user.email);
    // Recalc running days
    const fu=newData.firstUpdate?new Date(newData.firstUpdate):(old[C.FIRST_UPD-1] instanceof Date?old[C.FIRST_UPD-1]:null);
    const cd=old[C.CLOSE_DATE-1] instanceof Date?old[C.CLOSE_DATE-1]:null;
    const st=newData.status||String(old[C.STATUS-1]||'');
    sheet.getRange(row,C.RUN_DAYS).setValue(_runDays(fu,st,cd));
    changes.forEach(ch=>_log(caseId,'UPDATE',ch.field,ch.oldVal,ch.newVal,'',user.email));
    return{ok:true,message:`Case diupdate. ${changes.length} field diubah.`};
  } catch(e){return{ok:false,message:e.message};}
}

function deleteCase(caseId) {
  try {
    const user=getCurrentUser();if(user.role!=='superadmin')return{ok:false,message:'Unauthorized. Hanya superadmin.'};
    const sheet=_ss().getSheetByName(SH_CASES);
    const all=sheet.getDataRange().getValues();
    for(let i=1;i<all.length;i++){
      if(String(all[i][C.CASE_ID-1])===String(caseId)){sheet.deleteRow(i+1);_log(caseId,'DELETE','','','','',user.email);return{ok:true,message:'Case dihapus.'};}
    }
    return{ok:false,message:'Case tidak ditemukan.'};
  } catch(e){return{ok:false,message:e.message};}
}

// ═══════════════════════════════
//  SUB-STATUS & MOVE STATUS
// ═══════════════════════════════
function applySubStatus(caseId, subStatus, extraData) {
  try {
    const user=getCurrentUser();
    const sheet=_ss().getSheetByName(SH_CASES);
    const all=sheet.getDataRange().getValues();
    for(let i=1;i<all.length;i++){
      if(String(all[i][C.CASE_ID-1])!==String(caseId))continue;
      const row=i+1,now=new Date();
      const curStatus=String(all[i][C.STATUS-1]);
      let newStatus=curStatus;

      // Open sub-statuses → move to On Investigation
      if(OPEN_SUBS.includes(subStatus)){newStatus='On Investigation';}
      // LUNAS (PAID OFF) → move to Case Close
      if(subStatus==='LUNAS (PAID OFF)'){newStatus='Case Close';}

      sheet.getRange(row,C.SUB_STATUS).setValue(subStatus);
      if(newStatus!==curStatus){
        sheet.getRange(row,C.STATUS).setValue(newStatus);
        if(newStatus==='Case Close'){sheet.getRange(row,C.CLOSE_DATE).setValue(now);sheet.getRange(row,C.CLOSE_DATE).setNumberFormat('dd mmm yyyy');}
        _log(caseId,'MOVE_STATUS','status',curStatus,newStatus,'',user.email);
      }
      sheet.getRange(row,C.LAST_UPD).setValue(now);sheet.getRange(row,C.LAST_UPD).setNumberFormat('dd mmm yyyy');
      sheet.getRange(row,C.UPDATED_BY).setValue(user.email);
      const fu=all[i][C.FIRST_UPD-1] instanceof Date?all[i][C.FIRST_UPD-1]:null;
      sheet.getRange(row,C.RUN_DAYS).setValue(_runDays(fu,newStatus,newStatus==='Case Close'?now:null));
      const extra=extraData?JSON.stringify(extraData):'';
      _log(caseId,'SUB_STATUS','subStatus','',subStatus,extra,user.email);
      return{ok:true,message:'Sub-status: '+subStatus+(newStatus!==curStatus?' → Status: '+newStatus:''),newStatus,subStatus};
    }
    return{ok:false,message:'Case tidak ditemukan.'};
  } catch(e){return{ok:false,message:e.message};}
}

function closeCase(caseId, reason, notes) {
  try {
    const user=getCurrentUser();
    const sheet=_ss().getSheetByName(SH_CASES);
    const all=sheet.getDataRange().getValues();
    for(let i=1;i<all.length;i++){
      if(String(all[i][C.CASE_ID-1])!==String(caseId))continue;
      const row=i+1,now=new Date();
      sheet.getRange(row,C.STATUS).setValue('Case Close');
      sheet.getRange(row,C.CLOSE_REASON).setValue(reason||'');
      sheet.getRange(row,C.CLOSE_DATE).setValue(now);sheet.getRange(row,C.CLOSE_DATE).setNumberFormat('dd mmm yyyy');
      sheet.getRange(row,C.LAST_UPD).setValue(now);sheet.getRange(row,C.LAST_UPD).setNumberFormat('dd mmm yyyy');
      sheet.getRange(row,C.UPDATED_BY).setValue(user.email);
      const fu=all[i][C.FIRST_UPD-1] instanceof Date?all[i][C.FIRST_UPD-1]:null;
      sheet.getRange(row,C.RUN_DAYS).setValue(_runDays(fu,'Case Close',now));
      if(notes){
        const ns=_ss().getSheetByName(SH_NOTES);
        if(ns)ns.appendRow([now,caseId,notes,user.email]);
        sheet.getRange(row,C.LAST_NOTE).setValue(notes);
      }
      _log(caseId,'CLOSE_CASE','status',all[i][C.STATUS-1],'Case Close',reason||'',user.email);
      return{ok:true,message:'Case ditutup: '+reason};
    }
    return{ok:false,message:'Case tidak ditemukan.'};
  } catch(e){return{ok:false,message:e.message};}
}

// ═══════════════════════════════
//  NOTES & EVIDENCE
// ═══════════════════════════════
function addNote(caseId, note) {
  try {
    const user=getCurrentUser();
    const now=new Date();
    const ns=_ss().getSheetByName(SH_NOTES);
    if(!ns)return{ok:false,message:'Sheet CASE_NOTES tidak ditemukan.'};
    ns.appendRow([now,caseId,note,user.email]);
    const cs=_ss().getSheetByName(SH_CASES);
    const all=cs.getDataRange().getValues();
    for(let i=1;i<all.length;i++){
      if(String(all[i][C.CASE_ID-1])===String(caseId)){
        cs.getRange(i+1,C.LAST_NOTE).setValue(note);
        cs.getRange(i+1,C.LAST_UPD).setValue(now);cs.getRange(i+1,C.LAST_UPD).setNumberFormat('dd mmm yyyy');
        cs.getRange(i+1,C.UPDATED_BY).setValue(user.email);
        break;
      }
    }
    _log(caseId,'ADD_NOTE','','',note,'',user.email);
    return{ok:true,message:'Note ditambahkan.'};
  } catch(e){return{ok:false,message:e.message};}
}

function addEvidence(caseId, type, url, filename, description) {
  try {
    const user=getCurrentUser();
    const sh=_ss().getSheetByName(SH_EVIDENCE);
    if(!sh)return{ok:false,message:'Sheet CASE_EVIDENCE tidak ditemukan.'};
    sh.appendRow([caseId,type,url,filename||'',description||'',user.email,new Date()]);
    _log(caseId,'ADD_EVIDENCE','','',url,filename||'',user.email);
    return{ok:true,message:'Evidence ditambahkan.'};
  } catch(e){return{ok:false,message:e.message};}
}

// ═══════════════════════════════
//  USERS
// ═══════════════════════════════
function getUsers() {
  try {
    const user=getCurrentUser();if(user.role!=='superadmin')return{ok:false,message:'Unauthorized.'};
    const sh=_ss().getSheetByName(SH_USERS);
    if(!sh||sh.getLastRow()<2)return{ok:true,data:[]};
    return{ok:true,data:sh.getRange(2,1,sh.getLastRow()-1,U._TOTAL).getValues().filter(r=>r[U.EMAIL-1]!=='').map(r=>({email:r[U.EMAIL-1],role:r[U.ROLE-1],area:r[U.AREA-1],addedBy:r[U.ADDED_BY-1],addedAt:_fmtD(r[U.ADDED_AT-1])}))};
  } catch(e){return{ok:false,message:e.message};}
}
function addUser(data) {
  try {
    const user=getCurrentUser();if(user.role!=='superadmin')return{ok:false,message:'Unauthorized.'};
    if(!data.email||!data.role)return{ok:false,message:'Email dan Role wajib diisi.'};
    const sh=_ss().getSheetByName(SH_USERS);
    if(sh.getLastRow()>1){const ex=sh.getRange(2,1,sh.getLastRow()-1,1).getValues().flat();if(ex.some(e=>String(e).toLowerCase()===data.email.toLowerCase()))return{ok:false,message:'Email sudah terdaftar.'};}
    sh.appendRow([data.email,data.role||'user',data.area||'',user.email,new Date()]);
    return{ok:true,message:'User '+data.email+' ditambahkan.'};
  } catch(e){return{ok:false,message:e.message};}
}
function removeUser(email) {
  try {
    const user=getCurrentUser();if(user.role!=='superadmin')return{ok:false,message:'Unauthorized.'};
    const sh=_ss().getSheetByName(SH_USERS);
    const all=sh.getDataRange().getValues();
    for(let i=1;i<all.length;i++){if(String(all[i][U.EMAIL-1]).toLowerCase()===email.toLowerCase()){sh.deleteRow(i+1);return{ok:true,message:'User dihapus.'};}}
    return{ok:false,message:'User tidak ditemukan.'};
  } catch(e){return{ok:false,message:e.message};}
}

// ═══════════════════════════════
//  HELPERS
// ═══════════════════════════════
function _genId(sheet){
  const n=new Date(),yy=n.getFullYear(),mm=String(n.getMonth()+1).padStart(2,'0'),dd=String(n.getDate()).padStart(2,'0');
  const px=`AFP-${yy}${mm}${dd}`;let max=0;
  const last=sheet.getLastRow();
  if(last>1)sheet.getRange(2,C.CASE_ID,last-1).getValues().flat().forEach(id=>{if(String(id).startsWith(px)){const s=parseInt(String(id).split('-').pop())||0;if(s>max)max=s;}});
  return`${px}-${String(max+1).padStart(3,'0')}`;
}
function _log(caseId,action,field,ov,nv,extra,user){
  try{const s=_ss().getSheetByName(SH_HISTORY);if(s)s.appendRow([new Date(),caseId,action,field,ov,nv,extra||'',user||'System']);}catch(e){}
}
function _runDays(first,status,close){
  try{const s=first?new Date(first):new Date();const e=(String(status)==='Case Close'&&close)?new Date(close):new Date();return Math.max(0,Math.floor((e-s)/86400000));}catch(e){return 0;}
}
function _rowToCase(row){
  return{
    no:row[C.NO-1],caseId:row[C.CASE_ID-1],dateIn:_fmtD(row[C.DATE_IN-1]),
    courierId:row[C.COURIER_ID-1],courierName:row[C.COURIER_NM-1],hub:row[C.HUB-1],
    amount:row[C.AMOUNT-1],rekAmt:row[C.REK_AMT-1],
    // lookup
    statPend:row[C.STAT_PEND-1],fdsHub:row[C.FDS_HUB-1],nik:row[C.NIK-1],
    email:row[C.EMAIL_K-1],telepon:row[C.TELEPON-1],alamat:row[C.ALAMAT-1],ktp:row[C.KTP-1],
    namaKondar:row[C.NAMA_KONDAR-1],noKondar:row[C.NO_KONDAR-1],hubKondar:row[C.HUB_KONDAR-1],
    leadRegion:row[C.LEAD_REGION-1],pic:row[C.PIC-1],korlap:row[C.KORLAP-1],city:row[C.CITY-1],province:row[C.PROVINCE-1],
    // case mgmt
    status:row[C.STATUS-1],subStatus:row[C.SUB_STATUS-1],closeReason:row[C.CLOSE_REASON-1],
    firstUpdate:_fmtD(row[C.FIRST_UPD-1]),lastUpdate:_fmtD(row[C.LAST_UPD-1]),closeDate:_fmtD(row[C.CLOSE_DATE-1]),runningDays:row[C.RUN_DAYS-1],
    priority:row[C.PRIORITY-1],category:row[C.CATEGORY-1],area:row[C.AREA-1],assignedTo:row[C.ASSIGNED-1],
    lastNote:row[C.LAST_NOTE-1],createdBy:row[C.CREATED_BY-1],updatedBy:row[C.UPDATED_BY-1]
  };
}
function _fmtD(v){if(!v)return'';if(v instanceof Date){if(isNaN(v.getTime()))return'';return Utilities.formatDate(v,Session.getScriptTimeZone(),'dd MMM yyyy');}return String(v);}
function _fmtDT(v){if(!v)return'';if(v instanceof Date){if(isNaN(v.getTime()))return'';return Utilities.formatDate(v,Session.getScriptTimeZone(),'dd MMM yyyy HH:mm');}return String(v);}
function checkSetup(){try{const ss=_ss();return{ok:true,spreadsheet:ss.getName(),webAppUrl:_getWebAppUrl(),user:Session.getActiveUser().getEmail()};}catch(e){return{ok:false,message:e.message};}}
