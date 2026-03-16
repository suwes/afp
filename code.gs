// ╔══════════════════════════════════════════════════════╗
//  ANTI FRAUD PORTAL v2.0 — Code.gs
//  Deploy: Extensions → Apps Script → Deploy → Web App
//  Execute as: Me | Access: Anyone with Google account
// ╚══════════════════════════════════════════════════════╝

// ── Sheet names
const SH_CASES    = 'CASES';
const SH_HISTORY  = 'CASE_HISTORY';
const SH_NOTES    = 'CASE_NOTES';
const SH_USERS    = 'USERS';
const SH_PRIORITY = 'PRIORITY_CONFIG';

// ── Status flow (immutable order)
const STATUS_FLOW = ['Open', 'On Investigation', 'Case Close'];

// ── CASES columns (1-based)
const C = {
  NO:1, CASE_ID:2, DATE_IN:3, COURIER_ID:4, COURIER_NM:5,
  HUB:6, AMOUNT:7, REK:8, FIRST_UPD:9, STAT_PEND:10,
  STATUS:11, LAST_UPD:12, CLOSE_DATE:13, RUN_DAYS:14,
  PRIORITY:15, CATEGORY:16, ASSIGNED:17, EVIDENCE:18,
  LAST_NOTE:19, CREATED_BY:20, AREA:21, UPDATED_BY:22,
  _TOTAL:22
};

// ── CASE_HISTORY columns (1-based)
const H = {
  TIMESTAMP:1, CASE_ID:2, ACTION:3, FIELD:4,
  OLD_VAL:5, NEW_VAL:6, NOTES:7, BY:8, _TOTAL:8
};

// ── CASE_NOTES columns (1-based)
const N = { TIMESTAMP:1, CASE_ID:2, NOTE:3, BY:4, _TOTAL:4 };

// ── USERS columns (1-based)
const U = { EMAIL:1, ROLE:2, AREA:3, ADDED_BY:4, ADDED_AT:5, _TOTAL:5 };

// ── PRIORITY_CONFIG columns (1-based)
const P = { AMOUNT_MIN:1, LABEL:2, SORT_ORDER:3, _TOTAL:3 };

// ════════════════════════════════════════════
//  SPREADSHEET RESOLVER
// ════════════════════════════════════════════
function _ss() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (ss) return ss;
  } catch(e) {}
  const id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  if (!id) throw new Error('Spreadsheet belum dikonfigurasi. Jalankan Setup dari Google Sheets terlebih dahulu.');
  return SpreadsheetApp.openById(id);
}

// ════════════════════════════════════════════
//  WEB APP ENTRY
// ════════════════════════════════════════════
function doGet(e) {
  return HtmlService
    .createHtmlOutputFromFile('Portal')
    .setTitle('Anti Fraud Portal')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
}

// ════════════════════════════════════════════
//  MENU
// ════════════════════════════════════════════
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🛡️ Anti Fraud Portal')
    .addItem('🌐 Buka Portal (Tab Baru)', 'openPortalInBrowser')
    .addItem('🔗 Lihat URL Web App',      'showWebAppUrl')
    .addSeparator()
    .addItem('🔧 Setup / Inisialisasi',   'setupSheets')
    .addToUi();
}

function openPortalInBrowser() {
  const url = _getWebAppUrl();
  if (!url) { SpreadsheetApp.getUi().alert('Deploy Web App terlebih dahulu.\nDeploy → New deployment → Web app.'); return; }
  const safe = url.replace(/'/g,"\\'");
  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutput(
      `<html><head><style>body{background:#0f0f0f;color:#fff;font-family:sans-serif;display:flex;align-items:center;justify-content:center;flex-direction:column;height:100vh;gap:12px;text-align:center;}a{color:#90caf9;font-weight:700;}</style></head>
       <body><div style="font-size:32px">🛡️</div><p style="color:#aaa;font-size:13px">Membuka Anti Fraud Portal…</p>
       <a href="${safe}" target="_blank">Klik jika tidak terbuka otomatis</a>
       <script>window.open('${safe}','_blank');setTimeout(()=>google.script.host.close(),1500);</script></body></html>`
    ).setWidth(340).setHeight(180), '🛡️ Membuka Portal…'
  );
}

function showWebAppUrl() {
  const url = _getWebAppUrl();
  const ui = SpreadsheetApp.getUi();
  if (!url) { ui.alert('Web App belum di-deploy.'); return; }
  ui.alert('🔗 URL Portal', url, ui.ButtonSet.OK);
}

function _getWebAppUrl() {
  try { const u = ScriptApp.getService().getUrl(); return (u&&u.length>20)?u:null; } catch(e){return null;}
}

// ════════════════════════════════════════════
//  AUTH — GET CURRENT USER
// ════════════════════════════════════════════
function getCurrentUser() {
  const email = Session.getActiveUser().getEmail() || '';
  try {
    const sheet = _ss().getSheetByName(SH_USERS);
    // If USERS sheet doesn't exist or empty → treat as superadmin (setup phase)
    if (!sheet || sheet.getLastRow() < 2) {
      return { email, role: 'superadmin', area: 'ALL', authorized: true };
    }
    const rows = sheet.getRange(2, 1, sheet.getLastRow()-1, U._TOTAL).getValues();
    const found = rows.find(r => String(r[U.EMAIL-1]).trim().toLowerCase() === email.toLowerCase());
    if (!found) return { email, role: null, area: null, authorized: false };
    return { email, role: found[U.ROLE-1], area: found[U.AREA-1], authorized: true };
  } catch(e) {
    return { email, role: 'superadmin', area: 'ALL', authorized: true };
  }
}

// ════════════════════════════════════════════
//  SETUP
// ════════════════════════════════════════════
function setupSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  PropertiesService.getScriptProperties().setProperty('SPREADSHEET_ID', ss.getId());

  // ── CASES
  let cs = ss.getSheetByName(SH_CASES);
  if (cs) {
    const r = ui.alert('Sheet CASES sudah ada.','Reset? (Data hilang!) Pilih No = hanya update config.',ui.ButtonSet.YES_NO);
    if (r === ui.Button.YES) { cs.clearContents(); cs.clearFormats(); cs.clearConditionalFormatRules(); }
    else { _ensureExtras(ss); ui.alert('✅ Config disimpan.\n'+(_getWebAppUrl()||'⚠ Belum deploy Web App.')); return; }
  } else { cs = ss.insertSheet(SH_CASES); }

  _ensureExtras(ss);

  const hdrs=['No','Case ID','Tanggal Input','ID Kurir','Nama Kurir','Hub','Amount',
    'REK Penampung','First Update','Status Pending','Status','Last Update',
    'Case Close Date','Running Days','Priority','Kategori','Assigned To',
    'Evidence Link','Last Note','Created By','Area','Updated By'];
  cs.getRange(1,1,1,hdrs.length).setValues([hdrs]);
  _fmtHdr(cs, hdrs.length, '#0f172a');

  [40,160,110,90,160,130,110,110,100,90,140,110,110,80,80,140,130,160,200,130,100,130]
    .forEach((w,i)=>cs.setColumnWidth(i+1,w));

  _dropdown(cs,C.STATUS,    1000,['Open','On Investigation','Case Close','Escalated']);
  _dropdown(cs,C.PRIORITY,  1000,['High','Medium','Low']);
  _dropdown(cs,C.STAT_PEND, 1000,['LIVE','INACTIVE']);
  _dropdown(cs,C.CATEGORY,  1000,['COD Fraud','Driver Fraud','Data Manipulation','Unauthorized Transaction','Double Collection','Lainnya']);

  const sc={'Open':{bg:'#fef9c3',fc:'#713f12'},'On Investigation':{bg:'#ffedd5',fc:'#7c2d12'},
            'Case Close':{bg:'#dcfce7',fc:'#14532d'},'Escalated':{bg:'#fee2e2',fc:'#7f1d1d'}};
  cs.setConditionalFormatRules(Object.entries(sc).map(([v,c])=>
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(v)
      .setBackground(c.bg).setFontColor(c.fc)
      .setRanges([cs.getRange(2,C.STATUS,1000)]).build()
  ));

  const prClr={'High':{bg:'#fee2e2',fc:'#7f1d1d'},'Medium':{bg:'#fef3c7',fc:'#78350f'},'Low':{bg:'#dcfce7',fc:'#14532d'}};
  const existRules = cs.getConditionalFormatRules();
  Object.entries(prClr).forEach(([v,c])=>{
    existRules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(v)
      .setBackground(c.bg).setFontColor(c.fc)
      .setRanges([cs.getRange(2,C.PRIORITY,1000)]).build());
  });
  cs.setConditionalFormatRules(existRules);

  cs.getRange(2,C.AMOUNT,1000).setNumberFormat('#,##0');
  [C.DATE_IN,C.FIRST_UPD,C.LAST_UPD,C.CLOSE_DATE].forEach(col=>cs.getRange(2,col,1000).setNumberFormat('dd mmm yyyy'));

  ui.alert('✅ Setup Selesai!\n\n'+(_getWebAppUrl()?'🌐 '+_getWebAppUrl():'⚠ Deploy Web App untuk mendapat URL.'));
}

function _ensureExtras(ss) {
  // CASE_HISTORY
  let sh = ss.getSheetByName(SH_HISTORY);
  if (!sh) { sh=ss.insertSheet(SH_HISTORY); sh.getRange(1,1,1,H._TOTAL).setValues([['Timestamp','Case ID','Action','Field','Old Value','New Value','Notes','Changed By']]); _fmtHdr(sh,H._TOTAL,'#0f172a'); }

  // CASE_NOTES
  let sn = ss.getSheetByName(SH_NOTES);
  if (!sn) { sn=ss.insertSheet(SH_NOTES); sn.getRange(1,1,1,N._TOTAL).setValues([['Timestamp','Case ID','Note','By']]); _fmtHdr(sn,N._TOTAL,'#172554'); }

  // USERS
  let su = ss.getSheetByName(SH_USERS);
  if (!su) {
    su=ss.insertSheet(SH_USERS);
    su.getRange(1,1,1,U._TOTAL).setValues([['Email','Role','Area','Added By','Added At']]);
    _fmtHdr(su,U._TOTAL,'#172554');
    su.setColumnWidth(1,220); su.setColumnWidth(2,100); su.setColumnWidth(3,120);
    _dropdown(su,U.ROLE,500,['superadmin','user']);
  }

  // PRIORITY_CONFIG
  let sp = ss.getSheetByName(SH_PRIORITY);
  if (!sp) {
    sp=ss.insertSheet(SH_PRIORITY);
    sp.getRange(1,1,1,P._TOTAL).setValues([['Amount Min (>=)','Priority Label','Sort Order']]);
    _fmtHdr(sp,P._TOTAL,'#172554');
    sp.getRange(2,1,3,3).setValues([[10000000,'High',1],[3000000,'Medium',2],[0,'Low',3]]);
    sp.getRange(2,1,100).setNumberFormat('#,##0');
  }
}

function _fmtHdr(sheet,n,bg){
  sheet.getRange(1,1,1,n).setBackground(bg||'#0f172a').setFontColor('#fff')
    .setFontWeight('bold').setFontSize(10).setHorizontalAlignment('center')
    .setVerticalAlignment('middle').setWrap(false);
  sheet.setRowHeight(1,36); sheet.setFrozenRows(1);
}

function _dropdown(sheet,col,n,vals){
  sheet.getRange(2,col,n).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(vals,true).setAllowInvalid(false).build());
}

// ════════════════════════════════════════════
//  AUTO PRIORITY
// ════════════════════════════════════════════
function _autoPriority(amount) {
  try {
    const sheet = _ss().getSheetByName(SH_PRIORITY);
    if (!sheet || sheet.getLastRow() < 2) {
      if (amount >= 10000000) return 'High';
      if (amount >= 3000000) return 'Medium';
      return 'Low';
    }
    const rows = sheet.getRange(2,1,sheet.getLastRow()-1,P._TOTAL).getValues()
      .filter(r => r[P.AMOUNT_MIN-1] !== '' && !isNaN(r[P.AMOUNT_MIN-1]))
      .sort((a,b)=>b[P.AMOUNT_MIN-1]-a[P.AMOUNT_MIN-1]);
    for (const row of rows) {
      if (amount >= row[P.AMOUNT_MIN-1]) return row[P.LABEL-1] || 'Low';
    }
    return 'Low';
  } catch(e) { return 'Medium'; }
}

function getThresholds() {
  try {
    const sheet = _ss().getSheetByName(SH_PRIORITY);
    if (!sheet || sheet.getLastRow() < 2) return { ok:true, data:[{amountMin:10000000,label:'High',sortOrder:1},{amountMin:3000000,label:'Medium',sortOrder:2},{amountMin:0,label:'Low',sortOrder:3}]};
    return { ok:true, data: sheet.getRange(2,1,sheet.getLastRow()-1,P._TOTAL).getValues()
      .filter(r=>r[P.AMOUNT_MIN-1]!=='')
      .map(r=>({amountMin:r[P.AMOUNT_MIN-1],label:r[P.LABEL-1],sortOrder:r[P.SORT_ORDER-1]}))
      .sort((a,b)=>b.amountMin-a.amountMin)};
  } catch(e){return{ok:false,message:e.message};}
}

function saveThresholds(thresholds) {
  try {
    const user = getCurrentUser();
    if (user.role !== 'superadmin') return {ok:false,message:'Unauthorized.'};
    const sheet = _ss().getSheetByName(SH_PRIORITY);
    if (sheet.getLastRow() > 1) sheet.getRange(2,1,sheet.getLastRow()-1,P._TOTAL).clearContent();
    thresholds.forEach((t,i)=>sheet.appendRow([t.amountMin,t.label,i+1]));
    sheet.getRange(2,1,Math.max(sheet.getLastRow()-1,1)).setNumberFormat('#,##0');
    return {ok:true,message:'Threshold disimpan.'};
  } catch(e){return{ok:false,message:e.message};}
}

// ════════════════════════════════════════════
//  READ — CASES
// ════════════════════════════════════════════
function getCases() {
  try {
    const user = getCurrentUser();
    const sheet = _ss().getSheetByName(SH_CASES);
    if (!sheet) return {ok:false,message:'Sheet CASES tidak ditemukan. Jalankan Setup.'};
    const last = sheet.getLastRow();
    if (last < 2) return {ok:true,data:[],user};
    let rows = sheet.getRange(2,1,last-1,C._TOTAL).getValues()
      .filter(r=>r[C.CASE_ID-1]!==''&&r[C.CASE_ID-1]!=null)
      .map(_rowToCase);
    if (user.role!=='superadmin' && user.area && user.area!=='ALL') {
      rows = rows.filter(c=>String(c.area||'').trim().toLowerCase()===String(user.area||'').trim().toLowerCase());
    }
    return {ok:true,data:rows,user};
  } catch(e){return{ok:false,message:e.message};}
}

function getDashboardStats() {
  try {
    const r = getCases(); if(!r.ok) return r;
    const d=r.data;
    const priOrder={'High':3,'Medium':2,'Low':1,'':0};
    return {ok:true, user:r.user, data:{
      total:d.length,
      open:d.filter(c=>c.status==='Open').length,
      onInvestigation:d.filter(c=>c.status==='On Investigation').length,
      caseClosed:d.filter(c=>c.status==='Case Close').length,
      escalated:d.filter(c=>c.status==='Escalated').length,
      highPriority:d.filter(c=>c.priority==='High'&&c.status!=='Case Close').length,
      totalAmount:d.filter(c=>c.status!=='Case Close').reduce((s,c)=>s+(parseFloat(c.amount)||0),0),
      recoveredAmount:d.filter(c=>c.status==='Case Close').reduce((s,c)=>s+(parseFloat(c.amount)||0),0),
      currentUser:r.user
    }};
  } catch(e){return{ok:false,message:e.message};}
}

function getAreaStats() {
  try {
    const r = getCases(); if(!r.ok) return r;
    const m={};
    r.data.forEach(c=>{
      const a=String(c.area||'Unassigned').trim()||'Unassigned';
      if(!m[a])m[a]={total:0,open:0,invest:0,closed:0,escalated:0,amount:0};
      m[a].total++;
      if(c.status==='Open')m[a].open++;
      else if(c.status==='On Investigation')m[a].invest++;
      else if(c.status==='Case Close')m[a].closed++;
      else if(c.status==='Escalated')m[a].escalated++;
      m[a].amount+=parseFloat(c.amount)||0;
    });
    return {ok:true,data:m};
  } catch(e){return{ok:false,message:e.message};}
}

// ════════════════════════════════════════════
//  READ — HISTORY & NOTES
// ════════════════════════════════════════════
function getCaseHistory(caseId) {
  try {
    const sheet=_ss().getSheetByName(SH_HISTORY);
    if(!sheet||sheet.getLastRow()<2)return{ok:true,data:[]};
    return {ok:true, data:
      sheet.getRange(2,1,sheet.getLastRow()-1,H._TOTAL).getValues()
        .filter(r=>String(r[H.CASE_ID-1])===String(caseId))
        .map(r=>({timestamp:_fmtDT(r[H.TIMESTAMP-1]),action:r[H.ACTION-1],
          field:r[H.FIELD-1],oldVal:r[H.OLD_VAL-1],newVal:r[H.NEW_VAL-1],
          notes:r[H.NOTES-1],by:r[H.BY-1]}))
        .reverse()
    };
  } catch(e){return{ok:false,message:e.message};}
}

function getCaseNotes(caseId) {
  try {
    const sheet=_ss().getSheetByName(SH_NOTES);
    if(!sheet||sheet.getLastRow()<2)return{ok:true,data:[]};
    return {ok:true, data:
      sheet.getRange(2,1,sheet.getLastRow()-1,N._TOTAL).getValues()
        .filter(r=>String(r[N.CASE_ID-1])===String(caseId))
        .map(r=>({timestamp:_fmtDT(r[N.TIMESTAMP-1]),note:r[N.NOTE-1],by:r[N.BY-1]}))
        .reverse()
    };
  } catch(e){return{ok:false,message:e.message};}
}

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
        cs.getRange(i+1,C.LAST_UPD).setValue(now);
        cs.getRange(i+1,C.UPDATED_BY).setValue(user.email);
        break;
      }
    }
    _log(caseId,'ADD_NOTE','','',note,note,user.email);
    return{ok:true,message:'Note berhasil ditambahkan.'};
  } catch(e){return{ok:false,message:e.message};}
}

// ════════════════════════════════════════════
//  WRITE — ADD / UPDATE / DELETE CASE
// ════════════════════════════════════════════
function addCase(data) {
  try {
    const user=getCurrentUser();
    const sheet=_ss().getSheetByName(SH_CASES);
    if(!sheet)return{ok:false,message:'Sheet CASES tidak ditemukan.'};
    const caseId=_genId(sheet);
    const now=new Date();
    const amt=parseFloat(data.amount)||0;
    const priority=_autoPriority(amt);
    const no=Math.max(sheet.getLastRow(),1);
    const area = user.role==='superadmin' ? (data.area||'') : (user.area||'');
    sheet.appendRow([no,caseId,now,
      data.courierId||'',data.courierName||'',data.hub||'',
      amt,data.rek||'',
      data.firstUpdate?new Date(data.firstUpdate):now,
      data.statusPending||'LIVE',data.status||'Open',now,
      data.closeDate?new Date(data.closeDate):'',
      _runDays(data.firstUpdate,'Open',null),
      priority,data.category||'',data.assignedTo||'',data.evidence||'',
      data.notes||'',user.email,area,user.email]);
    const nr=sheet.getLastRow();
    sheet.getRange(nr,C.AMOUNT).setNumberFormat('#,##0');
    [C.DATE_IN,C.FIRST_UPD,C.LAST_UPD].forEach(col=>sheet.getRange(nr,col).setNumberFormat('dd mmm yyyy'));
    if(data.closeDate)sheet.getRange(nr,C.CLOSE_DATE).setNumberFormat('dd mmm yyyy');
    if(data.notes){const ns=_ss().getSheetByName(SH_NOTES);if(ns)ns.appendRow([now,caseId,data.notes,user.email]);}
    _log(caseId,'CREATE','','','',data.notes||'',user.email);
    return{ok:true,caseId,priority,message:'Case '+caseId+' berhasil dibuat. Priority: '+priority};
  } catch(e){return{ok:false,message:e.message};}
}

function updateCase(caseId, newData) {
  try {
    const user=getCurrentUser();
    const sheet=_ss().getSheetByName(SH_CASES);
    if(!sheet)return{ok:false,message:'Sheet CASES tidak ditemukan.'};
    const all=sheet.getDataRange().getValues();
    let row=-1,old=null;
    for(let i=1;i<all.length;i++){if(String(all[i][C.CASE_ID-1])===String(caseId)){row=i+1;old=all[i];break;}}
    if(row<0)return{ok:false,message:'Case ID tidak ditemukan.'};

    // Permission check for restricted fields
    const isSuperadmin=user.role==='superadmin';
    const changes=[];
    const fm={courierId:'COURIER_ID',courierName:'COURIER_NM',hub:'HUB',amount:'AMOUNT',rek:'REK',
      firstUpdate:'FIRST_UPD',statusPending:'STAT_PEND',status:'STATUS',closeDate:'CLOSE_DATE',
      priority:'PRIORITY',category:'CATEGORY',assignedTo:'ASSIGNED',evidence:'EVIDENCE',
      lastNote:'LAST_NOTE',area:'AREA'};
    const df=new Set(['firstUpdate','closeDate']);
    const adminOnly=new Set(['courierId','courierName','area']);

    Object.keys(newData).forEach(k=>{
      if(!(k in fm))return;
      if(adminOnly.has(k)&&!isSuperadmin)return;
      const col=C[fm[k]];let nv=newData[k];const ov=old[col-1];
      if(nv===undefined||nv===null)return;
      if(df.has(k)&&nv)nv=new Date(nv);
      const os=ov instanceof Date?_fmtD(ov):String(ov??'');
      const ns=nv instanceof Date?_fmtD(nv):String(nv??'');
      if(os!==ns){sheet.getRange(row,col).setValue(nv!==''?nv:'');changes.push({field:k,oldVal:os,newVal:ns});}
    });

    // Auto-recalculate priority if amount changed
    if(newData.amount!==undefined){
      const newPriority=_autoPriority(parseFloat(newData.amount)||0);
      if(String(old[C.PRIORITY-1])!==newPriority){
        sheet.getRange(row,C.PRIORITY).setValue(newPriority);
        changes.push({field:'priority',oldVal:String(old[C.PRIORITY-1]||''),newVal:newPriority});
      }
    }

    const now=new Date();
    sheet.getRange(row,C.LAST_UPD).setValue(now);
    sheet.getRange(row,C.UPDATED_BY).setValue(user.email);
    const cd=newData.closeDate?new Date(newData.closeDate):(old[C.CLOSE_DATE-1] instanceof Date?old[C.CLOSE_DATE-1]:null);
    const fu=newData.firstUpdate?new Date(newData.firstUpdate):(old[C.FIRST_UPD-1] instanceof Date?old[C.FIRST_UPD-1]:null);
    sheet.getRange(row,C.RUN_DAYS).setValue(_runDays(fu,newData.status||String(old[C.STATUS-1]||''),cd));
    changes.forEach(ch=>_log(caseId,'UPDATE',ch.field,ch.oldVal,ch.newVal,'',user.email));
    return{ok:true,message:'Case diupdate. '+changes.length+' field diubah.'};
  } catch(e){return{ok:false,message:e.message};}
}

function deleteCase(caseId) {
  try {
    const user=getCurrentUser();
    if(user.role!=='superadmin')return{ok:false,message:'Unauthorized. Hanya superadmin.'};
    const sheet=_ss().getSheetByName(SH_CASES);
    const all=sheet.getDataRange().getValues();
    for(let i=1;i<all.length;i++){
      if(String(all[i][C.CASE_ID-1])===String(caseId)){sheet.deleteRow(i+1);_log(caseId,'DELETE','','','','',user.email);return{ok:true,message:'Case dihapus.'};}
    }
    return{ok:false,message:'Case tidak ditemukan.'};
  } catch(e){return{ok:false,message:e.message};}
}

// ════════════════════════════════════════════
//  MOVE STATUS
// ════════════════════════════════════════════
function moveStatus(caseId) {
  try {
    const user=getCurrentUser();
    const sheet=_ss().getSheetByName(SH_CASES);
    const all=sheet.getDataRange().getValues();
    for(let i=1;i<all.length;i++){
      if(String(all[i][C.CASE_ID-1])===String(caseId)){
        const cur=all[i][C.STATUS-1];
        const idx=STATUS_FLOW.indexOf(cur);
        if(idx<0||idx>=STATUS_FLOW.length-1)return{ok:false,message:'Status sudah di tahap akhir (Case Close).'};
        const next=STATUS_FLOW[idx+1];
        const row=i+1;const now=new Date();
        sheet.getRange(row,C.STATUS).setValue(next);
        sheet.getRange(row,C.LAST_UPD).setValue(now);
        sheet.getRange(row,C.UPDATED_BY).setValue(user.email);
        if(next==='Case Close'){sheet.getRange(row,C.CLOSE_DATE).setValue(now);sheet.getRange(row,C.CLOSE_DATE).setNumberFormat('dd mmm yyyy');}
        const fu=all[i][C.FIRST_UPD-1];
        const cd=next==='Case Close'?now:null;
        sheet.getRange(row,C.RUN_DAYS).setValue(_runDays(fu,next,cd));
        _log(caseId,'MOVE_STATUS','status',cur,next,'',user.email);
        return{ok:true,message:cur+' → '+next,oldStatus:cur,newStatus:next};
      }
    }
    return{ok:false,message:'Case tidak ditemukan.'};
  } catch(e){return{ok:false,message:e.message};}
}

// ════════════════════════════════════════════
//  USERS MANAGEMENT
// ════════════════════════════════════════════
function getUsers() {
  try {
    const user=getCurrentUser();
    if(user.role!=='superadmin')return{ok:false,message:'Unauthorized.'};
    const sheet=_ss().getSheetByName(SH_USERS);
    if(!sheet||sheet.getLastRow()<2)return{ok:true,data:[]};
    return{ok:true,data:sheet.getRange(2,1,sheet.getLastRow()-1,U._TOTAL).getValues()
      .filter(r=>r[U.EMAIL-1]!=='')
      .map(r=>({email:r[U.EMAIL-1],role:r[U.ROLE-1],area:r[U.AREA-1],addedBy:r[U.ADDED_BY-1],addedAt:_fmtD(r[U.ADDED_AT-1])}))};
  } catch(e){return{ok:false,message:e.message};}
}

function addUser(data) {
  try {
    const user=getCurrentUser();
    if(user.role!=='superadmin')return{ok:false,message:'Unauthorized.'};
    if(!data.email||!data.role)return{ok:false,message:'Email dan Role wajib diisi.'};
    const sheet=_ss().getSheetByName(SH_USERS);
    // Check duplicate
    if(sheet.getLastRow()>1){
      const existing=sheet.getRange(2,1,sheet.getLastRow()-1,1).getValues().flat();
      if(existing.some(e=>String(e).toLowerCase()===data.email.toLowerCase()))return{ok:false,message:'Email sudah terdaftar.'};
    }
    sheet.appendRow([data.email,data.role||'user',data.area||'',user.email,new Date()]);
    return{ok:true,message:'User '+data.email+' berhasil ditambahkan.'};
  } catch(e){return{ok:false,message:e.message};}
}

function removeUser(email) {
  try {
    const user=getCurrentUser();
    if(user.role!=='superadmin')return{ok:false,message:'Unauthorized.'};
    const sheet=_ss().getSheetByName(SH_USERS);
    const all=sheet.getDataRange().getValues();
    for(let i=1;i<all.length;i++){
      if(String(all[i][U.EMAIL-1]).toLowerCase()===email.toLowerCase()){sheet.deleteRow(i+1);return{ok:true,message:'User dihapus.'};}
    }
    return{ok:false,message:'User tidak ditemukan.'};
  } catch(e){return{ok:false,message:e.message};}
}

// ════════════════════════════════════════════
//  HELPERS
// ════════════════════════════════════════════
function _genId(sheet){
  const n=new Date(),yy=n.getFullYear(),mm=String(n.getMonth()+1).padStart(2,'0'),dd=String(n.getDate()).padStart(2,'0');
  const px=`AFP-${yy}${mm}${dd}`;let max=0;
  const last=sheet.getLastRow();
  if(last>1)sheet.getRange(2,C.CASE_ID,last-1).getValues().flat().forEach(id=>{
    if(String(id).startsWith(px)){const s=parseInt(String(id).split('-').pop())||0;if(s>max)max=s;}
  });
  return`${px}-${String(max+1).padStart(3,'0')}`;
}

function _log(caseId,action,field,ov,nv,notes,user){
  try{const s=_ss().getSheetByName(SH_HISTORY);if(s)s.appendRow([new Date(),caseId,action,field,ov,nv,notes,user||'System']);}catch(e){}
}

function _runDays(first,status,close){
  try{const s=first?new Date(first):new Date();const e=(String(status)==='Case Close'&&close)?new Date(close):new Date();return Math.max(0,Math.floor((e-s)/86400000));}catch(e){return 0;}
}

function _rowToCase(row){
  return{
    no:row[C.NO-1],caseId:row[C.CASE_ID-1],
    dateIn:_fmtD(row[C.DATE_IN-1]),
    courierId:row[C.COURIER_ID-1],courierName:row[C.COURIER_NM-1],hub:row[C.HUB-1],
    amount:row[C.AMOUNT-1],rek:row[C.REK-1],
    firstUpdate:_fmtD(row[C.FIRST_UPD-1]),
    statusPending:row[C.STAT_PEND-1],status:row[C.STATUS-1],
    lastUpdate:_fmtD(row[C.LAST_UPD-1]),closeDate:_fmtD(row[C.CLOSE_DATE-1]),
    runningDays:row[C.RUN_DAYS-1],priority:row[C.PRIORITY-1],
    category:row[C.CATEGORY-1],assignedTo:row[C.ASSIGNED-1],
    evidence:row[C.EVIDENCE-1],lastNote:row[C.LAST_NOTE-1],
    createdBy:row[C.CREATED_BY-1],area:row[C.AREA-1],updatedBy:row[C.UPDATED_BY-1]
  };
}

function _fmtD(v){
  if(!v)return'';
  if(v instanceof Date){if(isNaN(v.getTime()))return'';return Utilities.formatDate(v,Session.getScriptTimeZone(),'dd MMM yyyy');}
  return String(v);
}

function _fmtDT(v){
  if(!v)return'';
  if(v instanceof Date){if(isNaN(v.getTime()))return'';return Utilities.formatDate(v,Session.getScriptTimeZone(),'dd MMM yyyy HH:mm');}
  return String(v);
}

function checkSetup(){
  try{
    const ss=_ss(),sheets=ss.getSheets().map(s=>s.getName());
    return{ok:true,spreadsheet:ss.getName(),sheets,webAppUrl:_getWebAppUrl(),user:Session.getActiveUser().getEmail()};
  }catch(e){return{ok:false,message:e.message};}
}
