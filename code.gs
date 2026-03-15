// ╔══════════════════════════════════════════════════════════════╗
//  ANTI FRAUD PORTAL — Apps Script Backend
//  File : Code.gs
//
//  ── CARA DEPLOY SEBAGAI WEB APP ───────────────────────────────
//
//  LANGKAH 1 — Pertama kali:
//    • Extensions → Apps Script
//    • Paste Code.gs ini, buat file HTML "Portal", paste Portal.html
//    • Simpan (Ctrl+S) → kembali ke Sheets → reload halaman
//    • Menu "🛡️ Anti Fraud Portal" → "🔧 Setup / Inisialisasi Sheets"
//
//  LANGKAH 2 — Deploy:
//    • Di Apps Script editor → klik tombol biru "Deploy" (kanan atas)
//    • Pilih "New deployment"
//    • Klik ikon ⚙️ → pilih type "Web app"
//    • Execute as      : Me  (email Anda)
//    • Who has access  : Anyone with Google account
//    • Klik "Deploy" → authorize → COPY URL-nya → bookmark!
//
//  LANGKAH 3 — Buka Portal:
//    • Kembali ke Sheets → menu "🌐 Buka Portal (Tab Baru)"
//    • Atau langsung paste URL tadi di browser → full screen!
//
//  UPDATE KODE:
//    Deploy → "Manage deployments" → ✏️ Edit
//    → Version: "New version" → Update. URL tetap sama.
// ╚══════════════════════════════════════════════════════════════╝

// ── Sheet names
const SH_CASES   = 'CASES';
const SH_HISTORY = 'CASE_HISTORY';

// ── Column map — CASES (1-based)
const C = {
  NO:1, CASE_ID:2, DATE_IN:3, COURIER_ID:4, COURIER_NM:5,
  HUB:6, AMOUNT:7, REK:8, FIRST_UPD:9, STAT_PEND:10,
  STATUS:11, LAST_UPD:12, CLOSE_DATE:13, RUN_DAYS:14,
  PRIORITY:15, CATEGORY:16, ASSIGNED:17, EVIDENCE:18,
  NOTES:19, CREATED_BY:20, _TOTAL:20
};

// ── Column map — CASE_HISTORY (1-based)
const H = {
  TIMESTAMP:1, CASE_ID:2, ACTION:3, FIELD:4,
  OLD_VAL:5, NEW_VAL:6, NOTES:7, BY:8, _TOTAL:8
};

// ════════════════════════════════════════════
//  RESOLVE SPREADSHEET
//  Saat dipanggil dari Sheets  → getActiveSpreadsheet() tersedia
//  Saat dipanggil dari Web App → tidak ada "active" spreadsheet,
//  baca ID yang disimpan saat Setup via PropertiesService
// ════════════════════════════════════════════
function _getSpreadsheet() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (ss) return ss;
  } catch(e) {}
  const id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  if (!id) throw new Error(
    'Spreadsheet belum dikonfigurasi. Buka Google Sheets → ' +
    'menu "🛡️ Anti Fraud Portal" → "🔧 Setup" terlebih dahulu.'
  );
  return SpreadsheetApp.openById(id);
}

// ════════════════════════════════════════════
//  WEB APP ENTRY POINT — doGet()
//  Fungsi ini WAJIB ada. Dipanggil otomatis
//  ketika URL web app dibuka di browser.
// ════════════════════════════════════════════
function doGet(e) {
  return HtmlService
    .createHtmlOutputFromFile('Portal')
    .setTitle('Anti Fraud Portal — Risk Management')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
}

// ════════════════════════════════════════════
//  MENU GOOGLE SHEETS
// ════════════════════════════════════════════
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🛡️ Anti Fraud Portal')
    .addItem('🌐 Buka Portal (Tab Baru)',      'openPortalInBrowser')
    .addItem('🔗 Lihat / Copy URL Web App',    'showWebAppUrl')
    .addSeparator()
    .addItem('🔧 Setup / Inisialisasi Sheets', 'setupSheets')
    .addToUi();
}

// ────────────────────────────────────────────
//  openPortalInBrowser()
//  Tampilkan mini-dialog yang memicu window.open()
//  di sisi klien → tab baru terbuka dengan URL web app
// ────────────────────────────────────────────
function openPortalInBrowser() {
  const url = _getWebAppUrl();

  if (!url) {
    SpreadsheetApp.getUi().alert(
      '⚠️  Web App Belum Di-deploy',
      'Portal belum bisa dibuka.\n\n' +
      'Cara deploy:\n' +
      '  1. Apps Script editor → klik "Deploy"\n' +
      '  2. Pilih "New deployment"\n' +
      '  3. Type: Web app\n' +
      '  4. Execute as: Me\n' +
      '  5. Who has access: Anyone with Google account\n' +
      '  6. Klik Deploy → authorize → salin URL\n\n' +
      'Setelah deploy, klik menu ini lagi.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return;
  }

  const safeUrl = url.replace(/'/g, "\\'");
  const html = HtmlService.createHtmlOutput(`
    <!DOCTYPE html><html>
    <head><style>
      *{box-sizing:border-box;margin:0;padding:0;}
      body{background:#0d0d10;color:#fafafa;font-family:-apple-system,sans-serif;
           display:flex;flex-direction:column;align-items:center;justify-content:center;
           height:100vh;gap:14px;text-align:center;padding:24px;}
      .icon{font-size:38px;}
      .title{font-size:14px;font-weight:700;color:#e2e8f0;}
      .sub{font-size:11px;color:#71717a;line-height:1.7;}
      .link{color:#c7d2fe;font-size:12px;font-weight:600;text-decoration:none;}
      .link:hover{text-decoration:underline;}
      .bar{width:180px;height:4px;background:#27272a;border-radius:99px;overflow:hidden;}
      .fill{height:100%;background:linear-gradient(90deg,#4f46e5,#7c3aed);
            animation:grow 1.5s ease forwards;}
      @keyframes grow{from{width:0}to{width:100%}}
    </style></head>
    <body>
      <div class="icon">🛡️</div>
      <div class="title">Membuka Anti Fraud Portal…</div>
      <div class="bar"><div class="fill"></div></div>
      <div class="sub">Jika tidak terbuka otomatis,<br>pop-up blocker mungkin aktif.</div>
      <a class="link" href="${safeUrl}" target="_blank">↗ Klik di sini untuk buka manual</a>
      <script>
        window.open('${safeUrl}','_blank');
        setTimeout(function(){try{google.script.host.close();}catch(e){}},1500);
      </script>
    </body></html>
  `).setWidth(340).setHeight(210);

  SpreadsheetApp.getUi().showModalDialog(html, '🛡️ Membuka Portal…');
}

function showWebAppUrl() {
  const url = _getWebAppUrl();
  const ui  = SpreadsheetApp.getUi();
  if (!url) {
    ui.alert('Web App belum di-deploy.\n\nGunakan: Deploy → New deployment → Web app.');
    return;
  }
  ui.alert('🔗 URL Anti Fraud Portal', url +
    '\n\nBookmark URL ini untuk akses langsung tanpa membuka Sheets.',
    ui.ButtonSet.OK);
}

function _getWebAppUrl() {
  try {
    const url = ScriptApp.getService().getUrl();
    return (url && url.length > 20) ? url : null;
  } catch(e) { return null; }
}

// ════════════════════════════════════════════
//  SETUP SHEETS
// ════════════════════════════════════════════
function setupSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  // WAJIB: Simpan ID spreadsheet agar web app bisa akses tanpa active context
  PropertiesService.getScriptProperties().setProperty('SPREADSHEET_ID', ss.getId());

  let cs = ss.getSheetByName(SH_CASES);
  if (cs) {
    const resp = ui.alert('Sheet CASES sudah ada.',
      'Reset & buat ulang? (Data AKAN HILANG)\nPilih "No" untuk hanya simpan ulang config.',
      ui.ButtonSet.YES_NO);
    if (resp !== ui.Button.YES) {
      _ensureHistorySheet(ss);
      ui.alert('✅ Config disimpan.\n\n' +
        (_getWebAppUrl() ? '🌐 URL: ' + _getWebAppUrl() : '⚠ Belum di-deploy sebagai Web App.'));
      return;
    }
    cs.clearContents(); cs.clearFormats(); cs.clearConditionalFormatRules();
  } else {
    cs = ss.insertSheet(SH_CASES);
  }

  _ensureHistorySheet(ss);

  const hdrs = ['No','Case ID','Tanggal Input','ID Kurir','Nama Kurir','Hub','Amount',
    'REK Penampung','First Update','Status Pending','Status','Last Update',
    'Case Close Date','Running Days','Priority','Kategori','Assigned To',
    'Evidence Link','Notes Update','Created By'];
  cs.getRange(1,1,1,hdrs.length).setValues([hdrs]);
  _fmtHdr(cs, hdrs.length, '#0f3460');

  [40,150,100,90,160,130,110,110,100,90,130,100,110,80,80,130,120,140,220,120]
    .forEach((w,i) => cs.setColumnWidth(i+1, w));

  _dropdown(cs, C.STATUS,    1000, ['Open','On Investigation','Case Close','Escalated']);
  _dropdown(cs, C.PRIORITY,  1000, ['High','Medium','Low']);
  _dropdown(cs, C.STAT_PEND, 1000, ['LIVE','INACTIVE']);
  _dropdown(cs, C.CATEGORY,  1000, ['COD Fraud','Driver Fraud','Data Manipulation',
    'Unauthorized Transaction','Double Collection','Lainnya']);

  const sc = {'Open':{bg:'#fff9c4',fc:'#6d5c00'},'On Investigation':{bg:'#ffe0b2',fc:'#7d3a00'},
              'Case Close':{bg:'#c8e6c9',fc:'#1b5e20'},'Escalated':{bg:'#ffcdd2',fc:'#7f0000'}};
  cs.setConditionalFormatRules(Object.entries(sc).map(([v,c]) =>
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo(v).setBackground(c.bg).setFontColor(c.fc)
      .setRanges([cs.getRange(2,C.STATUS,1000)]).build()));

  cs.getRange(2,C.AMOUNT,1000).setNumberFormat('#,##0');
  [C.DATE_IN,C.FIRST_UPD,C.LAST_UPD,C.CLOSE_DATE].forEach(col =>
    cs.getRange(2,col,1000).setNumberFormat('dd mmm yyyy'));

  const webUrl = _getWebAppUrl();
  ui.alert('✅ Setup Selesai!',
    'Sheet CASES & CASE_HISTORY berhasil disiapkan.\n\n' +
    (webUrl ? '🌐 URL Portal:\n' + webUrl
            : '⚠ Deploy sebagai Web App untuk mendapat URL.'), ui.ButtonSet.OK);
}

function _ensureHistorySheet(ss) {
  let sh = ss.getSheetByName(SH_HISTORY);
  if (!sh) {
    sh = ss.insertSheet(SH_HISTORY);
    const h=['Timestamp','Case ID','Action','Field','Old Value','New Value','Notes','Changed By'];
    sh.getRange(1,1,1,h.length).setValues([h]);
    _fmtHdr(sh, h.length, '#1a2744');
  }
  return sh;
}

function _fmtHdr(sheet, n, bg) {
  sheet.getRange(1,1,1,n)
    .setBackground(bg||'#0f3460').setFontColor('#fff').setFontWeight('bold')
    .setFontSize(10).setHorizontalAlignment('center').setVerticalAlignment('middle').setWrap(false);
  sheet.setRowHeight(1,36); sheet.setFrozenRows(1);
}

function _dropdown(sheet, col, n, vals) {
  sheet.getRange(2,col,n).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(vals,true).setAllowInvalid(false).build());
}

// ════════════════════════════════════════════
//  READ — ALL CASES
// ════════════════════════════════════════════
function getCases() {
  try {
    const sheet = _getSpreadsheet().getSheetByName(SH_CASES);
    if (!sheet) return {ok:false, message:'Sheet CASES tidak ditemukan. Jalankan Setup.'};
    const last = sheet.getLastRow();
    if (last < 2) return {ok:true, data:[]};
    return {ok:true, data:
      sheet.getRange(2,1,last-1,C._TOTAL).getValues()
        .filter(r => r[C.CASE_ID-1] !== '' && r[C.CASE_ID-1] != null)
        .map(_rowToCase)};
  } catch(e) { return {ok:false, message:e.message}; }
}

// ════════════════════════════════════════════
//  READ — CASE HISTORY
// ════════════════════════════════════════════
function getCaseHistory(caseId) {
  try {
    const sheet = _getSpreadsheet().getSheetByName(SH_HISTORY);
    if (!sheet) return {ok:true, data:[]};
    const last = sheet.getLastRow();
    if (last < 2) return {ok:true, data:[]};
    return {ok:true, data:
      sheet.getRange(2,1,last-1,H._TOTAL).getValues()
        .filter(r => String(r[H.CASE_ID-1]) === String(caseId))
        .map(r => ({
          timestamp:_fmtDT(r[H.TIMESTAMP-1]), caseId:r[H.CASE_ID-1],
          action:r[H.ACTION-1], field:r[H.FIELD-1],
          oldVal:r[H.OLD_VAL-1], newVal:r[H.NEW_VAL-1],
          notes:r[H.NOTES-1], by:r[H.BY-1]
        })).reverse()};
  } catch(e) { return {ok:false, message:e.message}; }
}

// ════════════════════════════════════════════
//  READ — DASHBOARD STATS
// ════════════════════════════════════════════
function getDashboardStats() {
  try {
    const r = getCases(); if (!r.ok) return r;
    const d = r.data;
    return {ok:true, data:{
      total           : d.length,
      open            : d.filter(c=>c.status==='Open').length,
      onInvestigation : d.filter(c=>c.status==='On Investigation').length,
      caseClosed      : d.filter(c=>c.status==='Case Close').length,
      escalated       : d.filter(c=>c.status==='Escalated').length,
      highPriority    : d.filter(c=>c.priority==='High').length,
      totalAmount     : d.filter(c=>c.status!=='Case Close').reduce((s,c)=>s+(parseFloat(c.amount)||0),0),
      recoveredAmount : d.filter(c=>c.status==='Case Close').reduce((s,c)=>s+(parseFloat(c.amount)||0),0),
      currentUser     : Session.getActiveUser().getEmail()||'—'
    }};
  } catch(e) { return {ok:false, message:e.message}; }
}

// ════════════════════════════════════════════
//  CHECK SETUP
// ════════════════════════════════════════════
function checkSetup() {
  try {
    const ss=_getSpreadsheet(), sheets=ss.getSheets().map(s=>s.getName());
    return {ok:sheets.includes(SH_CASES)&&sheets.includes(SH_HISTORY),
      spreadsheet:ss.getName(), sheets,
      hasCases:sheets.includes(SH_CASES), hasHistory:sheets.includes(SH_HISTORY),
      webAppUrl:_getWebAppUrl(), user:Session.getActiveUser().getEmail(),
      message:!sheets.includes(SH_CASES)?'Sheet CASES tidak ada — jalankan Setup.'
             :!sheets.includes(SH_HISTORY)?'CASE_HISTORY tidak ada — jalankan Setup.':'Semua OK.'};
  } catch(e) { return {ok:false, message:e.message}; }
}

// ════════════════════════════════════════════
//  WRITE — ADD CASE
// ════════════════════════════════════════════
function addCase(data) {
  try {
    const sheet = _getSpreadsheet().getSheetByName(SH_CASES);
    if (!sheet) return {ok:false, message:'Sheet CASES tidak ditemukan.'};
    const caseId=_genId(sheet), now=new Date(), no=Math.max(sheet.getLastRow(),1);
    sheet.appendRow([no,caseId,now,
      data.courierId||'',data.courierName||'',data.hub||'',
      parseFloat(data.amount)||0,data.rek||'',
      data.firstUpdate?new Date(data.firstUpdate):now,
      data.statusPending||'LIVE',data.status||'Open',now,
      data.closeDate?new Date(data.closeDate):'',
      _runDays(data.firstUpdate,data.status,data.closeDate),
      data.priority||'Medium',data.category||'',data.assignedTo||'',
      data.evidence||'',data.notes||'',
      Session.getActiveUser().getEmail()||'System']);
    const nr=sheet.getLastRow();
    sheet.getRange(nr,C.AMOUNT).setNumberFormat('#,##0');
    [C.DATE_IN,C.FIRST_UPD,C.LAST_UPD].forEach(col=>sheet.getRange(nr,col).setNumberFormat('dd mmm yyyy'));
    if(data.closeDate)sheet.getRange(nr,C.CLOSE_DATE).setNumberFormat('dd mmm yyyy');
    _log(caseId,'CREATE','','','',data.notes||'',Session.getActiveUser().getEmail());
    return {ok:true,caseId,message:'Case berhasil ditambahkan.'};
  } catch(e){return {ok:false,message:e.message};}
}

// ════════════════════════════════════════════
//  WRITE — UPDATE CASE
// ════════════════════════════════════════════
function updateCase(caseId, newData) {
  try {
    const sheet=_getSpreadsheet().getSheetByName(SH_CASES);
    if(!sheet)return{ok:false,message:'Sheet CASES tidak ditemukan.'};
    const all=sheet.getDataRange().getValues(); let row=-1,old=null;
    for(let i=1;i<all.length;i++){
      if(String(all[i][C.CASE_ID-1])===String(caseId)){row=i+1;old=all[i];break;}
    }
    if(row<0)return{ok:false,message:'Case ID tidak ditemukan.'};
    const user=Session.getActiveUser().getEmail()||'System', changes=[];
    const fm={courierId:'COURIER_ID',courierName:'COURIER_NM',hub:'HUB',amount:'AMOUNT',rek:'REK',
      firstUpdate:'FIRST_UPD',statusPending:'STAT_PEND',status:'STATUS',closeDate:'CLOSE_DATE',
      priority:'PRIORITY',category:'CATEGORY',assignedTo:'ASSIGNED',evidence:'EVIDENCE',notes:'NOTES'};
    const df=new Set(['firstUpdate','closeDate']);
    Object.keys(newData).forEach(k=>{
      if(!(k in fm))return;
      const col=C[fm[k]];let nv=newData[k];const ov=old[col-1];
      if(nv===undefined||nv===null)return;
      if(df.has(k)&&nv)nv=new Date(nv);
      const os=ov instanceof Date?_fmtD(ov):String(ov??'');
      const ns=nv instanceof Date?_fmtD(nv):String(nv??'');
      if(os!==ns){sheet.getRange(row,col).setValue(nv!==''?nv:'');changes.push({field:k,oldVal:os,newVal:ns});}
    });
    sheet.getRange(row,C.LAST_UPD).setValue(new Date());
    const cd=newData.closeDate?new Date(newData.closeDate):(old[C.CLOSE_DATE-1] instanceof Date?old[C.CLOSE_DATE-1]:null);
    const fu=newData.firstUpdate?new Date(newData.firstUpdate):(old[C.FIRST_UPD-1] instanceof Date?old[C.FIRST_UPD-1]:null);
    sheet.getRange(row,C.RUN_DAYS).setValue(_runDays(fu,newData.status||String(old[C.STATUS-1]||''),cd));
    changes.forEach(ch=>_log(caseId,'UPDATE',ch.field,ch.oldVal,ch.newVal,newData.notes||'',user));
    if(!changes.length&&newData.notes)_log(caseId,'UPDATE','notes',String(old[C.NOTES-1]||''),newData.notes,newData.notes,user);
    return{ok:true,message:`Case diupdate. ${changes.length} field diubah.`};
  } catch(e){return{ok:false,message:e.message};}
}

// ════════════════════════════════════════════
//  PRIVATE HELPERS
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
  try{const s=_getSpreadsheet().getSheetByName(SH_HISTORY);
    if(s)s.appendRow([new Date(),caseId,action,field,ov,nv,notes,user||'System']);}catch(e){}
}
function _runDays(first,status,close){
  try{const s=first?new Date(first):new Date();
    const e=(String(status)==='Case Close'&&close)?new Date(close):new Date();
    return Math.max(0,Math.floor((e-s)/86400000));}catch(e){return 0;}
}
function _rowToCase(row){
  return{no:row[C.NO-1],caseId:row[C.CASE_ID-1],dateIn:_fmtD(row[C.DATE_IN-1]),
    courierId:row[C.COURIER_ID-1],courierName:row[C.COURIER_NM-1],hub:row[C.HUB-1],
    amount:row[C.AMOUNT-1],rek:row[C.REK-1],firstUpdate:_fmtD(row[C.FIRST_UPD-1]),
    statusPending:row[C.STAT_PEND-1],status:row[C.STATUS-1],lastUpdate:_fmtD(row[C.LAST_UPD-1]),
    closeDate:_fmtD(row[C.CLOSE_DATE-1]),runningDays:row[C.RUN_DAYS-1],
    priority:row[C.PRIORITY-1],category:row[C.CATEGORY-1],assignedTo:row[C.ASSIGNED-1],
    evidence:row[C.EVIDENCE-1],notes:row[C.NOTES-1],createdBy:row[C.CREATED_BY-1]};
}
function _fmtD(v){
  if(!v)return'';
  if(v instanceof Date){if(isNaN(v.getTime()))return'';
    return Utilities.formatDate(v,Session.getScriptTimeZone(),'dd MMM yyyy');}
  return String(v);
}
function _fmtDT(v){
  if(!v)return'';
  if(v instanceof Date){if(isNaN(v.getTime()))return'';
    return Utilities.formatDate(v,Session.getScriptTimeZone(),'dd MMM yyyy HH:mm');}
  return String(v);
}
