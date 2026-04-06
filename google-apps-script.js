/**
 * ============================================================
 * KNNDCmdb – Google Apps Script Backend  v1.3
 * ============================================================
 * DEPLOYMENT GUIDE:
 *
 * ── FIRST TIME SETUP ─────────────────────────────────────
 * 1. Open your Google Sheet → Extensions → Apps Script
 * 2. Delete all existing code, paste this entire file
 * 3. Click 💾 Save (name: "KNNDCmdb Backend")
 * 4. Run → setupSpreadsheet   (creates all tabs with correct columns)
 *    → Authorise when prompted → Click Allow
 *    → You'll see "Setup complete!" alert
 * 5. Deploy → New Deployment → Type: Web App
 *    → Execute as: Me  |  Who has access: Anyone
 *    → Click Deploy → Authorise → Copy the Web App URL
 * 6. Paste the URL into KNNDCmdb → Settings → Google Sheets → Script URL
 *
 * ── UPGRADING FROM v1.1 or v1.2 (adds Zone & Gender) ────
 * 1. Paste this file over the old one in Apps Script
 * 2. Run → migrateAddZoneAndGender
 *    → This adds Zone column to Polling Stations sheet
 *    → And Gender column to Members Database sheet
 *    → Safe to run multiple times (skips if columns exist)
 * 3. Re-deploy: Deploy → Manage Deployments → Edit → New version → Deploy
 * ============================================================
 */

const SHEETS = {
  MEMBERS:          'Members Database',
  POLLING_STATIONS: 'Polling Stations',
  USERS:            'Users',
  AUDIT:            'Audit Log',
  SUMMARY:          'Summary',
};

const NDC_GREEN  = '#1a6b3a';
const NDC_GREEN2 = '#134d2a';
const NDC_RED    = '#c8102e';
const WHITE      = '#ffffff';
const LIGHT_GRN  = '#e8f5ec';


// ════════════════════════════════════════════════════════════
//  ONE-TIME SETUP — Run this once after pasting code
// ════════════════════════════════════════════════════════════
function setupSpreadsheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.setSpreadsheetName('KNNDCmdb – Ketu North NDC Members Database');
  _setupMembersSheet(ss);
  _setupPollingStationsSheet(ss);
  _setupUsersSheet(ss);
  _setupAuditSheet(ss);
  _setupSummarySheet(ss);
  const order = [SHEETS.MEMBERS, SHEETS.POLLING_STATIONS, SHEETS.USERS, SHEETS.AUDIT, SHEETS.SUMMARY];
  order.forEach((name,i)=>{ const s=ss.getSheetByName(name); if(s) ss.setActiveSheet(s).moveActiveSheet(i+1); });
  Logger.log('✅ Setup complete!');
  SpreadsheetApp.getUi().alert('✅ KNNDCmdb v1.3 Setup Complete!\n\nTabs created:\n• Members Database (with Zone & Gender columns)\n• Polling Stations (with Zone column)\n• Users\n• Audit Log\n• Summary\n\nNext: Deploy as Web App.');
}


// ════════════════════════════════════════════════════════════
//  MIGRATION — Run this if upgrading from v1.1 or v1.2
//  Adds Zone to Polling Stations + Gender to Members Database
// ════════════════════════════════════════════════════════════
function migrateAddZoneAndGender() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const log = [];

  // ── 1. Polling Stations: Add Zone column ──
  const psSheet = ss.getSheetByName(SHEETS.POLLING_STATIONS);
  if (psSheet) {
    const psHdrs = psSheet.getRange(2, 1, 1, psSheet.getLastColumn()).getValues()[0];
    if (!psHdrs.includes('Zone')) {
      // Insert Zone as column A (shift existing right)
      psSheet.insertColumnBefore(1);
      psSheet.getRange(2, 1).setValue('Zone').setFontWeight('bold').setBackground(NDC_GREEN2).setFontColor(WHITE);
      psSheet.setColumnWidth(1, 120);
      log.push('✅ Added "Zone" column to Polling Stations (column A).');
      log.push('   ⚠️  Please fill in the Zone values for each station.');
    } else {
      log.push('ℹ️  Polling Stations already has a "Zone" column — skipped.');
    }
  } else {
    log.push('⚠️  Polling Stations sheet not found — running full setup for it.');
    _setupPollingStationsSheet(ss);
  }

  // ── 2. Members Database: Add Gender column after Telephone Number ──
  const mSheet = ss.getSheetByName(SHEETS.MEMBERS);
  if (mSheet) {
    const mHdrs = mSheet.getRange(4, 1, 1, mSheet.getLastColumn()).getValues()[0];
    if (!mHdrs.includes('Gender')) {
      // Find position of "Telephone Number" (col E, index 4) — insert after it
      const phoneIdx = mHdrs.indexOf('Telephone Number');
      const insertAfterCol = phoneIdx >= 0 ? phoneIdx + 2 : 6; // 1-based
      mSheet.insertColumnAfter(insertAfterCol);
      const genderCol = insertAfterCol + 1;
      mSheet.getRange(4, genderCol).setValue('Gender').setFontWeight('bold').setBackground(NDC_GREEN).setFontColor(WHITE);
      mSheet.setColumnWidth(genderCol, 100);
      // Also add Zone if missing
      if (!mHdrs.includes('Zone')) {
        mSheet.insertColumnAfter(genderCol);
        const zoneCol = genderCol + 1;
        mSheet.getRange(4, zoneCol).setValue('Zone').setFontWeight('bold').setBackground(NDC_GREEN).setFontColor(WHITE);
        mSheet.setColumnWidth(zoneCol, 110);
        log.push('✅ Added "Gender" and "Zone" columns to Members Database.');
      } else {
        log.push('✅ Added "Gender" column to Members Database.');
      }
    } else {
      log.push('ℹ️  Members Database already has a "Gender" column — skipped.');
      // Check Zone separately
      if (!mHdrs.includes('Zone')) {
        const genderIdx = mHdrs.indexOf('Gender');
        mSheet.insertColumnAfter(genderIdx + 1);
        const zoneCol = genderIdx + 2;
        mSheet.getRange(4, zoneCol).setValue('Zone').setFontWeight('bold').setBackground(NDC_GREEN).setFontColor(WHITE);
        mSheet.setColumnWidth(zoneCol, 110);
        log.push('✅ Added "Zone" column to Members Database.');
      } else {
        log.push('ℹ️  Members Database already has a "Zone" column — skipped.');
      }
    }
  } else {
    log.push('⚠️  Members Database sheet not found — run setupSpreadsheet() first.');
  }

  const summary = log.join('\n');
  Logger.log(summary);
  SpreadsheetApp.getUi().alert('Migration Complete!\n\n' + summary + '\n\nDone. Remember to re-deploy your Web App after this.');
}


// ════════════════════════════════════════════════════════════
//  SHEET SETUP HELPERS
// ════════════════════════════════════════════════════════════

function _setupMembersSheet(ss) {
  let sheet = ss.getSheetByName(SHEETS.MEMBERS);
  if (!sheet) sheet = ss.insertSheet(SHEETS.MEMBERS);
  sheet.clear();

  // Row 1: Title (matches Excel template format)
  sheet.getRange('A1:E1').merge().setValue('MEMBERSHIP DATABASE')
    .setFontFamily('Arial').setFontSize(16).setFontWeight('bold')
    .setHorizontalAlignment('center').setBackground(NDC_GREEN2).setFontColor(WHITE);
  sheet.getRange('A2').setValue('Polling Station / Branch Name:').setFontWeight('bold');
  sheet.getRange('E2').setValue('Constituency: Ketu North').setHorizontalAlignment('right');

  // Row 4: Headers — core 5 (matching Excel template) + Gender + Zone + tracking
  const coreHdrs  = ['First Name','Surname','Party ID Number','Voter ID Number','Telephone Number','Gender','Zone'];
  const extraHdrs = ['Ward Name','Polling Station','Station Code','Branch Name','Branch Code','Other Names','Officer ID','Officer Name','Date/Time Added','Record ID'];
  const allHdrs   = [...coreHdrs, ...extraHdrs];

  const hdrRange = sheet.getRange(4, 1, 1, allHdrs.length);
  hdrRange.setValues([allHdrs]).setFontWeight('bold').setFontColor(WHITE).setBackground(NDC_GREEN);
  sheet.getRange(4, 1, 1, 5).setBackground(NDC_GREEN2); // darker for core 5 cols

  sheet.setFrozenRows(4);
  sheet.setFrozenColumns(2);

  const widths = [130,130,160,160,150,100,110, 120,150,100,140,110,130,100,160,180,160];
  widths.forEach((w,i) => sheet.setColumnWidth(i+1, w));
  sheet.setTabColor(NDC_GREEN);
}

function _setupPollingStationsSheet(ss) {
  let sheet = ss.getSheetByName(SHEETS.POLLING_STATIONS);
  if (!sheet) sheet = ss.insertSheet(SHEETS.POLLING_STATIONS);
  sheet.clear();

  sheet.getRange('A1:F1').merge().setValue('POLLING STATIONS REGISTER')
    .setFontSize(14).setFontWeight('bold').setHorizontalAlignment('center')
    .setBackground(NDC_GREEN).setFontColor(WHITE);

  const hdrs = ['Zone','Ward Name','Polling Station Name','Branch Name','Station Code','Branch Code'];
  sheet.getRange(2, 1, 1, hdrs.length).setValues([hdrs])
    .setFontWeight('bold').setBackground(NDC_GREEN2).setFontColor(WHITE);

  const demo = [
    ['Zone A','Aflao Ward',    'Aflao A Polling Station',   'Aflao Branch',    'PS-001','BR-001'],
    ['Zone A','Aflao Ward',    'Aflao B Polling Station',   'Aflao Branch',    'PS-002','BR-001'],
    ['Zone B','Denu Ward',     'Denu Polling Station',      'Denu Branch',     'PS-003','BR-002'],
    ['Zone B','Agbozume Ward', 'Agbozume Polling Station',  'Agbozume Branch', 'PS-004','BR-003'],
    ['Zone C','Klikor Ward',   'Klikor Polling Station',    'Klikor Branch',   'PS-005','BR-004'],
    ['Zone C','Adafienu Ward', 'Adafienu Polling Station',  'Adafienu Branch', 'PS-006','BR-005'],
  ];
  sheet.getRange(3, 1, demo.length, 6).setValues(demo);
  sheet.setFrozenRows(2);
  [120,150,200,160,120,120].forEach((w,i)=>sheet.setColumnWidth(i+1,w));
  sheet.setTabColor('#2563eb');
}

function _setupUsersSheet(ss) {
  let sheet = ss.getSheetByName(SHEETS.USERS);
  if (!sheet) sheet = ss.insertSheet(SHEETS.USERS);
  sheet.clear();
  sheet.getRange('A1:J1').merge().setValue('SYSTEM USERS')
    .setFontSize(14).setFontWeight('bold').setHorizontalAlignment('center').setBackground(NDC_RED).setFontColor(WHITE);
  const hdrs=['User ID','Full Name','Username','Password','Role','Ward','Station Code','Branch','Assigned Stations','Status'];
  sheet.getRange(2,1,1,hdrs.length).setValues([hdrs]).setFontWeight('bold').setBackground('#7f1d1d').setFontColor(WHITE);
  sheet.setFrozenRows(2);
  [80,180,120,120,80,120,100,130,150,80].forEach((w,i)=>sheet.setColumnWidth(i+1,w));
  sheet.setTabColor(NDC_RED);
}

function _setupAuditSheet(ss) {
  let sheet = ss.getSheetByName(SHEETS.AUDIT);
  if (!sheet) sheet = ss.insertSheet(SHEETS.AUDIT);
  sheet.clear();
  sheet.getRange('A1:E1').merge().setValue('SYSTEM AUDIT LOG')
    .setFontSize(14).setFontWeight('bold').setHorizontalAlignment('center').setBackground('#1e3a5f').setFontColor(WHITE);
  sheet.getRange(2,1,1,5).setValues([['Timestamp','Action','User','Details','Extra']]).setFontWeight('bold').setBackground('#1e3a5f').setFontColor(WHITE);
  sheet.setFrozenRows(2);
  [160,140,120,350,200].forEach((w,i)=>sheet.setColumnWidth(i+1,w));
  sheet.setTabColor('#1e3a5f');
}

function _setupSummarySheet(ss) {
  let sheet = ss.getSheetByName(SHEETS.SUMMARY);
  if (!sheet) sheet = ss.insertSheet(SHEETS.SUMMARY);
  sheet.clear();
  sheet.getRange('A1:D1').merge().setValue('MEMBERSHIP SUMMARY BY STATION')
    .setFontSize(14).setFontWeight('bold').setHorizontalAlignment('center').setBackground(NDC_GREEN).setFontColor(WHITE);
  sheet.getRange('A2').setValue('Auto-updated on export from KNNDCmdb.').setFontColor('#6b7280').setFontStyle('italic');
  sheet.getRange(3,1,1,4).setValues([['Ward','Polling Station','Branch','Total Members']]).setFontWeight('bold').setBackground(NDC_GREEN2).setFontColor(WHITE);
  sheet.setFrozenRows(3);
  [160,220,160,120].forEach((w,i)=>sheet.setColumnWidth(i+1,w));
  sheet.setTabColor('#d97706');
}


// ════════════════════════════════════════════════════════════
//  HTTP HANDLERS
// ════════════════════════════════════════════════════════════

function doGet(e) {
  const action = e?.parameter?.action || 'ping';
  try {
    if (action==='getMembers')        return _json(_getMembers());
    if (action==='getStations')       return _json(_getPollingStations());
    if (action==='ping')              return _json({status:'ok',app:'KNNDCmdb',version:'1.3'});
    return _json({error:'Unknown action'});
  } catch(err) { return _json({error:err.message}); }
}

function doPost(e) {
  try {
    const data=JSON.parse(e.postData.contents);
    const action=data.action||'addMember';
    if (action==='addMember')    return _json(_addMember(data));
    if (action==='updateMember') return _json(_updateMember(data));
    if (action==='deleteMember') return _json(_deleteMember(data));
    if (action==='logAudit')     return _json(_logAudit(data));
    if (action==='syncSummary')  return _json(_refreshSummary(SpreadsheetApp.getActiveSpreadsheet()));
    return _json(_addMember(data)); // default
  } catch(err) { return _json({success:false,error:err.message}); }
}

function _json(data) {
  return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON);
}


// ════════════════════════════════════════════════════════════
//  CRUD
// ════════════════════════════════════════════════════════════

function _addMember(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEETS.MEMBERS);
  if (!sheet) { setupSpreadsheet(); sheet = ss.getSheetByName(SHEETS.MEMBERS); }

  // Build row matching header order in _setupMembersSheet
  // Core: First Name, Surname, Party ID, Voter ID, Phone, Gender, Zone
  // Extra: Ward, Station, StationCode, Branch, BranchCode, OtherNames, OfficerID, OfficerName, Timestamp, RecordID
  sheet.appendRow([
    data.firstName   || '',
    data.lastName    || '',
    data.partyId     || '',
    data.voterId     || '',
    data.phone       || '',
    data.gender      || '',   // NEW
    data.zone        || '',   // NEW
    data.ward        || '',
    data.station     || '',
    data.stationCode || '',
    data.branch      || '',
    data.branchCode  || '',
    data.otherNames  || '',
    data.officer     || '',
    data.officerName || '',
    data.timestamp   || new Date().toLocaleString(),
    data.id          || 'm' + Date.now(),
  ]);
  _refreshSummary(ss);
  return { success:true };
}

function _updateMember(data) {
  const ss=SpreadsheetApp.getActiveSpreadsheet();
  const sheet=ss.getSheetByName(SHEETS.MEMBERS); if(!sheet) return{success:false,error:'Sheet not found'};
  const rows=sheet.getDataRange().getValues();
  // Record ID is column 17 (index 16, header row 4 = data start row 5)
  for(let i=4;i<rows.length;i++){
    if(String(rows[i][16])===String(data.id)){
      const r=i+1;
      if(data.firstName)  sheet.getRange(r,1).setValue(data.firstName);
      if(data.lastName)   sheet.getRange(r,2).setValue(data.lastName);
      if(data.partyId)    sheet.getRange(r,3).setValue(data.partyId);
      if(data.voterId)    sheet.getRange(r,4).setValue(data.voterId);
      if(data.phone)      sheet.getRange(r,5).setValue(data.phone);
      if(data.gender)     sheet.getRange(r,6).setValue(data.gender);
      return{success:true};
    }
  }
  return{success:false,error:'Record not found'};
}

function _deleteMember(data) {
  const ss=SpreadsheetApp.getActiveSpreadsheet();
  const sheet=ss.getSheetByName(SHEETS.MEMBERS); if(!sheet) return{success:false,error:'Sheet not found'};
  const rows=sheet.getDataRange().getValues();
  for(let i=4;i<rows.length;i++){
    if(String(rows[i][16])===String(data.id)){sheet.deleteRow(i+1);_refreshSummary(ss);return{success:true};}
  }
  return{success:false,error:'Record not found'};
}

function _logAudit(data) {
  const ss=SpreadsheetApp.getActiveSpreadsheet();
  let sheet=ss.getSheetByName(SHEETS.AUDIT);
  if(!sheet){_setupAuditSheet(ss);sheet=ss.getSheetByName(SHEETS.AUDIT);}
  sheet.appendRow([data.timestamp||new Date().toLocaleString(),data.action||'',data.user||'',data.details||'',data.extra||'']);
  return{success:true};
}

function _getMembers() {
  const ss=SpreadsheetApp.getActiveSpreadsheet();
  const sheet=ss.getSheetByName(SHEETS.MEMBERS); if(!sheet) return{members:[]};
  const all=sheet.getDataRange().getValues();
  if(all.length<5) return{members:[]};
  const headers=all[3]; // row 4 = index 3
  const members=[];
  for(let i=4;i<all.length;i++){
    const row=all[i];
    // Skip completely empty rows
    if(!row[0]&&!row[1]&&!row[16]) continue;
    const m={};
    headers.forEach((h,j)=>{ m[h]=row[j]; });
    members.push(m);
  }
  return{members,total:members.length};
}

function _getPollingStations() {
  const ss=SpreadsheetApp.getActiveSpreadsheet();
  const sheet=ss.getSheetByName(SHEETS.POLLING_STATIONS); if(!sheet) return{stations:[]};
  const rows=sheet.getDataRange().getValues();
  // Headers row 2 (index 1), data from row 3 (index 2)
  const hdrs=rows[1];
  const stations=[];
  for(let i=2;i<rows.length;i++){
    const r=rows[i]; if(!r[0]&&!r[1]) continue;
    // Build station object from named columns
    const s={};
    hdrs.forEach((h,j)=>{ s[h]=r[j]; });
    // Normalise keys
    stations.push({
      zone:       s['Zone']||s[0]||'',
      ward:       s['Ward Name']||s[1]||'',
      name:       s['Polling Station Name']||s[2]||'',
      branch:     s['Branch Name']||s[3]||'',
      code:       s['Station Code']||s[4]||'',
      branchCode: s['Branch Code']||s[5]||'',
    });
  }
  return{stations};
}

function _refreshSummary(ss) {
  let sheet=ss.getSheetByName(SHEETS.SUMMARY);
  if(!sheet){_setupSummarySheet(ss);sheet=ss.getSheetByName(SHEETS.SUMMARY);}
  const lastRow=sheet.getLastRow();
  if(lastRow>=4) sheet.getRange(4,1,lastRow-3,4).clearContent();
  const mSheet=ss.getSheetByName(SHEETS.MEMBERS); if(!mSheet) return{success:true};
  const all=mSheet.getDataRange().getValues();
  const byStation={};
  for(let i=4;i<all.length;i++){
    const r=all[i]; if(!r[8]) continue; // station col index 8 (Polling Station)
    const key=r[8];
    if(!byStation[key]) byStation[key]={ward:r[7]||'',station:r[8]||'',branch:r[10]||'',count:0};
    byStation[key].count++;
  }
  const rows=Object.values(byStation).sort((a,b)=>b.count-a.count);
  const total=rows.reduce((s,r)=>s+r.count,0);
  const toWrite=rows.map(r=>[r.ward,r.station,r.branch,r.count]);
  toWrite.push(['','','GRAND TOTAL',total]);
  if(toWrite.length) sheet.getRange(4,1,toWrite.length,4).setValues(toWrite);
  const gtRow=4+toWrite.length-1;
  sheet.getRange(gtRow,1,1,4).setFontWeight('bold').setBackground(LIGHT_GRN);
  sheet.getRange('D2').setValue('Updated: '+new Date().toLocaleString()).setFontColor('#6b7280');
  return{success:true};
}
