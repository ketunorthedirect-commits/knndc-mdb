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
  SETTINGS:         'App Settings',   // ← new: stores shared settings
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
  ss.rename('KNNDCmdb – Ketu North NDC Members Database');
  _setupMembersSheet(ss);
  _setupPollingStationsSheet(ss);
  _setupUsersSheet(ss);
  _setupAuditSheet(ss);
  _setupSummarySheet(ss);
  _setupAppSettingsSheet(ss);   // ← new

  const order = [SHEETS.MEMBERS, SHEETS.POLLING_STATIONS, SHEETS.USERS, SHEETS.AUDIT, SHEETS.SUMMARY, SHEETS.SETTINGS];
  order.forEach((name, i) => {
    const s = ss.getSheetByName(name);
    if (s) { ss.setActiveSheet(s); ss.moveActiveSheet(i + 1); }
  });
  Logger.log('✅ Setup complete!');
  SpreadsheetApp.getUi().alert(
    '✅ KNNDCmdb v1.9 Setup Complete!\n\n' +
    'Tabs created:\n' +
    '• Members Database\n• Polling Stations\n• Users\n• Audit Log\n• Summary\n• App Settings\n\n' +
    'Next: Deploy as Web App.\n\n' +
    'IMPORTANT: After deploying, paste the Web App URL into the App Settings tab (cell B2), ' +
    'then save it from the web app Settings page — all future logins on any device will ' +
    'automatically receive the correct settings.'
  );
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
//  MIGRATION v1.9 — Add App Settings tab to existing sheets
//  Run if upgrading from v1.8 or earlier
// ════════════════════════════════════════════════════════════
function migrateAddSettingsSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (ss.getSheetByName(SHEETS.SETTINGS)) {
    SpreadsheetApp.getUi().alert('ℹ️  App Settings tab already exists — nothing to do.');
    return;
  }
  _setupAppSettingsSheet(ss);
  SpreadsheetApp.getUi().alert(
    '✅ App Settings tab added!\n\n' +
    'Next: create a New Deployment for this script (Deploy → New Deployment).\n\n' +
    'Then log in as admin → Settings → Google Sheets → 💾 Save. ' +
    'All future logins on any device will automatically receive the correct settings.'
  );
}


// ─── App Settings Sheet ──────────────────────────────────────
function _setupAppSettingsSheet(ss) {
  let sheet = ss.getSheetByName(SHEETS.SETTINGS);
  if (!sheet) sheet = ss.insertSheet(SHEETS.SETTINGS);
  sheet.clear();

  sheet.getRange('A1').setValue('KNNDCmdb — Shared App Settings')
    .setFontWeight('bold').setFontSize(13).setBackground(NDC_GREEN).setFontColor(WHITE);
  sheet.getRange('B1').setValue('').setBackground(NDC_GREEN);
  sheet.getRange('A2').setValue('Last Updated By').setFontWeight('bold');
  sheet.getRange('B2').setValue('').setBackground(LIGHT_GRN);

  const rows = [
    ['scriptUrl',    ''],
    ['appName',      'Ketu North NDC Members Database'],
    ['constituency', 'Ketu North'],
    ['sheetId',      ss.getId()],
    ['apiKey',       ''],
    ['demoCleared',  ''],
  ];
  sheet.getRange(3, 1, rows.length, 2).setValues(rows);
  sheet.getRange(3, 1, rows.length, 1).setFontWeight('bold').setBackground(LIGHT_GRN);
  sheet.getRange(3, 2, rows.length, 1).setBackground(WHITE);
  sheet.setColumnWidth(1, 180);
  sheet.setColumnWidth(2, 420);
  sheet.setFrozenRows(2);
  sheet.setTabColor('#7c3aed');
  Logger.log('✅ App Settings sheet created.');
}


// ════════════════════════════════════════════════════════════
//  SHEET SETUP HELPERS
// ════════════════════════════════════════════════════════════

function _setupMembersSheet(ss) {
  let sheet = ss.getSheetByName(SHEETS.MEMBERS);
  if (!sheet) sheet = ss.insertSheet(SHEETS.MEMBERS);
  sheet.clear();

  // Row 1: Title — merge only columns A:A (no merge conflict with freeze)
  sheet.getRange('A1').setValue('MEMBERSHIP DATABASE')
    .setFontFamily('Arial').setFontSize(16).setFontWeight('bold')
    .setHorizontalAlignment('left').setBackground(NDC_GREEN2).setFontColor(WHITE);
  sheet.getRange('B1').setValue('').setBackground(NDC_GREEN2);
  sheet.getRange('C1').setValue('').setBackground(NDC_GREEN2);
  sheet.getRange('D1').setValue('').setBackground(NDC_GREEN2);
  sheet.getRange('E1').setValue('').setBackground(NDC_GREEN2);

  sheet.getRange('A2').setValue('Polling Station / Branch Name:').setFontWeight('bold');
  sheet.getRange('E2').setValue('Constituency: Ketu North').setHorizontalAlignment('right');

  // Row 4: Headers — core 5 (matching Excel template) + Gender + Zone + tracking
  const coreHdrs  = ['First Name','Surname','Party ID Number','Voter ID Number','Telephone Number','Gender','Zone'];
  const extraHdrs = ['Ward Name','Polling Station','Station Code','Branch Name','Branch Code','Other Names','Officer ID','Officer Name','Date/Time Added','Record ID'];
  const allHdrs   = [...coreHdrs, ...extraHdrs];

  const hdrRange = sheet.getRange(4, 1, 1, allHdrs.length);
  hdrRange.setValues([allHdrs]).setFontWeight('bold').setFontColor(WHITE).setBackground(NDC_GREEN);
  sheet.getRange(4, 1, 1, 5).setBackground(NDC_GREEN2); // darker for core 5 cols

  // Freeze rows only — no column freeze (avoids merge conflict)
  sheet.setFrozenRows(4);

  const widths = [130,130,160,160,150,100,110,120,150,100,140,110,130,100,160,180,160];
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
  sheet.getRange('A1:L1').merge().setValue('SYSTEM USERS')
    .setFontSize(14).setFontWeight('bold').setHorizontalAlignment('center').setBackground(NDC_RED).setFontColor(WHITE);
  const hdrs = [
    'User ID','Full Name','Username','Password','Role',
    'Ward','Station Code','Branch','Assigned Stations',
    'Status','Must Change Password','Last Updated'
  ];
  sheet.getRange(2,1,1,hdrs.length).setValues([hdrs])
    .setFontWeight('bold').setBackground('#7f1d1d').setFontColor(WHITE);
  sheet.setFrozenRows(2);
  [90,180,130,130,90,130,110,140,200,80,140,160].forEach((w,i)=>sheet.setColumnWidth(i+1,w));
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
    if (action==='getMembers')   return _json(_getMembers());
    if (action==='getStations')  return _json(_getPollingStations());
    if (action==='getSettings')  return _json(_getAppSettings());
    if (action==='getUsers')     return _json(_getUsers());
    if (action==='ping')         return _json({status:'ok',app:'KNNDCmdb',version:'1.9'});
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
    if (action==='saveSettings') return _json(_saveAppSettings(data));
    if (action==='saveUsers')    return _json(_saveUsers(data));   // full users array replace
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
  const all=sheet.getDataRange().getValues();
  if(all.length<5) return{success:false,error:'No data'};

  // Locate columns by header (row 4 = index 3)
  const hdrs=all[3].map(h=>String(h).trim().toLowerCase());
  const c=(names)=>{ for(const n of names){const i=hdrs.indexOf(n.toLowerCase());if(i>=0)return i;} return -1; };
  const iId    =c(['record id','id']);
  const iFirst =c(['first name','firstname']);
  const iLast  =c(['surname','last name','lastname']);
  const iParty =c(['party id number','party id']);
  const iVoter =c(['voter id number','voter id']);
  const iPhone =c(['telephone number','phone']);
  const iGender=c(['gender']);

  if(iId<0) return{success:false,error:'Record ID column not found'};

  for(let i=4;i<all.length;i++){
    if(String(all[i][iId]||'').trim()===String(data.id).trim()){
      const r=i+1;
      if(data.firstName  && iFirst>=0)  sheet.getRange(r,iFirst+1).setValue(data.firstName);
      if(data.lastName   && iLast>=0)   sheet.getRange(r,iLast+1).setValue(data.lastName);
      if(data.partyId    && iParty>=0)  sheet.getRange(r,iParty+1).setValue(data.partyId);
      if(data.voterId    && iVoter>=0)  sheet.getRange(r,iVoter+1).setValue(data.voterId);
      if(data.phone      && iPhone>=0)  sheet.getRange(r,iPhone+1).setValue(data.phone);
      if(data.gender     && iGender>=0) sheet.getRange(r,iGender+1).setValue(data.gender);
      return{success:true};
    }
  }
  return{success:false,error:'Record not found'};
}

function _deleteMember(data) {
  const ss=SpreadsheetApp.getActiveSpreadsheet();
  const sheet=ss.getSheetByName(SHEETS.MEMBERS); if(!sheet) return{success:false,error:'Sheet not found'};
  const all=sheet.getDataRange().getValues();
  if(all.length<5) return{success:false,error:'No data'};
  // Find Record ID column by header name
  const hdrs=all[3].map(h=>String(h).trim().toLowerCase());
  const iId=hdrs.findIndex(h=>h==='record id'||h==='id');
  if(iId<0) return{success:false,error:'Record ID column not found'};
  for(let i=4;i<all.length;i++){
    if(String(all[i][iId]||'').trim()===String(data.id).trim()){
      sheet.deleteRow(i+1);_refreshSummary(ss);return{success:true};
    }
  }
  return{success:false,error:'Record not found'};
}

function _logAudit(data) {
  const ss=SpreadsheetApp.getActiveSpreadsheet();
  let sheet=ss.getSheetByName(SHEETS.AUDIT);
  if(!sheet){_setupAuditSheet(ss);sheet=ss.getSheetByName(SHEETS.AUDIT);}
  // auditAction = the event name; action = the HTTP routing key (always 'logAudit')
  const auditAction = data.auditAction || data.action || '';
  sheet.appendRow([
    data.timestamp || new Date().toLocaleString(),
    auditAction,
    data.user    || '',
    data.details || '',
    data.extra   || '',
  ]);
  return{success:true};
}

function _getMembers() {
  const ss=SpreadsheetApp.getActiveSpreadsheet();
  const sheet=ss.getSheetByName(SHEETS.MEMBERS); if(!sheet) return{members:[]};
  const all=sheet.getDataRange().getValues();
  if(all.length<5) return{members:[]};

  // Headers are on row 4 (index 3)
  const rawHdrs = all[3];
  const hdrs = rawHdrs.map(h => String(h).trim());
  const col = (names) => {
    for(const n of names){
      const i = hdrs.findIndex(h => h.toLowerCase()===n.toLowerCase());
      if(i>=0) return i;
    }
    return -1;
  };

  // Map every column by name — handles any column order
  const iFirst   = col(['first name','firstname']);
  const iLast    = col(['surname','last name','lastname']);
  const iParty   = col(['party id number','party id','partyid']);
  const iVoter   = col(['voter id number','voter id','voterid']);
  const iPhone   = col(['telephone number','phone','telephone']);
  const iGender  = col(['gender']);
  const iZone    = col(['zone']);
  const iWard    = col(['ward name','ward']);
  const iStation = col(['polling station','station']);
  const iStCode  = col(['station code','stationcode']);
  const iBranch  = col(['branch name','branch']);
  const iBrCode  = col(['branch code','branchcode']);
  const iOther   = col(['other names','othernames']);
  const iOfficer = col(['officer id','officer']);
  const iOfficerN= col(['officer name','officername']);
  const iTime    = col(['date/time added','timestamp','date added']);
  const iId      = col(['record id','id']);

  const g = (row, i) => i>=0 ? String(row[i]||'').trim() : '';

  const members=[];
  for(let i=4;i<all.length;i++){
    const row=all[i];
    const firstName = g(row,iFirst);
    const lastName  = g(row,iLast);
    const id        = g(row,iId);
    if(!firstName && !lastName && !id) continue; // skip empty rows
    members.push({
      id:          id,
      firstName:   firstName,
      lastName:    lastName,
      otherNames:  g(row,iOther),
      gender:      g(row,iGender),
      zone:        g(row,iZone),
      partyId:     g(row,iParty),
      voterId:     g(row,iVoter),
      phone:       g(row,iPhone),
      ward:        g(row,iWard),
      station:     g(row,iStation),
      stationCode: g(row,iStCode),
      branch:      g(row,iBranch),
      branchCode:  g(row,iBrCode),
      officer:     g(row,iOfficer),
      officerName: g(row,iOfficerN),
      timestamp:   g(row,iTime),
    });
  }
  return{members,total:members.length};
}

function _getPollingStations() {
  const ss=SpreadsheetApp.getActiveSpreadsheet();
  const sheet=ss.getSheetByName(SHEETS.POLLING_STATIONS); if(!sheet) return{stations:[]};
  const rows=sheet.getDataRange().getValues();
  if(rows.length<2) return{stations:[]};

  // Headers are on row 2 (index 1) — find each column by name, case-insensitive
  const hdrs = rows[1].map(h => String(h).trim().toLowerCase());
  const col = (names) => {
    for(const n of names){
      const i = hdrs.indexOf(n.toLowerCase());
      if(i>=0) return i;
    }
    return -1;
  };

  // Map all known header name variants for each field
  const iZone   = col(['zone']);
  const iWard   = col(['ward name','ward']);
  const iName   = col(['polling station name','station name','name']);
  const iBranch = col(['branch name','branch']);
  const iCode   = col(['station code','code','polling station code']);
  const iBCode  = col(['branch code','branchcode']);

  const stations=[];
  for(let i=2;i<rows.length;i++){
    const r=rows[i];
    const nameVal  = iName>=0  ? String(r[iName]||'').trim()  : '';
    const codeVal  = iCode>=0  ? String(r[iCode]||'').trim()  : '';
    if(!nameVal && !codeVal) continue; // skip truly empty rows
    stations.push({
      zone:       iZone>=0   ? String(r[iZone]  ||'').trim() : '',
      ward:       iWard>=0   ? String(r[iWard]  ||'').trim() : '',
      name:       nameVal,
      branch:     iBranch>=0 ? String(r[iBranch]||'').trim() : '',
      code:       codeVal,
      branchCode: iBCode>=0  ? String(r[iBCode] ||'').trim() : '',
    });
  }
  return{stations};
}

// ─── Read App Settings from Sheet ────────────────────────────
function _getAppSettings() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  let sheet   = ss.getSheetByName(SHEETS.SETTINGS);
  if (!sheet) {
    // Settings sheet doesn't exist yet — return empty object so app falls back to defaults
    return { settings: {}, exists: false };
  }
  const rows  = sheet.getDataRange().getValues();
  const cfg   = {};
  // Data rows start at row 3 (index 2): key in col A, value in col B
  for (let i = 2; i < rows.length; i++) {
    const key = String(rows[i][0]||'').trim();
    const val = String(rows[i][1]||'').trim();
    if (key) cfg[key] = val;
  }
  return { settings: cfg, exists: true };
}

// ─── Write App Settings to Sheet ─────────────────────────────
function _saveAppSettings(data) {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEETS.SETTINGS);
  if (!sheet) { _setupAppSettingsSheet(ss); sheet = ss.getSheetByName(SHEETS.SETTINGS); }

  // Update "Last Updated By" row
  sheet.getRange('B2').setValue((data.updatedBy || 'admin') + '  |  ' + new Date().toLocaleString());

  // Keys we persist to the sheet (never store the apiKey server-side for security)
  const keysToSave = ['scriptUrl','appName','constituency','sheetId','demoCleared'];

  const rows = sheet.getDataRange().getValues();
  keysToSave.forEach(key => {
    if (data[key] === undefined) return;
    // Find existing row with this key and update value, or append
    let found = false;
    for (let i = 2; i < rows.length; i++) {
      if (String(rows[i][0]).trim() === key) {
        sheet.getRange(i + 1, 2).setValue(String(data[key]||''));
        found = true; break;
      }
    }
    if (!found) {
      sheet.appendRow([key, String(data[key]||'')]);
    }
  });

  return { success: true };
}


// ─── Get All Users ────────────────────────────────────────────
function _getUsers() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.USERS);
  if (!sheet) return { users: [] };
  const all = sheet.getDataRange().getValues();
  if (all.length < 3) return { users: [] };

  // Headers on row 2 (index 1), data from row 3 (index 2)
  const hdrs = all[1].map(h => String(h).trim().toLowerCase());
  const col = (names) => {
    for (const n of names) {
      const i = hdrs.indexOf(n.toLowerCase());
      if (i >= 0) return i;
    }
    return -1;
  };

  const iId       = col(['user id','id']);
  const iName     = col(['full name','name']);
  const iUsername = col(['username']);
  const iPassword = col(['password']);
  const iRole     = col(['role']);
  const iWard     = col(['ward']);
  const iStation  = col(['station code','station']);
  const iBranch   = col(['branch']);
  const iAssigned = col(['assigned stations','assignedstations']);
  const iStatus   = col(['status']);
  const iMustChg  = col(['must change password','mustchangepassword']);

  const g = (row, i) => i >= 0 ? String(row[i] || '').trim() : '';

  const users = [];
  for (let i = 2; i < all.length; i++) {
    const row = all[i];
    const id  = g(row, iId);
    const usr = g(row, iUsername);
    if (!id && !usr) continue; // skip empty rows
    const assignedRaw = g(row, iAssigned);
    const assignedStations = assignedRaw
      ? assignedRaw.split(',').map(s => s.trim()).filter(Boolean)
      : [];
    users.push({
      id:                 id,
      name:               g(row, iName),
      username:           usr,
      password:           g(row, iPassword),
      role:               g(row, iRole),
      ward:               g(row, iWard),
      station:            g(row, iStation),
      branch:             g(row, iBranch),
      assignedStations:   assignedStations,
      active:             g(row, iStatus).toLowerCase() !== 'inactive',
      mustChangePassword: g(row, iMustChg).toLowerCase() === 'true' || g(row, iMustChg) === '1',
    });
  }
  return { users, total: users.length };
}


// ─── Save Full Users Array (replaces all rows from row 3 down) ──
function _saveUsers(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEETS.USERS);
  if (!sheet) { _setupUsersSheet(ss); sheet = ss.getSheetByName(SHEETS.USERS); }

  const users = data.users;
  if (!Array.isArray(users)) return { success: false, error: 'users must be an array' };

  // Clear all existing data rows (keep title row 1 and header row 2)
  const lastRow = sheet.getLastRow();
  if (lastRow >= 3) {
    sheet.getRange(3, 1, lastRow - 2, sheet.getLastColumn()).clearContent();
  }

  if (users.length === 0) return { success: true };

  const now = new Date().toLocaleString();
  const rows = users.map(u => [
    u.id            || '',
    u.name          || '',
    u.username      || '',
    u.password      || '',
    u.role          || '',
    u.ward          || '',
    u.station       || '',
    u.branch        || '',
    Array.isArray(u.assignedStations) ? u.assignedStations.join(', ') : (u.assignedStations || ''),
    u.active === false ? 'Inactive' : 'Active',
    u.mustChangePassword ? 'true' : 'false',
    now,
  ]);

  sheet.getRange(3, 1, rows.length, 12).setValues(rows);
  return { success: true, saved: users.length };
}


function _refreshSummary(ss) {
  let sheet=ss.getSheetByName(SHEETS.SUMMARY);
  if(!sheet){_setupSummarySheet(ss);sheet=ss.getSheetByName(SHEETS.SUMMARY);}
  const lastRow=sheet.getLastRow();
  if(lastRow>=4) sheet.getRange(4,1,lastRow-3,4).clearContent();

  const mSheet=ss.getSheetByName(SHEETS.MEMBERS); if(!mSheet) return{success:true};
  const all=mSheet.getDataRange().getValues();
  if(all.length<5) return{success:true};

  // Locate columns by header name (row 4 = index 3)
  const hdrs=all[3].map(h=>String(h).trim().toLowerCase());
  const c=(names)=>{ for(const n of names){const i=hdrs.indexOf(n.toLowerCase());if(i>=0)return i;} return -1; };
  const iWard=c(['ward name','ward']);
  const iStn =c(['polling station','station']);
  const iBr  =c(['branch name','branch']);

  const byStation={};
  for(let i=4;i<all.length;i++){
    const r=all[i];
    const stn=iStn>=0?String(r[iStn]||'').trim():'';
    if(!stn) continue;
    if(!byStation[stn]) byStation[stn]={ward:iWard>=0?String(r[iWard]||'').trim():'',station:stn,branch:iBr>=0?String(r[iBr]||'').trim():'',count:0};
    byStation[stn].count++;
  }
  const rows=Object.values(byStation).sort((a,b)=>b.count-a.count);
  const total=rows.reduce((s,r)=>s+r.count,0);
  const toWrite=rows.map(r=>[r.ward,r.station,r.branch,r.count]);
  toWrite.push(['','','GRAND TOTAL',total]);
  if(toWrite.length){
    sheet.getRange(4,1,toWrite.length,4).setValues(toWrite);
    const gtRow=4+toWrite.length-1;
    sheet.getRange(gtRow,1,1,4).setFontWeight('bold').setBackground(LIGHT_GRN);
  }
  sheet.getRange('D2').setValue('Updated: '+new Date().toLocaleString()).setFontColor('#6b7280');
  return{success:true};
}
