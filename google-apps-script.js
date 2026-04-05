/**
 * ============================================================
 * KNNDCmdb – Google Apps Script Backend  v1.2
 * ============================================================
 * HOW TO DEPLOY (STEP BY STEP):
 *
 * 1. Open your Google Sheet
 * 2. Click  Extensions → Apps Script
 * 3. Delete all existing code
 * 4. Paste THIS ENTIRE FILE
 * 5. Click 💾 Save (name the project "KNNDCmdb Backend")
 * 6. Click Run → setupSpreadsheet  (ONE-TIME SETUP — creates all tabs)
 *    → Authorise when prompted (click "Allow")
 *    → You will see "Setup complete!" in the log
 * 7. Click Deploy → New Deployment
 *    → Type: Web App
 *    → Execute as: Me
 *    → Who has access: Anyone
 *    → Click Deploy → Authorise → Allow
 * 8. Copy the Web App URL (looks like: https://script.google.com/macros/s/.../exec)
 * 9. Paste that URL into the KNNDCmdb app → Settings → Google Sheets → Script URL
 * ============================================================
 */

// ─── Sheet Names ─────────────────────────────────────────────
const SHEETS = {
  MEMBERS:          'Members Database',
  POLLING_STATIONS: 'Polling Stations',
  USERS:            'Users',
  AUDIT:            'Audit Log',
  SUMMARY:          'Summary',
};

// ─── NDC Brand Colours ───────────────────────────────────────
const NDC_GREEN  = '#1a6b3a';
const NDC_GREEN2 = '#134d2a';
const NDC_RED    = '#c8102e';
const WHITE      = '#ffffff';
const LIGHT_GRN  = '#e8f5ec';


// ════════════════════════════════════════════════════════════
//  ONE-TIME SETUP — Run this manually ONCE after pasting code
// ════════════════════════════════════════════════════════════
function setupSpreadsheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.setSpreadsheetName('KNNDCmdb – Ketu North NDC Members Database');

  _setupMembersSheet(ss);
  _setupPollingStationsSheet(ss);
  _setupUsersSheet(ss);
  _setupAuditSheet(ss);
  _setupSummarySheet(ss);

  // Reorder tabs
  const order = [SHEETS.MEMBERS, SHEETS.POLLING_STATIONS, SHEETS.USERS, SHEETS.AUDIT, SHEETS.SUMMARY];
  order.forEach((name, i) => {
    const s = ss.getSheetByName(name);
    if (s) ss.setActiveSheet(s).moveActiveSheet(i + 1);
  });

  Logger.log('✅ Setup complete! All tabs created successfully.');
  SpreadsheetApp.getUi().alert('✅ KNNDCmdb Setup Complete!\n\nAll tabs have been created:\n• Members Database\n• Polling Stations\n• Users\n• Audit Log\n• Summary\n\nNow deploy this script as a Web App (Deploy → New Deployment).');
}


// ─── Members Database Sheet ──────────────────────────────────
function _setupMembersSheet(ss) {
  let sheet = ss.getSheetByName(SHEETS.MEMBERS);
  if (!sheet) sheet = ss.insertSheet(SHEETS.MEMBERS);
  sheet.clear();

  // Row 1: Title (matching Excel template)
  sheet.getRange('A1:E1').merge().setValue('MEMBERSHIP DATABASE')
    .setFontFamily('Outfit').setFontSize(16).setFontWeight('bold')
    .setHorizontalAlignment('center').setBackground(NDC_GREEN2).setFontColor(WHITE);

  // Row 2: Info line
  sheet.getRange('A2').setValue('Polling Station / Branch Name:').setFontWeight('bold');
  sheet.getRange('E2').setValue('Constituency: Ketu North').setHorizontalAlignment('right');

  // Row 3: Blank
  sheet.getRange('A3').setValue('');

  // Row 4: Column headers (core 5 match the Excel template exactly)
  const coreHdrs  = ['First Name','Surname','Party ID Number','Voter ID Number','Telephone Number'];
  const extraHdrs = ['Ward Name','Polling Station','Station Code','Branch Name','Branch Code','Other Names','Officer ID','Officer Name','Date/Time Added','Record ID'];
  const allHdrs   = [...coreHdrs, ...extraHdrs];

  const hdrRange = sheet.getRange(4, 1, 1, allHdrs.length);
  hdrRange.setValues([allHdrs]).setFontWeight('bold').setFontColor(WHITE).setBackground(NDC_GREEN);
  sheet.getRange(4, 1, 1, 5).setBackground(NDC_GREEN2); // darker for core cols

  sheet.setFrozenRows(4);
  sheet.setFrozenColumns(2);

  // Alternating row colours (starting row 5 — formula driven via conditional format)
  // Column widths
  const widths = [130,130,160,160,150,120,150,100,140,110,130,100,160,180,160];
  widths.forEach((w,i) => sheet.setColumnWidth(i+1, w));

  sheet.setTabColor(NDC_GREEN);
  Logger.log('✅ Members Database sheet ready.');
}


// ─── Polling Stations Sheet ──────────────────────────────────
function _setupPollingStationsSheet(ss) {
  let sheet = ss.getSheetByName(SHEETS.POLLING_STATIONS);
  if (!sheet) sheet = ss.insertSheet(SHEETS.POLLING_STATIONS);
  sheet.clear();

  sheet.getRange('A1:E1').merge().setValue('POLLING STATIONS REGISTER')
    .setFontFamily('Outfit').setFontSize(14).setFontWeight('bold')
    .setHorizontalAlignment('center').setBackground(NDC_GREEN).setFontColor(WHITE);

  const hdrs = ['Ward Name','Polling Station Name','Branch Name','Station Code','Branch Code'];
  sheet.getRange(2, 1, 1, hdrs.length).setValues([hdrs])
    .setFontWeight('bold').setBackground(NDC_GREEN2).setFontColor(WHITE);

  // Seed demo data
  const demo = [
    ['Aflao Ward',    'Aflao A Polling Station',   'Aflao Branch',    'PS-001','BR-001'],
    ['Aflao Ward',    'Aflao B Polling Station',   'Aflao Branch',    'PS-002','BR-001'],
    ['Denu Ward',     'Denu Polling Station',      'Denu Branch',     'PS-003','BR-002'],
    ['Agbozume Ward', 'Agbozume Polling Station',  'Agbozume Branch', 'PS-004','BR-003'],
    ['Klikor Ward',   'Klikor Polling Station',    'Klikor Branch',   'PS-005','BR-004'],
    ['Adafienu Ward', 'Adafienu Polling Station',  'Adafienu Branch', 'PS-006','BR-005'],
  ];
  sheet.getRange(3, 1, demo.length, 5).setValues(demo);

  sheet.setFrozenRows(2);
  [150,200,160,120,120].forEach((w,i) => sheet.setColumnWidth(i+1, w));
  sheet.setTabColor('#2563eb');
  Logger.log('✅ Polling Stations sheet ready with demo data.');
}


// ─── Users Sheet ─────────────────────────────────────────────
function _setupUsersSheet(ss) {
  let sheet = ss.getSheetByName(SHEETS.USERS);
  if (!sheet) sheet = ss.insertSheet(SHEETS.USERS);
  sheet.clear();

  sheet.getRange('A1:H1').merge().setValue('SYSTEM USERS')
    .setFontFamily('Outfit').setFontSize(14).setFontWeight('bold')
    .setHorizontalAlignment('center').setBackground(NDC_RED).setFontColor(WHITE);

  const hdrs = ['User ID','Full Name','Username','Password','Role','Ward','Station Code','Branch','Assigned Stations','Status'];
  sheet.getRange(2, 1, 1, hdrs.length).setValues([hdrs])
    .setFontWeight('bold').setBackground('#7f1d1d').setFontColor(WHITE);

  const seed = [
    ['u001','System Administrator','admin','Admin@2026','admin','','','','','Active'],
    ['u002','Constituency Executive','exec','Exec@2026','exec','','','','','Active'],
    ['u003','Ward Coordinator (Aflao)','ward1','Ward@2026','ward','Aflao Ward','PS-001','Aflao Branch','PS-001,PS-002','Active'],
    ['u004','Data Entry Officer 1','officer1','Off1@2026','officer','Aflao Ward','PS-001','Aflao Branch','PS-001','Active'],
    ['u005','Data Entry Officer 2','officer2','Off2@2026','officer','Denu Ward','PS-003','Denu Branch','PS-003','Active'],
  ];
  sheet.getRange(3, 1, seed.length, 10).setValues(seed);

  sheet.setFrozenRows(2);
  [80,180,120,120,80,120,100,130,150,80].forEach((w,i) => sheet.setColumnWidth(i+1, w));
  sheet.setTabColor(NDC_RED);
  Logger.log('✅ Users sheet ready.');
}


// ─── Audit Log Sheet ─────────────────────────────────────────
function _setupAuditSheet(ss) {
  let sheet = ss.getSheetByName(SHEETS.AUDIT);
  if (!sheet) sheet = ss.insertSheet(SHEETS.AUDIT);
  sheet.clear();

  sheet.getRange('A1:E1').merge().setValue('SYSTEM AUDIT LOG')
    .setFontFamily('Outfit').setFontSize(14).setFontWeight('bold')
    .setHorizontalAlignment('center').setBackground('#1e3a5f').setFontColor(WHITE);

  const hdrs = ['Timestamp','Action','User','Details','Extra'];
  sheet.getRange(2, 1, 1, hdrs.length).setValues([hdrs])
    .setFontWeight('bold').setBackground('#1e3a5f').setFontColor(WHITE);

  sheet.setFrozenRows(2);
  [160,140,120,350,200].forEach((w,i) => sheet.setColumnWidth(i+1, w));
  sheet.setTabColor('#1e3a5f');
  Logger.log('✅ Audit Log sheet ready.');
}


// ─── Summary Sheet ───────────────────────────────────────────
function _setupSummarySheet(ss) {
  let sheet = ss.getSheetByName(SHEETS.SUMMARY);
  if (!sheet) sheet = ss.insertSheet(SHEETS.SUMMARY);
  sheet.clear();

  sheet.getRange('A1:D1').merge().setValue('MEMBERSHIP SUMMARY BY STATION')
    .setFontFamily('Outfit').setFontSize(14).setFontWeight('bold')
    .setHorizontalAlignment('center').setBackground(NDC_GREEN).setFontColor(WHITE);

  sheet.getRange('A2').setValue('This sheet is auto-updated when records are exported from KNNDCmdb.')
    .setFontColor('#6b7280').setFontStyle('italic');

  const hdrs = ['Ward','Polling Station','Branch','Total Members'];
  sheet.getRange(3, 1, 1, hdrs.length).setValues([hdrs])
    .setFontWeight('bold').setBackground(NDC_GREEN2).setFontColor(WHITE);

  sheet.setFrozenRows(3);
  [160,220,160,120].forEach((w,i) => sheet.setColumnWidth(i+1, w));
  sheet.setTabColor('#d97706');
  Logger.log('✅ Summary sheet ready.');
}


// ════════════════════════════════════════════════════════════
//  HTTP HANDLERS
// ════════════════════════════════════════════════════════════

function doGet(e) {
  const action = e?.parameter?.action || 'ping';
  let result;
  try {
    if (action === 'getMembers')         result = _getMembers();
    else if (action === 'getStations')   result = _getPollingStations();
    else if (action === 'ping')          result = { status:'ok', app:'KNNDCmdb', version:'1.2' };
    else                                  result = { error: 'Unknown action: ' + action };
  } catch(err) {
    result = { error: err.message };
  }
  return _json(result);
}

function doPost(e) {
  let result;
  try {
    const data   = JSON.parse(e.postData.contents);
    const action = data.action || 'addMember';
    if      (action === 'addMember')    result = _addMember(data);
    else if (action === 'updateMember') result = _updateMember(data);
    else if (action === 'deleteMember') result = _deleteMember(data);
    else if (action === 'logAudit')     result = _logAudit(data);
    else if (action === 'syncSummary')  result = _syncSummary(data);
    else                                result = _addMember(data); // default
  } catch(err) {
    result = { success: false, error: err.message };
  }
  return _json(result);
}

function _json(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}


// ════════════════════════════════════════════════════════════
//  CRUD FUNCTIONS
// ════════════════════════════════════════════════════════════

function _addMember(data) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  let sheet   = ss.getSheetByName(SHEETS.MEMBERS);
  if (!sheet) { setupSpreadsheet(); sheet = ss.getSheetByName(SHEETS.MEMBERS); }

  sheet.appendRow([
    data.firstName    || '',
    data.lastName     || '',
    data.partyId      || '',
    data.voterId      || '',
    data.phone        || '',
    data.ward         || '',
    data.station      || '',
    data.stationCode  || '',
    data.branch       || '',
    data.branchCode   || '',
    data.otherNames   || '',
    data.officer      || '',
    data.officerName  || '',
    data.timestamp    || new Date().toLocaleString(),
    data.id           || ('m' + Date.now()),
  ]);

  // Update summary
  _refreshSummary(ss);
  return { success: true };
}

function _updateMember(data) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.MEMBERS);
  if (!sheet) return { success:false, error:'Members sheet not found' };
  const rows  = sheet.getDataRange().getValues();
  for (let i = 4; i < rows.length; i++) { // data starts row 5 (index 4)
    if (rows[i][14] === data.id) {
      const r = i + 1;
      sheet.getRange(r,1).setValue(data.firstName  || rows[i][0]);
      sheet.getRange(r,2).setValue(data.lastName   || rows[i][1]);
      sheet.getRange(r,3).setValue(data.partyId    || rows[i][2]);
      sheet.getRange(r,4).setValue(data.voterId    || rows[i][3]);
      sheet.getRange(r,5).setValue(data.phone      || rows[i][4]);
      return { success: true };
    }
  }
  return { success:false, error:'Record not found' };
}

function _deleteMember(data) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.MEMBERS);
  if (!sheet) return { success:false, error:'Members sheet not found' };
  const rows  = sheet.getDataRange().getValues();
  for (let i = 4; i < rows.length; i++) {
    if (rows[i][14] === data.id) { sheet.deleteRow(i + 1); _refreshSummary(ss); return { success: true }; }
  }
  return { success:false, error:'Record not found' };
}

function _logAudit(data) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  let sheet   = ss.getSheetByName(SHEETS.AUDIT);
  if (!sheet) { _setupAuditSheet(ss); sheet = ss.getSheetByName(SHEETS.AUDIT); }
  sheet.appendRow([data.timestamp||new Date().toLocaleString(), data.action||'', data.user||'', data.details||'', data.extra||'']);
  return { success: true };
}

function _getMembers() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.MEMBERS);
  if (!sheet) return { members:[] };
  const all = sheet.getDataRange().getValues();
  // Headers on row 4 (index 3); data from row 5 (index 4)
  const headers = all[3];
  const members = [];
  for (let i = 4; i < all.length; i++) {
    const row = all[i];
    if (!row[0] && !row[1] && !row[14]) continue;
    const m = {};
    headers.forEach((h, j) => { m[h] = row[j]; });
    members.push(m);
  }
  return { members, total: members.length };
}

function _getPollingStations() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.POLLING_STATIONS);
  if (!sheet) return { stations:[] };
  const rows = sheet.getDataRange().getValues();
  // headers row 2 (index 1), data from row 3 (index 2)
  const stations = [];
  for (let i = 2; i < rows.length; i++) {
    const r = rows[i];
    if (!r[0] && !r[1]) continue;
    stations.push({ ward:r[0], name:r[1], branch:r[2], code:r[3], branchCode:r[4] });
  }
  return { stations };
}

function _syncSummary(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  _refreshSummary(ss, data.summaryData);
  return { success: true };
}

function _refreshSummary(ss, summaryData) {
  let sheet = ss.getSheetByName(SHEETS.SUMMARY);
  if (!sheet) { _setupSummarySheet(ss); sheet = ss.getSheetByName(SHEETS.SUMMARY); }

  // Clear existing data rows
  const lastRow = sheet.getLastRow();
  if (lastRow >= 4) sheet.getRange(4, 1, lastRow - 3, 4).clearContent();

  // Compute from members sheet
  const membersSheet = ss.getSheetByName(SHEETS.MEMBERS);
  if (!membersSheet) return;
  const all = membersSheet.getDataRange().getValues();
  const byStation = {};
  for (let i = 4; i < all.length; i++) {
    const r = all[i];
    if (!r[6]) continue; // station col (index 6)
    const key = r[6];
    if (!byStation[key]) byStation[key] = { ward:r[5]||'', station:r[6]||'', branch:r[8]||'', count:0 };
    byStation[key].count++;
  }
  const rows   = Object.values(byStation).sort((a,b) => b.count - a.count);
  const total  = rows.reduce((s,r) => s+r.count, 0);
  const dataToWrite = rows.map(r => [r.ward, r.station, r.branch, r.count]);
  dataToWrite.push(['','','GRAND TOTAL', total]);

  if (dataToWrite.length) {
    sheet.getRange(4, 1, dataToWrite.length, 4).setValues(dataToWrite);
    // Bold grand total row
    const gtRow = 4 + dataToWrite.length - 1;
    sheet.getRange(gtRow, 1, 1, 4).setFontWeight('bold').setBackground(LIGHT_GRN);
  }
  sheet.getRange('D2').setValue('Last updated: ' + new Date().toLocaleString()).setFontColor('#6b7280');
}
