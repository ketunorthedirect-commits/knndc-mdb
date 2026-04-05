/**
 * KNNDCmdb – Google Apps Script Backend
 * Deploy as a Web App (Execute as: Me, Access: Anyone)
 * 
 * Paste this ENTIRE file into your Google Apps Script editor.
 */

const SHEET_NAME = 'Members Database';
const AUDIT_SHEET = 'Audit Log';

// ─── GET: Read all members ───────────────────────────────────
function doGet(e) {
  const action = e.parameter.action || 'getMembers';
  let result;
  
  try {
    if (action === 'getMembers') {
      result = getMembers();
    } else if (action === 'getStats') {
      result = getStats();
    } else {
      result = { error: 'Unknown action' };
    }
  } catch(err) {
    result = { error: err.message };
  }
  
  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// ─── POST: Add/Update/Delete members ────────────────────────
function doPost(e) {
  let result;
  
  try {
    const data = JSON.parse(e.postData.contents);
    const action = data.action || 'addMember';
    
    if (action === 'addMember') {
      result = addMember(data);
    } else if (action === 'updateMember') {
      result = updateMember(data);
    } else if (action === 'deleteMember') {
      result = deleteMember(data);
    } else if (action === 'logAudit') {
      result = logAuditEntry(data);
    } else {
      // Default: treat POST as addMember for backward compat
      result = addMember(data);
    }
  } catch(err) {
    result = { success: false, error: err.message };
  }
  
  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON)
    .setHeader ? ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON) : ContentService.createTextOutput(JSON.stringify(result));
}

// ─── Get Members ─────────────────────────────────────────────
function getMembers() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) return { members: [] };
  
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return { members: [] };
  
  // Row 1 = title, Row 2 = headers, Row 3+ = data (matching your Excel format)
  const headers = data[1]; // index 1 = row 2
  const members = [];
  
  for (let i = 2; i < data.length; i++) {
    const row = data[i];
    if (!row[0] && !row[1]) continue; // skip empty rows
    const member = {};
    headers.forEach((h, j) => { member[h] = row[j]; });
    members.push(member);
  }
  
  return { members, total: members.length };
}

// ─── Add Member ──────────────────────────────────────────────
function addMember(data) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  let sheet   = ss.getSheetByName(SHEET_NAME);
  
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    setupSheet(sheet);
  }
  
  const lastRow = sheet.getLastRow();
  
  // Append the member row (matching Excel columns: First Name, Surname, Party ID, Voter ID, Phone)
  sheet.appendRow([
    data.firstName   || '',
    data.lastName    || '',
    data.partyId     || '',
    data.voterId     || '',
    data.phone       || '',
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
  
  return { success: true, row: lastRow + 1 };
}

// ─── Update Member ───────────────────────────────────────────
function updateMember(data) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) return { success: false, error: 'Sheet not found' };
  
  const rows = sheet.getDataRange().getValues();
  // Find row by ID (column 14, index 13)
  for (let i = 2; i < rows.length; i++) {
    if (rows[i][13] === data.id) {
      const row = i + 1; // 1-indexed
      sheet.getRange(row, 1).setValue(data.firstName   || rows[i][0]);
      sheet.getRange(row, 2).setValue(data.lastName    || rows[i][1]);
      sheet.getRange(row, 3).setValue(data.partyId     || rows[i][2]);
      sheet.getRange(row, 4).setValue(data.voterId     || rows[i][3]);
      sheet.getRange(row, 5).setValue(data.phone       || rows[i][4]);
      return { success: true };
    }
  }
  
  return { success: false, error: 'Member not found' };
}

// ─── Delete Member ───────────────────────────────────────────
function deleteMember(data) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) return { success: false, error: 'Sheet not found' };
  
  const rows = sheet.getDataRange().getValues();
  for (let i = 2; i < rows.length; i++) {
    if (rows[i][13] === data.id) {
      sheet.deleteRow(i + 1);
      return { success: true };
    }
  }
  
  return { success: false, error: 'Member not found' };
}

// ─── Log Audit ───────────────────────────────────────────────
function logAuditEntry(data) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  let sheet   = ss.getSheetByName(AUDIT_SHEET);
  
  if (!sheet) {
    sheet = ss.insertSheet(AUDIT_SHEET);
    sheet.appendRow(['Timestamp','Action','User','Details']);
    sheet.getRange(1,1,1,4).setFontWeight('bold');
  }
  
  sheet.appendRow([
    data.timestamp || new Date().toLocaleString(),
    data.action    || '',
    data.user      || '',
    data.details   || '',
  ]);
  
  return { success: true };
}

// ─── Setup Sheet Headers ─────────────────────────────────────
function setupSheet(sheet) {
  // Row 1: Title
  sheet.getRange('A1').setValue('MEMBERSHIP DATABASE');
  sheet.getRange('A1').setFontWeight('bold').setFontSize(14);
  sheet.getRange('E1').setValue('Constituency: Ketu North');
  
  // Row 2: Polling Station info
  sheet.getRange('A2').setValue('Polling Station / Branch Name:');
  
  // Row 3: Empty
  
  // Row 4: Column headers (matching Excel format)
  const headers = [
    'First Name','Surname','Party ID Number','Voter ID Number','Telephone Number',
    'Polling Station','Station Code','Branch Name','Branch Code','Other Names',
    'Officer ID','Officer Name','Date/Time Added','Record ID'
  ];
  sheet.getRange(4, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(4, 1, 1, headers.length).setFontWeight('bold').setBackground('#1a6b3a').setFontColor('white');
  
  // Freeze header rows
  sheet.setFrozenRows(4);
  
  // Column widths
  sheet.setColumnWidth(1, 130); // First Name
  sheet.setColumnWidth(2, 130); // Surname
  sheet.setColumnWidth(3, 150); // Party ID
  sheet.setColumnWidth(4, 150); // Voter ID
  sheet.setColumnWidth(5, 140); // Phone
}

// ─── Stats ───────────────────────────────────────────────────
function getStats() {
  const { members } = getMembers();
  const today = new Date().toLocaleDateString();
  const todayCount = members.filter(m => {
    const ts = m['Date/Time Added'] || '';
    return ts.toString().includes(today);
  }).length;
  
  return {
    total: members.length,
    today: todayCount,
  };
}
