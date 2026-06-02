// ==========================================
// CHECK-IN SYSTEM — Google Apps Script
// ==========================================

const MASTER_SHEET = "📋 Log ทั้งหมด";
const PLAN_SHEET   = "📅 แผนงาน";
const USER_SHEET   = "👥 ผู้ใช้งาน";
const PHOTO_FOLDER = "CheckIn Photos";
const HDR_BG       = "#0d1117";
const HDR_FG       = "#58a6ff";

const LOG_HEADERS = [
  'วันที่','เวลา Check-in','ชื่อ','สาขา',
  'รอบ','ช่วงเวลา','Latitude','Longitude',
  'ความแม่นยำ (m)','Google Maps','รูปภาพ URL','เวลาจบงาน','แจ้งเตือน GPS','รูปภาพจบงาน'
];
const PLAN_HEADERS = [
  'ID','ชื่อพนักงาน','วันที่','เวลาเริ่ม','เวลาสิ้นสุด',
  'สาขา','ประเภทงาน','หมายเหตุ','สร้างเมื่อ','Latitude','Longitude'
];
const USER_HEADERS = ['ID','Password','ชื่อ','Role'];

// ──────────────────────────────────────────────────
// POST
// ──────────────────────────────────────────────────
function doPost(e) {
  try {
    const data   = JSON.parse(e.postData.contents);
    const action = data.action || 'checkin';

    if (action === 'savePlan')   return savePlan(data);
    if (action === 'deletePlan') return deletePlan(data);
    if (action === 'checkout')   return saveCheckout(data);
    if (action === 'saveUser')   return saveUser(data);
    if (action === 'deleteUser') return deleteUser(data);
    return saveCheckin(data);   // default

  } catch(err) {
    return jsonOK({ success: false, error: err.message });
  }
}

// ── Check-in ───────────────────────────────────────
function saveCheckin(data) {
  let photoUrl = '';
  if (data.photo && data.photo.startsWith('data:image')) {
    try {
      const blob = Utilities.newBlob(
        Utilities.base64Decode(data.photo.split(',')[1]),
        'image/jpeg', `${data.name}_${data.round}_${Date.now()}.jpg`
      );
      const file = getFolder().createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      photoUrl = `https://drive.google.com/thumbnail?id=${file.getId()}&sz=w400`;
    } catch(err) { photoUrl = 'error: '+err.message; }
  }

  const now     = new Date();
  const dateStr = Utilities.formatDate(now, 'Asia/Bangkok', 'dd/MM/yyyy');
  const timeStr = data.timestamp || Utilities.formatDate(now, 'Asia/Bangkok', 'dd/MM/yyyy HH:mm:ss');
  const mapsUrl = (data.lat && data.lng) ? `https://maps.google.com/?q=${data.lat},${data.lng}` : (data.mapsLink||'');

  const row = [
    dateStr, timeStr, data.name||'', data.branch||'',
    data.round||'', data.roundTime||'', data.lat||'', data.lng||'',
    data.accuracy||'', mapsUrl, photoUrl, '', data.geofenceAlert||''
  ];

  appendRow(MASTER_SHEET, LOG_HEADERS, row);
  if (data.name) appendRow('👤 '+data.name.trim(), LOG_HEADERS, row);

  return jsonOK({ success: true });
}

// ── Checkout ───────────────────────────────────────
function saveCheckout(data) {
  const now      = new Date();
  const timeStr  = data.checkoutTime || Utilities.formatDate(now, 'Asia/Bangkok', 'dd/MM/yyyy HH:mm:ss');
  const dateStr  = data.date || Utilities.formatDate(now, 'Asia/Bangkok', 'dd/MM/yyyy');
  const colCheckout = 12; // เวลาจบงาน
  const colPhoto    = 14; // รูปภาพจบงาน (คอลัมน์ N)

  // อัปโหลดรูปจบงาน (ถ้ามี)
  let photoUrl = '';
  if (data.photo && data.photo.startsWith('data:image')) {
    try {
      const blob = Utilities.newBlob(
        Utilities.base64Decode(data.photo.split(',')[1]),
        'image/jpeg', `${data.name}_checkout_${Date.now()}.jpg`
      );
      const file = getFolder().createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      photoUrl = `https://drive.google.com/thumbnail?id=${file.getId()}&sz=w400`;
    } catch(err) { photoUrl = 'error: '+err.message; }
  }

  const sheets = [MASTER_SHEET];
  if (data.name) sheets.push('👤 ' + data.name.trim());

  sheets.forEach(sheetName => {
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return;
    const rows = sheet.getDataRange().getValues();
    for (let i = rows.length - 1; i >= 1; i--) {
      if (rows[i][2] === data.name &&
          rows[i][4] === data.round &&
          rows[i][0].toString().includes(dateStr.split('/')[0])) {
        sheet.getRange(i + 1, colCheckout).setValue(timeStr);
        if (photoUrl) sheet.getRange(i + 1, colPhoto).setValue(photoUrl);
        break;
      }
    }
  });

  return jsonOK({ success: true });
}

// ── Plan CRUD ──────────────────────────────────────
function savePlan(data) {
  const sheet  = getOrCreateSheet(PLAN_SHEET, PLAN_HEADERS);
  const rows   = sheet.getDataRange().getValues();
  const id     = data.id || ('p_'+Date.now());
  const now    = new Date().toLocaleString('th-TH');
  const newRow = [
    id, data.staffName||'', data.date||'', data.startTime||'',
    data.endTime||'', data.branch||'', data.taskType||'', data.note||'', now,
    data.lat||'', data.lng||''
  ];

  for (let i=1; i<rows.length; i++) {
    if (rows[i][0] === id) {
      sheet.getRange(i+1, 1, 1, newRow.length).setValues([newRow]);
      return jsonOK({ success: true, id });
    }
  }
  sheet.appendRow(newRow);
  return jsonOK({ success: true, id });
}

function deletePlan(data) {
  const sheet = getOrCreateSheet(PLAN_SHEET, PLAN_HEADERS);
  const rows  = sheet.getDataRange().getValues();
  for (let i=1; i<rows.length; i++) {
    if (rows[i][0] === data.id) {
      sheet.deleteRow(i+1);
      return jsonOK({ success: true });
    }
  }
  return jsonOK({ success: false, error: 'not found' });
}

// ── User CRUD ──────────────────────────────────────
function saveUser(data) {
  const sheet  = getOrCreateSheet(USER_SHEET, USER_HEADERS);
  const rows   = sheet.getDataRange().getValues();

  // อัปเดตถ้ามี id อยู่แล้ว
  for (let i=1; i<rows.length; i++) {
    if (rows[i][0].toString() === data.id) {
      const pass = data.pass || rows[i][1]; // ถ้าไม่ส่ง pass ให้ใช้เดิม
      sheet.getRange(i+1, 1, 1, 4).setValues([[data.id, pass, data.name||'', data.role||'staff']]);
      return jsonOK({ success: true });
    }
  }
  // เพิ่มใหม่
  if (!data.pass) return jsonOK({ success: false, error: 'password required' });
  sheet.appendRow([data.id, data.pass, data.name||'', data.role||'staff']);
  return jsonOK({ success: true });
}

function deleteUser(data) {
  const sheet = getOrCreateSheet(USER_SHEET, USER_HEADERS);
  const rows  = sheet.getDataRange().getValues();
  for (let i=1; i<rows.length; i++) {
    if (rows[i][0].toString() === data.id) {
      sheet.deleteRow(i+1);
      return jsonOK({ success: true });
    }
  }
  return jsonOK({ success: false, error: 'not found' });
}

// ──────────────────────────────────────────────────
// GET
// ──────────────────────────────────────────────────
function doGet(e) {
  const p      = e.parameter || {};
  const action = p.action || '';

  // Login
  if (action === 'login') {
    const id   = (p.id   || '').toString().trim();
    const pass = (p.pass || '').toString().trim();
    const ss   = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(USER_SHEET);
    if (!sheet) return jsonOK({ success: false, error: 'no user sheet' });
    const rows = sheet.getDataRange().getValues();
    for (let i=1; i<rows.length; i++) {
      if (rows[i][0].toString().trim() === id && rows[i][1].toString().trim() === pass) {
        return jsonOK({ success: true, name: rows[i][2]+'', role: rows[i][3]+'' });
      }
    }
    return jsonOK({ success: false, error: 'invalid' });
  }

  if (action === 'getAll')   return jsonOK({ data: readLog(MASTER_SHEET) });
  if (action === 'getEmp')   return jsonOK({ data: readLog('👤 '+(p.name||'').trim()) });
  if (action === 'getPlans') return jsonOK({ data: readPlans() });
  if (action === 'getUsers') return jsonOK({ data: readUsers() });

  if (action === 'getPlan') {
    const name  = p.name || '';
    const date  = p.date || '';
    const all   = readPlans();
    const parts = date.split('-');
    const dateTh = parts.length===3 ? `${parts[2]}/${parts[1]}/${parts[0]}` : date;
    const filtered = all.filter(t =>
      (!name || t.staffName===name) &&
      (!date || t.date===date || t.date===dateTh)
    );
    return jsonOK({ data: filtered });
  }

  if (action === 'listEmp') {
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const names = ss.getSheets()
      .map(s => s.getName()).filter(n => n.startsWith('👤 '))
      .map(n => n.replace('👤 ',''));
    return jsonOK({ employees: names });
  }

  return jsonOK({ status: 'ok', time: new Date().toString() });
}

// ──────────────────────────────────────────────────
// Helpers
// ──────────────────────────────────────────────────
function appendRow(sheetName, headers, row) {
  const sheet = getOrCreateSheet(sheetName, headers);
  sheet.appendRow(row);
  const last = sheet.getLastRow();
  if (last > 1) {
    sheet.getRange(last,1,1,headers.length)
         .setBackground(last%2===0 ? '#f6f8fa' : '#ffffff');
  }
}

function readLog(sheetName) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return [];
  const rows = sheet.getDataRange().getValues();
  if (rows.length<=1) return [];
  return rows.slice(1).map(r => ({
    date:r[0]+'', timestamp:r[1]+'', name:r[2]+'', branch:r[3]+'',
    round:r[4]+'', roundTime:r[5]+'', lat:r[6]+'', lng:r[7]+'',
    accuracy:r[8]+'', mapsLink:r[9]+'', photo:r[10]+'',
    checkoutTime:r[11]+'', geofenceAlert:r[12]+'', checkoutPhoto:r[13]+''
  })).reverse();
}

function readPlans() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(PLAN_SHEET);
  if (!sheet) return [];
  const rows = sheet.getDataRange().getValues();
  if (rows.length<=1) return [];
  return rows.slice(1).map(r => ({
    id:r[0]+'',
    staffName:r[1]+'',
    date: r[2] instanceof Date
      ? Utilities.formatDate(r[2], 'Asia/Bangkok', 'yyyy-MM-dd')
      : r[2]+'',
    startTime: r[3] instanceof Date ? Utilities.formatDate(r[3],'Asia/Bangkok','HH:mm') : r[3]+'',
    endTime:   r[4] instanceof Date ? Utilities.formatDate(r[4],'Asia/Bangkok','HH:mm') : r[4]+'',
    branch:r[5]+'', taskType:r[6]+'', note:r[7]+'',
    lat:r[9]+'', lng:r[10]+''
  }));
}

function readUsers() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(USER_SHEET);
  if (!sheet) return [];
  const rows = sheet.getDataRange().getValues();
  if (rows.length<=1) return [];
  // ไม่ส่ง password กลับมา
  return rows.slice(1).map(r => ({
    id:r[0]+'', name:r[2]+'', role:r[3]+''
  }));
}

function getOrCreateSheet(name, headers) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  let   sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    const hr = sheet.getRange(1,1,1,headers.length);
    hr.setValues([headers]).setBackground(HDR_BG).setFontColor(HDR_FG).setFontWeight('bold');
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(2, 160);
  }
  return sheet;
}

function getFolder() {
  const f = DriveApp.getFoldersByName(PHOTO_FOLDER);
  return f.hasNext() ? f.next() : DriveApp.createFolder(PHOTO_FOLDER);
}

function jsonOK(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
