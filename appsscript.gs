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
  'ความแม่นยำ (m)','Google Maps','รูปภาพ URL','เวลาจบงาน','แจ้งเตือน GPS','รูปภาพจบงาน',
  'แจ้งเตือน GPS จบงาน','Lat จบงาน','Lng จบงาน','Google Maps จบงาน','ประเภทงาน'
];
const PLAN_HEADERS = [
  'ID','ชื่อพนักงาน','วันที่','เวลาเริ่ม','เวลาสิ้นสุด',
  'สาขา','ประเภทงาน','หมายเหตุ','สร้างเมื่อ','Latitude','Longitude'
];
const USER_HEADERS    = ['ID','Password','ชื่อ','Role'];
const DELETED_SHEET   = "🗑️ Deleted Log";
const DELETED_HEADERS = [
  'วันที่','เวลา Check-in','ชื่อ','สาขา',
  'รอบ','ช่วงเวลา','Latitude','Longitude',
  'ความแม่นยำ (m)','Google Maps','รูปภาพ URL','เวลาจบงาน','แจ้งเตือน GPS','รูปภาพจบงาน',
  'แจ้งเตือน GPS จบงาน','Lat จบงาน','Lng จบงาน','Google Maps จบงาน','ประเภทงาน',
  'ลบเมื่อ','ลบโดย'
];

// ──────────────────────────────────────────────────
// POST
// ──────────────────────────────────────────────────
const LINE_TOKEN    = 'kUEQye3hUx6Wtr4DIqz+mp8K9y9xQmPCEOhi7F4Zh8t6tslNG0aMalVKHPkeEbr7hI4TukGnA9mWnTjsw6p3a+JvwEmCEWsKCnGZOVyiYXPpL3T6KFT0ZVmrEue78uraP2RpfPuJjpdY1WFFUYByHAdB04t89/1O/w1cDnyilFU=';
const LINE_GROUP_ID_KEY = 'LINE_GROUP_ID'; // เก็บใน PropertiesService

function doPost(e) {
  try {
    const raw  = e.postData.contents;
    const data = JSON.parse(raw);

    // รับ LINE Webhook (มี events array)
    if (data.events) {
      data.events.forEach(ev => {
        // บันทึก Group ID อัตโนมัติ
        if (ev.source && ev.source.type === 'group' && ev.source.groupId) {
          PropertiesService.getScriptProperties().setProperty(LINE_GROUP_ID_KEY, ev.source.groupId);
        }
      });
      return jsonOK({ ok: true });
    }

    const action = data.action || 'checkin';
    if (action === 'savePlan')   return savePlan(data);
    if (action === 'deletePlan') return deletePlan(data);
    if (action === 'checkout')   return saveCheckout(data);
    if (action === 'saveUser')   return saveUser(data);
    if (action === 'deleteUser') return deleteUser(data);
    if (action === 'deleteRows') return deleteRows(data);
    return saveCheckin(data);   // default

  } catch(err) {
    return jsonOK({ success: false, error: err.message });
  }
}

// ── LINE Notify ────────────────────────────────────
function sendLineNotify(msg) {
  const groupId = PropertiesService.getScriptProperties().getProperty(LINE_GROUP_ID_KEY);
  if (!groupId) return;
  try {
    UrlFetchApp.fetch('https://api.line.me/v2/bot/message/push', {
      method: 'post',
      headers: {
        'Content-Type': 'application/json',
        'Authorization': 'Bearer ' + LINE_TOKEN
      },
      payload: JSON.stringify({
        to: groupId,
        messages: [{ type: 'text', text: msg }]
      }),
      muteHttpExceptions: true
    });
  } catch(e) {}
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
    data.accuracy||'', mapsUrl, photoUrl, '', data.geofenceAlert||'',
    '', '', '', '', '',   // placeholder cols 14-18 (checkout fields)
    data.taskType||''     // col 19 = ประเภทงาน
  ];

  appendRow(MASTER_SHEET, LOG_HEADERS, row);
  if (data.name) appendRow('👤 '+data.name.trim(), LOG_HEADERS, row);

  // แจ้งเตือน LINE
  const mapsLine   = mapsUrl ? `\n📍 GPS: ${mapsUrl}` : '';
  const alertLine  = data.geofenceAlert ? `\n🚨 ${data.geofenceAlert}` : '';
  const taskLine   = data.taskType ? `\n📋 ${data.taskType}` : '';
  sendLineNotify(`🏁 เช็กอิน\n👤 ${data.name||''}\n📍 ${data.branch||''} (${data.round||''})\n🕐 ${timeStr}${taskLine}${mapsLine}${alertLine}`);

  return jsonOK({ success: true });
}

// ── Checkout ───────────────────────────────────────
function saveCheckout(data) {
  const now      = new Date();
  const timeStr  = data.checkoutTime || Utilities.formatDate(now, 'Asia/Bangkok', 'dd/MM/yyyy HH:mm:ss');
  const dateStr  = data.date || Utilities.formatDate(now, 'Asia/Bangkok', 'dd/MM/yyyy');
  const colCheckout = 12; // เวลาจบงาน
  const colPhoto    = 14; // รูปภาพจบงาน (คอลัมน์ N)
  const colGeoOut   = 15; // แจ้งเตือน GPS จบงาน (คอลัมน์ O)
  const colLatOut   = 16; // Lat จบงาน (คอลัมน์ P)
  const colLngOut   = 17; // Lng จบงาน (คอลัมน์ Q)
  const colMapsOut  = 18; // Google Maps จบงาน (คอลัมน์ R)

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
        if (data.checkoutGeofenceAlert) sheet.getRange(i + 1, colGeoOut).setValue(data.checkoutGeofenceAlert);
        if (data.checkoutLat) sheet.getRange(i + 1, colLatOut).setValue(data.checkoutLat);
        if (data.checkoutLng) sheet.getRange(i + 1, colLngOut).setValue(data.checkoutLng);
        if (data.checkoutLat && data.checkoutLng) {
          const mapsOut = `https://maps.google.com/?q=${data.checkoutLat},${data.checkoutLng}`;
          sheet.getRange(i + 1, colMapsOut).setValue(mapsOut);
        }
        break;
      }
    }
  });

  // ดึง taskType จากชีท
  let savedTaskType = '';
  try {
    const ss2 = SpreadsheetApp.getActiveSpreadsheet();
    const m2  = ss2.getSheetByName(MASTER_SHEET);
    if (m2) {
      const rows2 = m2.getDataRange().getValues();
      for (let i = rows2.length - 1; i >= 1; i--) {
        if (rows2[i][2] === data.name && rows2[i][4] === data.round &&
            rows2[i][0].toString().includes(dateStr.split('/')[0])) {
          savedTaskType = rows2[i][18] + '';
          break;
        }
      }
    }
  } catch(e) {}

  // แจ้งเตือน LINE จบงาน
  const mapsOutLine  = (data.checkoutLat && data.checkoutLng)
    ? `\n📍 GPS จบงาน: https://maps.google.com/?q=${data.checkoutLat},${data.checkoutLng}` : '';
  const alertOutLine = data.checkoutGeofenceAlert ? `\n🚨 ${data.checkoutGeofenceAlert}` : '';
  const taskOutLine  = savedTaskType ? `\n📋 ${savedTaskType}` : '';
  const checkoutStatus = data.checkoutGeofenceAlert ? '⚠️ จบงานผิดปกติ' : '✅ จบงานแล้ว';
  sendLineNotify(`${checkoutStatus}\n👤 ${data.name||''}\n📍 ${data.branch||''} (${data.round||''})\n🕐 ${timeStr}${taskOutLine}${mapsOutLine}${alertOutLine}`);

  return jsonOK({ success: true });
}

// ── Delete Rows (ย้ายไป Deleted Log) ──────────────
function deleteRows(data) {
  // data.keys = array of "timestamp|name" เพื่อ identify แต่ละ row
  // data.deletedBy = ชื่อ admin ที่ลบ
  const keys      = data.keys || [];
  const deletedBy = data.deletedBy || '';
  const deletedAt = Utilities.formatDate(new Date(), 'Asia/Bangkok', 'dd/MM/yyyy HH:mm:ss');
  const ss        = SpreadsheetApp.getActiveSpreadsheet();
  const master    = ss.getSheetByName(MASTER_SHEET);
  if (!master) return jsonOK({ success: false, error: 'no master sheet' });

  const delSheet = getOrCreateSheet(DELETED_SHEET, DELETED_HEADERS);
  const rows     = master.getDataRange().getValues();
  const toDelete = []; // row indices (1-based) in master to delete

  // หา rows ที่ตรงกับ keys
  for (let i = rows.length - 1; i >= 1; i--) {
    const ts   = rows[i][1] + '';
    const name = rows[i][2] + '';
    const key  = ts + '|' + name;
    if (keys.includes(key)) {
      // copy ไป Deleted Log
      const newRow = [...rows[i], deletedAt, deletedBy];
      delSheet.appendRow(newRow);
      toDelete.push(i + 1); // 1-based
    }
  }

  // ลบจาก master (จากล่างขึ้นบน ป้องกัน index เลื่อน)
  toDelete.sort((a,b) => b - a).forEach(r => master.deleteRow(r));

  // ลบจากชีทพนักงานด้วย
  ss.getSheets().filter(s => s.getName().startsWith('👤 ')).forEach(empSheet => {
    const empRows = empSheet.getDataRange().getValues();
    const empDel  = [];
    for (let i = empRows.length - 1; i >= 1; i--) {
      const ts   = empRows[i][1] + '';
      const name = empRows[i][2] + '';
      if (keys.includes(ts + '|' + name)) empDel.push(i + 1);
    }
    empDel.sort((a,b) => b - a).forEach(r => empSheet.deleteRow(r));
  });

  return jsonOK({ success: true, deleted: toDelete.length });
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
    const id   = (p.id   || '').toString().trim().toLowerCase();
    const pass = (p.pass || '').toString().trim().toLowerCase();
    const ss   = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(USER_SHEET);
    if (!sheet) return jsonOK({ success: false, error: 'no user sheet' });
    const rows = sheet.getDataRange().getValues();
    for (let i=1; i<rows.length; i++) {
      const rowId   = rows[i][0].toString().trim().toLowerCase();
      const rowPass = rows[i][1].toString().trim().toLowerCase();
      if (rowId === id && rowPass === pass) {
        return jsonOK({ success: true, name: rows[i][2]+'', role: rows[i][3]+'' });
      }
    }
    return jsonOK({ success: false, error: 'invalid' });
  }

  if (action === 'getAll')     return jsonOK({ data: readLog(MASTER_SHEET) });
  if (action === 'getEmp')     return jsonOK({ data: readLog('👤 '+(p.name||'').trim()) });
  if (action === 'getDeleted') return jsonOK({ data: readDeleted() });
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
    date: r[0] instanceof Date ? Utilities.formatDate(r[0], 'Asia/Bangkok', 'dd/MM/yyyy') : r[0]+'',
    timestamp:r[1]+'', name:r[2]+'', branch:r[3]+'',
    round:r[4]+'', roundTime:r[5]+'', lat:r[6]+'', lng:r[7]+'',
    accuracy:r[8]+'', mapsLink:r[9]+'', photo:r[10]+'',
    checkoutTime:r[11]+'', geofenceAlert:r[12]+'', checkoutPhoto:r[13]+'',
    checkoutGeofenceAlert:r[14]+'', checkoutLat:r[15]+'', checkoutLng:r[16]+'',
    checkoutMapsLink:r[17]+'', taskType:r[18]+''
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

function readDeleted() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(DELETED_SHEET);
  if (!sheet) return [];
  const rows = sheet.getDataRange().getValues();
  if (rows.length <= 1) return [];
  return rows.slice(1).map(r => ({
    date: r[0] instanceof Date ? Utilities.formatDate(r[0],'Asia/Bangkok','dd/MM/yyyy') : r[0]+'',
    timestamp:r[1]+'', name:r[2]+'', branch:r[3]+'',
    round:r[4]+'', roundTime:r[5]+'', lat:r[6]+'', lng:r[7]+'',
    accuracy:r[8]+'', mapsLink:r[9]+'', photo:r[10]+'',
    checkoutTime:r[11]+'', geofenceAlert:r[12]+'', checkoutPhoto:r[13]+'',
    checkoutGeofenceAlert:r[14]+'', checkoutLat:r[15]+'', checkoutLng:r[16]+'',
    checkoutMapsLink:r[17]+'', deletedAt:r[18]+'', deletedBy:r[19]+''
  })).reverse();
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
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(2, 160);
  }
  // อัปเดต header เสมอ (รองรับคอลัมน์ใหม่หลัง deploy)
  const hr = sheet.getRange(1, 1, 1, headers.length);
  hr.setValues([headers]).setBackground(HDR_BG).setFontColor(HDR_FG).setFontWeight('bold');
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
