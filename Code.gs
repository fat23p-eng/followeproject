// ══════════════════════════════════════════════════
//  ระบบติดตามผลการปฏิบัติงาน 2569 — Code.gs
// ══════════════════════════════════════════════════

var SHEET_DATA   = 'Data';
var SHEET_USERS  = 'Users';
var SHEET_MASTER = 'MasterData';

var DATA_HEADERS = [
  'Timestamp', 'แผนงาน', 'หน่วย', 'เป้า', 'ผล',
  'ร้อยละ', 'ผู้รับผิดชอบ', 'สถานะ'
];

// ── Entry Point ──
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('ระบบติดตามผลการปฏิบัติงาน 2569')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ── Login ──
function login(email, password) {
  try {
    var sheet = getOrCreateSheet(SHEET_USERS);
    var inputEmail = String(email || '').trim().toLowerCase();
    var inputPass  = String(password || '').trim();
    var rows = sheet.getDataRange().getValues();
    for (var i = 1; i < rows.length; i++) {
      var rowEmail  = String(rows[i][1]).trim().toLowerCase();
      var rowPass   = String(rows[i][2]).trim();
      var rowActive = String(rows[i][4]).toLowerCase();
      if (rowActive !== 'true') continue;
      if (rowEmail === inputEmail && rowPass === inputPass) {
        return {
          success: true,
          name: String(rows[i][5]).trim(),
          role: String(rows[i][3]).trim()
        };
      }
    }
    return { success: false, message: 'อีเมลหรือรหัสผ่านไม่ถูกต้อง' };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

// ── Get All Data ──
function getAllData() {
  try {
    var sheet = getOrCreateSheet(SHEET_DATA);
    if (sheet.getLastRow() <= 1) return [];
    var rows    = sheet.getDataRange().getValues();
    var headers = rows[0];
    return rows.slice(1).map(function(row, i) {
      var obj = { _rowIndex: i + 2 };
      headers.forEach(function(h, j) { obj[h] = row[j]; });

      // คำนวณร้อยละและสถานะใหม่ทุกครั้ง (ไม่ขึ้นกับที่บันทึกไว้)
      var target = Number(obj['เป้า'])  || 0;
      var actual = Number(obj['ผล'])   || 0;
      var pct    = target > 0 ? (actual / target) * 100 : 0;
      obj['ร้อยละ'] = pct.toFixed(2);
      obj['สถานะ']  = pct < 100 ? 'ต่ำกว่าแผน' : (pct === 100 ? 'ตามแผน' : 'สูงกว่าแผน');

      // แปลง Timestamp เป็น string กันปัญหา JSON
      if (obj['Timestamp'] instanceof Date) {
        obj['Timestamp'] = Utilities.formatDate(
          obj['Timestamp'],
          Session.getScriptTimeZone(),
          'dd/MM/yyyy HH:mm'
        );
      }
      return obj;
    });
  } catch (e) {
    Logger.log('getAllData error: ' + e.message);
    return [];
  }
}

// ── Save Data (Add / Edit) ──
function saveData(data, role) {
  if (role !== 'Admin') {
    return { success: false, message: 'ไม่มีสิทธิ์จัดการข้อมูล' };
  }
  try {
    var sheet  = getOrCreateSheet(SHEET_DATA);
    var target = Number(data['เป้า']) || 0;
    var actual = Number(data['ผล'])  || 0;
    var pct    = target > 0 ? (actual / target) * 100 : 0;
    var status = pct < 100 ? 'ต่ำกว่าแผน' : (pct === 100 ? 'ตามแผน' : 'สูงกว่าแผน');
    var plan   = String(data['แผนงาน']      || '').trim();
    var unit   = String(data['หน่วย']        || '').trim();
    var resp   = String(data['ผู้รับผิดชอบ']  || '').trim();

    if (!plan || !unit || !resp) {
      return { success: false, message: 'ข้อมูลไม่ครบถ้วน' };
    }

    var newRow = [
      new Date(), plan, unit, target, actual,
      pct.toFixed(2), resp, status
    ];

    if (data._rowIndex && Number(data._rowIndex) > 1) {
      sheet.getRange(Number(data._rowIndex), 1, 1, newRow.length).setValues([newRow]);
    } else {
      sheet.appendRow(newRow);
    }
    return { success: true };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

// ── Delete Data ──
function deleteData(rowIndex, role) {
  if (role !== 'Admin') {
    return { success: false, message: 'ไม่มีสิทธิ์ลบข้อมูล' };
  }
  try {
    var idx = Number(rowIndex);
    if (!idx || idx < 2) return { success: false, message: 'ระบุแถวไม่ถูกต้อง' };
    getOrCreateSheet(SHEET_DATA).deleteRow(idx);
    return { success: true };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

// ── Master Data (Dropdown lists) ──
function getMasterData() {
  try {
    var sheet  = getOrCreateSheet(SHEET_MASTER);
    var rows   = sheet.getDataRange().getValues();
    var result = {};
    for (var i = 1; i < rows.length; i++) {
      var type = String(rows[i][0]).trim();
      var name = String(rows[i][1]).trim();
      if (type && name) {
        if (!result[type]) result[type] = [];
        if (result[type].indexOf(name) === -1) result[type].push(name);
      }
    }
    return result;
  } catch (e) {
    Logger.log('getMasterData error: ' + e.message);
    return {};
  }
}

// ── Utility: Get or Create Sheet ──
function getOrCreateSheet(name) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    if (name === SHEET_DATA) {
      sheet.appendRow(DATA_HEADERS);
      sheet.setFrozenRows(1);
      sheet.getRange(1, 1, 1, DATA_HEADERS.length)
        .setBackground('#1d4ed8').setFontColor('white').setFontWeight('bold');
    }
    if (name === SHEET_MASTER) {
      sheet.appendRow(['ประเภท', 'ชื่อรายการ']);
      sheet.setFrozenRows(1);
    }
    if (name === SHEET_USERS) {
      sheet.appendRow(['Timestamp', 'Email', 'Password', 'Role', 'Active', 'Name']);
      sheet.setFrozenRows(1);
    }
  }
  return sheet;
}

// ── Setup (รันครั้งแรก) ──
function setup() {
  // สร้าง Data sheet
  var dataSheet = getOrCreateSheet(SHEET_DATA);

  // สร้าง MasterData พร้อมตัวอย่าง
  var masterSheet = getOrCreateSheet(SHEET_MASTER);
  if (masterSheet.getLastRow() <= 1) {
    var sampleMaster = [
      ['แผนงาน', 'งานส่งเสริมสุขภาพ'],
      ['แผนงาน', 'งานควบคุมโรค'],
      ['แผนงาน', 'งานอนามัยสิ่งแวดล้อม'],
      ['แผนงาน', 'งานบริหารทั่วไป'],
      ['แผนงาน', 'งานพัฒนาบุคลากร'],
      ['ผู้รับผิดชอบ', 'นายสมชาย ใจดี'],
      ['ผู้รับผิดชอบ', 'นางสาวสมหญิง รักงาน'],
      ['ผู้รับผิดชอบ', 'นายวิชัย มุ่งมั่น'],
      ['ผู้รับผิดชอบ', 'นางสาวนิดา ขยันดี'],
    ];
    masterSheet.getRange(2, 1, sampleMaster.length, 2).setValues(sampleMaster);
    Logger.log('✅ สร้าง MasterData ตัวอย่างแล้ว');
  }

  // สร้าง Users พร้อม Admin
  var userSheet = getOrCreateSheet(SHEET_USERS);
  if (userSheet.getLastRow() <= 1) {
    userSheet.appendRow([new Date(), 'admin', 'admin1234', 'Admin', 'true', 'ผู้ดูแลระบบ']);
    Logger.log('✅ สร้าง Admin user แล้ว (email: admin / password: admin1234)');
  }

  // สร้างข้อมูลตัวอย่าง
  if (dataSheet.getLastRow() <= 1) {
    var now = new Date();
    var sampleData = [
      [now, 'งานส่งเสริมสุขภาพ', 'ราย', 500, 480, '96.00', 'นายสมชาย ใจดี', 'ต่ำกว่าแผน'],
      [now, 'งานควบคุมโรค',     'ครั้ง', 120, 130, '108.33', 'นางสาวสมหญิง รักงาน', 'สูงกว่าแผน'],
      [now, 'งานอนามัยสิ่งแวดล้อม', 'โครงการ', 20, 20, '100.00', 'นายวิชัย มุ่งมั่น', 'ตามแผน'],
      [now, 'งานบริหารทั่วไป', 'ครั้ง', 50, 42, '84.00', 'นางสาวนิดา ขยันดี', 'ต่ำกว่าแผน'],
      [now, 'งานพัฒนาบุคลากร', 'คน', 80, 90, '112.50', 'นายสมชาย ใจดี', 'สูงกว่าแผน'],
    ];
    dataSheet.getRange(2, 1, sampleData.length, sampleData[0].length).setValues(sampleData);
    Logger.log('✅ สร้างข้อมูลตัวอย่างแล้ว');
  }

  return "✅ Setup สำเร็จ — เปิดใช้งานได้เลย!";
}
