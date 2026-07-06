/**
 * ระบบรับคำสั่งซื้อเสื้อกีฬา SMTE
 * - บันทึกข้อมูลลง Google Sheets (ชีทที่ผูกกับสคริปต์นี้)
 * - อัปโหลดรูปสลิปเข้าโฟลเดอร์ Google Drive แล้วแปะลิงก์ไว้ในชีท
 * - ป้องกันการส่งซ้ำ: ล็อกด้วย LockService + เช็ก submissionId / ชั้น+เลขที่ / เบอร์โทร
 *
 * วิธีติดตั้งอยู่ใน README.md ของ repo
 */

var CONFIG = {
  SHEET_NAME: 'คำสั่งซื้อเสื้อ',
  FOLDER_NAME: 'SMTE Shirt - สลิปโอนเงิน',
  // กันซ้ำจากข้อมูลจริง: true = ห้ามชั้น+เลขที่เดิม หรือเบอร์โทรเดิม ส่งซ้ำ
  BLOCK_DUPLICATE_PERSON: true
};

var HEADERS = [
  'เวลาที่ส่ง', 'คำนำหน้า', 'ชื่อ', 'นามสกุล', 'ชั้น', 'เลขที่',
  'เบอร์โทรศัพท์', 'ชื่อบนเสื้อ', 'เลขบนเสื้อ', 'ไซส์', 'จำนวน (ตัว)',
  'ยอดโอน (บาท)', 'ลิงก์สลิป', 'Submission ID'
];

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000); // กันเขียนชนกันเมื่อมีคนกดพร้อมกัน
  } catch (err) {
    return jsonOut({ status: 'error', message: 'ระบบมีผู้ใช้งานหนาแน่น กรุณาลองใหม่อีกครั้ง' });
  }

  try {
    var data = JSON.parse(e.postData.contents);

    // --- ตรวจข้อมูลจำเป็น ---
    var required = ['prefix', 'firstName', 'lastName', 'classroom', 'classNo', 'phone', 'size', 'quantity', 'slipBase64'];
    for (var i = 0; i < required.length; i++) {
      if (!data[required[i]]) {
        return jsonOut({ status: 'error', message: 'ข้อมูลไม่ครบถ้วน (' + required[i] + ')' });
      }
    }
    if (!/^0\d{8,9}$/.test(String(data.phone))) {
      return jsonOut({ status: 'error', message: 'รูปแบบเบอร์โทรศัพท์ไม่ถูกต้อง' });
    }

    var sheet = getOrCreateSheet_();
    var rows = sheet.getDataRange().getValues();

    // --- กันส่งซ้ำ ---
    // 1) submissionId เดิม (เช่น กดซ้ำ/เน็ตหลุดแล้ว retry) -> ตอบสำเร็จโดยไม่บันทึกเพิ่ม
    for (var r = 1; r < rows.length; r++) {
      if (String(rows[r][13]) === String(data.submissionId)) {
        return jsonOut({ status: 'duplicate-id', message: 'บันทึกคำสั่งซื้อนี้ไว้แล้ว' });
      }
    }
    // 2) คนเดิมส่งซ้ำ (ชั้น+เลขที่ หรือเบอร์โทรเดิม)
    if (CONFIG.BLOCK_DUPLICATE_PERSON) {
      var classroom = normalize_(data.classroom);
      var classNo = String(data.classNo).trim();
      var phone = String(data.phone).trim();
      for (var r2 = 1; r2 < rows.length; r2++) {
        var sameSeat = normalize_(rows[r2][4]) === classroom && String(rows[r2][5]).trim() === classNo;
        var samePhone = String(rows[r2][6]).trim() === phone;
        if (sameSeat || samePhone) {
          return jsonOut({
            status: 'duplicate',
            message: sameSeat
              ? 'ชั้น ' + data.classroom + ' เลขที่ ' + classNo + ' มีคำสั่งซื้อแล้ว หากต้องการแก้ไข/สั่งเพิ่ม กรุณาติดต่อผู้ประสานงาน'
              : 'เบอร์โทรศัพท์นี้มีคำสั่งซื้อแล้ว หากต้องการแก้ไข/สั่งเพิ่ม กรุณาติดต่อผู้ประสานงาน'
          });
        }
      }
    }

    // --- เก็บสลิปเข้า Google Drive ---
    var folder = getOrCreateFolder_();
    var ext = (data.slipMimeType === 'image/png') ? '.png' : '.jpg';
    var fileName = 'สลิป_' + data.classroom + '_เลขที่' + data.classNo + '_' +
      data.firstName + '_' + Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyyMMdd-HHmmss') + ext;
    var blob = Utilities.newBlob(
      Utilities.base64Decode(data.slipBase64),
      data.slipMimeType || 'image/jpeg',
      fileName
    );
    var file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    // --- บันทึกลงชีท ---
    sheet.appendRow([
      Utilities.formatDate(new Date(), 'Asia/Bangkok', 'dd/MM/yyyy HH:mm:ss'),
      data.prefix,
      data.firstName,
      data.lastName,
      data.classroom,
      "'" + data.classNo,
      "'" + data.phone,   // กันชีทตัดเลข 0 นำหน้า
      data.desiredName || 'SMTE',
      "'" + (data.desiredNumber || '10'),  // เก็บเป็นข้อความ กันเลข 0 นำหน้าหาย
      data.size,
      Number(data.quantity),
      data.total === '' ? '' : Number(data.total),
      file.getUrl(),
      data.submissionId || ''
    ]);

    return jsonOut({ status: 'success' });
  } catch (err) {
    return jsonOut({ status: 'error', message: 'เกิดข้อผิดพลาด: ' + err.message });
  } finally {
    lock.releaseLock();
  }
}

/** เปิดลิงก์ /exec ตรง ๆ เพื่อเช็กว่า deploy สำเร็จ */
function doGet() {
  return jsonOut({ status: 'ok', message: 'SMTE shirt order endpoint is running' });
}

/** รันฟังก์ชันนี้ 1 ครั้งหลังวางโค้ด เพื่อสร้างชีท+โฟลเดอร์ และอนุมัติสิทธิ์ */
function setup() {
  getOrCreateSheet_();
  getOrCreateFolder_();
}

function getOrCreateSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEET_NAME);
  }
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
    sheet.getRange(1, 1, 1, HEADERS.length)
      .setFontWeight('bold')
      .setBackground('#E8501A')
      .setFontColor('#FFFFFF');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function getOrCreateFolder_() {
  var it = DriveApp.getFoldersByName(CONFIG.FOLDER_NAME);
  return it.hasNext() ? it.next() : DriveApp.createFolder(CONFIG.FOLDER_NAME);
}

function normalize_(s) {
  return String(s).replace(/\s+/g, '').replace(/ม\./g, 'ม.').toLowerCase();
}

function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
