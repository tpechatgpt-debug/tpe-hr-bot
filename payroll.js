// payroll.js — อ่านไฟล์ Excel และบันทึก/ดึงข้อมูลเงินเดือน
const { google } = require('googleapis');
const { execSync } = require('child_process');
const fs      = require('fs');

// ── โครงสร้าง column ของ Sheet "ค่าแรง" ──────────────────
const COL = {
  name:       1,   // ชื่อ-นามสกุล
  position:   2,   // ตำแหน่ง
  workDays:   3,   // จน.วัน
  holidayD:   4,   // วันหยุด
  otH:        5,   // O.T (ชม.)
  leaveP:     6,   // ลากิจ
  leaveSick:  7,   // ลาป่วย
  absent:     8,   // ขาดงาน
  leaveVac:   9,   // พักร้อน
  leaveNoPay: 10,  // ไม่รับเงิน
  leaveMat:   11,  // ลาคลอด
  leaveBday:  12,  // วันเกิด
  late:       13,  // สาย
  baseWage:   14,  // ค่าแรง (ฐาน)
  basePay:    15,  // ค่าปกติ
  holidayPay: 16,  // ค่าวันหยุด
  otPay:      17,  // ค่า OT
  allowance:  18,  // เบี้ยเลี้ยง
  bonus1:     19,  // เบี้ยขยัน 1
  bonus2:     20,  // เบี้ยขยัน 2
  bonus3:     21,  // เบี้ยขยัน 3
  bonus4:     22,  // เบี้ยขยัน 4
  bonus5:     23,  // เบี้ยขยัน 5
  otherInc:   24,  // รายได้อื่นๆ
  totalInc:   25,  // รวมรายได้
  advance:    26,  // หักล่วงหน้า
  kot:        27,  // กยศ.
  soc:        28,  // ประกันสังคม
  tax:        29,  // ภาษี ณ ที่จ่าย
  absentDed:  30,  // วันขาด
  noPayDed:   31,  // ไม่รับเงิน
  otherDed:   32,  // หักอื่นๆ
  totalDed:   33,  // รวมเงินหัก
  netPay:     34,  // เงินได้สุทธิ
};

// ── แปลงชื่อให้ normalize (ตัดช่องว่างเกิน) ──────────────
function normName(s) {
  return (s || '').toString().replace(/\s+/g, ' ').trim();
}

// ── อ่านไฟล์ Excel แล้วคืน array ของพนักงาน ─────────────
async function parseExcel(filePath) {
  const wb = new ExcelJS.Workbook();
  // รองรับทั้ง .xls และ .xlsx
  if (filePath.endsWith('.xls')) {
    await wb.xlsx.readFile(filePath).catch(() => {});
    // fallback ใช้ xlrd-style reader
    return parseXls(filePath);
  }
  await wb.xlsx.readFile(filePath);
  return extractFromWorkbook(wb);
}

function parseXls(filePath) {
  return new Promise((resolve, reject) => {
    const lines = [
      "import xlrd, json",
      "wb = xlrd.open_workbook('" + filePath + "')",
      "sh = wb.sheet_by_name('\u0e04\u0e48\u0e32\u0e41\u0e23\u0e07')",
      "month = str(sh.cell_value(1, 1))",
      "rows = []",
      "for r in range(3, sh.nrows):",
      "    name = str(sh.cell_value(r, 1)).strip()",
      "    if not name: continue",
      "    rows.append({'name':name,'position':str(sh.cell_value(r,2)).strip(),",
      "        'workDays':sh.cell_value(r,3),'holidayD':sh.cell_value(r,4),",
      "        'otH':sh.cell_value(r,5),'leaveP':sh.cell_value(r,6),",
      "        'leaveSick':sh.cell_value(r,7),'absent':sh.cell_value(r,8),",
      "        'leaveVac':sh.cell_value(r,9),'leaveNoPay':sh.cell_value(r,10),",
      "        'leaveMat':sh.cell_value(r,11),'leaveBday':sh.cell_value(r,12),",
      "        'late':sh.cell_value(r,13),'baseWage':sh.cell_value(r,14),",
      "        'basePay':sh.cell_value(r,15),'holidayPay':sh.cell_value(r,16),",
      "        'otPay':sh.cell_value(r,17),'allowance':sh.cell_value(r,18),",
      "        'bonus':sh.cell_value(r,19)+sh.cell_value(r,20)+sh.cell_value(r,21)+sh.cell_value(r,22)+sh.cell_value(r,23),",
      "        'otherInc':sh.cell_value(r,24),'totalInc':sh.cell_value(r,25),",
      "        'advance':sh.cell_value(r,26),'kot':sh.cell_value(r,27),",
      "        'soc':sh.cell_value(r,28),'tax':sh.cell_value(r,29),",
      "        'absentDed':sh.cell_value(r,30),'noPayDed':sh.cell_value(r,31),",
      "        'otherDed':sh.cell_value(r,32),'totalDed':sh.cell_value(r,33),",
      "        'netPay':sh.cell_value(r,34)})",
      "print(json.dumps({'month':month,'rows':rows},ensure_ascii=False))",
    ];
    const pyCode = lines.join("\n");
    const tmpFile = '/tmp/pxls_' + Date.now() + '.py';
    fs.writeFileSync(tmpFile, pyCode, 'utf8');
    try {
      const out = execSync('python3 ' + tmpFile, { maxBuffer: 10*1024*1024 });
      try { fs.unlinkSync(tmpFile); } catch(_) {}
      resolve(JSON.parse(out.toString('utf8')));
    } catch(e) {
      try { fs.unlinkSync(tmpFile); } catch(_) {}
      reject(new Error('parseXls error: ' + e.stderr?.toString() || e.message));
    }
  });
}


// ── บันทึกไฟล์ Excel ลง Google Drive ─────────────────────
async function saveToGoogleDrive(fileBuffer, filename, mimeType) {
  const auth = new google.auth.GoogleAuth({
    credentials: JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT_JSON),
    scopes: ['https://www.googleapis.com/auth/drive'],
  });
  const drive = google.drive({ version: 'v3', auth });
  const folderId = process.env.PAYROLL_FOLDER_ID || process.env.DRIVE_FOLDER_ID;

  const { Readable } = require('stream');
  const res = await drive.files.create({
    supportsAllDrives: true,
    requestBody: {
      name: filename,
      parents: [folderId],
    },
    media: { mimeType, body: Readable.from([fileBuffer]) },
  });

  // ให้สิทธิ์อ่านแก่ทุกคน
  try {
    await drive.permissions.create({
      fileId: res.data.id,
      supportsAllDrives: true,
      requestBody: { role: 'reader', type: 'anyone' },
    });
  } catch(e) {}

  return res.data.id;
}

// ── บันทึกลง Google Sheets ────────────────────────────────
async function savePayrollToSheet(month, rows) {
  const auth = new google.auth.GoogleAuth({
    credentials: JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT_JSON),
    scopes: ['https://www.googleapis.com/auth/spreadsheets'],
  });
  const sheets = google.sheets({ version: 'v4', auth });
  const sheetId = process.env.LOG_SHEET_ID;
  const sheetName = 'เงินเดือน_' + month;

  // ตรวจว่า sheet มีอยู่แล้วไหม
  let sheetExists = false;
  try {
    const meta = await sheets.spreadsheets.get({ spreadsheetId: sheetId });
    sheetExists = (meta.data.sheets || []).some(s => s.properties.title === sheetName);
  } catch(e) {
    console.log('get sheets error:', e.message);
  }

  if (!sheetExists) {
    // สร้าง sheet ใหม่
    await sheets.spreadsheets.batchUpdate({
      spreadsheetId: sheetId,
      resource: { requests: [{ addSheet: { properties: { title: sheetName } } }] },
    });
    // ใส่ header
    await sheets.spreadsheets.values.update({
      spreadsheetId: sheetId,
      range: "'" + sheetName + "'!A1",
      valueInputOption: 'USER_ENTERED',
      resource: { values: [['ชื่อ','ตำแหน่ง','วันทำงาน','วันหยุด','OT','ลากิจ','ลาป่วย','ขาดงาน','พักร้อน','ไม่รับเงิน','ลาคลอด','วันเกิด','ค่าแรง','ค่าปกติ','วันหยุดPay','OT Pay','เบี้ยเลี้ยง','เบี้ยขยัน','รวมรายได้','หักล่วงหน้า','กยศ.','ปกส.','ภาษี','วันขาด','ไม่รับเงิน','อื่นๆ','รวมหัก','เงินสุทธิ']] },
    });
  } else {
    // ลบข้อมูลเดิมออกก่อน (เว้น header) แล้วเขียนใหม่
    await sheets.spreadsheets.values.clear({
      spreadsheetId: sheetId,
      range: "'" + sheetName + "'!A2:Z",
    });
  }

  // เพิ่มข้อมูล
  const values = rows.map(r => [
    r.name, r.position,
    r.workDays, r.holidayD, r.otH,
    r.leaveP, r.leaveSick, r.absent, r.leaveVac, r.leaveNoPay, r.leaveMat, r.leaveBday,
    r.baseWage, r.basePay, r.holidayPay, r.otPay, r.allowance, r.bonus,
    r.totalInc, r.advance, r.kot, r.soc, r.tax, r.absentDed, r.noPayDed, r.otherDed, r.totalDed, r.netPay,
  ]);

  await sheets.spreadsheets.values.append({
    spreadsheetId: sheetId,
    range: "'" + sheetName + "'!A2",
    valueInputOption: 'USER_ENTERED',
    resource: { values },
  });

  console.log('savePayrollToSheet OK:', sheetName, rows.length, 'rows');
  return sheetName;
}

// ── ดึงข้อมูลเงินเดือนของพนักงาน 1 คน จาก Sheet ──────────
async function getEmployeePayroll(empName, month) {
  const auth = new google.auth.GoogleAuth({
    credentials: JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT_JSON),
    scopes: ['https://www.googleapis.com/auth/spreadsheets'],
  });
  const sheets = google.sheets({ version: 'v4', auth });
  const sheetId = process.env.LOG_SHEET_ID;
  const sheetName = 'เงินเดือน_' + month;

  try {
    const res = await sheets.spreadsheets.values.get({
      spreadsheetId: sheetId,
      range: "'" + sheetName + "'!A:R",
    });
    const rows = res.data.values || [];

    // ตัดตำแหน่งออกจากชื่อ (กรณีมี " (" ต่อท้าย)
    const cleanName = n => normName((n || '').split('(')[0].split('（')[0]);
    const target = cleanName(empName);

    console.log('searching for:', target, 'in', sheetName);

    for (let i = 1; i < rows.length; i++) {
      if (cleanName(rows[i][0]) === target) {
        const r = rows[i];
        const toN = v => parseFloat(v) || 0;
        return {
          name: r[0], position: r[1],
          workDays:   toN(r[2]),  holidayD:   toN(r[3]),  otH:       toN(r[4]),
          leaveP:     toN(r[5]),  leaveSick:  toN(r[6]),  absent:    toN(r[7]),
          leaveVac:   toN(r[8]),  leaveNoPay: toN(r[9]),  leaveMat:  toN(r[10]),
          leaveBday:  toN(r[11]),
          baseWage:   toN(r[12]), basePay:    toN(r[13]), holidayPay:toN(r[14]),
          otPay:      toN(r[15]), allowance:  toN(r[16]), bonus:     toN(r[17]),
          totalInc:   toN(r[18]), advance:    toN(r[19]), kot:       toN(r[20]),
          soc:        toN(r[21]), tax:        toN(r[22]), absentDed: toN(r[23]),
          noPayDed:   toN(r[24]), otherDed:   toN(r[25]), totalDed:  toN(r[26]),
          netPay:     toN(r[27]), month,
        };
      }
    }
    return null; // ไม่พบ
  } catch(e) {
    console.error('getEmployeePayroll error:', e.message);
    return null;
  }
}

// ── ดึงรายการเดือนที่มีข้อมูลอยู่ ────────────────────────
async function getAvailableMonths() {
  const auth = new google.auth.GoogleAuth({
    credentials: JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT_JSON),
    scopes: ['https://www.googleapis.com/auth/spreadsheets'],
  });
  const sheets = google.sheets({ version: 'v4', auth });
  const meta = await sheets.spreadsheets.get({
    spreadsheetId: process.env.LOG_SHEET_ID,
  });
  return (meta.data.sheets || [])
    .map(s => s.properties.title)
    .filter(t => t.startsWith('เงินเดือน_'))
    .map(t => t.replace('เงินเดือน_', ''));
}

module.exports = { parseXls, saveToGoogleDrive, savePayrollToSheet, getEmployeePayroll, getAvailableMonths, normName };
