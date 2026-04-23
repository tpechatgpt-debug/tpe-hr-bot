// payroll.js — TPE HR Bot v2
// รองรับ Excel ทั้งรายเดือน และรายวัน (auto-detect จาก sheet name)

const { google } = require('googleapis');
const { execSync } = require('child_process');
const fs   = require('fs');
const os   = require('os');
const path = require('path');

// ─── normalize ชื่อ ─────────────────────────────────────────
function normName(s) {
  return (s || '').toString().replace(/\s+/g, ' ').trim();
}

// ─── Google Sheets auth ──────────────────────────────────────
function getSheetsClient() {
  const auth = new google.auth.GoogleAuth({
    credentials: JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT_JSON),
    scopes: ['https://www.googleapis.com/auth/spreadsheets'],
  });
  return google.sheets({ version: 'v4', auth });
}

// ─── parseXls ────────────────────────────────────────────────
function parseXls(filePath) {
  return new Promise((resolve, reject) => {
    const script = `
import xlrd, json, sys

wb = xlrd.open_workbook(sys.argv[1])
sheet_names = wb.sheet_names()

# ทั้งสองไฟล์มี sheet 'ค่าแรง' — ต้อง detect จากเนื้อหาใน row 0
pay_type = 'monthly'
if 'ค่าแรง' in sheet_names:
    sh_check = wb.sheet_by_name('ค่าแรง')
    if sh_check.nrows > 0:
        row0 = ' '.join(str(sh_check.cell_value(0, c)) for c in range(min(sh_check.ncols, 8)))
        pay_type = 'daily' if 'รายวัน' in row0 else 'monthly'

def n(sh, r, c):
    try:
        v = sh.cell_value(r, c)
        return float(v) if v != '' else 0.0
    except:
        return 0.0

if pay_type == 'daily':
    sh = wb.sheet_by_name('ค่าแรง')
    month = ''
    for ci in range(min(sh.ncols, 4)):
        val = str(sh.cell_value(1, ci)).strip()
        if val and val != '' and val != '0.0':
            month = val
            break
    rows = []
    for r in range(3, sh.nrows):
        name = str(sh.cell_value(r, 2)).strip()
        if not name or name in ['0.0', '']:
            continue
        prefix = str(sh.cell_value(r, 1)).strip()
        if prefix in ['0.0', '']: prefix = ''
        full_name = (prefix + ' ' + name).strip()
        rows.append({
            'name': full_name,
            'position': str(sh.cell_value(r, 3)).strip(),
            'baseWage':     n(sh,r,16),
            'workDays':     n(sh,r,4),
            'otH':          n(sh,r,5),
            'holidayD':     n(sh,r,6),
            'festivalD':    n(sh,r,7),
            'leaveP':       n(sh,r,8),
            'leaveSick':    n(sh,r,9),
            'absent':       n(sh,r,10),
            'leaveVac':     n(sh,r,11),
            'basePay':      n(sh,r,17),
            'otPay':        n(sh,r,18),
            'holidayPay':   n(sh,r,19),
            'festivalPay':  n(sh,r,20),
            'festivalExtra':n(sh,r,21),
            'allowance':    n(sh,r,22),
            'bonus':        n(sh,r,23)+n(sh,r,24)+n(sh,r,25)+n(sh,r,26)+n(sh,r,27),
            'otherInc':     n(sh,r,28),
            'totalInc':     n(sh,r,30),
            'soc':          n(sh,r,31),
            'socCompany':   n(sh,r,32),
            'advance':      n(sh,r,33),
            'kot':          n(sh,r,34),
            'late':         n(sh,r,35),
            'otherDed':     n(sh,r,36),
            'totalDed':     n(sh,r,37),
            'netPay':       n(sh,r,38),
        })
    print(json.dumps({'type':'daily','month':month,'rows':rows}, ensure_ascii=False))

else:
    sh = None
    for sname in sheet_names:
        s = wb.sheet_by_name(sname)
        if s.nrows > 3:
            sh = s
            break
    if not sh:
        print(json.dumps({'type':'monthly','month':'','rows':[]}))
        sys.exit(0)
    header_row = 0
    for r in range(min(sh.nrows, 10)):
        vals = [str(sh.cell_value(r, c)) for c in range(min(sh.ncols, 5))]
        if any('ชื่อ' in v for v in vals):
            header_row = r
            break
    headers = {}
    for c in range(sh.ncols):
        h = str(sh.cell_value(header_row, c)).strip()
        if h: headers[h] = c
    def col(nm, fb): return headers.get(nm, fb)
    month = ''
    months_th = ['มกราคม','กุมภาพันธ์','มีนาคม','เมษายน','พฤษภาคม','มิถุนายน',
                 'กรกฎาคม','สิงหาคม','กันยายน','ตุลาคม','พฤศจิกายน','ธันวาคม']
    for r2 in range(min(sh.nrows, 3)):
        for c2 in range(min(sh.ncols, 4)):
            v = str(sh.cell_value(r2, c2)).strip()
            if any(m in v for m in months_th):
                month = v
                break
    rows = []
    for r in range(header_row + 1, sh.nrows):
        # col 1 = ชื่อ, col 2 = ตำแหน่ง (รายเดือนไม่มีคำนำหน้าแยก)
        name_val = str(sh.cell_value(r, 1)).strip()
        if not name_val or name_val in ['0.0','ชื่อ','ชี่อ/นามสกุล','ชื่อ-นามสกุล','']:
            continue
        rows.append({
            'name':     name_val,
            'position': str(sh.cell_value(r, 2)).strip(),
            # สถิติการทำงาน
            # สถิติการทำงาน (ตรงกับ Sheet col 2-11)
            'workDays':   n(sh,r,3),   # จน.วัน
            'holidayD':   n(sh,r,4),   # วันหยุด
            'otH':        n(sh,r,5),   # O.T
            'leaveP':     n(sh,r,6),   # ลากิจ
            'leaveSick':  n(sh,r,7),   # ลาป่วย
            'absent':     n(sh,r,8),   # ขาดงาน
            'leaveVac':   n(sh,r,9),   # พักร้อน
            'leaveNoPay': n(sh,r,10),  # ไม่รับเงิน
            'leaveMat':   n(sh,r,11),  # ลาคลอด
            'leaveBday':  n(sh,r,12),  # วันเกิด
            'late':       n(sh,r,13),  # สาย
            # รายรับ (col 14-25)
            'baseWage':   n(sh,r,14),  # ค่าแรง (ฐาน)
            'basePay':    n(sh,r,15),  # ค่าปกติ
            'holidayPay': n(sh,r,16),  # วันหยุด Pay
            'otPay':      n(sh,r,17),  # ค่า OT
            'allowance':  n(sh,r,18),  # เบี้ยเลี้ยง
            'bonus':      n(sh,r,19)+n(sh,r,20)+n(sh,r,21)+n(sh,r,22)+n(sh,r,23),  # เบี้ยขยัน 1-5
            'otherInc':   n(sh,r,24),  # อื่นๆ
            'totalInc':   n(sh,r,25),  # รวม
            # รายจ่าย (col 26-34)
            'advance':    n(sh,r,26),  # ล่วงหน้า
            'kot':        n(sh,r,27),  # กยศ.
            'soc':        n(sh,r,28),  # ปกส.(ล/จ)
            'tax':        n(sh,r,29),  # ณที่จ่าย
            'absentDed':  n(sh,r,30),  # วันขาด
            'noPayDed':   n(sh,r,31),  # ไม่รับเงิน
            'otherDed':   n(sh,r,32),  # อื่น ๆ
            'totalDed':   n(sh,r,33),  # รวม
            'netPay':     n(sh,r,34),  # เงินได้สุทธิ
        })
    print(json.dumps({'type':'monthly','month':month,'rows':rows}, ensure_ascii=False))
`;
    const tmpFile = path.join(os.tmpdir(), `pxls_${Date.now()}.py`);
    fs.writeFileSync(tmpFile, script, 'utf8');
    try {
      const out = execSync(`python3 "${tmpFile}" "${filePath}"`, {
        maxBuffer: 10 * 1024 * 1024, encoding: 'utf8',
      });
      try { fs.unlinkSync(tmpFile); } catch (_) {}
      resolve(JSON.parse(out.trim()));
    } catch (e) {
      try { fs.unlinkSync(tmpFile); } catch (_) {}
      reject(new Error('parseXls error: ' + (e.stderr || e.message)));
    }
  });
}

// ─── savePayrollToSheet ──────────────────────────────────────
async function savePayrollToSheet(month, rows, payType) {
  const sheets    = getSheetsClient();
  const sheetId   = process.env.LOG_SHEET_ID;
  const prefix    = (payType === 'daily') ? 'เงินเดือนรายวัน_' : 'เงินเดือน_';
  const sheetName = prefix + month;

  const meta     = await sheets.spreadsheets.get({ spreadsheetId: sheetId });
  const existing = (meta.data.sheets || []).find(s => s.properties.title === sheetName);

  if (!existing) {
    await sheets.spreadsheets.batchUpdate({
      spreadsheetId: sheetId,
      requestBody: { requests: [{ addSheet: { properties: { title: sheetName } } }] },
    });
    const hdrs = payType === 'daily'
      ? ['ชื่อ-นามสกุล','ตำแหน่ง','ค่าแรง/วัน','วันปกติ','OT(ชม)','หยุดสัปดาห์','ประเพณี',
         'ลากิจ','ลาป่วย','ขาดงาน','พักร้อน','ค่าปกติ','ค่าOT','ค่าหยุดสัปดาห์','ค่าประเพณี',
         'เพิ่มประเพณี','เบี้ยเลี้ยง','เบี้ยขยัน','รายได้อื่นๆ','รายได้รวม',
         'ปกส.','ปกส(น/จ)','ล่วงหน้า','กยศ','สาย','รายจ่ายอื่นๆ','รายจ่ายรวม','เงินได้สุทธิ','type']
      : ['ชื่อ-นามสกุล','ตำแหน่ง','วันทำงาน','วันหยุด','OT',
         'ลากิจ','ลาป่วย','ขาดงาน','พักร้อน','ไม่รับเงิน','ลาคลอด','วันเกิด',
         'ค่าแรง','ค่าปกติ','ค่าหยุดPay','OTPay','เบี้ยเลี้ยง','เบี้ยขยัน','รวมรายได้',
         'หักล่วงหน้า','กยศ','ปกส.','ภาษี','วันขาด','ไม่รับเงิน(หัก)','รายจ่ายอื่นๆ','รวมหัก',
         'เงินได้สุทธิ','type'];
    await sheets.spreadsheets.values.update({
      spreadsheetId: sheetId, range: `'${sheetName}'!A1`,
      valueInputOption: 'RAW', requestBody: { values: [hdrs] },
    });
  } else {
    await sheets.spreadsheets.values.clear({
      spreadsheetId: sheetId, range: `'${sheetName}'!A2:AM`,
    });
  }

  const values = payType === 'daily'
    ? rows.map(r => [
        r.name, r.position, r.baseWage||0,
        r.workDays||0, r.otH||0, r.holidayD||0, r.festivalD||0,
        r.leaveP||0, r.leaveSick||0, r.absent||0, r.leaveVac||0,
        r.basePay||0, r.otPay||0, r.holidayPay||0, r.festivalPay||0, r.festivalExtra||0,
        r.allowance||0, r.bonus||0, r.otherInc||0, r.totalInc||0,
        r.soc||0, r.socCompany||0, r.advance||0, r.kot||0, r.late||0,
        r.otherDed||0, r.totalDed||0, r.netPay||0, 'daily',
      ])
    : rows.map(r => [
        r.name, r.position,
        r.workDays||0, r.holidayD||0, r.otH||0,
        r.leaveP||0, r.leaveSick||0, r.absent||0, r.leaveVac||0,
        r.leaveNoPay||0, r.leaveMat||0, r.leaveBday||0,
        r.baseWage||r.basePay||0, r.basePay||0, r.holidayPay||0, r.otPay||0,
        r.allowance||0, r.bonus||0, r.totalInc||0,
        r.advance||0, r.kot||0, r.soc||0, r.tax||0,
        r.absentDed||0, r.noPayDed||0, r.otherDed||0, r.totalDed||0,
        r.netPay||0, 'monthly',
      ]);

  await sheets.spreadsheets.values.append({
    spreadsheetId: sheetId, range: `'${sheetName}'!A2`,
    valueInputOption: 'RAW', requestBody: { values },
  });

  console.log(`savePayrollToSheet OK: ${sheetName} (${rows.length} rows)`);
  return sheetName;
}

// ─── getEmployeePayroll ──────────────────────────────────────
async function getEmployeePayroll(empName, month, payType) {
  const sheets    = getSheetsClient();
  const sheetId   = process.env.LOG_SHEET_ID;
  const prefix    = (payType === 'daily') ? 'เงินเดือนรายวัน_' : 'เงินเดือน_';
  const sheetName = prefix + month;

  try {
    const res  = await sheets.spreadsheets.values.get({
      spreadsheetId: sheetId, range: `'${sheetName}'!A:AM`,
    });
    const rows = res.data.values || [];
    if (rows.length < 2) return null;

    const cleanName = n => normName((n||'').split('(')[0].split('（')[0]);
    const target    = cleanName(empName);
    console.log(`getEmployeePayroll: "${target}" in ${sheetName}`);

    for (let i = 1; i < rows.length; i++) {
      if (cleanName(rows[i][0]) !== target) continue;
      const r   = rows[i];
      const toN = v => parseFloat(v) || 0;

      if (payType === 'daily') {
        return {
          name: r[0], position: r[1],
          baseWage: toN(r[2]), workDays: toN(r[3]), otH: toN(r[4]),
          holidayD: toN(r[5]), festivalD: toN(r[6]),
          leaveP: toN(r[7]), leaveSick: toN(r[8]), absent: toN(r[9]), leaveVac: toN(r[10]),
          basePay: toN(r[11]), otPay: toN(r[12]), holidayPay: toN(r[13]),
          festivalPay: toN(r[14]), festivalExtra: toN(r[15]),
          allowance: toN(r[16]), bonus: toN(r[17]),
          otherInc: toN(r[18]), totalInc: toN(r[19]),
          soc: toN(r[20]), socCompany: toN(r[21]),
          advance: toN(r[22]), kot: toN(r[23]), late: toN(r[24]),
          otherDed: toN(r[25]), totalDed: toN(r[26]), netPay: toN(r[27]),
          month, payType: 'daily',
          // alias สำหรับ payslip template เดิม
          leaveMat: 0, leaveBday: 0, leaveNoPay: 0,
          tax: 0, absentDed: 0, noPayDed: 0,
        };
      } else {
        // monthly col order (ตรงกับ Sheet จริง):
        // 0=ชื่อ, 1=ตำแหน่ง, 2=วันทำงาน, 3=วันหยุด, 4=OT
        // 5=ลากิจ, 6=ลาป่วย, 7=ขาดงาน, 8=พักร้อน
        // 9=ไม่รับเงิน, 10=ลาคลอด, 11=วันเกิด
        // 12=ค่าแรง, 13=ค่าปกติ, 14=ค่าหยุดPay, 15=OTPay
        // 16=เบี้ยเลี้ยง, 17=เบี้ยขยัน, 18=รวมรายได้
        // 19=หักล่วงหน้า, 20=กยศ, 21=ปกส., 22=ภาษี
        // 23=วันขาด, 24=ไม่รับเงิน(หัก), 25=รายจ่ายอื่นๆ, 26=รวมหัก
        // 27=เงินได้สุทธิ, 28=type
        return {
          name: r[0], position: r[1],
          workDays: toN(r[2]), holidayD: toN(r[3]), otH: toN(r[4]),
          leaveP: toN(r[5]), leaveSick: toN(r[6]), absent: toN(r[7]),
          leaveVac: toN(r[8]), leaveNoPay: toN(r[9]),
          leaveMat: toN(r[10]), leaveBday: toN(r[11]),
          baseWage: toN(r[12]), basePay: toN(r[13]),
          holidayPay: toN(r[14]), otPay: toN(r[15]),
          allowance: toN(r[16]), bonus: toN(r[17]),
          totalInc: toN(r[18]),
          advance: toN(r[19]), kot: toN(r[20]),
          soc: toN(r[21]), tax: toN(r[22]),
          absentDed: toN(r[23]), noPayDed: toN(r[24]),
          otherDed: toN(r[25]), totalDed: toN(r[26]),
          netPay: toN(r[27]),
          month, payType: 'monthly',
          // ไม่มีใน monthly
          otherInc: 0, late: 0,
          festivalD:0, festivalPay:0, festivalExtra:0,
        };
      }
    }
    return null;
  } catch (e) {
    console.error('getEmployeePayroll error:', e.message);
    return null;
  }
}

// ─── getAvailableMonths ──────────────────────────────────────
// คืน array of { month: string, payType: 'daily'|'monthly' }
async function getAvailableMonths() {
  const sheets = getSheetsClient();
  const meta   = await sheets.spreadsheets.get({ spreadsheetId: process.env.LOG_SHEET_ID });
  return (meta.data.sheets || [])
    .map(s => s.properties.title)
    .filter(t => t.startsWith('เงินเดือน_') || t.startsWith('เงินเดือนรายวัน_'))
    .map(t => {
      if (t.startsWith('เงินเดือนรายวัน_')) {
        return { month: t.replace('เงินเดือนรายวัน_', ''), payType: 'daily' };
      }
      return { month: t.replace('เงินเดือน_', ''), payType: 'monthly' };
    });
}

module.exports = { parseXls, savePayrollToSheet, getEmployeePayroll, getAvailableMonths, normName };
