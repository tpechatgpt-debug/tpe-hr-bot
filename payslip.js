const { google } = require('googleapis');
const { Readable } = require('stream');
const axios = require('axios');

async function createFromPayroll(d) {
  const fmt  = n => (parseFloat(n)||0).toLocaleString('th-TH', { minimumFractionDigits: 2 });
  const fmtN = n => { const v = parseFloat(n)||0; return v > 0 ? v.toString() : '0'; };
  const n    = x => parseFloat(x) || 0;
  const zero = v => n(v) > 0 ? fmt(v) : '0.00';

  const totalInc = n(d.totalInc);
  const totalDed = n(d.totalDed);
  const netPay   = n(d.netPay);

  const html = `<!DOCTYPE html>
<html><head><meta charset="utf-8">
<style>
@import url('https://fonts.googleapis.com/css2?family=Sarabun:wght@400;600;700&display=swap');
*{margin:0;padding:0;box-sizing:border-box}
body{font-family:'Sarabun',sans-serif;font-size:11.5px;color:#1a1a1a;background:#fff;padding:20px 24px}
.header{display:flex;align-items:stretch;margin-bottom:0;border:1.5px solid #C9952A;border-radius:6px;overflow:hidden}
.logo-box{background:linear-gradient(145deg,#2C1A00,#6B3F00,#C9952A,#8B5E00);width:96px;flex-shrink:0;display:flex;align-items:center;justify-content:center;padding:10px}
.logo-inner{background:rgba(255,255,255,0.12);border-radius:50%;width:72px;height:72px;display:flex;flex-direction:column;align-items:center;justify-content:center;border:2px solid rgba(255,255,255,0.3)}
.logo-tp{font-size:24px;font-weight:900;color:#fff;letter-spacing:-1px;line-height:1}
.logo-cb{font-size:16px;font-weight:700;color:#FFD580;line-height:1}
.header-info{flex:1;padding:10px 14px;background:#fff}
.doc-title-main{font-size:16px;font-weight:700;color:#1a1a1a;text-align:center;margin-bottom:4px;border-bottom:1px solid #E8D9C0;padding-bottom:4px}
.co-name{font-size:11px;color:#333;margin-bottom:1px}
.co-sub{font-size:10px;color:#666}
.header-right{width:120px;flex-shrink:0;background:#FDF5E8;display:flex;flex-direction:column;align-items:center;justify-content:center;padding:10px;border-left:1px solid #E8D9C0}
.period-label{font-size:9px;color:#888;margin-bottom:3px}
.period-value{font-size:13px;font-weight:700;color:#7A4F00;text-align:center;line-height:1.3}
.emp-row{display:flex;gap:0;background:#FDF5E8;border:1.5px solid #C9952A;border-top:none;padding:6px 14px;gap:30px}
.emp-item{font-size:11px}
.emp-item span{color:#888}
.emp-item b{color:#1a1a1a}
.table-wrap{border:1.5px solid #C9952A;border-top:none}
.section-header{display:grid;grid-template-columns:1fr 1fr;background:linear-gradient(90deg,#C9952A,#A07020);color:#fff;font-size:11.5px;font-weight:700;text-align:center;padding:5px 0}
.section-header .sec-left{border-right:1px solid rgba(255,255,255,0.3)}
table{width:100%;border-collapse:collapse;font-size:10.5px}
th{background:linear-gradient(180deg,#D4A030,#B8861A);color:#fff;padding:4px 6px;font-weight:600;border:1px solid #C9952A;font-size:10px;text-align:center}
td{padding:4px 6px;border:1px solid #E5D0A0}
td.n{text-align:right}td.c{text-align:center;width:50px}
tr:nth-child(even) td{background:#FDFAF4}
.tot-row td{background:#FDF0D0!important;font-weight:700;border-top:2px solid #C9952A;font-size:11px}
.net-row td{background:linear-gradient(90deg,#7A4F00,#C9952A)!important;color:#fff!important;font-size:13px;font-weight:700;padding:7px 8px;border:none}
.net-row td.r{text-align:right;font-size:15px}
.z{color:#bbb}
</style>
</head><body>
<div class="header">
  <div class="logo-box"><div class="logo-inner"><div class="logo-tp">TP</div><div class="logo-cb">CB</div></div></div>
  <div class="header-info">
    <div class="doc-title-main">ใบสลิปเงินเดือน (PAY SLIP)</div>
    <div class="co-name">บริษัท ธนพลเอ็นจิเนียริ่ง จำกัด / ที่อยู่ 2 ถ.คลองหมอ ต.บ้านพรุ อ.หาดใหญ่ จ.สงขลา</div>
    <div class="co-sub">เลขประจำตัวผู้เสียภาษี : 0905559005578</div>
    <div class="co-sub" style="margin-top:2px">ติดต่อโทร : 086 488 0822 / E-mail : tpe.paphatsanan@gmail.com</div>
  </div>
  <div class="header-right">
    <div class="period-label">รอบเงินเดือน</div>
    <div class="period-value">${d.month||''}</div>
  </div>
</div>
<div class="emp-row">
  <div class="emp-item"><span>ชื่อ - นามสกุล : </span><b>${d.name}</b></div>
  <div class="emp-item"><span>ตำแหน่ง : </span><b>${d.position||'—'}</b></div>
</div>
<div class="table-wrap">
  <div class="section-header">
    <div class="sec-left">รายการเงินได้</div>
    <div>รายการเงินหัก</div>
  </div>
  <table>
    <thead><tr>
      <th style="width:22%">รายละเอียด</th><th style="width:7%">จำนวน</th><th style="width:11%">จำนวนเงิน</th>
      <th style="width:16%">รายละเอียดวันลา</th><th style="width:7%">จำนวน</th>
      <th style="width:17%">รายการเงินหัก</th><th style="width:7%">จำนวน</th><th style="width:13%">จำนวนเงิน</th>
    </tr></thead>
    <tbody>
      <tr><td>วันทำงานปกติ</td><td class="c">${fmtN(d.workDays)}</td><td class="n">${fmt(d.basePay)}</td><td>ลาพักร้อน</td><td class="c">${fmtN(d.leaveVac)}</td><td>หักเบิกเงินล่วงหน้า</td><td class="c">${n(d.advance)>0?'1':'0'}</td><td class="n">${zero(d.advance)}</td></tr>
      <tr><td>ทำงานวันหยุด</td><td class="c">${fmtN(d.holidayD)}</td><td class="n">${zero(d.holidayPay)}</td><td>ลากิจ</td><td class="c">${fmtN(d.leaveP)}</td><td>หักประกันสังคม</td><td class="c">${n(d.soc)>0?'1':'0'}</td><td class="n">${fmt(d.soc)}</td></tr>
      <tr><td>ทำงานล่วงเวลา</td><td class="c">${fmtN(d.otH)}</td><td class="n">${zero(d.otPay)}</td><td>ลาป่วย</td><td class="c">${fmtN(d.leaveSick)}</td><td>หัก ณ ที่จ่าย</td><td class="c">${n(d.tax)>0?'1':'0'}</td><td class="n">${zero(d.tax)}</td></tr>
      <tr><td>เบี้ยเลี้ยง</td><td class="c">${n(d.allowance)>0?'1':'0'}</td><td class="n">${zero(d.allowance)}</td><td>ลาไม่รับค่าจ้าง</td><td class="c">${fmtN(d.leaveNoPay)}</td><td>หักขาดงาน</td><td class="c">${n(d.absentDed)>0?'1':'0'}</td><td class="n">${zero(d.absentDed)}</td></tr>
      <tr><td>เบี้ยขยัน</td><td class="c">${n(d.bonus)>0?'1':'0'}</td><td class="n">${zero(d.bonus)}</td><td>ลาคลอด</td><td class="c">${fmtN(d.leaveMat)}</td><td>หักขอลาไม่รับค่าจ้าง</td><td class="c">${n(d.noPayDed)>0?'1':'0'}</td><td class="n">${zero(d.noPayDed)}</td></tr>
      <tr><td>วันหยุดประเพณี</td><td class="c">${fmtN(d.holidayD)}</td><td class="n">${zero(d.holidayPay)}</td><td>ลาหยุดวันเกิด</td><td class="c">${fmtN(d.leaveBday)}</td><td>หักกยศ.</td><td class="c">${n(d.kot)>0?'1':''}</td><td class="n">${n(d.kot)>0?fmt(d.kot):'-'}</td></tr>
      <tr><td>อื่นๆ</td><td></td><td class="n">${zero(d.otherInc)}</td><td colspan="2"></td><td>อื่นๆ</td><td></td><td class="n">${zero(d.otherDed)}</td></tr>
      <tr class="tot-row"><td colspan="2" style="text-align:right">รวม รายการเงินได้</td><td class="n">${fmt(totalInc)}</td><td colspan="2"></td><td style="text-align:right">รวม รายการเงินหัก</td><td></td><td class="n">${fmt(totalDed)}</td></tr>
      <tr class="net-row"><td colspan="6">เงินเดือนสุทธิ/บาท</td><td colspan="2" class="r">${fmt(netPay)}</td></tr>
    </tbody>
  </table>
</div>
<div style="margin-top:6px;font-size:9px;color:#aaa;text-align:right">สร้างโดยระบบ HR อัตโนมัติ | ${new Date().toLocaleString('th-TH',{timeZone:'Asia/Bangkok'})}</div>
</body></html>`;

  return await htmlToPdfBuffer(html);
}

// แปลง HTML เป็น PDF buffer ผ่าน Google Docs
async function htmlToPdfBuffer(html) {
  const auth = new google.auth.GoogleAuth({
    credentials: JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT_JSON),
    scopes: ['https://www.googleapis.com/auth/drive'],
  });
  const drive = google.drive({ version: 'v3', auth });

  // อัปโหลด HTML เป็น Google Doc
  const uploaded = await drive.files.create({
    requestBody: { name: 'tmp_pdf_' + Date.now(), mimeType: 'application/vnd.google-apps.document' },
    media: { mimeType: 'text/html', body: Readable.from([html]) },
  });
  const fileId = uploaded.data.id;

  // Export เป็น PDF buffer
  const pdfRes = await drive.files.export(
    { fileId, mimeType: 'application/pdf' },
    { responseType: 'arraybuffer' }
  );

  // ลบ temp file
  await drive.files.delete({ fileId }).catch(() => {});

  return Buffer.from(pdfRes.data);
}

// ส่ง PDF ขึ้น LINE แล้วส่งให้ user
async function sendPdfToLine(userId, pdfBuffer, filename) {
  const LINE_TOKEN = process.env.LINE_ACCESS_TOKEN;

  // Step 1: Upload ไปยัง LINE Content API
  const uploadRes = await axios.post(
    'https://api-data.line.me/v2/bot/upload/multipart',
    pdfBuffer,
    {
      headers: {
        'Authorization': 'Bearer ' + LINE_TOKEN,
        'Content-Type': 'application/pdf',
        'Content-Length': pdfBuffer.length,
      }
    }
  );

  const messageId = uploadRes.data.messageId;

  // Step 2: ส่ง file message
  await axios.post(
    'https://api.line.me/v2/bot/message/push',
    {
      to: userId,
      messages: [{
        type: 'file',
        originalContentUrl: 'https://api-data.line.me/v2/bot/message/' + messageId + '/content',
        previewImageUrl: 'https://api-data.line.me/v2/bot/message/' + messageId + '/content/preview',
        fileName: filename,
        fileSize: pdfBuffer.length,
      }]
    },
    { headers: { 'Authorization': 'Bearer ' + LINE_TOKEN } }
  );
}

// htmlToDriveUrl ยังใช้อยู่ใน certificate.js
async function htmlToDriveUrl(html, filename) {
  const pdfBuffer = await htmlToPdfBuffer(html);
  // เก็บชั่วคราวใน Drive ของ Service Account แล้วคืน URL
  const auth = new google.auth.GoogleAuth({
    credentials: JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT_JSON),
    scopes: ['https://www.googleapis.com/auth/drive'],
  });
  const drive = google.drive({ version: 'v3', auth });
  const pdfFile = await drive.files.create({
    requestBody: { name: filename + '.pdf', mimeType: 'application/pdf' },
    media: { mimeType: 'application/pdf', body: Readable.from([pdfBuffer]) },
  });
  await drive.permissions.create({
    fileId: pdfFile.data.id,
    requestBody: { role: 'reader', type: 'anyone' },
  });
  return 'https://drive.google.com/file/d/' + pdfFile.data.id + '/view';
}

module.exports = { createFromPayroll, htmlToDriveUrl, htmlToPdfBuffer, sendPdfToLine };
