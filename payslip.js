const { google } = require('googleapis');
const { Readable } = require('stream');

async function create(d) {
  const fmt  = n => (n||0).toLocaleString('th-TH', { minimumFractionDigits: 2 });
  const fmtN = n => n > 0 ? String(n) : '';
  const totalIncome = (d.baseSalary||0)+(d.hPay||0)+(d.otPay||0)+(d.dilPay||0)+(d.tradPay||0)+(d.otInc||0);
  const totalDeduct = (d.dAdv||0)+(d.dSoc||0)+(d.dTax||0)+(d.dAbs||0)+(d.dNpL||0)+(d.dKt||0)+(d.dOth||0);
  const netPay = totalIncome - totalDeduct;
  const now   = new Date().toLocaleString('th-TH', { timeZone: 'Asia/Bangkok' });

  const html = `<!DOCTYPE html><html><head><meta charset="utf-8">
<style>
@import url('https://fonts.googleapis.com/css2?family=Sarabun:wght@400;600;700&display=swap');
*{margin:0;padding:0;box-sizing:border-box}
body{font-family:'Sarabun',sans-serif;font-size:11px;padding:16px}
.header{display:grid;grid-template-columns:90px 1fr auto;gap:10px;align-items:center;margin-bottom:8px}
.logo{width:80px;height:62px;background:linear-gradient(135deg,#8B6914,#C9952A);border-radius:7px;display:flex;flex-direction:column;align-items:center;justify-content:center;font-weight:900;color:#fff;line-height:1.1}
.logo .tp{font-size:22px}.logo .cb{font-size:16px}
.co-name{font-size:13px;font-weight:700}.co-sub{font-size:10px;color:#555;margin-top:1px}
.doc-title{font-size:13px;font-weight:700;color:#C9952A;border:2px solid #C9952A;padding:3px 10px;border-radius:3px;display:inline-block}
.doc-month{font-size:10px;color:#888;margin-top:4px;text-align:right}
.emp-row{border:1px solid #ddd;background:#fafafa;padding:6px 10px;margin-bottom:6px;border-radius:3px;font-size:10.5px}
.emp-row span{color:#888}
table{width:100%;border-collapse:collapse;font-size:10.5px}
th{background:#C9952A;color:#fff;padding:4px 5px;font-weight:600;border:1px solid #b8860b;font-size:10px}
td{padding:3.5px 5px;border:1px solid #ddd}
td.n{text-align:right}td.c{text-align:center}
.tot td{background:#fdf3e3;font-weight:700;border-top:2px solid #C9952A}
.net td{background:#1E3A5F;color:#fff;font-size:12px;font-weight:700;border:1px solid #1E3A5F}
.net td.n{text-align:right;font-size:14px}
.zero{color:#ccc}
</style></head><body>
<div class="header">
  <div class="logo"><div class="tp">TP</div><div class="cb">CB</div></div>
  <div>
    <div class="co-name">บริษัท ธนพลเอ็นจิเนียริ่ง จำกัด</div>
    <div class="co-sub">ที่อยู่ 2 ถ.คลองหมอ ต.บ้านพรุ อ.หาดใหญ่ จ.สงขลา 90250</div>
    <div class="co-sub">เลขประจำตัวผู้เสียภาษี : 0905559005578</div>
    <div class="co-sub">โทรศัพท์ : 086 488 0822  |  E-mail : tpe.paphatsanan@gmail.com</div>
  </div>
  <div><div class="doc-title">ใบสลิปเงินเดือน (PAY SLIP)</div><div class="doc-month">รอบเงินเดือน<br><b>${d.month||''}</b></div></div>
</div>
<div class="emp-row"><span>ชื่อ - นามสกุล : </span><b>${d.employeeName}</b></div>
<table><thead><tr>
  <th style="width:25%">รายละเอียด</th><th style="width:7%">จำนวน</th><th style="width:13%">จำนวนเงิน</th>
  <th style="width:17%">รายละเอียดวันลา</th><th style="width:7%">จำนวน</th>
  <th style="width:14%">รายการเงินหัก</th><th style="width:7%">จำนวน</th><th style="width:13%">จำนวนเงิน</th>
</tr></thead><tbody>
<tr><td>วันทำงานปกติ</td><td class="c">${fmtN(d.workDays)}</td><td class="n">${fmt(d.baseSalary)}</td><td>ลาพักร้อน</td><td class="c">${fmtN(d.lv)}</td><td>หักเบิกเงินล่วงหน้า</td><td class="c">${(d.dAdv||0)>0?'1':''}</td><td class="n">${(d.dAdv||0)>0?fmt(d.dAdv):'<span class="zero">0.00</span>'}</td></tr>
<tr><td>ทำงานวันหยุด</td><td class="c">${fmtN(d.hDays)}</td><td class="n">${(d.hPay||0)>0?fmt(d.hPay):'<span class="zero">0.00</span>'}</td><td>ลากิจ</td><td class="c">${fmtN(d.lp)}</td><td>หักประกันสังคม</td><td class="c">${(d.dSoc||0)>0?'1':''}</td><td class="n">${fmt(d.dSoc)}</td></tr>
<tr><td>ทำงานล่วงเวลา</td><td class="c">${fmtN(d.otH)}</td><td class="n">${(d.otPay||0)>0?fmt(d.otPay):'<span class="zero">0.00</span>'}</td><td>ลาป่วย</td><td class="c">${fmtN(d.ls)}</td><td>หัก ณ ที่จ่าย</td><td class="c">${(d.dTax||0)>0?'1':''}</td><td class="n">${(d.dTax||0)>0?fmt(d.dTax):'<span class="zero">0.00</span>'}</td></tr>
<tr><td>เบี้ยเลี้ยง</td><td class="c">${(d.al||0)>0?'1':''}</td><td class="n">${(d.al||0)>0?fmt(d.al):'<span class="zero">0.00</span>'}</td><td>ลาไม่รับค่าจ้าง</td><td class="c">${fmtN(d.lnp)}</td><td>หักขาดงาน</td><td class="c">${(d.dAbs||0)>0?'1':''}</td><td class="n">${(d.dAbs||0)>0?fmt(d.dAbs):'<span class="zero">0.00</span>'}</td></tr>
<tr><td>เบี้ยขยัน</td><td class="c">${(d.dilPay||0)>0?'1':''}</td><td class="n">${(d.dilPay||0)>0?fmt(d.dilPay):'<span class="zero">0.00</span>'}</td><td>ลาคลอด</td><td class="c">${fmtN(d.lm)}</td><td>หักขอลาไม่รับค่าจ้าง</td><td class="c">${(d.dNpL||0)>0?'1':''}</td><td class="n">${(d.dNpL||0)>0?fmt(d.dNpL):'<span class="zero">0.00</span>'}</td></tr>
<tr><td>วันหยุดตามประเพณี</td><td class="c">${fmtN(d.trad)}</td><td class="n">${(d.tradPay||0)>0?fmt(d.tradPay):'<span class="zero">0.00</span>'}</td><td>ลาหยุดวันเกิด</td><td class="c">${fmtN(d.lb)}</td><td>หักกยศ.</td><td class="c">${(d.dKt||0)>0?'1':''}</td><td class="n">${(d.dKt||0)>0?fmt(d.dKt):'-'}</td></tr>
<tr><td>อื่นๆ</td><td></td><td class="n">${(d.otInc||0)>0?fmt(d.otInc):'<span class="zero">0.00</span>'}</td><td colspan="2"></td><td>หักอื่นๆ</td><td></td><td class="n">${(d.dOth||0)>0?fmt(d.dOth):'<span class="zero">0.00</span>'}</td></tr>
<tr class="tot"><td colspan="2" style="text-align:right">รวม รายการเงินได้</td><td class="n">${fmt(totalIncome)}</td><td colspan="2"></td><td style="text-align:right">รวม รายการเงินหัก</td><td></td><td class="n">${fmt(totalDeduct)}</td></tr>
<tr class="net"><td colspan="6">เงินเดือนสุทธิ/บาท</td><td colspan="2" class="n">${fmt(netPay)}</td></tr>
</tbody></table>
<div style="margin-top:5px;font-size:9px;color:#bbb;text-align:right">สร้างโดยระบบ HR อัตโนมัติ | ${now}</div>
</body></html>`;

  return await htmlToDriveUrl(html, `PaySlip_${d.employeeName}_${(d.month||'').replace(/ /g,'_')}`);
}

async function htmlToDriveUrl(html, filename) {
  const auth = new google.auth.GoogleAuth({
    credentials: JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT_JSON),
    scopes: ['https://www.googleapis.com/auth/drive'],
  });
  const drive = google.drive({ version: 'v3', auth });

  // อัปโหลด HTML ไป Drive แล้วแปลงเป็น Google Doc
  const htmlStream = Readable.from([html]);
  const uploaded = await drive.files.create({
    requestBody: {
      name: filename,
      mimeType: 'application/vnd.google-apps.document',
      parents: [process.env.DRIVE_FOLDER_ID],
    },
    media: { mimeType: 'text/html', body: htmlStream },
  });

  const fileId = uploaded.data.id;

  // Export เป็น PDF
  const pdfRes = await drive.files.export(
    { fileId, mimeType: 'application/pdf' },
    { responseType: 'stream' }
  );

  // อัปโหลด PDF กลับ Drive
  const pdfFile = await drive.files.create({
    requestBody: {
      name: filename + '.pdf',
      mimeType: 'application/pdf',
      parents: [process.env.DRIVE_FOLDER_ID],
    },
    media: { mimeType: 'application/pdf', body: pdfRes.data },
  });

  const pdfId = pdfFile.data.id;

  // ลบ HTML doc ทิ้ง
  await drive.files.delete({ fileId });

  // ให้ permission อ่านได้
  await drive.permissions.create({
    fileId: pdfId,
    requestBody: { role: 'reader', type: 'anyone' },
  });

  return `https://drive.google.com/file/d/${pdfId}/view`;
}

module.exports = { create, htmlToDriveUrl };
