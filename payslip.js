const axios    = require('axios');
const pdfStore   = {};
const imageStore = {}; // เก็บ JPG buffer สำหรับส่งเป็นรูปใน LINE

async function createFromPayroll(d) {
  const fmt  = n => (parseFloat(n)||0).toLocaleString('th-TH', { minimumFractionDigits: 2 });
  const fmtN = n => { const v = parseFloat(n)||0; return v > 0 ? v.toString() : '0'; };
  const n    = x => parseFloat(x) || 0;
  const zero = v => n(v) > 0 ? fmt(v) : '0.00';
  const totalInc = n(d.totalInc);
  const totalDed = n(d.totalDed);
  const netPay   = n(d.netPay);
  const now = new Date().toLocaleString('th-TH', { timeZone: 'Asia/Bangkok' });

  const isDaily = d.payType === 'daily';

  // ── rows ฝั่งรายได้ ──
  // รายวัน: แสดง ค่าแรง/วัน×วัน, OT, หยุดสัปดาห์, ประเพณี, เบี้ยเลี้ยง, เบี้ยขยัน, อื่นๆ
  // รายเดือน: แสดง เงินเดือน, โบนัส, เบี้ยเลี้ยง, เบี้ยขยัน, อื่นๆ
  const incRows = isDaily ? `
      <tr>
        <td>${isDaily ? `ค่าแรง/วัน (${fmt(d.baseWage)} × ${fmtN(d.workDays)} วัน)` : 'วันทำงานปกติ'}</td>
        <td class="c">${fmtN(d.workDays)}</td>
        <td class="n">${fmt(d.basePay)}</td>
        <td>ลาพักร้อน</td><td class="c">${fmtN(d.leaveVac)}</td>
        <td>หักเบิกล่วงหน้า</td><td class="c">${n(d.advance)>0?'1':''}</td><td class="n">${zero(d.advance)}</td>
      </tr>
      <tr>
        <td>ทำงานวันหยุดสัปดาห์</td><td class="c">${fmtN(d.holidayD)}</td><td class="n">${zero(d.holidayPay)}</td>
        <td>ลากิจ</td><td class="c">${fmtN(d.leaveP)}</td>
        <td>หักประกันสังคม</td><td class="c">${n(d.soc)>0?'1':''}</td><td class="n">${fmt(d.soc)}</td>
      </tr>
      <tr>
        <td>ทำงานล่วงเวลา (OT)</td><td class="c">${fmtN(d.otH)} ชม.</td><td class="n">${zero(d.otPay)}</td>
        <td>ลาป่วย</td><td class="c">${fmtN(d.leaveSick)}</td>
        <td>หักกยศ.</td><td class="c">${n(d.kot)>0?'1':''}</td><td class="n">${zero(d.kot)}</td>
      </tr>
      <tr>
        <td>วันหยุดประเพณี</td><td class="c">${fmtN(d.festivalD)}</td><td class="n">${zero(d.festivalPay)}</td>
        <td>ขาดงาน</td><td class="c">${fmtN(d.absent)}</td>
        <td>สาย</td><td class="c">${n(d.late)>0?'1':''}</td><td class="n">${zero(d.late)}</td>
      </tr>
      <tr>
        <td>เพิ่มเติมประเพณี</td><td class="c"></td><td class="n">${zero(d.festivalExtra)}</td>
        <td></td><td></td>
        <td>รายจ่ายอื่นๆ</td><td></td><td class="n">${zero(d.otherDed)}</td>
      </tr>
      <tr>
        <td>เบี้ยเลี้ยง</td><td class="c">${n(d.allowance)>0?'1':''}</td><td class="n">${zero(d.allowance)}</td>
        <td></td><td></td><td></td><td></td><td></td>
      </tr>
      <tr>
        <td>เบี้ยขยัน</td><td class="c">${n(d.bonus)>0?'1':''}</td><td class="n">${zero(d.bonus)}</td>
        <td></td><td></td><td></td><td></td><td></td>
      </tr>
      <tr>
        <td>รายได้อื่นๆ</td><td></td><td class="n">${zero(d.otherInc)}</td>
        <td></td><td></td><td></td><td></td><td></td>
      </tr>` : `
      <tr>
        <td>วันทำงานปกติ</td>
        <td class="c">${fmtN(d.workDays)}</td>
        <td class="n">${fmt(d.basePay||d.baseWage)}</td>
        <td>ลาพักร้อน</td><td class="c">${fmtN(d.leaveVac)}</td>
        <td>หักเบิกเงินล่วงหน้า</td>
        <td class="c">${n(d.advance)>0?'1':''}</td>
        <td class="n">${zero(d.advance)}</td>
      </tr>
      <tr>
        <td>ทำงานวันหยุด</td>
        <td class="c">${fmtN(d.holidayD)}</td>
        <td class="n">${zero(d.holidayPay)}</td>
        <td>ลากิจ</td><td class="c">${fmtN(d.leaveP)}</td>
        <td>หักประกันสังคม</td>
        <td class="c">${n(d.soc)>0?'1':''}</td>
        <td class="n">${fmt(d.soc)}</td>
      </tr>
      <tr>
        <td>ทำงานล่วงเวลา</td>
        <td class="c">${n(d.otH)>0?fmtN(d.otH)+' ชม.':0}</td>
        <td class="n">${zero(d.otPay)}</td>
        <td>ลาป่วย</td><td class="c">${fmtN(d.leaveSick)}</td>
        <td>หัก ณ ที่จ่าย</td>
        <td class="c">${n(d.tax)>0?'1':''}</td>
        <td class="n">${zero(d.tax)}</td>
      </tr>
      <tr>
        <td>เบี้ยเลี้ยง</td>
        <td class="c">${n(d.allowance)>0?'1':''}</td>
        <td class="n">${zero(d.allowance)}</td>
        <td>ลาไม่รับค่าจ้าง</td><td class="c">${fmtN(d.leaveNoPay||0)}</td>
        <td>หักขาดงาน</td>
        <td class="c">${n(d.absentDed)>0?'1':''}</td>
        <td class="n">${zero(d.absentDed)}</td>
      </tr>
      <tr>
        <td>เบี้ยขยัน</td>
        <td class="c">${n(d.bonus)>0?'1':''}</td>
        <td class="n">${zero(d.bonus)}</td>
        <td>ลาคลอด</td><td class="c">${fmtN(d.leaveMat||0)}</td>
        <td>หักขอลาโดยไม่ขอรับ</td>
        <td class="c">${n(d.noPayDed)>0?'1':''}</td>
        <td class="n">${zero(d.noPayDed)}</td>
      </tr>
      <tr>
        <td>วันหยุดตามประเพณี</td>
        <td class="c">${fmtN(d.festivalD||0)}</td>
        <td class="n">${zero(d.festivalPay||0)}</td>
        <td>ลาหยุดวันเกิด</td><td class="c">${fmtN(d.leaveBday||0)}</td>
        <td>หักกยศ.</td>
        <td class="c">${n(d.kot)>0?'1':''}</td>
        <td class="n">${n(d.kot)>0?fmt(d.kot):'-'}</td>
      </tr>
      <tr>
        <td>อื่นๆ</td>
        <td></td>
        <td class="n">${zero(d.otherInc)}</td>
        <td></td><td></td>
        <td></td><td></td><td></td>
      </tr>`;

  const html = `<!DOCTYPE html><html><head><meta charset="utf-8">
<style>
@import url('https://fonts.googleapis.com/css2?family=Sarabun:wght@400;600;700&display=swap');
*{margin:0;padding:0;box-sizing:border-box}
body{font-family:'Sarabun',sans-serif;font-size:11.5px;color:#1a1a1a;background:#fff;padding:20px 24px}
.header{display:flex;align-items:stretch;border:1.5px solid #C9952A;border-radius:6px;overflow:hidden;margin-bottom:0}
.logo-box{background:transparent;width:90px;flex-shrink:0;display:flex;align-items:center;justify-content:center;padding:8px}
.logo-img{width:74px;height:74px;object-fit:contain}
.header-info{flex:1;padding:10px 14px}
.doc-title-main{font-size:15px;font-weight:700;text-align:center;margin-bottom:4px;border-bottom:1px solid #E8D9C0;padding-bottom:4px}
.co-sub{font-size:10px;color:#666;margin-top:1px}
.header-right{width:115px;flex-shrink:0;background:#FDF5E8;display:flex;flex-direction:column;align-items:center;justify-content:center;padding:10px;border-left:1px solid #E8D9C0}
.period-label{font-size:9px;color:#888;margin-bottom:3px}
.period-value{font-size:13px;font-weight:700;color:#7A4F00;text-align:center;line-height:1.3}
.type-badge{font-size:9px;background:${isDaily?'#1E3A5F':'#1B7F4E'};color:#fff;border-radius:4px;padding:2px 6px;margin-top:4px}
.emp-row{display:flex;gap:30px;background:#FDF5E8;border:1.5px solid #C9952A;border-top:none;padding:6px 14px}
.emp-item{font-size:11px}.emp-item span{color:#888}
.table-wrap{border:1.5px solid #C9952A;border-top:none}
.sec-hdr{display:grid;grid-template-columns:1fr 1fr;background:#C9952A;color:#fff;font-size:11px;font-weight:700;text-align:center;padding:4px 0}
.sec-hdr .sl{border-right:1px solid rgba(255,255,255,0.3)}
table{width:100%;border-collapse:collapse;font-size:10.5px}
th{background:#D4A030;color:#fff;padding:4px 5px;font-weight:600;border:1px solid #C9952A;font-size:10px;text-align:center}
td{padding:3.5px 5px;border:1px solid #E5D0A0}
td.n{text-align:right}td.c{text-align:center}
tr:nth-child(even) td{background:#FDFAF4}
.tot td{background:#FDF0D0!important;font-weight:700;border-top:2px solid #C9952A}
.net td{background:#7A4F00!important;color:#fff!important;font-size:13px;font-weight:700;padding:6px 7px;border:none}
.net td.r{text-align:right;font-size:14px}
</style></head><body>
<div class="header">
  <div class="logo-box"><img class="logo-img" src="data:image/png;base64,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"></div>
  <div class="header-info">
    <div class="doc-title-main">ใบสลิปเงินเดือน (PAY SLIP)</div>
    <div class="co-sub">บริษัท ธนพลเอ็นจิเนียริ่ง จำกัด / ที่อยู่ 2 ถ.คลองหมอ ต.บ้านพรุ อ.หาดใหญ่ จ.สงขลา</div>
    <div class="co-sub">เลขประจำตัวผู้เสียภาษี : 0905559005578 | โทร : 086 488 0822</div>
  </div>
  <div class="header-right">
    <div class="period-label">รอบเงินเดือน</div>
    <div class="period-value">${d.month || ''}</div>
    <div class="type-badge">${isDaily ? 'รายวัน' : 'รายเดือน'}</div>
  </div>
</div>
<div class="emp-row">
  <div class="emp-item"><span>ชื่อ - นามสกุล : </span><b>${d.name}</b></div>
  <div class="emp-item"><span>ตำแหน่ง : </span><b>${d.position || '—'}</b></div>
  ${isDaily ? `<div class="emp-item"><span>ค่าแรง/วัน : </span><b>${fmt(d.baseWage)} บาท</b></div>` : ''}
</div>
<div class="table-wrap">
  <div class="sec-hdr"><div class="sl">รายการเงินได้</div><div>รายการเงินหัก</div></div>
  <table>
    <thead><tr>
      <th style="width:22%">รายละเอียด</th><th style="width:6%">จำนวน</th><th style="width:11%">จำนวนเงิน</th>
      <th style="width:13%">รายละเอียดวันลา</th><th style="width:5%">จำนวน</th>
      <th style="width:19%">รายการเงินหัก</th><th style="width:6%">จำนวน</th><th style="width:11%">จำนวนเงิน</th>
    </tr></thead>
    <tbody>
      ${incRows}
      <tr class="tot">
        <td colspan="2" style="text-align:right">รวมรายได้</td><td class="n">${fmt(totalInc)}</td>
        <td colspan="2"></td>
        <td style="text-align:right">รวมเงินหัก</td><td></td><td class="n">${fmt(totalDed)}</td>
      </tr>
      <tr class="net">
        <td colspan="6">เงินได้สุทธิ/บาท</td>
        <td colspan="2" class="r">${fmt(netPay)}</td>
      </tr>
    </tbody>
  </table>
</div>
<div style="margin-top:5px;font-size:9px;color:#aaa;text-align:right">สร้างโดยระบบ HR อัตโนมัติ | ${now}</div>
</body></html>`;

  const pdfBuf = await htmlToPdfBuffer(html);
  // เก็บ html ไว้ใน pdfBuffer object เพื่อใช้ screenshot ภายหลัง
  pdfBuf._html = html;
  return pdfBuf;
}

async function htmlToPdfBuffer(html) {
  const chromium  = require('@sparticuz/chromium');
  const puppeteer = require('puppeteer-core');

  const browser = await puppeteer.launch({
    args: chromium.args,
    defaultViewport: chromium.defaultViewport,
    executablePath: await chromium.executablePath(),
    headless: chromium.headless,
  });

  const page = await browser.newPage();
  await page.setContent(html, { waitUntil: 'networkidle0', timeout: 30000 });

  const pdfBuffer = await page.pdf({
    format: 'A4',
    printBackground: true,
    margin: { top: '10mm', bottom: '10mm', left: '10mm', right: '10mm' },
  });

  await browser.close();
  return Buffer.from(pdfBuffer);
}

// ── แปลง HTML → JPG buffer (สำหรับส่งเป็นรูปใน LINE) ──────
async function htmlToImageBuffer(html) {
  const chromium  = require('@sparticuz/chromium');
  const puppeteer = require('puppeteer-core');

  const browser = await puppeteer.launch({
    args: chromium.args,
    defaultViewport: { width: 794, height: 1123, deviceScaleFactor: 2 },
    executablePath: await chromium.executablePath(),
    headless: chromium.headless,
  });

  const page = await browser.newPage();
  await page.setContent(html, { waitUntil: 'networkidle0', timeout: 30000 });

  // รอ font โหลด
  await page.evaluate(() => document.fonts.ready);

  const imgBuffer = await page.screenshot({
    type: 'jpeg',
    quality: 92,
    fullPage: true,
  });

  await browser.close();
  return Buffer.from(imgBuffer);
}

async function sendPdfToLine(userId, pdfBuffer, filename) {
  const LINE_TOKEN = process.env.LINE_ACCESS_TOKEN;
  const RENDER_URL = process.env.RENDER_URL || 'https://tpe-hr-bot.onrender.com';

  const token    = Math.random().toString(36).slice(2) + Date.now().toString(36);
  const imgToken = token + '_img';

  // เก็บ PDF
  pdfStore[token] = { buffer: pdfBuffer, filename, createdAt: Date.now() };
  const pdfUrl = RENDER_URL + '/pdf/' + token;
  console.log('sendPdf: serving at', pdfUrl);

  // push ลิงก์ PDF ก่อน (fallback ถ้า push image ไม่ได้)
  const result = await axios.post(
    'https://api.line.me/v2/bot/message/push',
    {
      to: userId,
      messages: [{
        type: 'text',
        text: `📄 เอกสารพร้อมแล้วครับ\n🔗 กดลิงก์เพื่อดาวน์โหลด PDF (ใช้ได้ 1 ชั่วโมง)\n\n${pdfUrl}`,
      }]
    },
    { headers: { 'Authorization': 'Bearer ' + LINE_TOKEN, 'Content-Type': 'application/json' } }
  ).then(() => ({ ok: true })).catch(e => {
    if (e.response?.status === 429) return { fallback: true };
    throw e;
  });

  setTimeout(() => { delete pdfStore[token]; delete imageStore[imgToken]; }, 60 * 60 * 1000);
  return result;
}

async function htmlToDriveUrl(html, filename) {
  return await htmlToPdfBuffer(html);
}

// ── ส่งเอกสารเป็นรูปภาพ + ลิงก์ PDF ──────────────────────
// html = HTML string ที่ใช้สร้าง PDF/รูป
// pdfBuffer = Buffer ของ PDF ที่สร้างแล้ว
async function sendDocToLine(userId, html, pdfBuffer, filename) {
  const LINE_TOKEN = process.env.LINE_ACCESS_TOKEN;
  const RENDER_URL = process.env.RENDER_URL || 'https://tpe-hr-bot.onrender.com';

  // สร้าง token
  const token    = Math.random().toString(36).slice(2) + Date.now().toString(36);
  const imgToken = token + '_img';

  // เก็บ PDF
  pdfStore[token] = { buffer: pdfBuffer, filename, createdAt: Date.now() };
  const pdfUrl = RENDER_URL + '/pdf/' + token;

  // สร้าง JPG จาก HTML
  let imgBuffer = null;
  try {
    imgBuffer = await htmlToImageBuffer(html);
    imageStore[imgToken] = { buffer: imgBuffer, createdAt: Date.now() };
  } catch(e) {
    console.error('htmlToImageBuffer error:', e.message);
  }

  const imgUrl = RENDER_URL + '/img/' + imgToken;
  console.log('sendDoc: pdf=', pdfUrl, 'img=', imgUrl);

  // สร้าง messages array
  const messages = [];

  // ถ้ามีรูป — ส่งรูปก่อน
  if (imgBuffer) {
    messages.push({
      type: 'image',
      originalContentUrl: imgUrl,
      previewImageUrl:    imgUrl,
    });
  }

  // ส่งลิงก์ PDF
  messages.push({
    type: 'text',
    text: `📄 กดลิงก์เพื่อดาวน์โหลด PDF (ใช้ได้ 1 ชั่วโมง)\n${pdfUrl}`,
  });

  // push
  try {
    await axios.post(
      'https://api.line.me/v2/bot/message/push',
      { to: userId, messages },
      { headers: { 'Authorization': 'Bearer ' + LINE_TOKEN, 'Content-Type': 'application/json' } }
    );
    setTimeout(() => { delete pdfStore[token]; delete imageStore[imgToken]; }, 60 * 60 * 1000);
    return { ok: true };
  } catch(e) {
    if (e.response?.status === 429) {
      console.log('sendDocToLine: 429 quota exhausted');
      return { fallback: true };
    }
    throw e;
  }
}

module.exports = { createFromPayroll, htmlToPdfBuffer, htmlToImageBuffer, htmlToDriveUrl, sendPdfToLine, sendDocToLine, pdfStore, imageStore };
