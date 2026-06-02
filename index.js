require('dotenv').config();
const path     = require('path');
const express  = require('express');
const axios    = require('axios');
const multer   = require('multer');
const fs       = require('fs');
const lark     = require('./lark');
const payroll  = require('./payroll');
const payslip  = require('./payslip');
const { sendDocToLine, imageStore } = payslip;
const cert     = require('./certificate');
const sheet    = require('./sheet');

const app    = express();
const upload = multer({ dest: '/tmp/uploads/' });
app.use(express.json());

const LINE_TOKEN = process.env.LINE_ACCESS_TOKEN;
const HR_USER_ID = process.env.HR_LINE_USER_ID;

const pending        = {};
const requestLog     = {};
const pendingNotify  = {}; // คำขอที่ push ไม่ได้ (quota หมด) รอ HR ดูใน Portal
const readyDocs      = {}; // PDF พร้อมส่ง รอพนักงาน reply มารับ
const portalPdfs     = {}; // PDF ที่ push ไม่ได้ (quota หมด) รอ HR download จาก Portal

app.get('/', (req, res) => res.send('TPE HR Bot v2 OK'));

// ════════════════════════════════════════════════════════
// LINE Webhook
// ════════════════════════════════════════════════════════
app.post('/webhook', async (req, res) => {
  res.sendStatus(200);
  try {
    const events = req.body.events || [];
    if (!events.length) return;
    const event = events[0];

    if (event.type === 'postback') { await handlePostback(event); return; }
    if (event.type !== 'message' || event.message.type !== 'text') return;

    const userId     = event.source.userId;
    const replyToken = event.replyToken;
    const msg        = event.message.text.trim();

    if (msg === 'id') { await reply(replyToken, 'User ID: ' + userId); return; }

    const larkToken = await lark.getToken();

    if (msg === 'ขอสลิปเงินเดือน' || msg === 'ขอสลีปเงินเดือน') {
      await handleDocRequest(replyToken, userId, larkToken, 'payslip'); return;
    }
    if (msg === 'ขอใบรับรองเงินเดือน') {
      await handleDocRequest(replyToken, userId, larkToken, 'certificate'); return;
    }

    if (msg.includes('เลือกเดือน:')) {
      const clean       = msg.trim();
      const firstColon  = clean.indexOf(':');
      const secondColon = clean.indexOf(':', firstColon + 1);
      const month = clean.substring(firstColon + 1, secondColon).trim();
      const reqId = clean.substring(secondColon + 1).trim();
      await handleMonthSelected(replyToken, userId, month, reqId); return;
    }

    const profile  = await getProfile(userId);
    const imgUrl   = profile?.pictureUrl || 'https://cdn-icons-png.flaticon.com/512/3135/3135715.png';
    const employee = await lark.findByLineId(larkToken, userId);

    if (employee) {
      await reply(replyToken, createLeaveCard(employee, imgUrl));
    } else {
      if (msg.length < 2) { await reply(replyToken, '⚠️ กรุณาพิมพ์ชื่อจริง เพื่อลงทะเบียนครับ'); return; }
      await reply(replyToken, await lark.register(larkToken, msg, userId));
    }

  } catch (err) { console.error('webhook error:', err.message); }
});

// ════════════════════════════════════════════════════════
// พนักงานขอเอกสาร → ถามเดือน
// ════════════════════════════════════════════════════════
async function handleDocRequest(replyToken, userId, larkToken, docType) {
  const emp = await lark.findByLineId(larkToken, userId);
  if (!emp) { await reply(replyToken, '❌ ไม่พบข้อมูลของคุณ กรุณาลงทะเบียนก่อนนะครับ'); return; }

  const rawName   = emp['ชื่อ - นามสกุล'] || emp['ชื่อ-นามสกุล'] || 'พนักงาน';
  const empName   = payroll.normName(rawName.split('(')[0].split('（')[0]);
  const requestId = (docType === 'payslip' ? 'PAY' : 'CERT') + '_' + userId + '_' + Date.now();

  // ── อ่าน field 'ประเภท' จาก Lark ─────────────────────────
  const rawType = (emp['ประเภท'] || 'รายเดือน').toString().trim();
  const payType = rawType.includes('รายวัน') ? 'daily' : 'monthly';
  console.log(`${empName} → ประเภท: "${rawType}" → payType: ${payType}`);

  pending[requestId]    = { empName, empLineId: userId, docType, larkToken, payType };
  requestLog[requestId] = { empName, empLineId: userId, docType, payType, status: 'pending', time: Date.now() };

  // ── ดึงเดือน กรองเฉพาะประเภทของพนักงานคนนี้ ──────────────
  const allMonths   = await payroll.getAvailableMonths();
  // เรียงเดือนจากเก่าไปใหม่ตามลำดับเวลาจริง
  const thMonths = ['มกราคม','กุมภาพันธ์','มีนาคม','เมษายน','พฤษภาคม','มิถุนายน',
                    'กรกฎาคม','สิงหาคม','กันยายน','ตุลาคม','พฤศจิกายน','ธันวาคม'];
  const parseMonthKey = s => {
    const m = thMonths.findIndex(n => s.includes(n));
    const y = parseInt((s.match(/25\d\d/) || ['0'])[0]) || 0;
    return y * 100 + m;
  };
  const validMonths = allMonths
    .filter(m => m.payType === payType)
    .map(m => m.month)
    .filter(m => m && m.trim().length > 0)
    .sort((a, b) => parseMonthKey(a) - parseMonthKey(b))
    .slice(-13);

  const docLabel  = docType === 'payslip' ? 'สลิปเงินเดือน' : 'ใบรับรองเงินเดือน';
  const typeLabel = rawType === 'รายวัน' ? 'รายวัน' : 'รายเดือน';

  if (validMonths.length === 0) {
    await reply(replyToken, `⚠️ ยังไม่มีข้อมูล${typeLabel}ในระบบ กรุณาให้ HR อัปโหลดไฟล์ Excel ก่อนนะครับ`);
    return;
  }

  // encode months เป็น index เพื่อหลีกเลี่ยงปัญหา | ใน Thai
  // validMonths = ['มีนาคม 2569','เมษายน 2569','พฤษภาคม 2569']
  // postback: MULTI|reqId|idx1,idx2,idx3
  const makeCard = (idxArr, label, emoji) => {
    const ms = idxArr.map(i => validMonths[i]);
    const body_items = ms.map(m => ({
      type: 'box', layout: 'horizontal',
      contents: [{ type: 'text', text: '• ' + m, size: 'sm', color: '#555555', wrap: true }]
    }));
    return {
      type: 'bubble', size: 'micro',
      header: {
        type: 'box', layout: 'vertical',
        backgroundColor: '#1E3A5F', paddingAll: '12px',
        contents: [{ type: 'text', text: emoji + ' ' + label, color: '#ffffff', weight: 'bold', size: 'sm', align: 'center' }]
      },
      body: {
        type: 'box', layout: 'vertical', spacing: 'xs', paddingAll: '12px',
        contents: [
          { type: 'text', text: docLabel, size: 'xs', color: '#888888', align: 'center' },
          { type: 'separator', margin: 'sm' },
          ...body_items
        ]
      },
      footer: {
        type: 'box', layout: 'vertical', paddingAll: '10px',
        contents: [{
          type: 'button', style: 'primary', color: '#C9952A', height: 'sm',
          action: { type: 'postback', label: 'เลือก ' + label, data: 'MULTI|' + requestId + '|' + idxArr.join(',') }
        }]
      }
    };
  };

  const n = validMonths.length;
  const bubbles = [];
  // 1 เดือนล่าสุด (ทุกประเภทเอกสาร)
  bubbles.push(makeCard([n-1], '1 เดือนล่าสุด', '📄'));
  // 3 และ 6 เดือนล่าสุด — เฉพาะสลิปเงินเดือนเท่านั้น
  if (docType === 'payslip') {
    if (n >= 2) {
      const idx3 = Array.from({length: Math.min(3,n)}, (_,i) => n - Math.min(3,n) + i);
      bubbles.push(makeCard(idx3, idx3.length + ' เดือนล่าสุด', '📋'));
    }
    if (n >= 4) {
      const idx6 = Array.from({length: Math.min(6,n)}, (_,i) => n - Math.min(6,n) + i);
      bubbles.push(makeCard(idx6, idx6.length + ' เดือนล่าสุด', '📁'));
    }
  }

  await reply(replyToken, {
    type: 'flex', altText: 'เลือกช่วงเวลา' + docLabel,
    contents: { type: 'carousel', contents: bubbles }
  });
}

// ════════════════════════════════════════════════════════
// พนักงานเลือกเดือน → แจ้ง HR
// ════════════════════════════════════════════════════════
async function handleMonthSelected(replyToken, userId, month, requestId, allMonths) {
  const req = pending[requestId];
  if (!req) { await reply(replyToken, '⚠️ คำขอหมดอายุ กรุณาขอใหม่อีกครั้งครับ'); return; }

  // เก็บ allMonths ถ้ามี (กรณีเลือกหลายเดือน)
  req.months = allMonths && allMonths.length > 1 ? allMonths : [month];
  req.month  = req.months.join(', ');
  if (requestLog[requestId]) requestLog[requestId].month = req.month;
  const docLabel = req.docType === 'payslip' ? 'สลิปเงินเดือน' : 'ใบรับรองเงินเดือน';

  await reply(replyToken, {
    type: 'flex', altText: 'ส่งคำขอแล้ว',
    contents: { type: 'bubble', body: { type: 'box', layout: 'vertical', spacing: 'md', paddingAll: '20px',
      contents: [
        { type: 'text', text: '✅ ส่งคำขอแล้ว', weight: 'bold', size: 'lg', color: '#06C755' },
        { type: 'text', text: docLabel, size: 'sm', color: '#555555', margin: 'sm' },
        { type: 'text', text: `เดือน: ${month}`, size: 'sm', color: '#888888', margin: 'xs' },
        { type: 'text', text: 'HR จะตรวจสอบและส่ง PDF กลับมาให้คุณเร็วๆ นี้ครับ', size: 'sm', color: '#888888', wrap: true, margin: 'md' },
      ]
    }}
  });

  await notifyHR(requestId, req.empName, docLabel, month, userId);
}

// ════════════════════════════════════════════════════════
// แจ้ง HR
// ════════════════════════════════════════════════════════
async function notifyHR(requestId, empName, docType, month, empUserId) {
  const now = new Date().toLocaleString('th-TH', { timeZone: 'Asia/Bangkok', hour: '2-digit', minute: '2-digit', day: '2-digit', month: '2-digit' });
  const result = await push(HR_USER_ID, {
    type: 'flex', altText: `คำขอ${docType} จาก ${empName}`,
    contents: {
      type: 'bubble',
      header: { type: 'box', layout: 'vertical', backgroundColor: '#1E3A5F', paddingAll: '16px',
        contents: [
          { type: 'text', text: 'คำขอเอกสาร HR', color: '#ffffff', weight: 'bold', size: 'md' },
          { type: 'text', text: docType, color: '#B8D4F0', size: 'sm', margin: 'xs' },
        ]},
      body: { type: 'box', layout: 'vertical', spacing: 'sm', paddingAll: '16px',
        contents: [
          { type: 'box', layout: 'horizontal', contents: [{ type: 'text', text: 'พนักงาน', size: 'sm', color: '#888888', flex: 3 }, { type: 'text', text: empName, size: 'sm', weight: 'bold', flex: 5, wrap: true }]},
          { type: 'box', layout: 'horizontal', contents: [{ type: 'text', text: 'เดือน', size: 'sm', color: '#888888', flex: 3 }, { type: 'text', text: month, size: 'sm', flex: 5 }]},
          { type: 'box', layout: 'horizontal', contents: [{ type: 'text', text: 'ประเภทเอกสาร', size: 'sm', color: '#888888', flex: 3 }, { type: 'text', text: docType, size: 'sm', flex: 5 }]},
          { type: 'box', layout: 'horizontal', contents: [{ type: 'text', text: 'เวลาขอ', size: 'sm', color: '#888888', flex: 3 }, { type: 'text', text: now, size: 'sm', flex: 5 }]},
        ]},
      footer: { type: 'box', layout: 'vertical', spacing: 'sm', paddingAll: '12px',
        contents: [
          { type: 'button', style: 'primary', color: '#1B7F4E', height: 'sm', action: { type: 'postback', label: 'อนุมัติ และส่ง PDF', data: 'A|' + requestId }},
          { type: 'button', style: 'secondary', height: 'sm', action: { type: 'postback', label: 'ปฏิเสธ', data: 'R|' + requestId }},
        ]}
    }
  });
  // ถ้า LINE quota หมด บันทึกใน pendingNotify เพื่อให้ Portal แสดง
  if (result?.fallback) {
    pendingNotify[requestId] = { requestId, empName, docType, month, empUserId, time: Date.now() };
    console.log('pendingNotify saved:', requestId);
  }
}

// ════════════════════════════════════════════════════════
// Postback: HR กดอนุมัติ/ปฏิเสธ
// ════════════════════════════════════════════════════════
async function handlePostback(event) {
  const hrUserId = event.source.userId;
  const data     = event.postback.data;
  const action   = data.startsWith('A|') ? 'approve' : data.startsWith('R|') ? 'reject' : null;
  const rid      = action ? data.slice(2) : null;

  // พนักงานเลือกหลายเดือน (Flex carousel)
  if (data.startsWith('MULTI|')) {
    // format: MULTI|requestId|idx1,idx2,idx3
    const parts   = data.split('|');
    const reqId   = parts[1];
    const idxList = parts[2].split(',').map(Number);
    // ดึง validMonths จาก pending
    const req2 = pending[reqId];
    if (!req2) { await reply(event.replyToken, '⚠️ คำขอหมดอายุ กรุณาขอใหม่ครับ'); return; }
    const allM  = await payroll.getAvailableMonths();
    const thMon = ['มกราคม','กุมภาพันธ์','มีนาคม','เมษายน','พฤษภาคม','มิถุนายน',
                   'กรกฎาคม','สิงหาคม','กันยายน','ตุลาคม','พฤศจิกายน','ธันวาคม'];
    const parseKey = s => { const mi = thMon.findIndex(n => s.includes(n)); const yr = parseInt((s.match(/25[0-9]{2}/) || ['0'])[0]); return yr*100+mi; };
    const vMonths = allM.filter(m => m.payType === req2.payType).map(m => m.month)
      .filter(m => m && m.trim()).sort((a,b) => parseKey(a)-parseKey(b));
    const months = idxList.map(i => vMonths[i]).filter(Boolean);
    if (!months.length) { await reply(event.replyToken, '⚠️ ไม่พบข้อมูลเดือนที่เลือก'); return; }
    await handleMonthSelected(event.replyToken, event.source.userId, months[months.length-1], reqId, months);
    return;
  }

  if (action === 'reject') {
    const req = pending[rid] || {};
    if (req.empLineId) await push(req.empLineId, '❌ คำขอ' + (req.docType === 'payslip' ? 'สลิปเงินเดือน' : 'ใบรับรองเงินเดือน') + 'ของคุณถูกปฏิเสธ กรุณาติดต่อ HR โดยตรงครับ');
    await push(hrUserId, '✅ ปฏิเสธคำขอของ ' + (req.empName || 'พนักงาน') + ' แล้ว');
    delete pending[rid];
    if (requestLog[rid]) requestLog[rid].status = 'rejected';
    return;
  }

  if (action === 'approve') {
    const req = pending[rid];
    if (!req) { await push(hrUserId, '⚠️ ไม่พบคำขอนี้ อาจดำเนินการไปแล้ว'); return; }

    await push(hrUserId, `⏳ กำลังสร้าง PDF สำหรับ ${req.empName} เดือน ${req.month}...`);
    try {
      const docLabel = req.docType === 'payslip' ? 'สลิปเงินเดือน' : 'ใบรับรองเงินเดือน';
      const safeName = req.empName.replace(/[^ก-๙a-zA-Z0-9 ]/g, '').trim();
      const months   = req.months || [req.month];
      const filename = docLabel + '_' + safeName + '_' + months.join('-') + '.pdf';

      // สร้าง PDF ทีละเดือน แล้วรวมเป็นไฟล์เดียว
      const pdfBuffers = [];
      for (const m of months) {
        const empData = await payroll.getEmployeePayroll(req.empName, m, req.payType || 'monthly');
        if (!empData) {
          await push(hrUserId, `❌ ไม่พบข้อมูลเงินเดือนของ ${req.empName} เดือน ${m}`);
          continue;
        }
        const buf = req.docType === 'payslip'
          ? await payslip.createFromPayroll(empData)
          : await cert.createFromPayroll(empData);
        pdfBuffers.push(buf);
      }

      if (!pdfBuffers.length) return;

      // รวม PDF หลายหน้า
      const htmlArr   = pdfBuffers.map(b => b._html || '');
      const pdfBuffer = pdfBuffers.length === 1
        ? pdfBuffers[0]
        : await mergePdfs(pdfBuffers);

      // ส่งรูปทุกเดือน + ลิงก์ PDF รวม
      const sendResult = await sendDocToLine(req.empLineId, htmlArr, pdfBuffer, filename);
      delete pending[rid];
      if (requestLog[rid]) requestLog[rid].status = 'sent';
      // log เดือนแรก (สำหรับ multi-month ใช้ข้อมูลรวม)
      await sheet.log({ name: req.empName, month: req.month, docType: req.docType });

    } catch(err) {
      console.error('approve error:', err.message);
      await push(hrUserId, `❌ เกิดข้อผิดพลาด: ${err.message}`);
    }
  }
}

// ════════════════════════════════════════════════════════
// HR อัปโหลด Excel
// ════════════════════════════════════════════════════════
app.post('/upload-payroll', upload.single('file'), async (req, res) => {
  res.json({ ok: true, message: 'กำลังประมวลผล...' });
  try {
    const file = req.file;
    if (!file) return;
    const data    = await payroll.parseXls(file.path);
    const { month, rows, type } = data;
    const payType = type || req.body.payType || 'monthly';
    console.log(`เดือน ${month} จำนวน ${rows.length} คน ประเภท: ${payType}`);
    await payroll.savePayrollToSheet(month, rows, payType);
    fs.unlinkSync(file.path);
    if (req.body.notifyUserId) {
      await push(req.body.notifyUserId, `✅ อัปโหลดข้อมูล${payType === 'daily' ? 'รายวัน' : 'รายเดือน'} เดือน ${month} สำเร็จ\nจำนวนพนักงาน: ${rows.length} คน`);
    }
  } catch(err) { console.error('upload-payroll error:', err.message); }
});

// ════════════════════════════════════════════════════════
// Flex Card วันลา
// ════════════════════════════════════════════════════════
function createLeaveCard(emp, imgUrl) {
  const d = { name: emp['ชื่อ - นามสกุล'] || 'พนักงาน', vacationTotal: emp['สิทธิ์พักร้อน']||'0', vacationLeft: emp['คงเหลือพักร้อน']||'0', personalTotal: emp['สิทธิ์ลากิจ']||'0', personalLeft: emp['คงเหลือลากิจ']||'0', sickTotal: emp['สิทธิ์ลาป่วย']||'0', sickLeft: emp['คงเหลือลาป่วย']||'0', birthdayTotal: emp['สิทธิ์วันเกิด']||'0', birthdayLeft: emp['คงเหลือลาวันเกิด']||'0', maternityTotal: emp['สิทธิ์ลาคลอด']||'0', maternityLeft: emp['คงเหลือลาคลอด']||'0' };
  return { type: 'flex', altText: `สรุปวันลาของ ${d.name}`, contents: { type: 'bubble', size: 'giga', body: { type: 'box', layout: 'vertical', paddingAll: '0px', contents: [ { type: 'box', layout: 'horizontal', backgroundColor: '#06C755', paddingAll: '20px', contents: [ { type: 'box', layout: 'vertical', flex: 0, width: '70px', height: '70px', cornerRadius: '100px', borderColor: '#ffffff', borderWidth: '2px', contents: [{ type: 'image', url: imgUrl, aspectMode: 'cover', size: 'full' }] }, { type: 'box', layout: 'vertical', flex: 1, paddingStart: '15px', justifyContent: 'center', contents: [ { type: 'text', text: d.name, color: '#ffffff', weight: 'bold', size: 'lg', wrap: true }, { type: 'text', text: 'พนักงาน', color: '#E5E5E5', size: 'sm' } ]} ] }, { type: 'box', layout: 'vertical', paddingAll: '20px', spacing: 'md', contents: [ { type: 'text', text: 'วันลาคงเหลือ', weight: 'bold', color: '#888888', size: 'xs' }, leaveRow('🏖','ลาพักร้อน',d.vacationTotal,d.vacationLeft,'#F39C12'), { type: 'separator', color: '#F0F0F0' }, leaveRow('🚗','ลากิจ',d.personalTotal,d.personalLeft,'#3498DB'), { type: 'separator', color: '#F0F0F0' }, leaveRow('😷','ลาป่วย',d.sickTotal,d.sickLeft,'#E74C3C'), { type: 'separator', color: '#F0F0F0' }, leaveRow('🎂','ลาวันเกิด',d.birthdayTotal,d.birthdayLeft,'#9B59B6'), { type: 'separator', color: '#F0F0F0' }, leaveRow('👶','ลาคลอด',d.maternityTotal,d.maternityLeft,'#FF69B4') ] }, { type: 'box', layout: 'vertical', backgroundColor: '#F9F9F9', paddingAll: '10px', contents: [{ type: 'text', text: 'อัปเดตข้อมูลล่าสุดจากระบบ HR', color: '#AAAAAA', size: 'xxs', align: 'center' }] } ] } } };
}
function leaveRow(icon, label, total, left, color) {
  return { type: 'box', layout: 'horizontal', alignItems: 'center', contents: [ { type: 'box', layout: 'horizontal', flex: 7, contents: [ { type: 'text', text: icon, size: 'xl', flex: 0, margin: 'sm' }, { type: 'box', layout: 'vertical', paddingStart: '10px', contents: [ { type: 'text', text: label, size: 'sm', weight: 'bold', color: '#333333' }, { type: 'text', text: `สิทธิ์ทั้งหมด ${total} วัน`, size: 'xxs', color: '#AAAAAA' } ] } ] }, { type: 'box', layout: 'vertical', flex: 3, alignItems: 'flex-end', contents: [ { type: 'text', text: String(left), size: 'xxl', weight: 'bold', color }, { type: 'text', text: 'วัน', size: 'xxs', color } ] } ] };
}

const sleep = ms => new Promise(r => setTimeout(r, ms));

async function reply(replyToken, msg) {
  if (!replyToken) return; // replyToken หมดอายุ ข้ามได้
  const messages = typeof msg === 'string' ? [{ type: 'text', text: msg }] : [msg];
  try {
    await axios.post('https://api.line.me/v2/bot/message/reply',
      { replyToken, messages },
      { headers: { Authorization: `Bearer ${LINE_TOKEN}` } }
    );
  } catch (err) {
    console.error('reply error:', err.response?.data || err.message);
  }
}

async function push(userId, msg) {
  const messages = typeof msg === 'string' ? [{ type: 'text', text: msg }] : [msg];
  await sleep(300);
  try {
    await axios.post('https://api.line.me/v2/bot/message/push',
      { to: userId, messages },
      { headers: { Authorization: `Bearer ${LINE_TOKEN}` } }
    );
    return { ok: true };
  } catch (err) {
    if (err.response?.status === 429) {
      console.log('LINE push quota exhausted (429) — using portal fallback');
      return { fallback: true };
    }
    throw err;
  }
}
async function getProfile(userId) {
  try { const r = await axios.get(`https://api.line.me/v2/bot/profile/${userId}`, { headers: { Authorization: `Bearer ${LINE_TOKEN}` } }); return r.data; } catch { return null; }
}

// ════════════════════════════════════════════════════════
// PDF / Portal endpoints
// ════════════════════════════════════════════════════════
app.get('/pdf/:token', (req, res) => {
  const entry = payslip.pdfStore[req.params.token];
  if (!entry) return res.status(404).send('ไฟล์หมดอายุหรือไม่พบ');
  res.setHeader('Content-Type', 'application/pdf');
  res.setHeader('Content-Disposition', 'attachment; filename="' + encodeURIComponent(entry.filename) + '"');
  res.setHeader('Content-Length', entry.buffer.length);
  res.send(entry.buffer);
});

// ── รูปภาพ JPG endpoint ────────────────────────────────────
app.get('/img/:token', (req, res) => {
  const entry = imageStore[req.params.token];
  if (!entry) return res.status(404).send('รูปหมดอายุหรือไม่พบ');
  res.setHeader('Content-Type', 'image/jpeg');
  res.setHeader('Content-Length', entry.buffer.length);
  res.send(entry.buffer);
});
app.get('/portal', (req, res) => res.sendFile(path.join(__dirname, 'portal.html')));
app.get('/portal/stats', async (req, res) => {
  try {
    const [emps, allMonths] = await Promise.all([lark.getAllEmployees(await lark.getToken()), payroll.getAvailableMonths()]);
    res.json({ employees: emps.length, pending: Object.keys(pending).length, latestMonth: allMonths.length ? allMonths[allMonths.length-1].month : '—' });
  } catch(e) { res.json({ employees: 0, pending: 0, latestMonth: '—' }); }
});
app.get('/portal/employees', async (req, res) => {
  try {
    const emps = await lark.getAllEmployees(await lark.getToken());
    res.json(emps.map(e => ({ name: e['ชื่อ - นามสกุล']||'', position: e['ตำแหน่ง']||'', type: e['ประเภท']||'—', lineId: e['Line ID']||'' })));
  } catch(e) { res.json([]); }
});
app.get('/portal/requests', (req, res) => {
  res.json(Object.entries(requestLog).sort((a,b)=>b[1].time-a[1].time).slice(0,parseInt(req.query.limit)||50).map(([id,r])=>({ requestId:id, name:r.empName, docType:r.docType, month:r.month||'—', status:r.status||'pending', time:new Date(r.time).toLocaleString('th-TH',{timeZone:'Asia/Bangkok',hour:'2-digit',minute:'2-digit',day:'2-digit',month:'2-digit'}) })));
});
app.get('/portal/months', async (req, res) => {
  try {
    const all = await payroll.getAvailableMonths();
    res.json(all.map(m => `${m.month}${m.payType === 'daily' ? ' (รายวัน)' : ''}`));
  } catch(e) { res.json([]); }
});

// ── Portal: retry สร้าง PDF ใหม่จาก requestLog ─────────
app.post('/portal/retry', express.json(), async (req, res) => {
  const { requestId } = req.body;
  if (!requestId) return res.status(400).json({ error: 'missing requestId' });

  const log = requestLog[requestId];
  if (!log) return res.status(404).json({ error: 'request not found or expired' });

  try {
    const empData = await payroll.getEmployeePayroll(
      log.empName, log.month,
      log.payType || (log.docType === 'payslip' ? 'monthly' : 'monthly')
    );
    if (!empData) return res.status(404).json({ error: 'payroll data not found' });

    const docLabel = log.docType === 'payslip' ? 'สลิปเงินเดือน' : 'ใบรับรองเงินเดือน';
    const safeName = log.empName.replace(/[^ก-๙a-zA-Z0-9 ]/g, '').trim();
    const filename  = docLabel + '_' + safeName + '_' + log.month + '.pdf';

    const pdfBuffer = log.docType === 'payslip'
      ? await payslip.createFromPayroll(empData)
      : await cert.createFromPayroll(empData);

    const pushResult = await sendDocToLine(log.empLineId, pdfBuffer._html || '', pdfBuffer, filename)
      .catch(e => ({ fallback: true, error: e.message }));

    if (pushResult?.fallback) {
      // เก็บใน portalPdfs
      const pdfToken = require('crypto').randomBytes(8).toString('hex');
      payslip.pdfStore[pdfToken] = { buffer: pdfBuffer, filename };
      const pdfUrl = `${process.env.RENDER_URL || 'https://tpe-hr-bot.onrender.com'}/pdf/${pdfToken}`;
      portalPdfs[pdfToken] = {
        token: pdfToken, url: pdfUrl, filename,
        empName: log.empName, empLineId: log.empLineId,
        label: docLabel, month: log.month, time: Date.now(),
      };
    }

    requestLog[requestId].status = 'sent';
    return res.json({ ok: true, pushOk: !pushResult?.fallback });
  } catch(err) {
    return res.status(500).json({ error: err.message });
  }
});

// ── Portal: ลบรายการคำขอออกจากประวัติ ───────────────────
app.delete('/portal/request/:id', (req, res) => {
  const id = req.params.id;
  delete requestLog[id];
  delete pending[id];
  delete pendingNotify[id];
  res.json({ ok: true });
});

// ── Portal: PDF ที่รอ HR download (quota หมด) ───────────
app.get('/portal/pdf-ready', (req, res) => {
  res.json(Object.values(portalPdfs).map(p => ({
    token:   p.token,
    url:     p.url,
    empName: p.empName,
    label:   p.label,
    month:   p.month,
    time:    new Date(p.time).toLocaleString('th-TH', { timeZone: 'Asia/Bangkok' }),
  })));
});

// ── Portal: HR ลบ PDF ออกหลังส่งเองแล้ว ─────────────────
app.delete('/portal/pdf-ready/:token', (req, res) => {
  delete portalPdfs[req.params.token];
  delete payslip.pdfStore[req.params.token];
  res.json({ ok: true });
});

// ── Portal: คำขอที่ยังไม่ได้แจ้ง HR (quota หมด) ──────────
app.get('/portal/pending-notify', (req, res) => {
  res.json(Object.values(pendingNotify).map(n => ({
    requestId: n.requestId,
    empName:   n.empName,
    docType:   n.docType,
    month:     n.month,
    time:      new Date(n.time).toLocaleString('th-TH', { timeZone: 'Asia/Bangkok' }),
  })));
});

// ── Portal: HR approve/reject ผ่าน Portal (ไม่ผ่าน LINE) ─
app.post('/portal/approve', express.json(), async (req, res) => {
  const { requestId, action } = req.body;
  if (!requestId || !action) return res.status(400).json({ error: 'missing params' });

  const req2 = pending[requestId];
  if (!req2) return res.status(404).json({ error: 'request not found or expired' });

  delete pendingNotify[requestId];

  if (action === 'reject') {
    delete pending[requestId];
    if (requestLog[requestId]) requestLog[requestId].status = 'rejected';
    // พยายาม push พนักงาน
    await push(req2.empLineId, '❌ คำขอ' + (req2.docType === 'payslip' ? 'สลิปเงินเดือน' : 'ใบรับรองเงินเดือน') + 'ของคุณถูกปฏิเสธ กรุณาติดต่อ HR โดยตรงครับ').catch(() => {});
    return res.json({ ok: true, action: 'rejected' });
  }

  if (action === 'approve') {
    try {
      const empData = await payroll.getEmployeePayroll(req2.empName, req2.month, req2.payType || 'monthly');
      if (!empData) return res.status(404).json({ error: 'payroll data not found' });

      const docLabel = req2.docType === 'payslip' ? 'สลิปเงินเดือน' : 'ใบรับรองเงินเดือน';
      const safeName = req2.empName.replace(/[^ก-๙a-zA-Z0-9 ]/g, '').trim();
      const filename  = docLabel + '_' + safeName + '_' + req2.month + '.pdf';

      const pdfBuffer = req2.docType === 'payslip'
        ? await payslip.createFromPayroll(empData)
        : await cert.createFromPayroll(empData);

      const pushResult = await payslip.sendPdfToLine(req2.empLineId, pdfBuffer, filename).catch(e => ({ fallback: true, error: e.message }));

      if (pushResult?.fallback) {
        const tok2 = require('crypto').randomBytes(8).toString('hex');
        payslip.pdfStore[tok2] = { buffer: pdfBuffer, filename };
        const pdfUrl2 = `${process.env.RENDER_URL || 'https://tpe-hr-bot.onrender.com'}/pdf/${tok2}`;
        portalPdfs[tok2] = {
          token: tok2, url: pdfUrl2, filename,
          empName: req2.empName, empLineId: req2.empLineId,
          label: docLabel, month: req2.month, time: Date.now(),
        };
        console.log(`portalPdfs saved for ${req2.empName}`);
      }

      delete pending[requestId];
      if (requestLog[requestId]) requestLog[requestId].status = 'sent';
      await sheet.log({ name: req2.empName, month: req2.month, docType: req2.docType });

      return res.json({ ok: true, action: 'approved', pushOk: !pushResult?.fallback });
    } catch (err) {
      return res.status(500).json({ error: err.message });
    }
  }

  res.status(400).json({ error: 'invalid action' });
});

// ── รวม PDF หลายไฟล์เป็นไฟล์เดียว ─────────────────────────
async function mergePdfs(buffers) {
  try {
    const { PDFDocument } = require('pdf-lib');
    const merged = await PDFDocument.create();
    for (const buf of buffers) {
      const doc   = await PDFDocument.load(buf);
      const pages = await merged.copyPages(doc, doc.getPageIndices());
      pages.forEach(p => merged.addPage(p));
    }
    const bytes = await merged.save();
    return Buffer.from(bytes);
  } catch(e) {
    // pdf-lib ไม่มี — ส่งแค่เดือนแรก
    console.error('mergePdfs error:', e.message);
    return buffers[0];
  }
}

// ════════════════════════════════════════════════════════
// Start server
// ════════════════════════════════════════════════════════
const PORT = parseInt(process.env.PORT) || 3000;
function startServer(port) {
  const srv = app.listen(port, '0.0.0.0', () => console.log(`TPE HR Bot v2 on port ${port}`));
  srv.on('error', err => {
    if (err.code === 'EADDRINUSE') { console.log(`Port ${port} busy, trying ${port+1}`); srv.close(() => startServer(port+1)); }
    else { console.error('Server error:', err); process.exit(1); }
  });
}

// ════════════════════════════════════════════════════════
// HR อัปโหลด Excel
// ════════════════════════════════════════════════════════

// ════════════════════════════════════════════════════════
// E-Slip LIFF Endpoints
// ════════════════════════════════════════════════════════
app.get('/eslip', (req, res) => res.sendFile(path.join(__dirname, 'eslip.html')));

// พนักงาน: ข้อมูลตัวเอง
app.get('/eslip/employee', async (req, res) => {
  try {
    const { lineId } = req.query;
    if (!lineId) return res.status(400).json({ error: 'missing lineId' });
    const larkToken = await lark.getToken();
    const emp = await lark.findByLineId(larkToken, lineId);
    if (!emp) return res.status(404).json({ error: 'not found' });
    const rawName = emp['ชื่อ - นามสกุล'] || emp['ชื่อ-นามสกุล'] || '';
    const rawType = (emp['ประเภท'] || 'รายเดือน').toString().trim();
    res.json({
      name: payroll.normName(rawName.split('(')[0]),
      position: emp['ตำแหน่ง'] || '',
      payType: rawType.includes('รายวัน') ? 'daily' : 'monthly',
      lineId,
    });
  } catch(e) { res.status(500).json({ error: e.message }); }
});

// เดือนที่มีข้อมูล + netPay
app.get('/eslip/months', async (req, res) => {
  try {
    const { lineId, payType } = req.query;
    if (!lineId) return res.status(400).json({ error: 'missing lineId' });
    const larkToken = await lark.getToken();
    const emp = await lark.findByLineId(larkToken, lineId);
    if (!emp) return res.status(404).json({ error: 'not found' });
    const rawName = emp['ชื่อ - นามสกุล'] || emp['ชื่อ-นามสกุล'] || '';
    const empName = payroll.normName(rawName.split('(')[0]);
    const pt = (payType || 'monthly');
    const allMonths = await payroll.getAvailableMonths();
    const thM = ['มกราคม','กุมภาพันธ์','มีนาคม','เมษายน','พฤษภาคม','มิถุนายน',
                 'กรกฎาคม','สิงหาคม','กันยายน','ตุลาคม','พฤศจิกายน','ธันวาคม'];
    const parseKey = s => { const mi = thM.findIndex(n=>s.includes(n)); const yr = parseInt((s.match(/25[0-9]{2}/)||['0'])[0]); return yr*100+mi; };
    const months = allMonths
      .filter(m => m.payType === pt)
      .map(m => m.month)
      .filter(Boolean)
      .sort((a,b) => parseKey(b)-parseKey(a)); // ใหม่ก่อน
    // ดึง netPay แต่ละเดือน
    const result = await Promise.all(months.map(async m => {
      try {
        const d = await payroll.getEmployeePayroll(empName, m, pt);
        return { month: m, netPay: d?.netPay || null };
      } catch { return { month: m, netPay: null }; }
    }));
    res.json(result);
  } catch(e) { res.status(500).json({ error: e.message }); }
});

// HTML ของเอกสาร (สำหรับแสดงใน LIFF)
app.get('/eslip/doc', async (req, res) => {
  try {
    const { lineId, docType, month } = req.query;
    const larkToken = await lark.getToken();
    const emp = await lark.findByLineId(larkToken, lineId);
    if (!emp) return res.status(404).json({ error: 'not found' });
    const rawName = emp['ชื่อ - นามสกุล'] || emp['ชื่อ-นามสกุล'] || '';
    const empName = payroll.normName(rawName.split('(')[0]);
    const rawType = (emp['ประเภท'] || 'รายเดือน').toString().trim();
    const pt = rawType.includes('รายวัน') ? 'daily' : 'monthly';
    const empData = await payroll.getEmployeePayroll(empName, month, pt);
    if (!empData) return res.status(404).json({ error: 'payroll not found' });
    const pdfBuf = docType === 'payslip'
      ? await payslip.createFromPayroll(empData)
      : await cert.createFromPayroll(empData);
    res.json({ html: pdfBuf._html || '' });
  } catch(e) { res.status(500).json({ error: e.message }); }
});

// JPG image
app.get('/eslip/image', async (req, res) => {
  try {
    const { lineId, docType, month } = req.query;
    const larkToken = await lark.getToken();
    const emp = await lark.findByLineId(larkToken, lineId);
    if (!emp) return res.status(404).send('not found');
    const rawName = emp['ชื่อ - นามสกุล'] || emp['ชื่อ-นามสกุล'] || '';
    const empName = payroll.normName(rawName.split('(')[0]);
    const rawType = (emp['ประเภท'] || 'รายเดือน').toString().trim();
    const pt = rawType.includes('รายวัน') ? 'daily' : 'monthly';
    const empData = await payroll.getEmployeePayroll(empName, month, pt);
    if (!empData) return res.status(404).send('not found');
    const pdfBuf = docType === 'payslip'
      ? await payslip.createFromPayroll(empData)
      : await cert.createFromPayroll(empData);
    const imgBuf = await payslip.htmlToImageBuffer(pdfBuf._html || '');
    res.setHeader('Content-Type', 'image/jpeg');
    res.setHeader('Content-Disposition', `attachment; filename="slip_${month}.jpg"`);
    res.send(imgBuf);
  } catch(e) { res.status(500).send(e.message); }
});

// JSON data สำหรับ renderDocCard (ไม่ต้อง puppeteer)
app.get('/eslip/data', async (req, res) => {
  try {
    const { lineId, docType, month } = req.query;
    const larkToken = await lark.getToken();
    const emp = await lark.findByLineId(larkToken, lineId);
    if (!emp) return res.status(404).json({ error: 'not found' });
    const rawName = emp['ชื่อ - นามสกุล'] || emp['ชื่อ-นามสกุล'] || '';
    const empName = payroll.normName(rawName.split('(')[0]);
    const rawType = (emp['ประเภท'] || 'รายเดือน').toString().trim();
    const pt = rawType.includes('รายวัน') ? 'daily' : 'monthly';
    const d = await payroll.getEmployeePayroll(empName, month, pt);
    if (!d) return res.status(404).json({ error: 'payroll not found' });
    res.json(d);
  } catch(e) { res.status(500).json({ error: e.message }); }
});

// PDF download
app.get('/eslip/pdf', async (req, res) => {
  try {
    const { lineId, docType, month } = req.query;
    const larkToken = await lark.getToken();
    const emp = await lark.findByLineId(larkToken, lineId);
    if (!emp) return res.status(404).send('not found');
    const rawName = emp['ชื่อ - นามสกุล'] || emp['ชื่อ-นามสกุล'] || '';
    const empName = payroll.normName(rawName.split('(')[0]);
    const rawType = (emp['ประเภท'] || 'รายเดือน').toString().trim();
    const pt = rawType.includes('รายวัน') ? 'daily' : 'monthly';
    const empData = await payroll.getEmployeePayroll(empName, month, pt);
    if (!empData) return res.status(404).send('not found');
    const pdfBuf = docType === 'payslip'
      ? await payslip.createFromPayroll(empData)
      : await cert.createFromPayroll(empData);
    const docLabel = docType === 'payslip' ? 'สลิปเงินเดือน' : 'ใบรับรองเงินเดือน';
    const filename = docLabel + '_' + empName + '_' + month + '.pdf';
    res.setHeader('Content-Type', 'application/pdf');
    res.setHeader('Content-Disposition', `attachment; filename="${encodeURIComponent(filename)}"`);
    res.send(pdfBuf);
  } catch(e) { res.status(500).send(e.message); }
});


startServer(PORT);
