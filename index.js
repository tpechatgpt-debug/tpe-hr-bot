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

const telegramBot = require('./telegram');
const gramJS      = require('./gramjs');
const app    = express();
const upload = multer({ dest: '/tmp/uploads/' });
app.use(express.json());
app.use((req, res, next) => {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET,POST,PUT,DELETE,OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type,Authorization');
  if (req.method === 'OPTIONS') return res.sendStatus(200);
  next();
});

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
      // ตรวจว่าพิมพ์ "เช็ควันลา" ตรงๆ เท่านั้น
      if (msg === 'งานวันนี้') {
        const team = (employee['ชุด'] || '').toString().trim();
        if (!team) {
          await reply(replyToken, '⚠️ ยังไม่ได้กำหนดชุดให้คุณ กรุณาติดต่อหัวหน้าครับ');
          return;
        }
        const jobs = await lark.getJobsToday(larkToken, team);
        if (!jobs.length) {
          await reply(replyToken, `📋 ${team} — วันนี้ไม่มีงานที่ได้รับมอบหมายครับ`);
          return;
        }
        const bubbles = jobs.map(j => ({
          type: 'bubble',
          header: {
            type: 'box', layout: 'vertical',
            backgroundColor: '#1E3A5F', paddingAll: '14px',
            contents: [{ type: 'text', text: `🔧 ${team}`, color: '#C9A227', weight: 'bold', size: 'md' }]
          },
          body: {
            type: 'box', layout: 'vertical', spacing: 'sm', paddingAll: '16px',
            contents: [
              { type: 'text', text: (j['JOB'] || '—'), weight: 'bold', wrap: true, size: 'sm', color: '#1E3A5F' },
              { type: 'separator', margin: 'sm' },
              { type: 'box', layout: 'horizontal', margin: 'sm', contents: [
                { type: 'text', text: '🏭', flex: 0, size: 'sm' },
                { type: 'text', text: (j['บริษัท'] || '—'), size: 'sm', color: '#555555', wrap: true, margin: 'sm' }
              ]},
              { type: 'box', layout: 'horizontal', contents: [
                { type: 'text', text: '📍', flex: 0, size: 'sm' },
                { type: 'text', text: (j['จังหวัด'] || '—'), size: 'sm', color: '#555555', margin: 'sm' }
              ]},
              { type: 'box', layout: 'horizontal', contents: [
                { type: 'text', text: '📝', flex: 0, size: 'sm' },
                { type: 'text', text: (j['รายละเอียดงาน'] || '—'), size: 'sm', color: '#888888', wrap: true, margin: 'sm' }
              ]},
              { type: 'box', layout: 'horizontal', contents: [
                { type: 'text', text: '🚗', flex: 0, size: 'sm' },
                { type: 'text', text: (j['รถที่ใช้ออกหน้างาน'] || '—'), size: 'sm', color: '#888888', margin: 'sm' }
              ]},
              { type: 'separator', margin: 'sm' },
              { type: 'text', text: '🔖 ' + (j['สถานะ'] || '—'), size: 'xs', color: '#888888', margin: 'sm' },
            ]
          }
        }));
        await reply(replyToken, {
          type: 'flex', altText: `งานวันนี้ ${team}`,
          contents: bubbles.length === 1 ? bubbles[0] : { type: 'carousel', contents: bubbles }
        });
        return;
      }
      if (msg === 'เช็ควันลา') {
        await reply(replyToken, createLeaveCard(employee, imgUrl));
      } else {
        await reply(replyToken, {
          type: 'flex', altText: 'กรุณาเลือกหัวข้อที่ต้องการ',
          contents: {
            type: 'bubble',
            body: { type: 'box', layout: 'vertical', spacing: 'md', paddingAll: '20px',
              contents: [
                { type: 'text', text: '📋 TPE HR Connect', weight: 'bold', size: 'lg', color: '#1E3A5F' },
                { type: 'text', text: 'กรุณาเลือกหัวข้อที่ต้องการครับ', size: 'sm', color: '#888888', margin: 'sm', wrap: true },
                { type: 'separator', margin: 'md' },
                { type: 'box', layout: 'vertical', spacing: 'sm', margin: 'md', contents: [
                  { type: 'button', style: 'primary', color: '#1E3A5F', action: { type: 'message', label: '💰 ขอสลิปเงินเดือน', text: 'ขอสลิปเงินเดือน' }},
                  { type: 'button', style: 'secondary', action: { type: 'message', label: '📋 ขอใบรับรองเงินเดือน', text: 'ขอใบรับรองเงินเดือน' }},
                  { type: 'button', style: 'secondary', action: { type: 'message', label: '📅 เช็ควันลา', text: 'เช็ควันลา' }},
                  { type: 'button', style: 'primary', color: '#C9A227', action: { type: 'message', label: '🔧 งานวันนี้', text: 'งานวันนี้' }},
                ]}
              ]
            }
          }
        });
      }
      
    } else {
      // ลงทะเบียน — ตรวจสอบ ลงเทียน ซ้ำ
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


// ════════════════════════════════════════════════════════
// E-Slip: ดึงข้อมูลเวลาเข้า-ออกของพนักงาน
// ════════════════════════════════════════════════════════
// Cache สำหรับ attendance แยกตาม lineId (5 นาที)
const attCache = {};
const LEAVE_CACHE = {}; // key = lineId

app.get('/eslip/attendance', async (req, res) => {
  try {
    const { lineId } = req.query;
    if (!lineId) return res.status(400).json({ error: 'missing lineId' });

    // ดึงชื่อพนักงานจาก Lark
    const larkToken = await lark.getToken();
    const emp = await lark.findByLineId(larkToken, lineId);
    if (!emp) return res.status(404).json({ error: 'not found' });

    const rawName = emp['ชื่อ - นามสกุล'] || emp['ชื่อ-นามสกุล'] || '';
    const posMatch = rawName.match(/[（(]([^)）]+)[)）]/);
    const position = posMatch ? posMatch[1].trim() : (emp['ตำแหน่ง'] || '');
    const empName = payroll.normName(rawName.split('(')[0].split('（')[0]);
    const empId   = (emp['รหัสพนักงาน'] || '').toString().trim();

    // ดึงข้อมูล Attendance จาก Google Sheets
    const { google } = require('googleapis');
    const auth = new google.auth.GoogleAuth({
      credentials: JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT_JSON),
      scopes: ['https://www.googleapis.com/auth/spreadsheets'],
    });
    const sheets = google.sheets({ version: 'v4', auth: await auth.getClient() });
    const now2 = Date.now();
    const CACHE_TTL2 = 5 * 60 * 1000;

    // ดึง Attendance + Leave พร้อมกัน
    const [attResult, leaveData] = await Promise.all([
      sheets.spreadsheets.values.get({
        spreadsheetId: process.env.LOG_SHEET_ID,
        range: 'Attendance!A:F',
      }),
      (async () => {
        const cached = LEAVE_CACHE[lineId];
        if (cached && (now2 - cached.ts) < CACHE_TTL2) {
          return cached.data;
        }
        const d = await getLeaveDates(larkToken, empName, empId).catch(() => ({}));
        LEAVE_CACHE[lineId] = { data: d, ts: now2 };
        return d;
      })()
    ]);

    const rows = (attResult.data.values || []).slice(1);

    // กรองเฉพาะของพนักงานคนนี้ — fuzzy match
    console.log(`[Attendance] empName="${empName}" empId="${empId}" totalRows=${rows.length}`);

    // กรองด้วย ID เป็นหลัก (แม่นยำ 100%)
    let finalRows = [];
    if (empId) {
      finalRows = rows.filter(row => (row[2]||'').toString().trim() === empId);
      console.log(`[Attendance] ID match: ${finalRows.length} rows`);
    }
    // ถ้าไม่มี ID ให้ exact match ชื่อเท่านั้น
    if (finalRows.length === 0) {
      const normName = (s) => (s||'').replace(/\s+/g,'').toLowerCase();
      const empNorm = normName(empName);
      finalRows = rows.filter(row => normName(row[3]||'') === empNorm);
      console.log(`[Attendance] Exact name match: ${finalRows.length} rows`);
    }

    // จัดกลุ่มตามวัน — เก็บทุก time ในวันเดียวกัน
    const byDate = {};
    finalRows.forEach(row => {
      const date = (row[0]||'').trim();
      const time = (row[1]||'').trim();
      if (!date || !time) return;
      if (!byDate[date]) byDate[date] = [];
      byDate[date].push(time);
    });
    console.log(`[Attendance] byDate keys: ${Object.keys(byDate).length} วัน`);
    // log ตัวอย่าง 3 วันแรก
    Object.entries(byDate).slice(0,3).forEach(([d,t]) => {
      console.log(`[Attendance]   ${d}: [${t.join(', ')}]`);
    });

    const toM = t => { const[h,m,s]=(t||'00:00:00').split(':').map(Number); return h*60+m+(s||0)/60; };

    const records = Object.entries(byDate).map(([date, times]) => {
      const sorted = [...times].sort();
      const first  = sorted[0];
      const last   = sorted[sorted.length-1];
      const fm     = toM(first);
      const lm     = toM(last);

      // สถานะเข้างาน
      let status = 'ok', lateMinutes = 0;
      if (fm <= toM('08:00:59'))       status = 'ok';
      else if (fm <= toM('11:59:59')) { status = 'late'; lateMinutes = Math.round(fm - toM('08:01:00')); }
      else if (fm <= toM('12:00:59')) status = 'half-am';
      else                            { status = 'late'; lateMinutes = Math.round(fm - toM('08:01:00')); }

      // OT
      let ot = 0;
      if (sorted.length > 1) {
        if      (lm >= toM('11:50:00') && lm <= toM('12:30:00')) status = 'half-pm';
        else if (lm >= toM('17:30:00') && lm < toM('17:55:00')) ot = 0.5;
        else if (lm >= toM('17:55:00') && lm < toM('18:25:00')) ot = 1;
        else if (lm >= toM('18:25:00') && lm < toM('18:55:00')) ot = 1.5;
        else if (lm >= toM('18:55:00') && lm < toM('19:25:00')) ot = 2;
        else if (lm >= toM('19:25:00') && lm < toM('19:55:00')) ot = 2.5;
        else if (lm >= toM('19:55:00') && lm < toM('20:25:00')) ot = 3;
        else if (lm >= toM('20:25:00'))                          ot = 3.5;
      }

      return {
        date, time: first, timeOut: sorted.length > 1 ? last : null,
        late: status === 'late', lateMinutes, status, ot
      };
    }).sort((a, b) => {
      const [ad,am,ay] = a.date.split('/').map(Number);
      const [bd,bm,by] = b.date.split('/').map(Number);
      return (by*10000+bm*100+bd) - (ay*10000+am*100+ad);
    });

    // คำนวณ summary รายเดือนและทั้งหมด
    const summaryByMonth = {};
    let totalWork = 0, totalLate = 0, totalOT = 0;

    records.forEach(r => {
      const [dd, mm, yyyy] = r.date.split('/');
      const key = `${mm}/${yyyy}`;
      if (!summaryByMonth[key]) {
        summaryByMonth[key] = { month: key, work: 0, late: 0, lateMinutes: 0, ot: 0, otHours: 0 };
      }
      const s = summaryByMonth[key];
      if (r.status !== 'half-am' && r.status !== 'half-pm') {
        s.work++; totalWork++;
      }
      if (r.late) { s.late++; s.lateMinutes += r.lateMinutes || 0; totalLate++; }
      if (r.ot > 0) { s.ot++; s.otHours += r.ot; totalOT += r.ot; }
    });

    const summary = {
      total: { work: totalWork, late: totalLate, ot: totalOT },
      byMonth: Object.values(summaryByMonth).sort((a, b) => {
        const [am, ay] = a.month.split('/').map(Number);
        const [bm, by] = b.month.split('/').map(Number);
        return (by * 100 + bm) - (ay * 100 + am);
      })
    };

    // ใช้ leaveData ที่ดึงมาแล้วพร้อมกัน
    const leaveMap = leaveData;

    // เพิ่มข้อมูลลาในวันที่ไม่มีสแกน
    const allDates = new Set(records.map(r => r.date));
    const leaveRecords = Object.entries(leaveMap)
      .filter(([date]) => !allDates.has(date))
      .map(([date, leaveType]) => ({
        date, time: null, timeOut: null,
        late: false, lateMinutes: 0, status: 'leave',
        ot: 0, leaveType
      }));

    // รวม records และ leaveRecords เรียงใหม่
    const allRecords = [...records, ...leaveRecords].sort((a, b) => {
      const [ad,am,ay] = a.date.split('/').map(Number);
      const [bd,bm,by] = b.date.split('/').map(Number);
      return (by*10000+bm*100+bd) - (ay*10000+am*100+ad);
    });

    // คำนวณ summary ใหม่รวม leave
    const summaryByMonth2 = {};
    let totalWork2 = 0, totalLate2 = 0, totalOT2 = 0;
    allRecords.forEach(r => {
      const [dd, mm, yyyy] = r.date.split('/');
      const key = `${mm}/${yyyy}`;
      if (!summaryByMonth2[key]) summaryByMonth2[key] = { month: key, work: 0, late: 0, lateMinutes: 0, ot: 0, otHours: 0, leave: 0 };
      const s = summaryByMonth2[key];
      if (r.status === 'leave') { s.leave++; return; }
      if (r.status !== 'half-am' && r.status !== 'half-pm') { s.work++; totalWork2++; }
      if (r.late) { s.late++; s.lateMinutes += r.lateMinutes || 0; totalLate2++; }
      if (r.ot > 0) { s.ot++; s.otHours += r.ot; totalOT2 += r.ot; }
    });

    const summary2 = {
      total: { work: totalWork2, late: totalLate2, ot: totalOT2 },
      byMonth: Object.values(summaryByMonth2).sort((a, b) => {
        const [am, ay] = a.month.split('/').map(Number);
        const [bm, by] = b.month.split('/').map(Number);
        return (by * 100 + bm) - (ay * 100 + am);
      })
    };

    res.json({ name: empName, position, records: allRecords, summary: summary2 });
  } catch(e) {
    console.error('/eslip/attendance error:', e.message);
    res.status(500).json({ error: e.message });
  }
});


// ════════════════════════════════════════════════════════
// ATTENDANCE ADMIN ENDPOINTS
// ════════════════════════════════════════════════════════

// ดึงข้อมูล attendance ทุกคน + holidays สำหรับ admin
app.get('/admin/attendance', async (req, res) => {
  const { password, year, month } = req.query;
  if (password !== 'tpe2569') return res.status(401).json({ error: 'unauthorized' });
  try {
    const { google } = require('googleapis');
    const auth = new google.auth.GoogleAuth({
      credentials: JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT_JSON),
      scopes: ['https://www.googleapis.com/auth/spreadsheets'],
    });
    const sheets = google.sheets({ version: 'v4', auth: await auth.getClient() });
    const sid = process.env.LOG_SHEET_ID;

    // ดึง Attendance
    const attR = await sheets.spreadsheets.values.get({ spreadsheetId: sid, range: 'Attendance!A:F' });
    const attRows = (attR.data.values || []).slice(1);

    // ดึง Holidays
    let holidays = [];
    try {
      const holR = await sheets.spreadsheets.values.get({ spreadsheetId: sid, range: 'Holidays!A:B' });
      holidays = (holR.data.values || []).slice(1).map(r => ({ date: r[0], name: r[1] || 'วันหยุดบริษัท' }));
    } catch(e) {}

    // ดึงพนักงานจาก Lark
    const larkToken = await lark.getToken();
    const emps = await lark.getAllEmployees(larkToken).catch(() => []);

    // กรองตามรอบเดือน (26 เดือนก่อน - 25 เดือนนี้)
    const y = parseInt(year) || new Date().getFullYear();
    const m = parseInt(month) || new Date().getMonth() + 1;
    // รอบ: 26 เดือนก่อน - 25 เดือนนี้
    const prevM = m === 1 ? 12 : m - 1;
    const prevY = m === 1 ? y - 1 : y;
    const startDate = `26/${String(prevM).padStart(2,'0')}/${prevY}`;
    const endDate   = `25/${String(m).padStart(2,'0')}/${y}`;

    // สร้าง array วันในรอบ
    const start = new Date(prevY, prevM-1, 26);
    const end   = new Date(y, m-1, 25);
    const days = [];
    for (let d = new Date(start); d <= end; d.setDate(d.getDate()+1)) {
      const dd = String(d.getDate()).padStart(2,'0');
      const mm = String(d.getMonth()+1).padStart(2,'0');
      const yyyy = d.getFullYear();
      days.push(`${dd}/${mm}/${yyyy}`);
    }

    // สร้าง map รหัสพนักงาน → ชื่อ Lark
    const normN = s => (s||'').replace(/\s+/g,'').toLowerCase();
    const larkEmps = emps.map(e => ({
      raw: e,
      clean: (e['ชื่อ - นามสกุล']||'').split('(')[0].trim(),
      norm: normN((e['ชื่อ - นามสกุล']||'').split('(')[0]),
      id: (e['รหัสพนักงาน']||'').toString().trim()
    }));

    // สร้าง ID → ชื่อ map
    const idToName = {};
    larkEmps.forEach(e => { if(e.id) idToName[e.id] = e.clean; });

    // จัด attendance โดย map ID/ชื่อ → ชื่อ Lark
    const attMap = {};
    attRows.forEach(row => {
      const date = row[0], time = row[1];
      const attId = (row[2]||'').toString().trim();
      const rawAtt = (row[3]||'').trim();
      if (!date || !time) return;

      // หาชื่อ Lark — ใช้ ID ก่อน ถ้าไม่มีค่อย fuzzy name
      let larkName = rawAtt;
      if (attId && idToName[attId]) {
        larkName = idToName[attId];
      } else {
        const attNorm = normN(rawAtt);
        for (const le of larkEmps) {
          const ln = le.norm;
          if (ln === attNorm || ln.includes(attNorm) || attNorm.includes(ln) ||
              (ln.length >= 4 && attNorm.includes(ln.slice(0,4)))) {
            larkName = le.clean; break;
          }
        }
      }

      if (!attMap[date]) attMap[date] = {};
      if (!attMap[date][larkName]) attMap[date][larkName] = [];
      attMap[date][larkName].push(time);
    });

    res.json({ days, attMap, holidays, emps: larkEmps.map(e => ({
      name: e.clean,
      type: (e.raw['ประเภท']||'').includes('รายวัน') ? 'daily' : 'monthly'
    })), startDate, endDate, year: y, month: m });
  } catch(e) {
    console.error('/admin/attendance error:', e.message);
    res.status(500).json({ error: e.message });
  }
});

// ดึง leave map สำหรับ admin
app.get('/admin/leave-map', async (req, res) => {
  const { password, year, month } = req.query;
  if (password !== 'tpe2569') return res.status(401).json({ error: 'unauthorized' });
  try {
    const larkToken = await lark.getToken();
    const axios = require('axios');
    const normN = s => (s||'').replace(/\s+/g,'').toLowerCase();

    // ดึงพนักงานทั้งหมด
    const emps = await lark.getAllEmployees(larkToken).catch(() => []);
    const larkEmps = emps.map(e => ({
      clean: (e['ชื่อ - นามสกุล']||'').split('(')[0].trim(),
      norm: normN((e['ชื่อ - นามสกุล']||'').split('(')[0])
    }));

    // ดึง leave records
    let allRecords = [], pageToken = '';
    for (let i = 0; i < 5; i++) {
      const url = `https://open.larksuite.com/open-apis/bitable/v1/apps/T1RhbpctWafjxGsoVVtlSJaGgJf/tables/tbl0fDzMNrGBOVwu/records?page_size=100${pageToken ? '&page_token=' + pageToken : ''}`;
      const r = await axios.get(url, { headers: { Authorization: `Bearer ${larkToken}` } });
      const data = r.data?.data;
      allRecords = allRecords.concat(data?.items || []);
      if (!data?.has_more) break;
      pageToken = data.page_token || '';
    }

    // สร้าง leaveMap: { empName: { 'dd/mm/yyyy': leaveType } }
    const leaveMap = {};
    allRecords.forEach(item => {
      const f = item.fields;
      const rawName = (f['ชื่อ-นามสกุล'] || '').split('(')[0].trim();
      const rn = normN(rawName);
      const type = f['ประเภทการลา'] || 'ลา';
      const start = f['ลาตั้งเเต่วันที่'];
      const end   = f['จนถึงวันที่'];
      if (!start) return;

      // หาชื่อ Lark ที่ match
      let larkName = rawName;
      for (const le of larkEmps) {
        if (le.norm === rn || le.norm.includes(rn) || rn.includes(le.norm) ||
            (le.norm.length >= 4 && rn.includes(le.norm.slice(0,4)))) {
          larkName = le.clean; break;
        }
      }

      if (!leaveMap[larkName]) leaveMap[larkName] = {};
      const toTD = ts => {
        const d = new Date(ts + 7*60*60*1000);
        return { dd: String(d.getUTCDate()).padStart(2,'0'), mm: String(d.getUTCMonth()+1).padStart(2,'0'), yyyy: d.getUTCFullYear() };
      };
      const s = toTD(start), e2 = end ? toTD(end) : s;
      const cur2 = new Date(Date.UTC(s.yyyy, parseInt(s.mm)-1, parseInt(s.dd)));
      const fin2 = new Date(Date.UTC(e2.yyyy, parseInt(e2.mm)-1, parseInt(e2.dd)));
      while (cur2 <= fin2) {
        const dd = String(cur2.getUTCDate()).padStart(2,'0');
        const mm = String(cur2.getUTCMonth()+1).padStart(2,'0');
        const yyyy = cur2.getUTCFullYear();
        leaveMap[larkName][`${dd}/${mm}/${yyyy}`] = type;
        cur2.setUTCDate(cur2.getUTCDate()+1);
      }
    });

    res.json({ leaveMap });
  } catch(e) {
    console.error('/admin/leave-map error:', e.message);
    res.status(500).json({ error: e.message });
  }
});

// ดูประเภทการลาทั้งหมด
app.get('/admin/leave-types', async (req, res) => {
  const { password } = req.query;
  if (password !== 'tpe2569') return res.status(401).json({ error: 'unauthorized' });
  try {
    const larkToken = await lark.getToken();
    const axios = require('axios');
    let allRecords = [], pageToken = '';
    for (let i = 0; i < 10; i++) {
      const url = `https://open.larksuite.com/open-apis/bitable/v1/apps/T1RhbpctWafjxGsoVVtlSJaGgJf/tables/tbl0fDzMNrGBOVwu/records?page_size=100${pageToken ? '&page_token=' + pageToken : ''}`;
      const r = await axios.get(url, { headers: { Authorization: `Bearer ${larkToken}` } });
      const data = r.data?.data;
      allRecords = allRecords.concat(data?.items || []);
      if (!data?.has_more) break;
      pageToken = data.page_token || '';
    }
    const types = [...new Set(allRecords.map(r => r.fields['ประเภทการลา']).filter(Boolean))];
    res.json({ types, total: allRecords.length });
  } catch(e) { res.status(500).json({ error: e.message }); }
});

// เพิ่ม/ลบวันหยุด
app.post('/admin/holidays', express.json(), async (req, res) => {
  const { password, action, date, name } = req.body;
  if (password !== 'tpe2569') return res.status(401).json({ error: 'unauthorized' });
  try {
    const { google } = require('googleapis');
    const auth = new google.auth.GoogleAuth({
      credentials: JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT_JSON),
      scopes: ['https://www.googleapis.com/auth/spreadsheets'],
    });
    const sheets = google.sheets({ version: 'v4', auth: await auth.getClient() });
    const sid = process.env.LOG_SHEET_ID;

    // ตรวจ/สร้าง Holidays sheet
    const meta = await sheets.spreadsheets.get({ spreadsheetId: sid });
    const exists = meta.data.sheets.some(s => s.properties.title === 'Holidays');
    if (!exists) {
      await sheets.spreadsheets.batchUpdate({ spreadsheetId: sid,
        requestBody: { requests: [{ addSheet: { properties: { title: 'Holidays' } } }] }
      });
      await sheets.spreadsheets.values.append({ spreadsheetId: sid, range: 'Holidays!A1',
        valueInputOption: 'RAW', requestBody: { values: [['วันที่', 'ชื่อวันหยุด']] }
      });
    }

    if (action === 'add') {
      await sheets.spreadsheets.values.append({ spreadsheetId: sid, range: 'Holidays!A:B',
        valueInputOption: 'RAW', requestBody: { values: [[date, name || 'วันหยุดบริษัท']] }
      });
    } else if (action === 'delete') {
      // ดึงทั้งหมดแล้วลบแถวที่ตรง
      const r = await sheets.spreadsheets.values.get({ spreadsheetId: sid, range: 'Holidays!A:B' });
      const rows = r.data.values || [];
      const sheetId = meta.data.sheets.find(s => s.properties.title === 'Holidays').properties.sheetId;
      const delIdx = rows.findIndex((row, i) => i > 0 && row[0] === date);
      if (delIdx > 0) {
        await sheets.spreadsheets.batchUpdate({ spreadsheetId: sid,
          requestBody: { requests: [{ deleteDimension: { range: {
            sheetId, dimension: 'ROWS', startIndex: delIdx, endIndex: delIdx + 1
          }}}]}
        });
      }
    }
    res.json({ ok: true });
  } catch(e) { res.status(500).json({ error: e.message }); }
});

// ════════════════════════════════════════════════════════
// Telegram Attendance Webhook
// ════════════════════════════════════════════════════════
app.post('/telegram-webhook', async (req, res) => {
  res.sendStatus(200); // ตอบ Telegram ก่อนเสมอ
  try {
    const { google } = require('googleapis');
    const auth = new google.auth.GoogleAuth({
      credentials: JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT_JSON),
      scopes: ['https://www.googleapis.com/auth/spreadsheets'],
    });
    const sheets = google.sheets({ version: 'v4', auth: await auth.getClient() });
    const spreadsheetId = process.env.LOG_SHEET_ID;
    await telegramBot.handleTelegramUpdate(req.body, sheets, spreadsheetId);
  } catch(e) {
    console.error('telegram webhook error:', e.message);
  }
});

// ── ตั้ง Telegram Webhook ──────────────────────────────
app.get('/telegram-setup', async (req, res) => {
  try {
    const token = process.env.TELEGRAM_TOKEN;
    // ลบ Webhook ออกก่อนเสมอ เพื่อให้ polling ทำงานได้
    await require('axios').post(
      `https://api.telegram.org/bot${token}/deleteWebhook`,
      { drop_pending_updates: true }
    );
    res.json({ ok: true, message: 'Webhook removed — polling is active' });
  } catch(e) {
    res.json({ ok: false, error: e.message });
  }
});

// ── ดูประวัติ Attendance ──────────────────────────────
app.get('/attendance', async (req, res) => {
  try {
    const { google } = require('googleapis');
    const auth = new google.auth.GoogleAuth({
      credentials: JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT_JSON),
      scopes: ['https://www.googleapis.com/auth/spreadsheets'],
    });
    const sheets = google.sheets({ version: 'v4', auth: await auth.getClient() });
    const r = await sheets.spreadsheets.values.get({
      spreadsheetId: process.env.LOG_SHEET_ID,
      range: 'Attendance!A:F',
    });
    const rows = r.data.values || [];
    res.json({ total: rows.length - 1, rows: rows.slice(1).slice(-100) });
  } catch(e) {
    res.status(500).json({ error: e.message });
  }
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
  const srv = app.listen(port, '0.0.0.0', () => {
    console.log(`TPE HR Bot v2 on port ${port}`);
    // เริ่ม background services หลัง server listen แล้ว
    setTimeout(() => initBackgroundServices(), 2000);
  });
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
app.get('/attendance-admin', (req, res) => res.sendFile(path.join(__dirname, 'attendance-admin.html')));

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
    // ดึงตำแหน่งจากชื่อ เช่น "อรุณ ช่วยจวน (Graphic Designer & DBA)"
    const nameRaw = rawName.trim();
    const posMatch = nameRaw.match(/[（(]([^)）]+)[)）]/);
    const position = posMatch ? posMatch[1].trim() : (emp['ตำแหน่ง'] || '');
    const cleanName = payroll.normName(nameRaw.split('(')[0].split('（')[0]);

    res.json({
      name: cleanName,
      position: position,
      payType: rawType.includes('รายวัน') ? 'daily' : 'monthly',
      startDate: emp['วันที่เริ่มงาน'] || emp['เริ่มงาน'] || '',
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

// HTML ของเอกสาร (สำหรับแสดงใน LIFF) — ไม่ใช้ puppeteer
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
    // สร้าง HTML โดยตรง ไม่ต้อง launch puppeteer
    let html;
    if (docType === 'payslip') {
      html = payslip.buildPayslipHtml(empData);
    } else {
      html = cert.buildCertHtml(empData);
    }
    if (!html) return res.status(500).json({ error: 'buildHtml returned empty' });
    res.json({ html });
  } catch(e) {
    console.error('/eslip/doc error:', e.message, e.stack?.split('\n')[1]);
    res.status(500).json({ error: e.message });
  }
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
    // สร้าง HTML โดยตรงไม่ต้อง puppeteer 2 รอบ
    let html;
    if (docType === 'payslip') {
      html = payslip.buildPayslipHtml(empData);
    } else {
      html = cert.buildCertHtml(empData);
    }
    if (!html) return res.status(500).send('build html failed');
    const imgBuf = await payslip.htmlToImageBuffer(html);
    res.setHeader('Content-Type', 'image/jpeg');
    res.setHeader('Content-Disposition', 'attachment; filename="slip.jpg"');
    res.send(imgBuf);
  } catch(e) { console.error('[eslip/image] ERROR:', e.message); res.status(500).send(e.message); }
});

// temp image store สำหรับ Android LIFF download
const tempImages = {};
app.post('/eslip/temp-image', upload.single('file'), (req, res) => {
  if (!req.file) return res.status(400).json({ error: 'no file' });
  const token = require('crypto').randomBytes(8).toString('hex');
  const buf = require('fs').readFileSync(req.file.path);
  require('fs').unlinkSync(req.file.path);
  tempImages[token] = { buffer: buf, filename: req.file.originalname, createdAt: Date.now() };
  // ลบหลัง 10 นาที
  setTimeout(() => delete tempImages[token], 10 * 60 * 1000);
  const RENDER_URL = process.env.RENDER_URL || 'https://tpe-hr-bot.onrender.com';
  res.json({ url: `${RENDER_URL}/eslip/temp-image/${token}` });
});
app.get('/eslip/temp-image/:token', (req, res) => {
  const entry = tempImages[req.params.token];
  if (!entry) return res.status(404).send('หมดอายุ');
  // ถ้ามี ?download=1 ให้ download เลย
  if (req.query.download === '1') {
    res.setHeader('Content-Type', 'image/jpeg');
    res.setHeader('Content-Disposition', `attachment; filename="${encodeURIComponent(entry.filename)}"`);
    res.send(entry.buffer);
    return;
  }
  // ไม่งั้นแสดงหน้า HTML พร้อมปุ่มบันทึก
  const RENDER_URL = process.env.RENDER_URL || 'https://tpe-hr-bot.onrender.com';
  const imgUrl = `${RENDER_URL}/eslip/temp-image-raw/${req.params.token}`;
  const dlUrl  = `${RENDER_URL}/eslip/temp-image/${req.params.token}?download=1`;
  res.send(`<!DOCTYPE html><html><head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width,initial-scale=1,maximum-scale=1">
    <title>สลิปเงินเดือน</title>
    <style>
      *{margin:0;padding:0;box-sizing:border-box}
      body{background:#111;font-family:sans-serif;min-height:100vh}
      .bar{background:linear-gradient(135deg,#1E3A5F,#2a4f80);padding:16px;text-align:center}
      .btn{display:block;background:#C9952A;color:#fff;font-size:16px;font-weight:700;
           padding:14px;text-align:center;border-radius:12px;margin:12px 16px;text-decoration:none}
      img{width:100%;display:block;margin-top:8px}
    </style></head><body>
    <div class="bar" style="color:#fff;font-size:14px">📄 สลิปเงินเดือน</div>
    <a class="btn" href="${dlUrl}" download>💾 บันทึกรูปลงเครื่อง</a>
    <img src="${imgUrl}" alt="สลิป">
  </body></html>`);
});

// override ด้วย raw endpoint
app.get('/eslip/temp-image-raw/:token', (req, res) => {
  const entry = tempImages[req.params.token];
  if (!entry) return res.status(404).send('หมดอายุ');
  res.setHeader('Content-Type', 'image/jpeg');
  res.setHeader('Content-Disposition', `inline; filename="${encodeURIComponent(entry.filename)}"`);
  res.send(entry.buffer);
});

// ── ดึงข้อมูลการลาจาก Lark Base ──────────────────────
async function getLeaveDates(larkToken, empName, empId) {
  try {
    const axios = require('axios');
    const normN = s => (s||'').replace(/\s+/g,'').toLowerCase();
    const empNorm = normN(empName);

    let allRecords = [], pageToken = '';
    // ดึงทั้งหมด (max 500)
    for (let i = 0; i < 5; i++) {
      const url = `https://open.larksuite.com/open-apis/bitable/v1/apps/T1RhbpctWafjxGsoVVtlSJaGgJf/tables/tbl0fDzMNrGBOVwu/records?page_size=100${pageToken ? '&page_token=' + pageToken : ''}`;
      const r = await axios.get(url, { headers: { Authorization: `Bearer ${larkToken}` } });
      const data = r.data?.data;
      allRecords = allRecords.concat(data?.items || []);
      if (!data?.has_more) break;
      pageToken = data.page_token || '';
    }

    // กรองเฉพาะพนักงานคนนี้
    const myLeaves = allRecords.filter(item => {
      const f = item.fields;
      const rawName = (f['ชื่อ-นามสกุล'] || '').split('(')[0].trim();
      const n = normN(rawName);
      return n === empNorm || n.includes(empNorm) || empNorm.includes(n) ||
             (n.length >= 4 && empNorm.includes(n.slice(0, 4)));
    });

    // แปลง timestamp → วันที่ dd/mm/yyyy
    const tsToDate = ts => {
      const d = new Date(ts);
      const dd = String(d.getDate()).padStart(2,'0');
      const mm = String(d.getMonth()+1).padStart(2,'0');
      const yyyy = d.getFullYear();
      return `${dd}/${mm}/${yyyy}`;
    };

    // สร้าง map วันที่ → ประเภทการลา
    const leaveMap = {};
    myLeaves.forEach(item => {
      const f = item.fields;
      const start = f['ลาตั้งเเต่วันที่'];
      const end   = f['จนถึงวันที่'];
      const type  = f['ประเภทการลา'] || 'ลา';
      if (!start) return;
      // แปลง timestamp → วันที่ไทย (UTC+7) แล้ววนทุกวันในช่วงลา
      const toThaiDate = ts => {
        const d = new Date(ts + 7 * 60 * 60 * 1000); // +7 ชั่วโมง
        const dd   = String(d.getUTCDate()).padStart(2,'0');
        const mm   = String(d.getUTCMonth()+1).padStart(2,'0');
        const yyyy = d.getUTCFullYear();
        return { dd, mm, yyyy, str: `${dd}/${mm}/${yyyy}` };
      };
      const startTh = toThaiDate(start);
      const endTh   = end ? toThaiDate(end) : startTh;
      // วนทุกวัน
      const cur = new Date(Date.UTC(parseInt(startTh.yyyy), parseInt(startTh.mm)-1, parseInt(startTh.dd)));
      const fin = new Date(Date.UTC(parseInt(endTh.yyyy),   parseInt(endTh.mm)-1,   parseInt(endTh.dd)));
      while (cur <= fin) {
        const dd   = String(cur.getUTCDate()).padStart(2,'0');
        const mm   = String(cur.getUTCMonth()+1).padStart(2,'0');
        const yyyy = cur.getUTCFullYear();
        leaveMap[`${dd}/${mm}/${yyyy}`] = type;
        cur.setUTCDate(cur.getUTCDate() + 1);
      }
    });
    return leaveMap;
  } catch(e) {
    console.error('getLeaveDates error:', e.message);
    return {};
  }
}

// Debug: ดู Leave table fields
app.get('/eslip/debug-leave', async (req, res) => {
  try {
    const larkToken = await lark.getToken();
    const r = await require('axios').get(
      'https://open.larksuite.com/open-apis/bitable/v1/apps/T1RhbpctWafjxGsoVVtlSJaGgJf/tables/tbl0fDzMNrGBOVwu/records?page_size=5',
      { headers: { 'Authorization': `Bearer ${larkToken}` } }
    );
    const items = r.data?.data?.items || [];
    if (!items.length) return res.json({ error: 'no records', raw: r.data });
    res.json({
      keys: Object.keys(items[0].fields),
      sample: items[0].fields,
      total: r.data?.data?.total
    });
  } catch(e) { res.status(500).json({ error: e.message }); }
});

// Debug: ดู raw Lark fields
app.get('/eslip/debug-emp', async (req, res) => {
  try {
    const { lineId } = req.query;
    const larkToken = await lark.getToken();
    const emp = await lark.findByLineId(larkToken, lineId);
    if (!emp) return res.json({ error: 'not found' });
    // ส่ง keys ทั้งหมดกลับมา
    res.json({ keys: Object.keys(emp), values: emp });
  } catch(e) { res.json({ error: e.message }); }
});

// E-Slip error log endpoint
app.post('/eslip/log', express.json(), (req, res) => {
  const { step, message, ua, time, lineId } = req.body;
  const device = /iPhone|iPad|iPod/.test(ua) ? 'iOS' : /Android/.test(ua) ? 'Android' : 'Desktop';
  console.log(`[ESLIP-LOG] ${device} | ${step} | ${message} | lineId:${lineId} | ${time}`);
  res.json({ ok: true });
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


async function initBackgroundServices() {
  try {
    const { google } = require('googleapis');
    const auth = new google.auth.GoogleAuth({
      credentials: JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT_JSON),
      scopes: ['https://www.googleapis.com/auth/spreadsheets'],
    });
    const sheetsClient = google.sheets({ version: 'v4', auth: await auth.getClient() });
    try { telegramBot.startPolling(sheetsClient, process.env.LOG_SHEET_ID); }
    catch(e) { console.error('Telegram polling error:', e.message); }
    try { gramJS.startGramJS(sheetsClient, process.env.LOG_SHEET_ID); }
    catch(e) { console.error('GramJS error:', e.message); }
  } catch(e) {
    console.error('initBackgroundServices error:', e.message);
  }
}

// ── Job Queue: ดึงงานวันนี้ของช่าง ──────────────────────
app.get('/eslip/jobs-today', async (req, res) => {
  try {
    const { lineId } = req.query;
    if (!lineId) return res.status(400).json({ error: 'missing lineId' });
    const larkToken = await lark.getToken();
    const emp = await lark.findByLineId(larkToken, lineId);
    if (!emp) return res.status(404).json({ error: 'not found' });
    const team = (emp['ชุด'] || '').toString().trim();
    if (!team) return res.json({ team: '', jobs: [] });
    const jobs = await lark.getJobsToday(larkToken, team);
    res.json({ team, jobs });
  } catch(e) {
    console.error('/eslip/jobs-today error:', e.message);
    res.status(500).json({ error: e.message });
  }
});

// ── Dashboard API ──────────────────────────────────────
app.get('/api/dashboard', async (req, res) => {
  try {
    const larkToken = await lark.getToken();

    // ดึง Assignments
    let assignments = [], pageToken = '';
    for (let i = 0; i < 10; i++) {
      const url = `https://open.larksuite.com/open-apis/bitable/v1/apps/${process.env.LARK_JOB_BASE_ID}/tables/${process.env.LARK_ASSIGN_TABLE_ID}/records?page_size=100${pageToken ? '&page_token=' + pageToken : ''}`;
      const r = await axios.get(url, { headers: { Authorization: `Bearer ${larkToken}` } });
      const data = r.data?.data;
      assignments = assignments.concat(data?.items || []);
      if (!data?.has_more) break;
      pageToken = data.page_token || '';
    }

    // ดึง JOB2026
    let jobs = [], jobPageToken = '';
    for (let i = 0; i < 10; i++) {
      const url = `https://open.larksuite.com/open-apis/bitable/v1/apps/${process.env.LARK_JOB_BASE_ID}/tables/tblyHWAlWKVwKOz9/records?page_size=100${jobPageToken ? '&page_token=' + jobPageToken : ''}`;
      const r = await axios.get(url, { headers: { Authorization: `Bearer ${larkToken}` } });
      const data = r.data?.data;
      jobs = jobs.concat(data?.items || []);
      if (!data?.has_more) break;
      jobPageToken = data.page_token || '';
    }

    // สร้าง job map
    const jobMap = {};
    jobs.forEach(j => { jobMap[j.record_id] = j.fields; });

    // แปลง assignments
    const result = assignments.map(a => {
      const f = a.fields;
      const jobId = Array.isArray(f['JOB']) ? f['JOB'][0]?.record_id : null;
      const jobFields = jobId ? jobMap[jobId] : null;
      return {
        id: a.record_id,
        team: (f['ชุด'] || '').toString(),
        jobNo: jobFields?.['JOB'] || f['JOB'] || '—',
        jobName: jobFields?.['งาน'] || f['รายละเอียดงาน'] || '—',
        company: f['บริษัท'] || jobFields?.['บริษัท'] || '—',
        province: f['จังหวัด'] || '—',
        startDate: f['วันที่เริ่ม'] || null,
        endDate: f['วันสิ้นสุด'] || null,
        status: f['สถานะ'] || 'วางแผน',
        car: f['รถที่ใช้ออกหน้างาน'] || '—',
        detail: f['รายละเอียดงาน'] || '—',
        note: f['หมายเหตุ'] || '',
      };
    });

    res.json({ ok: true, assignments: result, total: result.length });
  } catch(e) {
    console.error('/api/dashboard error:', e.message);
    res.status(500).json({ error: e.message });
  }
});

app.get('/eslip/debug-jobs', async (req, res) => {
  try {
    const larkToken = await lark.getToken();
    const url = `https://open.larksuite.com/open-apis/bitable/v1/apps/${process.env.LARK_JOB_BASE_ID}/tables/${process.env.LARK_ASSIGN_TABLE_ID}/records?page_size=3`;
    const r = await axios.get(url, { headers: { Authorization: `Bearer ${larkToken}` } });
    res.json(r.data);
  } catch(e) {
    res.json({ error: e.message, status: e.response?.status, data: e.response?.data });
  }
});

// ── แจ้งเตือน Assignment ใหม่ ──────────────────────────
app.post('/notify-assignment', async (req, res) => {
  res.json({ ok: true });
  try {
    const { team, jobNo, company, province, detail, startDate, endDate, car } = req.body;
    const larkToken = await lark.getToken();
    const emps = await lark.getAllEmployees(larkToken);

    // แจ้งเสมอทุก Assignment
    const ALWAYS_NOTIFY = [
      'เจ้าหน้าที่กราฟฟิคและฐานข้อมูล',
    ];

    // แจ้งเฉพาะเมื่อ Assignment ใช้ชุดนั้น
    const TEAM_ROLES = {
      'ฝ่ายติดตั้ง':  ['ผู้จัดการฝ่ายติดตั้ง'],
      'ฝ่ายบอยเลอร์': ['หัวหน้าบอยเลอร์'],
      'ฝ่ายผลิต A':   ['หัวหน้าผลิต A'],
      'ฝ่ายผลิต B':   ['หัวหน้าผลิต B'],
    };

    const rolesForThisTeam = [
      ...ALWAYS_NOTIFY,
      ...(TEAM_ROLES[team] || []),
    ];

    const targets = emps.filter(e => {
      const pos = (e['ตำแหน่ง'] || '').toString().trim();
      const lid = (e['Line ID'] || e['LineID'] || '').toString().trim();
      if (!lid) return false;
      return rolesForThisTeam.some(r => pos === r);
    });

    const now = new Date().toLocaleString('th-TH', { timeZone: 'Asia/Bangkok', hour: '2-digit', minute: '2-digit', day: '2-digit', month: '2-digit' });

    const msg = {
      type: 'flex',
      altText: `📋 งานใหม่ ${jobNo} → ${team}`,
      contents: {
        type: 'bubble',
        header: {
          type: 'box', layout: 'vertical',
          backgroundColor: '#1E3A5F', paddingAll: '16px',
          contents: [
            { type: 'text', text: '📋 มอบหมายงานใหม่', color: '#ffffff', weight: 'bold', size: 'md' },
            { type: 'text', text: team || '—', color: '#C9A227', size: 'sm', margin: 'xs', weight: 'bold' },
          ]
        },
        body: {
          type: 'box', layout: 'vertical', spacing: 'sm', paddingAll: '16px',
          contents: [
            { type: 'box', layout: 'horizontal', contents: [
              { type: 'text', text: 'JOB', size: 'sm', color: '#888888', flex: 3 },
              { type: 'text', text: jobNo || '—', size: 'sm', weight: 'bold', flex: 5, color: '#1E3A5F' }
            ]},
            { type: 'box', layout: 'horizontal', contents: [
              { type: 'text', text: 'บริษัท', size: 'sm', color: '#888888', flex: 3 },
              { type: 'text', text: company || '—', size: 'sm', flex: 5, wrap: true }
            ]},
            { type: 'box', layout: 'horizontal', contents: [
              { type: 'text', text: 'จังหวัด', size: 'sm', color: '#888888', flex: 3 },
              { type: 'text', text: province || '—', size: 'sm', flex: 5 }
            ]},
            { type: 'box', layout: 'horizontal', contents: [
              { type: 'text', text: 'รายละเอียด', size: 'sm', color: '#888888', flex: 3 },
              { type: 'text', text: detail || '—', size: 'sm', flex: 5, wrap: true }
            ]},
            { type: 'box', layout: 'horizontal', contents: [
              { type: 'text', text: 'ช่วงงาน', size: 'sm', color: '#888888', flex: 3 },
              { type: 'text', text: `${startDate || '—'} – ${endDate || '—'}`, size: 'sm', flex: 5 }
            ]},
            { type: 'box', layout: 'horizontal', contents: [
              { type: 'text', text: 'รถ', size: 'sm', color: '#888888', flex: 3 },
              { type: 'text', text: car || '—', size: 'sm', flex: 5 }
            ]},
          ]
        },
        footer: {
          type: 'box', layout: 'vertical', paddingAll: '10px',
          contents: [{
            type: 'text',
            text: `TPE Job Queue · ${now}`,
            size: 'xxs', color: '#AAAAAA', align: 'center'
          }]
        }
      }
    };

    for (const emp of targets) {
      const lid = (emp['Line ID'] || emp['LineID'] || '').toString().trim();
      if (lid) await push(lid, msg).catch(e => console.error('push error:', lid, e.message));
    }

    console.log(`[notify] ${jobNo} → ${team} → ส่ง ${targets.length} คน`);
  } catch(e) {
    console.error('/notify-assignment error:', e.message);
  }
});

startServer(PORT);

// เริ่ม GramJS แยก async block
(async () => {
  try {
    const { google } = require('googleapis');
    const auth2 = new google.auth.GoogleAuth({
      credentials: JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT_JSON),
      scopes: ['https://www.googleapis.com/auth/spreadsheets'],
    });
    const sheetsClient2 = google.sheets({ version: 'v4', auth: await auth2.getClient() });
    gramJS.startGramJS(sheetsClient2, process.env.LOG_SHEET_ID);
  } catch(e) {
    console.error('GramJS init error:', e.message);
  }
})();
