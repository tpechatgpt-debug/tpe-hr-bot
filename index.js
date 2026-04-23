require('dotenv').config();
const path     = require('path');
const express  = require('express');
const axios    = require('axios');
const multer   = require('multer');
const fs       = require('fs');
const lark     = require('./lark');
const payroll  = require('./payroll');
const payslip  = require('./payslip');
const cert     = require('./certificate');
const sheet    = require('./sheet');

const app    = express();
const upload = multer({ dest: '/tmp/uploads/' });
app.use(express.json());

const LINE_TOKEN = process.env.LINE_ACCESS_TOKEN;
const HR_USER_ID = process.env.HR_LINE_USER_ID;

const pending    = {};
const requestLog = {};

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
    const msg        = event.message.text.trim();

    if (msg === 'id') { await push(userId, 'User ID: ' + userId); return; }

    const larkToken = await lark.getToken();

    if (msg === 'ขอสลิปเงินเดือน' || msg === 'ขอสลีปเงินเดือน') {
      await handleDocRequest(userId, larkToken, 'payslip'); return;
    }
    if (msg === 'ขอใบรับรองเงินเดือน') {
      await handleDocRequest(userId, larkToken, 'certificate'); return;
    }

    if (msg.includes('เลือกเดือน:')) {
      const clean       = msg.trim();
      const firstColon  = clean.indexOf(':');
      const secondColon = clean.indexOf(':', firstColon + 1);
      const month = clean.substring(firstColon + 1, secondColon).trim();
      const reqId = clean.substring(secondColon + 1).trim();
      await handleMonthSelected(userId, month, reqId); return;
    }

    const profile  = await getProfile(userId);
    const imgUrl   = profile?.pictureUrl || 'https://cdn-icons-png.flaticon.com/512/3135/3135715.png';
    const employee = await lark.findByLineId(larkToken, userId);

    if (employee) {
      await push(userId, createLeaveCard(employee, imgUrl));
    } else {
      if (msg.length < 2) { await push(userId, '⚠️ กรุณาพิมพ์ชื่อจริง เพื่อลงทะเบียนครับ'); return; }
      await push(userId, await lark.register(larkToken, msg, userId));
    }

  } catch (err) { console.error('webhook error:', err.message); }
});

// ════════════════════════════════════════════════════════
// พนักงานขอเอกสาร → ถามเดือน
// ════════════════════════════════════════════════════════
async function handleDocRequest(userId, larkToken, docType) {
  const emp = await lark.findByLineId(larkToken, userId);
  if (!emp) { await push(userId, '❌ ไม่พบข้อมูลของคุณ กรุณาลงทะเบียนก่อนนะครับ'); return; }

  const rawName   = emp['ชื่อ - นามสกุล'] || emp['ชื่อ-นามสกุล'] || 'พนักงาน';
  const empName   = payroll.normName(rawName.split('(')[0].split('（')[0]);
  const requestId = (docType === 'payslip' ? 'PAY' : 'CERT') + '_' + userId + '_' + Date.now();

  // ── อ่าน field 'ประเภท' จาก Lark ─────────────────────────
  const rawType = (emp['ประเภท'] || 'รายเดือน').toString().trim();
  const payType = rawType.includes('รายวัน') ? 'daily' : 'monthly';
  console.log(`${empName} → ประเภท: "${rawType}" → payType: ${payType}`);

  pending[requestId]    = { empName, empLineId: userId, docType, larkToken, payType };
  requestLog[requestId] = { empName, empLineId: userId, docType, status: 'pending', time: Date.now() };

  // ── ดึงเดือน กรองเฉพาะประเภทของพนักงานคนนี้ ──────────────
  const allMonths   = await payroll.getAvailableMonths();
  const validMonths = allMonths
    .filter(m => m.payType === payType)
    .map(m => m.month)
    .filter(m => m && m.trim().length > 0)
    .slice(0, 13);

  const docLabel  = docType === 'payslip' ? 'สลิปเงินเดือน' : 'ใบรับรองเงินเดือน';
  const typeLabel = rawType === 'รายวัน' ? 'รายวัน' : 'รายเดือน';

  if (validMonths.length === 0) {
    await push(userId, `⚠️ ยังไม่มีข้อมูล${typeLabel}ในระบบ กรุณาให้ HR อัปโหลดไฟล์ Excel ก่อนนะครับ`);
    return;
  }

  await push(userId, {
    type: 'text',
    text: `📅 ต้องการ${docLabel} (${typeLabel}) เดือนไหนครับ?`,
    quickReply: {
      items: validMonths.map(m => ({
        type: 'action',
        action: { type: 'message', label: m.length > 20 ? m.slice(0, 20) : m, text: 'เลือกเดือน:' + m + ':' + requestId }
      }))
    }
  });
}

// ════════════════════════════════════════════════════════
// พนักงานเลือกเดือน → แจ้ง HR
// ════════════════════════════════════════════════════════
async function handleMonthSelected(userId, month, requestId) {
  const req = pending[requestId];
  if (!req) { await push(userId, '⚠️ คำขอหมดอายุ กรุณาขอใหม่อีกครั้งครับ'); return; }

  req.month = month;
  if (requestLog[requestId]) requestLog[requestId].month = month;
  const docLabel = req.docType === 'payslip' ? 'สลิปเงินเดือน' : 'ใบรับรองเงินเดือน';

  await push(userId, {
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
  await push(HR_USER_ID, {
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
}

// ════════════════════════════════════════════════════════
// Postback: HR กดอนุมัติ/ปฏิเสธ
// ════════════════════════════════════════════════════════
async function handlePostback(event) {
  const hrUserId = event.source.userId;
  const data     = event.postback.data;
  const action   = data.startsWith('A|') ? 'approve' : data.startsWith('R|') ? 'reject' : null;
  const rid      = action ? data.slice(2) : null;

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
      const empData = await payroll.getEmployeePayroll(req.empName, req.month, req.payType || 'monthly');
      if (!empData) {
        await push(hrUserId, `❌ ไม่พบข้อมูลเงินเดือนของ ${req.empName} เดือน ${req.month}`);
        return;
      }

      const docLabel = req.docType === 'payslip' ? 'สลิปเงินเดือน' : 'ใบรับรองเงินเดือน';
      const safeName = req.empName.replace(/[^ก-๙a-zA-Z0-9 ]/g, '').trim();
      const filename  = docLabel + '_' + safeName + '_' + req.month + '.pdf';

      const pdfBuffer = req.docType === 'payslip'
        ? await payslip.createFromPayroll(empData)
        : await cert.createFromPayroll(empData);

      await push(req.empLineId, '📄 ' + docLabel + ' เดือน ' + req.month + ' พร้อมแล้วครับ กำลังส่งไฟล์...');
      await payslip.sendPdfToLine(req.empLineId, pdfBuffer, filename);

      await push(hrUserId, `✅ ส่ง PDF ${docLabel} ให้ ${req.empName} แล้วครับ`);
      delete pending[rid];
      if (requestLog[rid]) requestLog[rid].status = 'sent';
      await sheet.log({ ...empData, docType: req.docType });

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

async function push(userId, msg) {
  const messages = typeof msg === 'string' ? [{ type: 'text', text: msg }] : [msg];
  await axios.post('https://api.line.me/v2/bot/message/push', { to: userId, messages }, { headers: { Authorization: `Bearer ${LINE_TOKEN}` } });
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
  res.json(Object.entries(requestLog).sort((a,b)=>b[1].time-a[1].time).slice(0,parseInt(req.query.limit)||50).map(([,r])=>({ name:r.empName, docType:r.docType, month:r.month||'—', status:r.status||'pending', time:new Date(r.time).toLocaleString('th-TH',{timeZone:'Asia/Bangkok',hour:'2-digit',minute:'2-digit',day:'2-digit',month:'2-digit'}) })));
});
app.get('/portal/months', async (req, res) => {
  try {
    const all = await payroll.getAvailableMonths();
    res.json(all.map(m => `${m.month}${m.payType === 'daily' ? ' (รายวัน)' : ''}`));
  } catch(e) { res.json([]); }
});

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
startServer(PORT);
