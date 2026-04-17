require('dotenv').config();
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
const FORM_URL   = process.env.FORM_URL;

// pending requests: { requestId -> { empName, empLineId, docType, month, token } }
const pending = {};
const requestLog = {};  // เก็บประวัติคำขอทั้งหมด


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

    // postback: HR กดอนุมัติ/ปฏิเสธ
    if (event.type === 'postback') {
      await handlePostback(event);
      return;
    }

    if (event.type !== 'message' || event.message.type !== 'text') return;

    const userId     = event.source.userId;
    const replyToken = event.replyToken;
    const msg        = event.message.text.trim();

    // ── utility: ดู User ID ───────────────────────────
    if (msg === 'id') {
      await push(userId, 'User ID: ' + userId);
      return;
    }

    const larkToken = await lark.getToken();

    // ── ขอสลิปเงินเดือน ──────────────────────────────
    if (msg === 'ขอสลิปเงินเดือน' || msg === 'ขอสลีปเงินเดือน') {
      await handleDocRequest(replyToken, userId, larkToken, 'payslip');
      return;
    }

    // ── ขอใบรับรองเงินเดือน ──────────────────────────
    if (msg === 'ขอใบรับรองเงินเดือน') {
      await handleDocRequest(replyToken, userId, larkToken, 'certificate');
      return;
    }

    // ── พนักงานเลือกเดือน (Quick Reply response) ─────
    // format: "เลือกเดือน:เมษายน 2569:PAY_Uxxxx_123"
    if (msg.startsWith('เลือกเดือน:')) {
      const parts   = msg.split(':');
      const month   = parts[1];
      const reqId   = parts[2];
      await handleMonthSelected(userId, month, reqId);
      return;
    }

    // ── ระบบลาเดิม ───────────────────────────────────
    const profile  = await getProfile(userId);
    const imgUrl   = profile?.pictureUrl || 'https://cdn-icons-png.flaticon.com/512/3135/3135715.png';
    const employee = await lark.findByLineId(larkToken, userId);

    if (employee) {
      await reply(replyToken, createLeaveCard(employee, imgUrl));
    } else {
      if (msg.length < 2) {
        await reply(replyToken, '⚠️ กรุณาพิมพ์ชื่อจริง เพื่อลงทะเบียนครับ');
        return;
      }
      const result = await lark.register(larkToken, msg, userId);
      await reply(replyToken, result);
    }

  } catch (err) {
    console.error('webhook error:', err.message);
  }
});

// ════════════════════════════════════════════════════════
// พนักงานขอเอกสาร → ถามเดือน
// ════════════════════════════════════════════════════════
async function handleDocRequest(replyToken, userId, larkToken, docType) {
  const emp = await lark.findByLineId(larkToken, userId);
  if (!emp) {
    await reply(replyToken, '❌ ไม่พบข้อมูลของคุณ กรุณาลงทะเบียนก่อนนะครับ');
    return;
  }
  const empName = payroll.normName(emp['ชื่อ - นามสกุล'] || 'พนักงาน');
  const requestId = (docType === 'payslip' ? 'PAY' : 'CERT') + '_' + userId + '_' + Date.now();

  // บันทึก pending request
  pending[requestId] = { empName, empLineId: userId, docType, larkToken };
  requestLog[requestId] = { empName, empLineId: userId, docType, status: 'pending', time: Date.now() };

  // ดึงเดือนที่มีข้อมูล
  const months = await payroll.getAvailableMonths();

  if (months.length === 0) {
    await reply(replyToken, '⚠️ ยังไม่มีข้อมูลเงินเดือนในระบบ กรุณาติดต่อ HR ครับ');
    return;
  }

  // ส่ง Quick Reply ให้เลือกเดือน
  const quickItems = months.slice(0, 13).map(m => ({
    type: 'action',
    action: {
      type: 'message',
      label: m,
      text: `เลือกเดือน:${m}:${requestId}`,
    }
  }));

  await reply(replyToken, {
    type: 'text',
    text: `📅 ต้องการ${docType === 'payslip' ? 'สลิปเงินเดือน' : 'ใบรับรองเงินเดือน'}เดือนไหนครับ?`,
    quickReply: { items: quickItems },
  });
}

// ════════════════════════════════════════════════════════
// พนักงานเลือกเดือนแล้ว → แจ้ง HR
// ════════════════════════════════════════════════════════
async function handleMonthSelected(userId, month, requestId) {
  const req = pending[requestId];
  if (!req) {
    await push(userId, '⚠️ คำขอหมดอายุ กรุณาขอใหม่อีกครั้งครับ');
    return;
  }

  req.month = month;
  if (requestLog[requestId]) { requestLog[requestId].month = month; }
  const docLabel = req.docType === 'payslip' ? 'สลิปเงินเดือน' : 'ใบรับรองเงินเดือน';

  // ยืนยันกับพนักงาน
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

  // แจ้ง HR พร้อมปุ่มอนุมัติ
  await notifyHR(requestId, req.empName, docLabel, month, userId);
}

// ════════════════════════════════════════════════════════
// แจ้ง HR
// ════════════════════════════════════════════════════════
async function notifyHR(requestId, empName, docType, month, empUserId) {
  const now = new Date().toLocaleString('th-TH', {
    timeZone: 'Asia/Bangkok', hour: '2-digit', minute: '2-digit',
    day: '2-digit', month: '2-digit'
  });

  await push(HR_USER_ID, {
    type: 'flex',
    altText: `คำขอ${docType} จาก ${empName}`,
    contents: {
      type: 'bubble',
      header: {
        type: 'box', layout: 'vertical',
        backgroundColor: '#1E3A5F', paddingAll: '16px',
        contents: [
          { type: 'text', text: 'คำขอเอกสาร HR', color: '#ffffff', weight: 'bold', size: 'md' },
          { type: 'text', text: docType, color: '#B8D4F0', size: 'sm', margin: 'xs' },
        ]
      },
      body: {
        type: 'box', layout: 'vertical', spacing: 'sm', paddingAll: '16px',
        contents: [
          { type: 'box', layout: 'horizontal', contents: [
            { type: 'text', text: 'พนักงาน', size: 'sm', color: '#888888', flex: 3 },
            { type: 'text', text: empName, size: 'sm', weight: 'bold', flex: 5, wrap: true },
          ]},
          { type: 'box', layout: 'horizontal', contents: [
            { type: 'text', text: 'เดือน', size: 'sm', color: '#888888', flex: 3 },
            { type: 'text', text: month, size: 'sm', flex: 5 },
          ]},
          { type: 'box', layout: 'horizontal', contents: [
            { type: 'text', text: 'ประเภท', size: 'sm', color: '#888888', flex: 3 },
            { type: 'text', text: docType, size: 'sm', flex: 5 },
          ]},
          { type: 'box', layout: 'horizontal', contents: [
            { type: 'text', text: 'เวลาขอ', size: 'sm', color: '#888888', flex: 3 },
            { type: 'text', text: now, size: 'sm', flex: 5 },
          ]},
        ]
      },
      footer: {
        type: 'box', layout: 'vertical', spacing: 'sm', paddingAll: '12px',
        contents: [
          { type: 'button', style: 'primary', color: '#1B7F4E', height: 'sm',
            action: { type: 'postback', label: 'อนุมัติ และส่ง PDF',
              data: `action=approve&rid=${requestId}` } },
          { type: 'button', style: 'secondary', height: 'sm',
            action: { type: 'postback', label: 'ปฏิเสธ',
              data: `action=reject&rid=${requestId}&uid=${empUserId}&name=${encodeURIComponent(empName)}&doc=${encodeURIComponent(docType)}` } },
        ]
      }
    }
  });
}

// ════════════════════════════════════════════════════════
// Postback: HR กดอนุมัติ/ปฏิเสธ
// ════════════════════════════════════════════════════════
async function handlePostback(event) {
  const hrUserId = event.source.userId;
  const params   = Object.fromEntries(
    event.postback.data.split('&').map(p => {
      const [k, v] = p.split('=');
      return [k, decodeURIComponent(v || '')];
    })
  );

  if (params.action === 'reject') {
    const empId  = params.uid || '';
    const name   = params.name || 'พนักงาน';
    const doc    = params.doc || 'เอกสาร';
    if (empId) await push(empId, `❌ คำขอ${doc}ของคุณถูกปฏิเสธ กรุณาติดต่อ HR โดยตรงครับ`);
    await push(hrUserId, `✅ ปฏิเสธคำขอของ ${name} แล้ว`);
    delete pending[params.rid];
    if (requestLog[params.rid]) requestLog[params.rid].status = 'rejected';
    return;
  }

  if (params.action === 'approve') {
    const req = pending[params.rid];
    if (!req) {
      await push(hrUserId, '⚠️ ไม่พบคำขอนี้ อาจดำเนินการไปแล้ว');
      return;
    }

    await push(hrUserId, `⏳ กำลังสร้าง PDF สำหรับ ${req.empName} เดือน ${req.month}...`);

    try {
      // ดึงข้อมูลเงินเดือนจาก Google Sheets
      const data = await payroll.getEmployeePayroll(req.empName, req.month);
      if (!data) {
        await push(hrUserId, `❌ ไม่พบข้อมูลเงินเดือนของ ${req.empName} เดือน ${req.month}\nกรุณาตรวจสอบว่าอัปโหลดไฟล์ Excel เดือนนั้นแล้ว`);
        return;
      }

      // สร้าง PDF
      let pdfUrl;
      if (req.docType === 'payslip') {
        pdfUrl = await payslip.createFromPayroll(data);
      } else {
        pdfUrl = await cert.createFromPayroll(data);
      }

      const docLabel = req.docType === 'payslip' ? 'สลิปเงินเดือน' : 'ใบรับรองเงินเดือน';

      // ส่งให้พนักงาน
      await push(req.empLineId, {
        type: 'flex',
        altText: `${docLabel} ${req.month} พร้อมแล้ว`,
        contents: {
          type: 'bubble',
          header: {
            type: 'box', layout: 'vertical',
            backgroundColor: '#1E3A5F', paddingAll: '16px',
            contents: [
              { type: 'text', text: `📄 ${docLabel}`, color: '#ffffff', weight: 'bold' },
              { type: 'text', text: req.empName, color: '#B8D4F0', size: 'sm', margin: 'xs' },
            ]
          },
          body: {
            type: 'box', layout: 'vertical', spacing: 'sm', paddingAll: '16px',
            contents: [
              { type: 'text', text: '✅ เอกสารพร้อมแล้วครับ', weight: 'bold', color: '#06C755' },
              { type: 'text', text: `ประจำเดือน ${req.month}`, size: 'sm', color: '#555555', margin: 'sm' },
            ]
          },
          footer: {
            type: 'box', layout: 'vertical', spacing: 'sm', paddingAll: '12px',
            contents: [
              { type: 'button', style: 'primary', color: '#1E3A5F', height: 'sm',
                action: { type: 'uri', label: 'เปิด PDF', uri: pdfUrl } },
            ]
          }
        }
      });

      await push(hrUserId, `✅ ส่ง PDF ${docLabel} ให้ ${req.empName} แล้วครับ`);
      delete pending[params.rid];
      if (requestLog[params.rid]) requestLog[params.rid].status = 'sent';

      // log ลง sheet
      await sheet.log({ ...data, docType: req.docType });

    } catch(err) {
      console.error('approve error:', err.message);
      await push(hrUserId, `❌ เกิดข้อผิดพลาด: ${err.message}`);
    }
  }
}

// ════════════════════════════════════════════════════════
// HR อัปโหลด Excel รายเดือน
// ════════════════════════════════════════════════════════
app.post('/upload-payroll', upload.single('file'), async (req, res) => {
  res.json({ ok: true, message: 'กำลังประมวลผล...' });

  try {
    const file = req.file;
    if (!file) return;

    console.log('upload-payroll:', file.originalname);

    // อ่านไฟล์ Excel
    const data = await payroll.parseXls(file.path);
    const { month, rows } = data;

    console.log(`อ่านข้อมูลเดือน ${month} จำนวน ${rows.length} คน`);

    // บันทึกลง Google Sheets
    await payroll.savePayrollToSheet(month, rows);

    // บันทึกไฟล์ต้นฉบับลง Drive
    const fileBuffer = fs.readFileSync(file.path);
    await payroll.saveToGoogleDrive(
      fileBuffer,
      `เงินเดือน_${month}_${file.originalname}`,
      file.mimetype
    );

    // ลบไฟล์ temp
    fs.unlinkSync(file.path);

    // แจ้ง HR
    if (req.body.notifyUserId) {
      await push(req.body.notifyUserId,
        `✅ อัปโหลดข้อมูลเงินเดือนเดือน ${month} สำเร็จ\nจำนวนพนักงาน: ${rows.length} คน`
      );
    }

  } catch(err) {
    console.error('upload-payroll error:', err.message);
  }
});

// ════════════════════════════════════════════════════════
// Flex Card วันลา (คงเดิม)
// ════════════════════════════════════════════════════════
function createLeaveCard(emp, imgUrl) {
  const d = {
    name:           emp['ชื่อ - นามสกุล'] || 'พนักงาน',
    vacationTotal:  emp['สิทธิ์พักร้อน']    || '0',
    vacationLeft:   emp['คงเหลือพักร้อน']   || '0',
    personalTotal:  emp['สิทธิ์ลากิจ']      || '0',
    personalLeft:   emp['คงเหลือลากิจ']     || '0',
    sickTotal:      emp['สิทธิ์ลาป่วย']     || '0',
    sickLeft:       emp['คงเหลือลาป่วย']    || '0',
    birthdayTotal:  emp['สิทธิ์วันเกิด']    || '0',
    birthdayLeft:   emp['คงเหลือลาวันเกิด'] || '0',
    maternityTotal: emp['สิทธิ์ลาคลอด']     || '0',
    maternityLeft:  emp['คงเหลือลาคลอด']    || '0',
  };
  return {
    type: 'flex', altText: `สรุปวันลาของ ${d.name}`,
    contents: {
      type: 'bubble', size: 'giga',
      body: {
        type: 'box', layout: 'vertical', paddingAll: '0px',
        contents: [
          { type: 'box', layout: 'horizontal', backgroundColor: '#06C755', paddingAll: '20px',
            contents: [
              { type: 'box', layout: 'vertical', flex: 0, width: '70px', height: '70px',
                cornerRadius: '100px', borderColor: '#ffffff', borderWidth: '2px',
                contents: [{ type: 'image', url: imgUrl, aspectMode: 'cover', size: 'full' }] },
              { type: 'box', layout: 'vertical', flex: 1, paddingStart: '15px', justifyContent: 'center',
                contents: [
                  { type: 'text', text: d.name, color: '#ffffff', weight: 'bold', size: 'lg', wrap: true },
                  { type: 'text', text: 'พนักงาน', color: '#E5E5E5', size: 'sm' },
                ]}
            ]},
          { type: 'box', layout: 'vertical', paddingAll: '20px', spacing: 'md',
            contents: [
              { type: 'text', text: 'วันลาคงเหลือ', weight: 'bold', color: '#888888', size: 'xs' },
              leaveRow('🏖','ลาพักร้อน',  d.vacationTotal,  d.vacationLeft,  '#F39C12'),
              { type: 'separator', color: '#F0F0F0' },
              leaveRow('🚗','ลากิจ',       d.personalTotal,  d.personalLeft,  '#3498DB'),
              { type: 'separator', color: '#F0F0F0' },
              leaveRow('😷','ลาป่วย',      d.sickTotal,      d.sickLeft,      '#E74C3C'),
              { type: 'separator', color: '#F0F0F0' },
              leaveRow('🎂','ลาวันเกิด',   d.birthdayTotal,  d.birthdayLeft,  '#9B59B6'),
              { type: 'separator', color: '#F0F0F0' },
              leaveRow('👶','ลาคลอด',      d.maternityTotal, d.maternityLeft, '#FF69B4'),
            ]},
          { type: 'box', layout: 'vertical', backgroundColor: '#F9F9F9', paddingAll: '10px',
            contents: [{ type: 'text', text: 'อัปเดตข้อมูลล่าสุดจากระบบ HR', color: '#AAAAAA', size: 'xxs', align: 'center' }] }
        ]
      }
    }
  };
}

function leaveRow(icon, label, total, left, color) {
  return {
    type: 'box', layout: 'horizontal', alignItems: 'center',
    contents: [
      { type: 'box', layout: 'horizontal', flex: 7, contents: [
        { type: 'text', text: icon, size: 'xl', flex: 0, margin: 'sm' },
        { type: 'box', layout: 'vertical', paddingStart: '10px', contents: [
          { type: 'text', text: label, size: 'sm', weight: 'bold', color: '#333333' },
          { type: 'text', text: `สิทธิ์ทั้งหมด ${total} วัน`, size: 'xxs', color: '#AAAAAA' },
        ]},
      ]},
      { type: 'box', layout: 'vertical', flex: 3, alignItems: 'flex-end', contents: [
        { type: 'text', text: String(left), size: 'xxl', weight: 'bold', color },
        { type: 'text', text: 'วัน', size: 'xxs', color },
      ]},
    ]
  };
}

async function reply(replyToken, msg) {
  const messages = typeof msg === 'string' ? [{ type: 'text', text: msg }] : [msg];
  await axios.post('https://api.line.me/v2/bot/message/reply',
    { replyToken, messages },
    { headers: { Authorization: `Bearer ${LINE_TOKEN}` } }
  );
}

async function push(userId, msg) {
  const messages = typeof msg === 'string' ? [{ type: 'text', text: msg }] : [msg];
  await axios.post('https://api.line.me/v2/bot/message/push',
    { to: userId, messages },
    { headers: { Authorization: `Bearer ${LINE_TOKEN}` } }
  );
}

async function getProfile(userId) {
  try {
    const r = await axios.get(`https://api.line.me/v2/bot/profile/${userId}`,
      { headers: { Authorization: `Bearer ${LINE_TOKEN}` } });
    return r.data;
  } catch { return null; }
}


// ════════════════════════════════════════════════════════
// Portal endpoints
// ════════════════════════════════════════════════════════
const path = require('path');

app.get('/portal', (req, res) => {
  res.sendFile(path.join(__dirname, 'portal.html'));
});

app.get('/portal/stats', async (req, res) => {
  try {
    const larkToken = await lark.getToken();
    const employees = await lark.getAllEmployees(larkToken);
    const months    = await payroll.getAvailableMonths();
    const pendCount = Object.keys(pending).length;
    res.json({
      employees:   employees.length,
      pending:     pendCount,
      latestMonth: months.length ? months[months.length-1] : '—',
    });
  } catch(e) {
    res.json({ employees: 0, pending: 0, latestMonth: '—' });
  }
});

app.get('/portal/employees', async (req, res) => {
  try {
    const token = await lark.getToken();
    const emps  = await lark.getAllEmployees(token);
    res.json(emps.map(e => ({
      name:     e['ชื่อ - นามสกุล'] || '',
      position: e['ตำแหน่ง'] || '',
      lineId:   e['Line ID'] || '',
    })));
  } catch(e) {
    res.json([]);
  }
});

app.get('/portal/requests', (req, res) => {
  const list = Object.entries(requestLog)
    .sort((a,b) => b[1].time - a[1].time)
    .slice(0, parseInt(req.query.limit) || 50)
    .map(([id, r]) => ({
      name:    r.empName,
      docType: r.docType,
      month:   r.month || '—',
      status:  r.status || 'pending',
      time:    new Date(r.time).toLocaleString('th-TH', { timeZone: 'Asia/Bangkok', hour: '2-digit', minute: '2-digit', day: '2-digit', month: '2-digit' }),
    }));
  res.json(list);
});

app.get('/portal/months', async (req, res) => {
  try {
    const months = await payroll.getAvailableMonths();
    res.json(months);
  } catch(e) {
    res.json([]);
  }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`TPE HR Bot v2 on port ${PORT}`));
