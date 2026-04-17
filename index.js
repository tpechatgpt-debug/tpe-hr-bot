const axios = require('axios');
axios.interceptors.response.use(
  res => res,
  err => {
    if (err.response && err.response.status === 400) {
      console.log("=== [FOUND 400 ERROR] ===");
      console.log("URL:", err.config.url);
      console.log("Response Data:", JSON.stringify(err.response.data, null, 2));
      console.log("==========================");
    }
    throw err;
  }
);
require('./logger');
require('dotenv').config();
const express = require('express');
const lark    = require('./lark');
const payslip = require('./payslip');
const cert    = require('./certificate');
const sheet   = require('./sheet');

const app  = express();
app.use(express.json());

const LINE_TOKEN = process.env.LINE_ACCESS_TOKEN;
const HR_USER_ID = process.env.HR_LINE_USER_ID;
const FORM_URL   = process.env.FORM_URL;

// ── health check ──────────────────────────────────────────
app.get('/', (req, res) => res.send('TPE HR Bot OK'));

// ── LINE Webhook ──────────────────────────────────────────
app.post('/webhook', async (req, res) => {
  res.sendStatus(200); // ตอบ LINE ทันที

  try {
    const events = req.body.events;
    if (!events || events.length === 0) return;
    const event = events[0];

    if (event.type === 'postback') {
      await handlePostback(event);
      return;
    }
    if (event.type !== 'message' || event.message.type !== 'text') return;

    const userId     = event.source.userId;
    const replyToken = event.replyToken;
    const msg        = event.message.text.trim();

    if (msg === 'id') {
      await push(userId, `User ID: ${userId}`);
      return;
    }

    const token = await lark.getToken();

    if (msg === 'ขอสลิปเงินเดือน' || msg === 'ขอสลีปเงินเดือน') {
      await handlePayslipRequest(userId, token);
      return;
    }
    if (msg === 'ขอใบรับรองเงินเดือน') {
      await handleCertRequest(userId, token);
      return;
    }

    // ── ระบบลาเดิม ─────────────────────────────────────
    const profile  = await getProfile(userId);
    const imgUrl   = profile?.pictureUrl || 'https://cdn-icons-png.flaticon.com/512/3135/3135715.png';
    const employee = await lark.findByLineId(token, userId);

    if (employee) {
      await reply(replyToken, createLeaveCard(employee, imgUrl));
    } else {
      if (msg.length < 2) {
        await reply(replyToken, '⚠️ กรุณาพิมพ์ชื่อจริง เพื่อลงทะเบียนครับ');
        return;
      }
      const result = await lark.register(token, msg, userId);
      await reply(replyToken, result);
    }

  } catch (err) {
    console.error('webhook error:', err.message);
  }
});

// ── HR Form endpoint — รับข้อมูลจากฟอร์มแล้วสร้าง PDF ───
app.post('/send-doc', async (req, res) => {
  res.json({ ok: true }); // ตอบ client ทันที

  try {
    const data = req.body;
    console.log('send-doc:', data.employeeName, data.docType);

    // บันทึก Sheet
    await sheet.log(data);

    // สร้าง PDF
    let pdfUrl;
    if (data.docType === 'payslip') {
      pdfUrl = await payslip.create(data);
    } else {
      pdfUrl = await cert.create(data);
    }

    // ส่งให้พนักงาน
    const docLabel = data.docType === 'payslip' ? 'สลิปเงินเดือน' : 'ใบรับรองเงินเดือน';
    await push(data.employeeLineId, {
      type: 'flex',
      altText: `${docLabel} พร้อมแล้ว`,
      contents: {
        type: 'bubble',
        header: {
          type: 'box', layout: 'vertical',
          backgroundColor: '#1E3A5F', paddingAll: '16px',
          contents: [
            { type: 'text', text: `📄 ${docLabel}`, color: '#ffffff', weight: 'bold' },
            { type: 'text', text: data.employeeName, color: '#B8D4F0', size: 'sm', margin: 'xs' },
          ]
        },
        body: {
          type: 'box', layout: 'vertical', spacing: 'sm', paddingAll: '16px',
          contents: [
            { type: 'text', text: '✅ เอกสารพร้อมแล้วครับ', weight: 'bold', color: '#06C755' },
            { type: 'text', text: `ประจำเดือน ${data.month || ''}`, size: 'sm', color: '#555', margin: 'sm' },
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

  } catch (err) {
    console.error('send-doc error:', err.message);
  }
});

// ── พนักงานขอสลิป ────────────────────────────────────────
async function handlePayslipRequest(userId, token) {
  const emp = await lark.findByLineId(token, userId);
  if (!emp) {
    await push(userId, '❌ ไม่พบข้อมูลของคุณ กรุณาลงทะเบียนก่อนนะครับ');
    return;
  }
  const empName   = (emp['ชื่อ - นามสกุล'] || 'พนักงาน').toString().trim();
  const requestId = `PAY_${userId}_${Date.now()}`;

  await push(userId, {
    type: 'flex', altText: 'ส่งคำขอสลิปเงินเดือนแล้ว',
    contents: { type: 'bubble', body: { type: 'box', layout: 'vertical', spacing: 'md', paddingAll: '20px',
      contents: [
        { type: 'text', text: '✅ ส่งคำขอแล้ว', weight: 'bold', size: 'lg', color: '#06C755' },
        { type: 'text', text: 'สลิปเงินเดือน', size: 'sm', color: '#555', margin: 'sm' },
        { type: 'text', text: 'HR จะดำเนินการและส่ง PDF กลับมาให้คุณเร็วๆ นี้ครับ', size: 'sm', color: '#888', wrap: true, margin: 'md' },
      ]
    }}
  });

  await notifyHR(requestId, empName, 'สลิปเงินเดือน', userId);
}

// ── พนักงานขอใบรับรอง ────────────────────────────────────
async function handleCertRequest(userId, token) {
  const emp = await lark.findByLineId(token, userId);
  if (!emp) {
    await push(userId, '❌ ไม่พบข้อมูลของคุณ กรุณาลงทะเบียนก่อนนะครับ');
    return;
  }
  const empName   = (emp['ชื่อ - นามสกุล'] || 'พนักงาน').toString().trim();
  const requestId = `CERT_${userId}_${Date.now()}`;

  await push(userId, {
    type: 'flex', altText: 'ส่งคำขอใบรับรองเงินเดือนแล้ว',
    contents: { type: 'bubble', body: { type: 'box', layout: 'vertical', spacing: 'md', paddingAll: '20px',
      contents: [
        { type: 'text', text: '✅ ส่งคำขอแล้ว', weight: 'bold', size: 'lg', color: '#06C755' },
        { type: 'text', text: 'ใบรับรองเงินเดือน', size: 'sm', color: '#555', margin: 'sm' },
        { type: 'text', text: 'HR จะดำเนินการและส่ง PDF กลับมาให้คุณเร็วๆ นี้ครับ', size: 'sm', color: '#888', wrap: true, margin: 'md' },
      ]
    }}
  });

  await notifyHR(requestId, empName, 'ใบรับรองเงินเดือน', userId);
}

// ── แจ้ง HR พร้อม Flex card + ปุ่มเปิดฟอร์ม ─────────────
async function notifyHR(requestId, empName, docType, empUserId) {
  const docTypeParam = docType.includes('สลิป') ? 'payslip' : 'certificate';
  const formUrl = `${FORM_URL}?rid=${encodeURIComponent(requestId)}&type=${docTypeParam}&name=${encodeURIComponent(empName)}&lineId=${encodeURIComponent(empUserId)}`;
  const now = new Date().toLocaleString('th-TH', { timeZone: 'Asia/Bangkok', hour: '2-digit', minute: '2-digit', day: '2-digit', month: '2-digit' });

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
            { type: 'text', text: 'ชื่อพนักงาน', size: 'sm', color: '#888888', flex: 3 },
            { type: 'text', text: empName, size: 'sm', weight: 'bold', flex: 5, wrap: true },
          ]},
          { type: 'box', layout: 'horizontal', contents: [
            { type: 'text', text: 'ประเภท', size: 'sm', color: '#888888', flex: 3 },
            { type: 'text', text: docType, size: 'sm', flex: 5 },
          ]},
          { type: 'box', layout: 'horizontal', contents: [
            { type: 'text', text: 'เวลาที่ขอ', size: 'sm', color: '#888888', flex: 3 },
            { type: 'text', text: now, size: 'sm', flex: 5 },
          ]},
        ]
      },
      footer: {
        type: 'box', layout: 'vertical', spacing: 'sm', paddingAll: '12px',
        contents: [
          { type: 'button', style: 'primary', color: '#C9952A', height: 'sm',
            action: { type: 'uri', label: 'กรอกข้อมูลและออกเอกสาร', uri: formUrl } },
          { type: 'button', style: 'secondary', height: 'sm',
            action: { type: 'postback', label: 'ปฏิเสธ', data: `action=reject&requestId=${requestId}&empId=${empUserId}&empName=${encodeURIComponent(empName)}&docType=${encodeURIComponent(docType)}` } },
        ]
      }
    }
  });
}

// ── HR กดปฏิเสธ ──────────────────────────────────────────
async function handlePostback(event) {
  const params = Object.fromEntries(event.postback.data.split('&').map(p => p.split('=')));
  if (params.action !== 'reject') return;

  const empId   = params.empId || '';
  const empName = decodeURIComponent(params.empName || 'พนักงาน');
  const docType = decodeURIComponent(params.docType || 'เอกสาร');

  if (empId) await push(empId, `❌ คำขอ${docType}ของคุณถูกปฏิเสธ กรุณาติดต่อ HR โดยตรงครับ`);
  await push(event.source.userId, `✅ ปฏิเสธคำขอของ ${empName} แล้ว`);
}

// ── Flex Card วันลา ───────────────────────────────────────
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
              { type: 'box', layout: 'vertical', flex: 0, width: '70px', height: '70px', cornerRadius: '100px', borderColor: '#ffffff', borderWidth: '2px',
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
              leaveRow('🏖️', 'ลาพักร้อน',  d.vacationTotal,  d.vacationLeft,  '#F39C12'),
              { type: 'separator', color: '#F0F0F0' },
              leaveRow('🚗', 'ลากิจ',       d.personalTotal,  d.personalLeft,  '#3498DB'),
              { type: 'separator', color: '#F0F0F0' },
              leaveRow('😷', 'ลาป่วย',      d.sickTotal,      d.sickLeft,      '#E74C3C'),
              { type: 'separator', color: '#F0F0F0' },
              leaveRow('🎂', 'ลาวันเกิด',   d.birthdayTotal,  d.birthdayLeft,  '#9B59B6'),
              { type: 'separator', color: '#F0F0F0' },
              leaveRow('👶', 'ลาคลอด',      d.maternityTotal, d.maternityLeft, '#FF69B4'),
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
    type: 'box', 
    layout: 'horizontal', 
    alignItems: 'center',
    contents: [
      {
        type: 'box',
        layout: 'horizontal',
        flex: 7,
        contents: [
          { type: 'text', text: icon, size: 'xl', flex: 0, margin: 'sm' },
          {
            type: 'box',
            layout: 'vertical',
            paddingStart: '10px',
            contents: [
              { type: 'text', text: label, size: 'sm', weight: 'bold', color: '#333333' },
              { type: 'text', text: `สิทธิ์ทั้งหมด ${total} วัน`, size: 'xxs', color: '#AAAAAA' }
            ]
          }
        ]
      },
      {
        type: 'box',
        layout: 'vertical',
        flex: 3,
        alignItems: 'flex-end',
        contents: [
          // แก้ตรงนี้: ย้าย color มาไว้ใน text เท่านั้น
          { type: 'text', text: String(left), size: 'xxl', weight: 'bold', color: color }, 
          { type: 'text', text: 'วัน', size: 'xxs', color: color }
        ]
      }
    ]
  };
}

// ── LINE helpers ──────────────────────────────────────────
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

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`TPE HR Bot running on port ${PORT}`));

// --- ส่วนที่เพิ่มใหม่เพื่อตรวจสอบ Error 400 ---
app.use((req, res, next) => {
    // เก็บ log ข้อมูลที่ส่งเข้ามาหาบอท
    console.log("------------------------------------");
    console.log("New Request:", req.method, req.url);
    console.log("Body:", JSON.stringify(req.body, null, 2));
    next();
});
