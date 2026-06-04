// ════════════════════════════════════════════════════
// TPE Telegram Attendance Bot
// Poll ดึงข้อความที่ bot ส่งออกไป → บันทึก Google Sheets
// ════════════════════════════════════════════════════
const axios = require('axios');

const TELEGRAM_TOKEN  = process.env.TELEGRAM_TOKEN;
const TELEGRAM_API    = `https://api.telegram.org/bot${TELEGRAM_TOKEN}`;
const OWNER_CHAT_ID   = process.env.TELEGRAM_OWNER_ID || '7870528980';

let lastUpdateId = -1; // -1 = ดึงทั้งหมดตั้งแต่ต้น

// Parse ข้อความจาก bot สแกนหน้า (รองรับทั้ง newline และ inline)
function parseAttendance(text) {
  if (!text) return null;
  const id_m   = text.match(/ID:\s*(\d+)/);
  const name_m = text.match(/ชื่อ:\s*([^\n]+?)(?:\n|ตรวจสอบ|เวลา|={3}|$)/);
  const mode_m = text.match(/ตรวจสอบโหมด:\s*([^\n]+?)(?:\n|เวลา|={3}|$)/);
  // รองรับ yyyy/mm/dd และ dd/mm/yyyy
  const time_m = text.match(/เวลา:\s*(\d{4}\/\d{2}\/\d{2}|\d{2}\/\d{2}\/\d{4})\s+(\d{2}:\d{2}:\d{2})/);
  if (!name_m || !time_m) return null;
  // แปลง yyyy/mm/dd → dd/mm/yyyy
  let rawDate = time_m[1];
  if (rawDate.indexOf('/') === 4) {
    const p = rawDate.split('/');
    rawDate = `${p[2]}/${p[1]}/${p[0]}`;
  }
  return {
    id:   id_m   ? id_m[1].trim()   : '',
    name: name_m[1].trim(),
    mode: mode_m ? mode_m[1].trim() : '',
    date: rawDate,
    time: time_m[2],
  };
}

// บันทึกลง Google Sheets sheet "Attendance"
async function saveAttendance(sheets, spreadsheetId, data) {
  const sheetName = 'Attendance';
  try {
    const meta = await sheets.spreadsheets.get({ spreadsheetId });
    const exists = meta.data.sheets.some(s => s.properties.title === sheetName);
    if (!exists) {
      await sheets.spreadsheets.batchUpdate({
        spreadsheetId,
        requestBody: { requests: [{ addSheet: { properties: { title: sheetName } } }] }
      });
      await sheets.spreadsheets.values.append({
        spreadsheetId, range: `${sheetName}!A1`, valueInputOption: 'RAW',
        requestBody: { values: [['วันที่', 'เวลา', 'ID', 'ชื่อ', 'โหมด', 'บันทึกเมื่อ']] }
      });
    }
  } catch(e) { console.error('sheet check:', e.message); }

  const now = new Date().toLocaleString('th-TH', { timeZone: 'Asia/Bangkok' });
  await sheets.spreadsheets.values.append({
    spreadsheetId, range: `${sheetName}!A:F`, valueInputOption: 'RAW',
    requestBody: { values: [[data.date, data.time, data.id, data.name, data.mode, now]] }
  });
  console.log(`[Attendance] ✅ ${data.name} | ${data.date} ${data.time}`);
}

// ดึงข้อความใหม่จาก Telegram (getUpdates)
async function pollTelegram(sheets, spreadsheetId) {
  try {
    const params = lastUpdateId === -1
      ? { limit: 100 }  // ครั้งแรก: ดึงทั้งหมด
      : { offset: lastUpdateId + 1, limit: 100 };
    const r = await axios.get(`${TELEGRAM_API}/getUpdates`, {
      params, timeout: 5000
    });
    const updates = r.data.result || [];
    for (const update of updates) {
      lastUpdateId = update.update_id;
      const msg = update.message;
      if (!msg) continue;

      // Log ทุก message เพื่อ debug
      console.log(`[Attendance] msg from chat_id=${msg.chat.id} type=${msg.chat.type} text=${(msg.text||'').slice(0,50)}`);

      // รับทุก chat ID (ไม่ filter) เพื่อ debug
      const data = parseAttendance(msg.text);
      if (!data) {
        console.log('[Attendance] parse failed — ไม่ใช่ข้อความ attendance');
        continue;
      }
      await saveAttendance(sheets, spreadsheetId, data);
    }
  } catch(e) {
    console.error('[Attendance] poll error:', e.message);
  }
}

// เริ่ม polling ทุก 30 วินาที
function startPolling(sheets, spreadsheetId) {
  console.log('[Attendance] เริ่ม polling Telegram ทุก 30 วินาที');
  pollTelegram(sheets, spreadsheetId); // poll ทันทีครั้งแรก
  setInterval(() => pollTelegram(sheets, spreadsheetId), 30 * 1000);
}

module.exports = { startPolling, parseAttendance };
