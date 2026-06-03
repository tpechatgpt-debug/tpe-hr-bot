// ════════════════════════════════════════════════════
// TPE Telegram Attendance Bot
// รับข้อความจาก bot สแกนหน้า → บันทึก Google Sheets
// ════════════════════════════════════════════════════
const axios = require('axios');

const TELEGRAM_TOKEN = process.env.TELEGRAM_TOKEN;
const TELEGRAM_API   = `https://api.telegram.org/bot${TELEGRAM_TOKEN}`;

// Parse ข้อความจาก bot สแกนหน้า
function parseAttendance(text) {
  const id_m   = text.match(/ID:\s*(\d+)/);
  const name_m = text.match(/ชื่อ:\s*(.+)/);
  const mode_m = text.match(/ตรวจสอบโหมด:\s*(.+)/);
  const time_m = text.match(/เวลา:\s*(\d{2}\/\d{2}\/\d{4})\s+(\d{2}:\d{2}:\d{2})/);
  if (!name_m || !time_m) return null;
  return {
    id:   id_m   ? id_m[1].trim()   : '',
    name: name_m[1].trim(),
    mode: mode_m ? mode_m[1].trim() : '',
    date: time_m[1],
    time: time_m[2],
  };
}

// บันทึกลง Google Sheets sheet "Attendance"
async function saveAttendance(sheets, spreadsheetId, data) {
  const sheetName = 'Attendance';

  // ตรวจว่ามี sheet ชื่อ Attendance ไหม ถ้าไม่มีสร้างใหม่
  try {
    const meta = await sheets.spreadsheets.get({ spreadsheetId });
    const exists = meta.data.sheets.some(s => s.properties.title === sheetName);
    if (!exists) {
      await sheets.spreadsheets.batchUpdate({
        spreadsheetId,
        requestBody: { requests: [{ addSheet: { properties: { title: sheetName } } }] }
      });
      // เพิ่ม header
      await sheets.spreadsheets.values.append({
        spreadsheetId,
        range: `${sheetName}!A1`,
        valueInputOption: 'RAW',
        requestBody: { values: [['วันที่', 'เวลา', 'ID', 'ชื่อ', 'โหมด', 'บันทึกเมื่อ']] }
      });
    }
  } catch(e) {
    console.error('sheet check error:', e.message);
  }

  const now = new Date().toLocaleString('th-TH', { timeZone: 'Asia/Bangkok' });
  await sheets.spreadsheets.values.append({
    spreadsheetId,
    range: `${sheetName}!A:F`,
    valueInputOption: 'RAW',
    requestBody: { values: [[data.date, data.time, data.id, data.name, data.mode, now]] }
  });
  console.log(`[Attendance] บันทึก: ${data.name} ${data.date} ${data.time}`);
}

// ส่งข้อความกลับไปที่ Telegram
async function telegramReply(chatId, text) {
  try {
    await axios.post(`${TELEGRAM_API}/sendMessage`, {
      chat_id: chatId,
      text,
      parse_mode: 'HTML',
    });
  } catch(e) {
    console.error('telegram reply error:', e.message);
  }
}

// Handler หลัก — เรียกจาก webhook
async function handleTelegramUpdate(update, sheets, spreadsheetId) {
  const msg = update.message || update.channel_post;
  if (!msg || !msg.text) return;

  const text   = msg.text;
  const chatId = msg.chat.id;

  const data = parseAttendance(text);
  if (!data) return; // ไม่ใช่ข้อความ attendance

  try {
    await saveAttendance(sheets, spreadsheetId, data);
    await telegramReply(chatId, `✅ บันทึกแล้ว\n👤 ${data.name}\n🕐 ${data.date} ${data.time}`);
  } catch(e) {
    console.error('handleTelegramUpdate error:', e.message);
    await telegramReply(chatId, `❌ บันทึกไม่สำเร็จ: ${e.message}`);
  }
}

module.exports = { handleTelegramUpdate, parseAttendance };
