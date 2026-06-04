// ════════════════════════════════════════════════════
// TPE GramJS — อ่านข้อความจาก Telegram User Account
// ดักข้อความที่ bot สแกนหน้าส่งมา → บันทึก Google Sheets
// ════════════════════════════════════════════════════
const { TelegramClient } = require('telegram');
const { StringSession }  = require('telegram/sessions');
const { NewMessage }     = require('telegram/events');

const API_ID   = parseInt(process.env.TELEGRAM_API_ID   || '37417945');
const API_HASH = process.env.TELEGRAM_API_HASH || 'a355eea7c1b5edaa55e600ac81975403';
const SESSION  = process.env.TELEGRAM_SESSION  || '';
const BOT_USERNAME = process.env.TELEGRAM_BOT_USERNAME || 'Thanaphon0822_bot';

// Parse ข้อความ attendance
function parseAttendance(text) {
  if (!text) return null;
  const id_m   = text.match(/ID:\s*(\d+)/);
  const name_m = text.match(/ชื่อ:\s*([^\n]+?)(?:\n|ตรวจสอบ|เวลา|={3}|$)/);
  const mode_m = text.match(/ตรวจสอบโหมด:\s*([^\n]+?)(?:\n|เวลา|={3}|$)/);
  const time_m = text.match(/เวลา:\s*(\d{4}\/\d{2}\/\d{2}|\d{2}\/\d{2}\/\d{4})\s+(\d{2}:\d{2}:\d{2})/);
  if (!name_m || !time_m) return null;
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

// queue สำหรับ batch write
const writeQueue = [];
let writing = false;

async function flushQueue(sheets, spreadsheetId) {
  if (writing || !writeQueue.length) return;
  writing = true;
  const sheetName = 'Attendance';
  try {
    // ตรวจ sheet
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
    // เขียนทีละ 20 rows
    while (writeQueue.length > 0) {
      const batch = writeQueue.splice(0, 20);
      await sheets.spreadsheets.values.append({
        spreadsheetId, range: `${sheetName}!A:F`, valueInputOption: 'RAW',
        requestBody: { values: batch }
      });
      if (writeQueue.length > 0) await new Promise(r => setTimeout(r, 2000));
    }
  } catch(e) {
    console.error('[GramJS] flush error:', e.message);
  }
  writing = false;
}

// บันทึก Google Sheets (queue-based)
async function saveAttendance(sheets, spreadsheetId, data) {
  const now = new Date().toLocaleString('th-TH', { timeZone: 'Asia/Bangkok' });
  writeQueue.push([data.date, data.time, data.id, data.name, data.mode, now]);
  console.log(`[GramJS] ✅ ${data.name} | ${data.date} ${data.time}`);
  // flush ทุก 3 วินาที
  setTimeout(() => flushQueue(sheets, spreadsheetId), 3000);
}

// เริ่ม GramJS client
async function startGramJS(sheets, spreadsheetId) {
  if (!SESSION) {
    console.log('[GramJS] ไม่มี TELEGRAM_SESSION — ข้าม');
    return;
  }
  try {
    const client = new TelegramClient(
      new StringSession(SESSION), API_ID, API_HASH,
      { connectionRetries: 5, useWSS: true }
    );
    await client.connect();
    console.log('[GramJS] เชื่อมต่อ Telegram สำเร็จ ✅');

    // ดักข้อความใหม่จาก bot สแกนหน้า
    client.addEventHandler(async (event) => {
      try {
        const msg    = event.message;
        const sender = await msg.getSender();
        const senderUsername = sender?.username || '';

        // รับเฉพาะจาก bot สแกนหน้า
        if (senderUsername !== BOT_USERNAME) return;

        const text = msg.text || '';
        console.log(`[GramJS] ข้อความจาก @${senderUsername}: ${text.slice(0, 60)}`);

        const data = parseAttendance(text);
        if (!data) { console.log('[GramJS] ไม่ใช่ข้อความ attendance'); return; }

        await saveAttendance(sheets, spreadsheetId, data);
      } catch(e) {
        console.error('[GramJS] handler error:', e.message);
      }
    }, new NewMessage({}));

    console.log(`[GramJS] รอข้อความจาก @${BOT_USERNAME}...`);

    // ดึงข้อความย้อนหลังตั้งแต่ 1 มิถุนายน 2026
    try {
      const sinceDate = new Date('2026-06-01T00:00:00+07:00');
      const sinceUnix = Math.floor(sinceDate.getTime() / 1000);
      console.log('[GramJS] กำลังดึงข้อความย้อนหลังตั้งแต่ 1 มิ.ย. 2569...');

      const botEntity = await client.getEntity(BOT_USERNAME);
      const messages  = await client.getMessages(botEntity, {
        limit: 500,
        offsetDate: Math.floor(Date.now() / 1000), // เริ่มจากปัจจุบัน
      });

      let count = 0;
      for (const msg of messages.reverse()) {
        if (msg.date < sinceUnix) continue;
        const text = msg.text || '';
        const data = parseAttendance(text);
        if (!data) continue;
        await saveAttendance(sheets, spreadsheetId, data);
        count++;
      }
      console.log(`[GramJS] queue ${count} รายการ กำลังเขียน Sheets...`);
      // flush ทั้งหมดหลัง queue เสร็จ
      await new Promise(r => setTimeout(r, 4000));
      await flushQueue(sheets, spreadsheetId);
      console.log('[GramJS] ดึงย้อนหลังเสร็จ ✅');
    } catch(e) {
      console.error('[GramJS] ดึงย้อนหลังไม่สำเร็จ:', e.message);
    }
  } catch(e) {
    console.error('[GramJS] เชื่อมต่อไม่สำเร็จ:', e.message);
  }
}

module.exports = { startGramJS };
