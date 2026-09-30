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

let _attendanceCapacityCheckedAt = 0;
// บันทึกลง Google Sheets sheet "Attendance" — throw ถ้าล้มเหลว (ให้ pollTelegram ตัดสินใจว่าจะ retry ไหม)
async function saveAttendance(sheets, spreadsheetId, data) {
  const sheetName = 'Attendance';
  const meta = await sheets.spreadsheets.get({ spreadsheetId });
  let sheetProps = meta.data.sheets.find(s => s.properties.title === sheetName);
  if (!sheetProps) {
    await sheets.spreadsheets.batchUpdate({
      spreadsheetId,
      requestBody: { requests: [{ addSheet: { properties: { title: sheetName, gridProperties: { rowCount: 5000, columnCount: 6 } } } }] }
    });
    await sheets.spreadsheets.values.append({
      spreadsheetId, range: `${sheetName}!A1`, valueInputOption: 'RAW',
      requestBody: { values: [['วันที่', 'เวลา', 'ID', 'ชื่อ', 'โหมด', 'บันทึกเมื่อ']] }
    });
  } else if (Date.now() - _attendanceCapacityCheckedAt > 30 * 60 * 1000) {
    // ── auto-ขยายแถวกัน "ชีตเต็ม" เช็คไม่เกินทุก 30 นาที ──
    _attendanceCapacityCheckedAt = Date.now();
    try {
      const rowCount = sheetProps.properties.gridProperties.rowCount;
      const valuesRes = await sheets.spreadsheets.values.get({ spreadsheetId, range: `${sheetName}!A:A` });
      const usedRows = (valuesRes.data.values || []).length;
      if (rowCount - usedRows < 200) {
        await sheets.spreadsheets.batchUpdate({
          spreadsheetId,
          requestBody: { requests: [{ appendDimension: { sheetId: sheetProps.properties.sheetId, dimension: 'ROWS', length: 3000 } }] }
        });
        console.log(`[Attendance] ⚠️ ใกล้เต็ม (ใช้ ${usedRows}/${rowCount} แถว) → ขยายเพิ่ม 3000 แถวอัตโนมัติ`);
      }
    } catch (e) { console.error('[Attendance] capacity check error:', e.message); }
  }

  const now = new Date().toLocaleString('th-TH', { timeZone: 'Asia/Bangkok' });
  // ไม่ครอบ try/catch ตรงนี้ — ถ้า append พังต้องปล่อยให้ error หลุดออกไป
  // เพื่อให้ pollTelegram รู้ว่าบันทึกไม่สำเร็จ แล้วไม่ขยับ lastUpdateId (กันข้อมูลหาย)
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
      const msg = update.message;
      if (!msg) { lastUpdateId = update.update_id; continue; } // ไม่ใช่ message (เช่น edited_message) ข้ามได้ปลอดภัย

      // Log ทุก message เพื่อ debug
      console.log(`[Attendance] msg from chat_id=${msg.chat.id} type=${msg.chat.type} text=${(msg.text||'').slice(0,50)}`);

      const data = parseAttendance(msg.text);
      if (!data) {
        console.log('[Attendance] parse failed — ไม่ใช่ข้อความ attendance');
        lastUpdateId = update.update_id; // ไม่ใช่ข้อความสแกน ข้ามได้ปลอดภัย ไม่มีข้อมูลจะหาย
        continue;
      }

      // ── สำคัญ: ขยับ lastUpdateId ก็ต่อเมื่อบันทึกสำเร็จเท่านั้น ──
      // กันเคส Sheet เขียนไม่ได้ (เช่นแถวเต็ม) แล้ว Telegram ทำเหมือนข้อความนี้ถูกอ่านไปแล้ว
      // ทั้งที่ไม่เคยถูกบันทึกจริง (นี่คือสาเหตุที่ทำให้ข้อมูลหายไปตั้งแต่ 24/09)
      try {
        await saveAttendance(sheets, spreadsheetId, data);
        lastUpdateId = update.update_id;
      } catch (e) {
        console.error(`[Attendance] ❌ บันทึกไม่สำเร็จ (${data.name} | ${data.date} ${data.time}):`, e.message, '→ จะลองใหม่รอบถัดไป');
        break; // หยุด loop รอบนี้ทันที ไม่ขยับ offset ต่อ กันข้ามข้อความที่ยังไม่สำเร็จ
      }
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
