const axios = require('axios');
const BASE_ID         = process.env.LARK_BASE_ID;
const TABLE_ID        = process.env.LARK_TABLE_ID;
const APP_ID          = process.env.LARK_APP_ID;
const SECRET          = process.env.LARK_APP_SECRET;
const JOB_BASE_ID     = process.env.LARK_JOB_BASE_ID;
const ASSIGN_TABLE_ID = process.env.LARK_ASSIGN_TABLE_ID;

async function getToken() {
  const r = await axios.post('https://open.larksuite.com/open-apis/auth/v3/tenant_access_token/internal',
    { app_id: APP_ID, app_secret: SECRET });
  return r.data.tenant_access_token;
}

async function getRecords(token) {
  const url = `https://open.larksuite.com/open-apis/bitable/v1/apps/${BASE_ID}/tables/${TABLE_ID}/records?page_size=100`;
  const r = await axios.get(url, { headers: { Authorization: `Bearer ${token}` } });
  return r.data.data.items || [];
}

async function findByLineId(token, lineId) {
  const items = await getRecords(token);
  for (const item of items) {
    const f   = item.fields;
    const lid = f['Line ID'] || f['LineID'] || f['Line_ID'];
    if (lid && lid.toString().trim() === lineId) return f;
  }
  return null;
}

// ── ดึง records พร้อม record_id ──────────────────────────
async function getRecordsWithId(token) {
  const url = `https://open.larksuite.com/open-apis/bitable/v1/apps/${BASE_ID}/tables/${TABLE_ID}/records?page_size=100`;
  const r = await axios.get(url, { headers: { Authorization: `Bearer ${token}` } });
  return r.data.data.items || [];
}

async function register(token, nameToFind, lineId) {
  const items = await getRecords(token);
  for (const item of items) {
    const dbName = item.fields['ชื่อ - นามสกุล'];
    if (dbName && dbName.includes(nameToFind)) {
      await axios.put(
        `https://open.larksuite.com/open-apis/bitable/v1/apps/${BASE_ID}/tables/${TABLE_ID}/records/${item.record_id}`,
        { fields: { 'Line ID': lineId } },
        { headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' } }
      );
      return `✅ ลงทะเบียนสำเร็จ!\nยินดีต้อนรับคุณ "${dbName}" ครับ 🎉`;
    }
  }
  return `❌ ไม่พบชื่อ "${nameToFind}" ครับ\nลองพิมพ์ชื่อจริงใหม่อีกครั้งนะครับ`;
}

async function getAllEmployees(token) {
  const items = await getRecords(token);
  return items.map(item => item.fields || item);
}

// ── helper: แปลง Lark field (single/multi select หรือ text) → array of strings ──
function toStringArray(val) {
  if (!val) return [];
  if (Array.isArray(val)) return val.map(v => typeof v === 'object' ? (v.text || v.value || '') : v.toString()).filter(Boolean);
  return [val.toString()].filter(Boolean);
}

// ── getJobsToday: รองรับ ชุด และ สมาชิก แบบ Multiple select ──────────────────
async function getJobsToday(token, team, empName) {
  const now = Date.now();
  const url = `https://open.larksuite.com/open-apis/bitable/v1/apps/${JOB_BASE_ID}/tables/${ASSIGN_TABLE_ID}/records?page_size=100`;
  const r   = await axios.get(url, { headers: { Authorization: `Bearer ${token}` } });
  const items = r.data?.data?.items || [];

  const normN = s => (s || '').replace(/\s+/g, '').toLowerCase();
  const empNorm = empName ? normN(empName) : '';

  return items.filter(item => {
    const f     = item.fields;
    const start = f['วันที่เริ่ม'] || 0;
    const end   = f['วันสิ้นสุด'] || 0;
    if (start > now || now > end + 86400000) return false;

    // ชุด → Multiple select → Array
    const teams = toStringArray(f['ชุด']);

    // สมาชิก → Multiple select → Array
    const members = toStringArray(f['สมาชิก']);

    // ถ้ามี สมาชิก → เช็คว่าชื่อช่างอยู่ใน list ไหม
    if (members.length && empNorm) {
      const inList = members.some(m => {
        const mn = normN(m);
        return mn && (mn.includes(empNorm) || empNorm.includes(mn));
      });
      if (inList) return true;
      // มี สมาชิก แต่ชื่อไม่อยู่ใน list → ไม่แสดง
      return false;
    }

    // ถ้าไม่มี สมาชิก → fallback เช็ค ชุด (รองรับหลายชุด)
    return teams.some(t => normN(t) === normN(team));
  }).map(item => item.fields);
}

module.exports = { getToken, findByLineId, register, getAllEmployees, getJobsToday, getRecordsWithId };
