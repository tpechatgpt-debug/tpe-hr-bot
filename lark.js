const axios = require('axios');
const BASE_ID  = process.env.LARK_BASE_ID;
const TABLE_ID = process.env.LARK_TABLE_ID;
const APP_ID   = process.env.LARK_APP_ID;
const SECRET   = process.env.LARK_APP_SECRET;

// ─── Token Cache ─────────────────────────────────────────────
let _tokenCache = null;
let _tokenExpiry = 0;

async function getToken() {
  const now = Date.now();
  if (_tokenCache && now < _tokenExpiry) return _tokenCache;
  const r = await axios.post(
    'https://open.larksuite.com/open-apis/auth/v3/tenant_access_token/internal',
    { app_id: APP_ID, app_secret: SECRET }
  );
  _tokenCache = r.data.tenant_access_token;
  _tokenExpiry = now + (r.data.expire - 60) * 1000;
  return _tokenCache;
}

// ─── Retry helper ─────────────────────────────────────────────
async function axiosWithRetry(fn, retries = 3) {
  for (let i = 0; i < retries; i++) {
    try {
      return await fn();
    } catch (err) {
      const status = err?.response?.status;
      if (status === 429 && i < retries - 1) {
        const wait = (i + 1) * 2000;
        console.warn(`Lark 429 — รอ ${wait}ms แล้ว retry (${i + 1}/${retries})`);
        await new Promise(r => setTimeout(r, wait));
      } else {
        throw err;
      }
    }
  }
}

// ─── getRecords (ไม่มี cache — ลด call แทน) ──────────────────
async function getRecords(token) {
  const url = `https://open.larksuite.com/open-apis/bitable/v1/apps/${BASE_ID}/tables/${TABLE_ID}/records?page_size=100`;
  const r = await axiosWithRetry(() =>
    axios.get(url, { headers: { Authorization: `Bearer ${token}` } })
  );
  return r.data.data.items || [];
}

// ─── findByLineId (รับ items จากภายนอกได้) ───────────────────
async function findByLineId(token, lineId, items = null) {
  const records = items || await getRecords(token);
  for (const item of records) {
    const f   = item.fields;
    const lid = f['Line ID'] || f['LineID'] || f['Line_ID'];
    if (lid && lid.toString().trim() === lineId) return f;
  }
  return null;
}

// ─── register (รับ items จากภายนอกได้) ───────────────────────
async function register(token, nameToFind, lineId, items = null) {
  const records = items || await getRecords(token);
  for (const item of records) {
    const dbName = item.fields['ชื่อ - นามสกุล'];
    if (dbName && dbName.includes(nameToFind)) {
      await axiosWithRetry(() =>
        axios.put(
          `https://open.larksuite.com/open-apis/bitable/v1/apps/${BASE_ID}/tables/${TABLE_ID}/records/${item.record_id}`,
          { fields: { 'Line ID': lineId } },
          { headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' } }
        )
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

module.exports = { getToken, getRecords, findByLineId, register, getAllEmployees };
