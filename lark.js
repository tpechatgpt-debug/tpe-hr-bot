const axios = require('axios');

const BASE_ID         = process.env.LARK_BASE_ID;
const TABLE_ID        = process.env.LARK_TABLE_ID;
const APP_ID          = process.env.LARK_APP_ID;
const SECRET          = process.env.LARK_APP_SECRET;
const JOB_BASE_ID     = process.env.LARK_JOB_BASE_ID;
const ASSIGN_TABLE_ID = process.env.LARK_ASSIGN_TABLE_ID;

// ── Token cache (expires ~2 hours, refresh 5 min early) ─────────────────────
let _tokenCache = null;
async function getToken() {
  const now = Date.now();
  if (_tokenCache && now < _tokenCache.expiresAt) return _tokenCache.token;
  const r = await axios.post(
    'https://open.larksuite.com/open-apis/auth/v3/tenant_access_token/internal',
    { app_id: APP_ID, app_secret: SECRET }
  );
  const token = r.data.tenant_access_token;
  const ttl   = (r.data.expire || 7200) * 1000;
  _tokenCache = { token, expiresAt: now + ttl - 5 * 60 * 1000 };
  return token;
}

// ── Paginated record fetcher ─────────────────────────────────────────────────
async function fetchAllRecords(token, baseId, tableId) {
  const records = [];
  let pageToken;
  do {
    const params = new URLSearchParams({ page_size: '100' });
    if (pageToken) params.set('page_token', pageToken);
    const url = `https://open.larksuite.com/open-apis/bitable/v1/apps/${baseId}/tables/${tableId}/records?${params}`;
    const r   = await axios.get(url, { headers: { Authorization: `Bearer ${token}` } });
    const data = r.data?.data || {};
    records.push(...(data.items || []));
    pageToken = data.has_more ? data.page_token : undefined;
  } while (pageToken);
  return records;
}

const getRecords       = (token) => fetchAllRecords(token, BASE_ID, TABLE_ID);
const getAssignRecords = (token) => fetchAllRecords(token, JOB_BASE_ID, ASSIGN_TABLE_ID);

// ── helper: normalise Lark select field → string[] ──────────────────────────
function toStringArray(val) {
  if (!val) return [];
  const arr = Array.isArray(val) ? val : [val];
  return arr
    .map(v => (typeof v === 'object' ? (v.text ?? v.value ?? '') : String(v)))
    .filter(Boolean);
}

const normN = s => (s || '').replace(/\s+/g, '').toLowerCase();

// ── findByLineId ─────────────────────────────────────────────────────────────
async function findByLineId(token, lineId) {
  const items = await getRecords(token);
  for (const item of items) {
    const f   = item.fields;
    const lid = f['Line ID'] ?? f['LineID'] ?? f['Line_ID'];
    if (lid?.toString().trim() === lineId) return f;
  }
  return null;
}

// alias ที่ใช้ใน eslip/fieldwork-history
const getEmployeeByLineId = async (token, lineId) => {
  const items = await getRecords(token);
  for (const item of items) {
    const f   = item.fields;
    const lid = f['Line ID'] ?? f['LineID'] ?? f['Line_ID'];
    if (lid?.toString().trim() === lineId) return f;
  }
  return null;
};

// ── register ─────────────────────────────────────────────────────────────────
async function register(token, nameToFind, lineId) {
  const items = await getRecords(token);
  for (const item of items) {
    const dbName = item.fields['ชื่อ - นามสกุล'];
    if (dbName?.includes(nameToFind)) {
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

// ── getAllEmployees ───────────────────────────────────────────────────────────
async function getAllEmployees(token) {
  const items = await getRecords(token);
  return items.map(item => item.fields);
}

// ── getJobsToday ─────────────────────────────────────────────────────────────
async function getJobsToday(token, team, empName) {
  const now     = Date.now();
  const items   = await getAssignRecords(token);
  const empNorm = empName ? normN(empName) : '';

  return items
    .filter(item => {
      const f     = item.fields;
      const start = f['วันที่เริ่ม'] || 0;
      const end   = f['วันสิ้นสุด']  || 0;
      if (start > now || now > end + 86_400_000) return false;

      const members = toStringArray(f['สมาชิก']);
      if (members.length) {
        return empNorm
          ? members.some(m => { const mn = normN(m); return mn && (mn.includes(empNorm) || empNorm.includes(mn)); })
          : false;
      }
      const teams = toStringArray(f['ชุด']);
      return teams.some(t => normN(t) === normN(team));
    })
    .map(item => item.fields);
}

const getRecordsWithId = getRecords;

module.exports = {
  getToken, findByLineId, getEmployeeByLineId,
  register, getAllEmployees, getJobsToday, getRecordsWithId,
};
