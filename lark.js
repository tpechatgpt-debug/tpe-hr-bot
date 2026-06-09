const axios = require('axios');
const BASE_ID        = process.env.LARK_BASE_ID;
const TABLE_ID       = process.env.LARK_TABLE_ID;
const APP_ID         = process.env.LARK_APP_ID;
const SECRET         = process.env.LARK_APP_SECRET;
const JOB_BASE_ID    = process.env.LARK_JOB_BASE_ID;
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
async function getJobsToday(token, team) {
  const now   = Date.now();
  const url   = `https://open.larksuite.com/open-apis/bitable/v1/apps/${JOB_BASE_ID}/tables/${ASSIGN_TABLE_ID}/records?page_size=100`;
  const r     = await axios.get(url, { headers: { Authorization: `Bearer ${token}` } });
  const items = r.data?.data?.items || [];
  return items.filter(item => {
    const f     = item.fields;
    const start = f['วันที่เริ่ม']  || 0;
    const end   = f['วันสิ้นสุด'] || 0;
    const t     = (f['ชุด'] || '').toString();
    return t === team && start <= now && now <= end;
  }).map(item => item.fields);
}

module.exports = { getToken, findByLineId, register, getAllEmployees, getJobsToday };
