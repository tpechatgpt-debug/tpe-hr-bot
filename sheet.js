const { google } = require('googleapis');

async function log(d) {
  try {
    const auth = new google.auth.GoogleAuth({
      credentials: JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT_JSON),
      scopes: ['https://www.googleapis.com/auth/spreadsheets'],
    });
    const sheets  = google.sheets({ version: 'v4', auth });
    const sheetId = process.env.LOG_SHEET_ID;

    // ชื่อ tab ตาม docType
    const tabName = d.docType === 'payslip' ? 'สลิปเงินเดือน' : 'ใบรับรองเงินเดือน';
    const now     = new Date().toLocaleString('th-TH', { timeZone: 'Asia/Bangkok' });

    // ตรวจว่า tab มีอยู่ไหม ถ้าไม่มีให้สร้าง
    const meta = await sheets.spreadsheets.get({ spreadsheetId: sheetId });
    const exists = (meta.data.sheets || []).some(s => s.properties.title === tabName);

    if (!exists) {
      await sheets.spreadsheets.batchUpdate({
        spreadsheetId: sheetId,
        requestBody: {
          requests: [{ addSheet: { properties: { title: tabName } } }],
        },
      });
      // ใส่ header
      const headers = d.docType === 'payslip'
        ? [['วันที่', 'ชื่อ-นามสกุล', 'ตำแหน่ง', 'เดือน', 'ประเภท', 'เงินเดือน', 'รายได้รวม', 'รายจ่ายรวม', 'เงินได้สุทธิ']]
        : [['วันที่', 'ชื่อ-นามสกุล', 'ตำแหน่ง', 'เดือน', 'เงินเดือน']];
      await sheets.spreadsheets.values.update({
        spreadsheetId: sheetId,
        range: `'${tabName}'!A1`,
        valueInputOption: 'RAW',
        requestBody: { values: headers },
      });
    }

    // append ข้อมูล
    const values = d.docType === 'payslip'
      ? [[now, d.name||'', d.position||'', d.month||'', d.payType||'monthly',
          d.baseWage||d.basePay||0, d.totalInc||0, d.totalDed||0, d.netPay||0]]
      : [[now, d.name||'', d.position||'', d.month||'', d.baseWage||d.basePay||0]];

    await sheets.spreadsheets.values.append({
      spreadsheetId: sheetId,
      range: `'${tabName}'!A1`,
      valueInputOption: 'USER_ENTERED',
      requestBody: { values },
    });

  } catch (err) {
    console.error('sheet.log error:', err.message);
  }
}

module.exports = { log };
