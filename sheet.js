const { google } = require('googleapis');

async function log(d) {
  try {
    const auth = new google.auth.GoogleAuth({
      credentials: JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT_JSON),
      scopes: ['https://www.googleapis.com/auth/spreadsheets'],
    });
    const sheets    = google.sheets({ version: 'v4', auth });
    const sheetId   = process.env.LOG_SHEET_ID;
    const sheetName = d.docType === 'payslip' ? 'สลิปเงินเดือน' : 'ใบรับรองเงินเดือน';
    const now       = new Date().toLocaleString('th-TH', { timeZone: 'Asia/Bangkok' });

    const values = d.docType === 'payslip'
      ? [[now, d.month, d.employeeName,
          d.workDays||0, d.baseSalary||0, d.hDays||0, d.hPay||0,
          d.otH||0, d.otPay||0, d.al||0, d.dilPay||0,
          d.trad||0, d.tradPay||0, d.otInc||0,
          (d.baseSalary||0)+(d.hPay||0)+(d.otPay||0)+(d.dilPay||0)+(d.tradPay||0)+(d.otInc||0),
          d.lv||0, d.lp||0, d.ls||0, d.lnp||0, d.lm||0, d.lb||0,
          d.dAdv||0, d.dSoc||0, d.dTax||0, d.dAbs||0, d.dNpL||0, d.dKt||0, d.dOth||0,
          (d.dAdv||0)+(d.dSoc||0)+(d.dTax||0)+(d.dAbs||0)+(d.dNpL||0)+(d.dKt||0)+(d.dOth||0),
          (d.baseSalary||0)+(d.hPay||0)+(d.otPay||0)+(d.dilPay||0)+(d.tradPay||0)+(d.otInc||0)
          -(d.dAdv||0)-(d.dSoc||0)-(d.dTax||0)-(d.dAbs||0)-(d.dNpL||0)-(d.dKt||0)-(d.dOth||0)
        ]]
      : [[now, d.employeeName, d.baseSalary||0, d.month||'']];

    await sheets.spreadsheets.values.append({
      spreadsheetId: sheetId,
      range: `${sheetName}!A1`,
      valueInputOption: 'USER_ENTERED',
      resource: { values },
    });
  } catch (err) {
    console.error('sheet.log error:', err.message);
  }
}

module.exports = { log };
