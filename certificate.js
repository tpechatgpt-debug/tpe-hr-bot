const { htmlToDriveUrl } = require('./payslip');

async function createFromPayroll(d) {
  const fmt   = n => (parseFloat(n)||0).toLocaleString('th-TH', { minimumFractionDigits: 2 });
  const today = new Date();
  const thM   = ['มกราคม','กุมภาพันธ์','มีนาคม','เมษายน','พฤษภาคม','มิถุนายน','กรกฎาคม','สิงหาคม','กันยายน','ตุลาคม','พฤศจิกายน','ธันวาคม'];
  const dateStr  = `${today.getDate()} เดือน${thM[today.getMonth()]} พ.ค.${today.getFullYear()+543}`;
  const docNo    = `FM-HR-09-${today.getFullYear()+543}/${String(Math.floor(Math.random()*89+10)).padStart(2,'0')}`;
  const salary   = parseFloat(d.baseWage) || parseFloat(d.basePay) || 0;

  // คำนวณอายุงาน (ถ้ามี startDate)
  let workDuration = '';
  if (d.startDate) {
    try {
      const parts = d.startDate.split('/');
      const start = new Date(parseInt(parts[2])-543, parseInt(parts[1])-1, parseInt(parts[0]));
      const months = Math.floor((today - start) / (1000*60*60*24*30.44));
      workDuration = `${Math.floor(months/12)} ปี ${months%12} เดือน`;
    } catch(e) {}
  }

  const html = `<!DOCTYPE html>
<html><head><meta charset="utf-8">
<style>
@import url('https://fonts.googleapis.com/css2?family=Sarabun:ital,wght@0,300;0,400;0,500;0,600;0,700;1,400&display=swap');
*{margin:0;padding:0;box-sizing:border-box}
body{font-family:'Sarabun',sans-serif;font-size:14px;color:#1a1a1a;background:#fff}
.page{padding:28px 40px 20px;min-height:100vh;position:relative}

/* HEADER */
.header{display:flex;align-items:center;gap:0;margin-bottom:0}
.logo-block{display:flex;align-items:center;gap:0;background:linear-gradient(135deg,#2C1A00,#6B3F00,#C9952A);padding:14px 16px;border-radius:6px 0 0 6px}
.logo-img{width:70px;height:70px;object-fit:contain}
.co-block{flex:1;background:#1E3A5F;padding:14px 20px;border-radius:0 6px 6px 0}
.co-en{font-size:20px;font-weight:700;color:#ffffff;letter-spacing:.03em}
.co-addr{font-size:11px;color:#B8D4F0;margin-top:3px}

/* DIVIDER */
.divider-gold{height:4px;background:linear-gradient(90deg,#8B5E00,#C9952A,#FFD580,#C9952A,#8B5E00);margin:12px 0}
.divider-dark{height:3px;background:#333;margin-bottom:20px}

/* DOC META */
.doc-no{text-align:right;font-size:11px;color:#888;margin-bottom:16px}

/* TITLE */
.doc-title{font-size:16px;font-weight:700;text-align:center;margin-bottom:28px;color:#1a1a1a;letter-spacing:.02em}

/* BODY */
.body-para{line-height:2.5;font-size:14px;text-align:justify;text-indent:3em;margin-bottom:10px}
.hl{font-weight:700;color:#1a1a1a}
.salary-ul{font-weight:700;border-bottom:1.5px solid #1a1a1a;padding:0 4px;font-size:15px}
.blank-line{display:inline-block;min-width:160px;border-bottom:1.5px solid #888;text-align:center}

/* ISSUE DATE */
.issue-date{text-align:center;margin:28px 0 32px;font-size:14px;line-height:1.8;color:#333}

/* SIGNATURE */
.sign-area{display:flex;justify-content:flex-end;margin-bottom:32px}
.sign-box{text-align:center;width:220px}
.sign-dots{color:#888;font-size:14px;margin-bottom:4px;border-bottom:1px solid #ccc;padding-bottom:6px;margin-top:48px}
.sign-name{font-size:13px;font-weight:600;margin-top:5px}
.sign-pos{font-size:12px;color:#666;margin-top:2px}

/* FOOTER */
.footer{border-top:3px solid #555;padding-top:10px;display:flex;justify-content:space-between;align-items:center;font-size:11px;color:#555}
.footer-contact{display:flex;gap:20px}
.footer-item{display:flex;align-items:center;gap:5px}
.footer-qr{width:52px;height:52px;background:#eee;border-radius:4px;display:flex;align-items:center;justify-content:center;font-size:8px;color:#888}

/* WATERMARK */
.watermark{position:absolute;top:50%;left:50%;transform:translate(-50%,-50%) rotate(-25deg);font-size:80px;font-weight:900;color:#C9952A;opacity:.04;pointer-events:none;white-space:nowrap;letter-spacing:6px}
</style>
</head><body>
<div class="page">
<div class="watermark">TPCB</div>

<!-- HEADER -->
<div class="header">
  <div class="logo-block">
    <img class="logo-img" src="data:image/png;base64,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">
  </div>
  <div class="co-block">
    <div class="co-en">THANAPHON ENGINEERING CO.,LTD.</div>
    <div class="co-addr">2 ก.คลองหมอ, ต.บ้านพรุ, อ.หาดใหญ่, จ.สงขลา 90250 &nbsp;(Tax ID : 0905559005578)</div>
  </div>
</div>

<div class="divider-gold"></div>
<div class="divider-dark"></div>

<div class="doc-no">${docNo}</div>

<div class="doc-title">หนังสือรับรองค่าจ้าง</div>

<p class="body-para">
  หนังสือฉบับนี้ให้ไว้เพื่อรับรองว่า <span class="hl">${d.name}</span>
  เป็นพนักงานบริษัท ธนพลเอ็นจิเนียริ่ง จำกัด
  ${d.startDate ? `ตั้งแต่วันที่ <span class="hl">${d.startDate}</span> จนถึงปัจจุบัน` : ''}
  ${workDuration ? `รวมระยะเวลาในการทำงานคือ <span class="hl">${workDuration}</span>` : ''}
  ซึ่งดำรงตำแหน่งงาน ณ ปัจจุบัน คือ
  <span class="hl">"${d.position||'—'}"</span>
  โดยได้รับอัตราเงินเดือน
  <span class="salary-ul">${fmt(salary)}</span>
  บาท/เดือน
  (<span class="blank-line">${thaiMoney(salary)} บาทถ้วน</span>)
  โดยยังไม่รวมค่าจ้างอื่น ๆ เมื่อพนักงานทำงานล่วงเวลา หรือเบี้ยเลี้ยงการออกหน้างาน
</p>

<div class="issue-date">ออกให้ ณ. วันที่ ${today.getDate()} เดือน${thM[today.getMonth()]} พ.ศ.${today.getFullYear()+543}</div>

<div class="sign-area">
  <div class="sign-box">
    <div class="sign-dots">ลงชื่อ............................................รับรอง</div>
    <div class="sign-name">${process.env.SIGNER_NAME || '(ปภัสสนันท์ เรืองฤทธิวรรณ)'}</div>
    <div class="sign-pos">ตำแหน่ง ${process.env.SIGNER_POSITION || 'เจ้าหน้าที่ทรัพยากรมนุษย์'}</div>
  </div>
</div>

<!-- FOOTER -->
<div class="footer">
  <div class="footer-contact">
    <div class="footer-item">📞 081-132-8878</div>
    <div class="footer-item">✉ sales@thanaphon.tech</div>
    <div class="footer-item">🌐 https://thanaphon.tech</div>
  </div>
  <div class="footer-qr">QR</div>
</div>

</div>
</body></html>`;

  const yr = today.getFullYear();
  const mo = String(today.getMonth()+1).padStart(2,'0');
  const { htmlToPdfBuffer } = require('./payslip');
  return await htmlToPdfBuffer(html);
}

function thaiMoney(n) {
  if (!n) return '—';
  const d=['','หนึ่ง','สอง','สาม','สี่','ห้า','หก','เจ็ด','แปด','เก้า'];
  const p=['','สิบ','ร้อย','พัน','หมื่น','แสน','ล้าน'];
  const s=Math.floor(n).toString(); let r='';
  for(let i=0;i<s.length;i++){
    const dg=parseInt(s[i]),ps=s.length-1-i;
    if(!dg) continue;
    if(ps===1&&dg===2) r+='ยี่'; else if(ps===1&&dg===1) r+=''; else r+=d[dg];
    r+=p[ps];
  }
  return r;
}

module.exports = { createFromPayroll };
