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
.logo-circle{background:rgba(255,255,255,0.15);border-radius:50%;width:60px;height:60px;display:flex;flex-direction:column;align-items:center;justify-content:center;border:2px solid rgba(255,255,255,0.3);flex-shrink:0}
.logo-tp{font-size:20px;font-weight:900;color:#fff;line-height:1}
.logo-cb{font-size:14px;font-weight:700;color:#FFD580;line-height:1}
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
    <div class="logo-circle">
      <div class="logo-tp">TP</div>
      <div class="logo-cb">CB</div>
    </div>
  </div>
  <div class="co-block">
    <div class="co-en">THANAPHON ENGINEERING CO.,LTD.</div>
    <div class="co-addr">2 ก.คลองหมอ, ต.บ้านพรุ, อ.หาดใหญ่, จ.สงขลา 90250 &nbsp;(Tax ID : 0905559005578)</div>
  </div>
</div>

<div class="divider-gold"></div>
<div class="divider-dark"></div>

<div class="doc-no">${docNo}</div>

<div class="doc-title">หนังสือรับรองเงินเดือน</div>

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
    <div class="sign-name">(ปภัสสนันท์ เรืองฤทธิวรรณ)</div>
    <div class="sign-pos">ตำแหน่ง เจ้าหน้าที่ทรัพยากรมนุษย์</div>
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
