const { htmlToDriveUrl } = require('./payslip');

async function create(d) {
  const fmt  = n => (n||0).toLocaleString('th-TH', { minimumFractionDigits: 2 });
  const today = new Date();
  const thM   = ['มกราคม','กุมภาพันธ์','มีนาคม','เมษายน','พฤษภาคม','มิถุนายน','กรกฎาคม','สิงหาคม','กันยายน','ตุลาคม','พฤศจิกายน','ธันวาคม'];
  const dateStr = `${today.getDate()} ${thM[today.getMonth()]} พ.ศ. ${today.getFullYear()+543}`;
  const docNo   = `FM-HR-09-${today.getFullYear()+543}/${String(Math.floor(Math.random()*90+10))}`;

  const html = `<!DOCTYPE html><html><head><meta charset="utf-8">
<style>
@import url('https://fonts.googleapis.com/css2?family=Sarabun:wght@400;600;700&display=swap');
*{margin:0;padding:0;box-sizing:border-box}
body{font-family:'Sarabun',sans-serif;font-size:14px}
.page{padding:36px 52px;position:relative;min-height:100vh}
.border-outer{position:absolute;top:10px;left:10px;right:10px;bottom:10px;border:2.5px solid #1E3A5F;border-radius:3px;pointer-events:none}
.border-inner{position:absolute;top:16px;left:16px;right:16px;bottom:16px;border:1px solid #1E3A5F;border-radius:2px;pointer-events:none;opacity:.3}
.header{display:grid;grid-template-columns:80px 1fr;gap:12px;align-items:center;padding-bottom:10px;border-bottom:3px solid #C9952A;margin-bottom:6px}
.logo{width:70px;height:56px;background:linear-gradient(135deg,#8B6914,#C9952A);border-radius:6px;display:flex;flex-direction:column;align-items:center;justify-content:center;font-weight:900;color:#fff;line-height:1.1}
.logo .tp{font-size:20px}.logo .cb{font-size:14px}
.co-en{font-size:18px;font-weight:700;color:#1E3A5F}.co-addr{font-size:10.5px;color:#666;margin-top:2px}
.gray-bar{background:#555;height:5px;margin-bottom:20px}
.doc-no{text-align:right;font-size:11px;color:#888;margin-bottom:20px}
.doc-title{font-size:15px;font-weight:700;text-align:center;margin-bottom:26px}
.body-text{line-height:2.4;font-size:14px;text-align:justify}
.body-text p{text-indent:3em;margin-bottom:8px}
.bold{font-weight:700}
.salary-line{font-weight:700;font-size:15px;border-bottom:1.5px solid #1a1a1a;padding:0 3px}
.blank{display:inline-block;min-width:120px;border-bottom:1px solid #999;text-align:center;font-weight:600}
.issue{text-align:center;margin:26px 0 28px;line-height:2}
.sign{display:flex;justify-content:flex-end;margin-bottom:28px}
.sign-box{text-align:center;width:190px}
.sign-line{border-top:1px solid #888;padding-top:7px;margin-top:44px;font-size:12px;color:#555}
.footer-bar{border-top:3px solid #555;padding-top:8px;display:grid;grid-template-columns:1fr 1fr 1fr;font-size:10.5px;color:#555}
.watermark{position:absolute;top:50%;left:50%;transform:translate(-50%,-50%) rotate(-25deg);font-size:70px;font-weight:900;color:#C9952A;opacity:.04;pointer-events:none;white-space:nowrap}
</style></head><body>
<div class="page">
<div class="border-outer"></div><div class="border-inner"></div>
<div class="watermark">TPCB</div>
<div class="header">
  <div class="logo"><div class="tp">TP</div><div class="cb">CB</div></div>
  <div>
    <div class="co-en">THANAPHON ENGINEERING CO.,LTD.</div>
    <div class="co-addr">2 ถ.คลองหมอ, ต.บ้านพรุ, อ.หาดใหญ่, จ.สงขลา 90250 &nbsp;(Tax ID : 0905559005578)</div>
  </div>
</div>
<div class="gray-bar"></div>
<div class="doc-no">${docNo}</div>
<div class="doc-title">หนังสือรับรองค่าจ้าง</div>
<div class="body-text">
  <p>หนังสือฉบับนี้ให้ไว้เพื่อรับรองว่า <span class="bold">${d.employeeName}</span> เป็นพนักงานบริษัท ธนพลเอ็นจิเนียริ่ง จำกัด ซึ่งได้รับอัตราเงินเดือน <span class="salary-line">${fmt(d.baseSalary)}</span> บาท/เดือน (<span class="blank">${thaiMoney(d.baseSalary)} บาทถ้วน</span>) โดยยังไม่รวมค่าจ้างอื่น ๆ เมื่อพนักงานทำงานล่วงเวลา หรือเบี้ยเลี้ยงการออกหน้างาน</p>
</div>
<div class="issue">ออกให้ ณ วันที่ ${dateStr}</div>
<div class="sign">
  <div class="sign-box">
    <div class="sign-line">
      <div style="font-size:12.5px;font-weight:600">............................................รับรอง</div>
      <div style="margin-top:3px">(นางสาวปภัสนันท์ เรืองฤทธิ์วรรณ)</div>
      <div style="margin-top:2px">ตำแหน่ง เจ้าหน้าที่ทรัพยากรมนุษย์</div>
    </div>
  </div>
</div>
<div class="footer-bar">
  <div>081-132-8878</div>
  <div style="text-align:center">sales@thanaphon.tech</div>
  <div style="text-align:right">https://thanaphon.tech</div>
</div>
</div></body></html>`;

  return await htmlToDriveUrl(html, `Certificate_${d.employeeName}_${today.getFullYear()}${String(today.getMonth()+1).padStart(2,'0')}`);
}

function thaiMoney(n) {
  if (!n) return 'ศูนย์';
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

module.exports = { create };
