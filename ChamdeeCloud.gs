/**
 * ชำดี POS Cloud — ตัวเชื่อม Google ของร้าน (ฟรี)
 *  1) ซิงก์ข้อมูลร้านระหว่างมือถือ/แท็บเล็ตหลายเครื่อง (เก็บใน Google Sheet ของคุณเอง)
 *  2) เชื่อมระบบสมาชิก Lucky Points (ชีต "สมาชิก")
 *
 * ติดตั้งครั้งแรก
 *  1. script.google.com → โครงการใหม่ → ตั้งชื่อ "ChamdeeCloud" → ลบโค้ดเดิม วางไฟล์นี้ → บันทึก
 *  2. เลือกฟังก์ชัน setup → เรียกใช้ → อนุญาตสิทธิ์ → ดู "รหัส POS" ในบันทึกการดำเนินการ
 *  3. ทำให้ใช้งานได้ → การทำให้ใช้งานได้รายการใหม่ → เว็บแอป
 *       เรียกใช้ในฐานะ: ฉัน · ผู้มีสิทธิ์เข้าถึง: ทุกคน → คัดลอก URL (/exec)
 *  4. แอปชำดี: การเชื่อมต่อ → วาง URL + รหัส POS → บันทึกและทดสอบ
 *
 * เคยติดตั้งตัวเชื่อมเวอร์ชันเก่า (POS Bridge) แล้ว
 *  วางโค้ดนี้ทับ → บันทึก → เรียกใช้ setup อีกครั้ง →
 *  ทำให้ใช้งานได้ → จัดการการทำให้ใช้งานได้ → ✎ แก้ไข → เวอร์ชัน: เวอร์ชันใหม่ → ทำให้ใช้งานได้ (URL เดิม)
 */
const MEMBER_SHEET_ID = '1K2LTSrxoprjbm1zKAF5lMMWgzoA1f4rcqQR_lt5aSJs'; // ชีต "สมาชิก" (Lucky Points)
const TZ = 'Asia/Bangkok';
const CHUNK = 40000; // ตัวอักษรต่อช่อง (Google Sheet รับได้ 50,000)

function props_() { return PropertiesService.getScriptProperties(); }

function setup() {
  const p = props_();
  let k = p.getProperty('POS_KEY');
  if (!k) { k = Utilities.getUuid().replace(/-/g, '').slice(0, 16); p.setProperty('POS_KEY', k); }
  let ss = null, id = p.getProperty('DATA_ID');
  if (id) { try { ss = SpreadsheetApp.openById(id); } catch (e) { ss = null; } }
  if (!ss) { ss = SpreadsheetApp.create('ชำดี POS – ข้อมูลร้าน (ห้ามแก้ไขเอง)'); p.setProperty('DATA_ID', ss.getId()); }
  ensureSheets_(ss);
  try { SpreadsheetApp.openById(MEMBER_SHEET_ID).getSheetByName('Purchases'); } catch (e) { Logger.log('ยังเปิดชีตสมาชิกไม่ได้ (ข้ามได้ถ้าไม่ใช้ Lucky Points): ' + e); }
  Logger.log('รหัส POS ของคุณ: ' + k);
  Logger.log('ชีตเก็บข้อมูลร้าน: ' + ss.getUrl());
}

function ensureSheets_(ss) {
  let o = ss.getSheetByName('Ops');
  if (!o) { o = ss.insertSheet('Ops'); o.appendRow(['id', 'time', 'device', 'type', 'parts', 'data']); }
  let s = ss.getSheetByName('Snapshot');
  if (!s) { s = ss.insertSheet('Snapshot'); s.appendRow(['chunk']); }
  ss.getSheets().forEach(sh => { const n = sh.getName(); if (n !== 'Ops' && n !== 'Snapshot' && sh.getLastRow() === 0 && ss.getSheets().length > 2) ss.deleteSheet(sh); });
}
function data_() { const ss = SpreadsheetApp.openById(props_().getProperty('DATA_ID')); ensureSheets_(ss); return ss; }

function doGet(e) {
  const p = (e && e.parameter) || {};
  if (p.page === 'admin') return HtmlService.createHtmlOutput(ADMIN_HTML).setTitle('จัดการคะแนน Lucky Points').addMetaTag('viewport', 'width=device-width, initial-scale=1');
  if (!p.action && !p.key) { // มีคนเปิดลิงก์นี้ตรงๆ (เช่น สแกน QR) — ไม่ใช่หน้าสมัครสมาชิก
    return HtmlService.createHtmlOutput('<meta name="viewport" content="width=device-width,initial-scale=1"><div style="font-family:sans-serif;padding:24px;line-height:1.6"><h2>ตัวเชื่อม ชำดี POS</h2><p>ลิงก์นี้ใช้เชื่อมแอปขายของกับ Google Sheet ของร้านเท่านั้น<br>ไม่ใช่หน้าสมัครสมาชิก</p><p>เจ้าของร้าน: ในแอปไปที่ การเชื่อมต่อ › ช่อง "ลิงก์สมัครสมาชิก" ใส่ลิงก์หน้าสมัครของ Lucky Points แทน</p></div>').setTitle('ชำดี POS');
  }
  return handle_(p);
}
/* หน้าจัดการคะแนน (เปิด URL ตัวเชื่อม + ?page=admin) — ต้องใส่รหัส POS ก่อนใช้งาน */
function adminApi(p) { try { return JSON.parse(handle_(p || {}).getContent()); } catch (e) { return { ok: false, error: String(e) }; } }
const ADMIN_HTML = `<!doctype html><html lang="th"><head><meta charset="utf-8"><style>
*{box-sizing:border-box}body{margin:0;background:#140c2c;color:#f3eeff;font:16px/1.5 system-ui,"Noto Sans Thai",sans-serif}
.wrap{max-width:720px;margin:0 auto;padding:16px}h1{font-size:1.35rem;margin:6px 0 2px}.mut{color:#b9a8e6;font-size:.9rem}
.box{background:#22174a;border:1px solid #3d2d7a;border-radius:18px;padding:14px;margin-top:12px}
input,select,button{font:inherit;border-radius:12px;border:1px solid #4a3a8c;background:#1a1240;color:#f3eeff;padding:10px 12px}
button{cursor:pointer;background:#2f2366}button.pri{background:linear-gradient(180deg,#ff63b2,#e8157e);border:0;color:#fff;font-weight:700}
button.sm{padding:7px 12px;font-size:.92rem}button:disabled{opacity:.45}.row{display:flex;gap:8px;flex-wrap:wrap;align-items:center}.grow{flex:1;min-width:140px}
.m{background:#2a1e58;border-radius:16px;padding:12px 14px;margin-top:10px}.m b.n{font-size:1.1rem}.big{font-size:1.05rem}.t{color:#ffd23f;font-weight:700}
.edit{display:none;margin-top:10px;gap:8px;flex-direction:column}
button.add{background:#1f6b4a;border-color:#2e9a6b;color:#eafff4;font-weight:700}button.sub{background:#7a2440;border-color:#b23a60;color:#ffeaf1;font-weight:700}
.lbl{font-weight:700;margin-top:4px}.lbl.add{color:#7ff0b5}.lbl.sub{color:#ff9cbc}
.pend{margin-top:10px;background:#3a2a10;border:1px solid #ffd23f;border-radius:14px;padding:10px 12px;display:flex;flex-direction:column;gap:8px}
.up{color:#7ff0b5;font-weight:700}.dn{color:#ff9cbc;font-weight:700}.m.open .edit{display:flex}
#toast{position:fixed;left:12px;right:12px;bottom:16px;background:#fff;color:#222;border-radius:12px;padding:12px;text-align:center;display:none;font-weight:600}
</style></head><body><div class="wrap">
<h1>จัดการคะแนน / สิทธิ์ Lucky Points</h1><div class="mut">เพิ่มสิทธิ์ ดูและแก้คะแนนสมาชิก · ทุกการแก้ไขบันทึกเป็นรายการ "ปรับคะแนน" ในชีต Purchases</div>
<div class="box" id="login"><div class="row"><input id="key" class="grow" placeholder="รหัส POS (ดูในแอป › การเชื่อมต่อ)" autocomplete="off"><button class="pri" onclick="login()">เข้าใช้งาน</button></div><div class="mut" id="lmsg"></div></div>
<div id="app" style="display:none"><div class="box"><div class="row"><input id="q" class="grow" placeholder="ค้นหาชื่อ แผนก ที่ทำงาน เบอร์" oninput="draw()"><select id="mon" onchange="load()"></select></div><div class="mut" id="sum" style="margin-top:8px"></div></div><div id="list"></div></div>
</div><div id="toast"></div><script>
let KEY='',D=null,open='';const $=s=>document.querySelector(s);
const esc=s=>String(s==null?'':s).replace(/[&<>"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]));
function toast(t){const e=$('#toast');e.textContent=t;e.style.display='block';clearTimeout(e._t);e._t=setTimeout(()=>e.style.display='none',2600)}
function api(p){return new Promise((res,rej)=>google.script.run.withSuccessHandler(j=>j&&j.ok?res(j):rej(new Error(j&&j.error||'ผิดพลาด'))).withFailureHandler(rej).adminApi(Object.assign({key:KEY},p)))}
try{KEY=localStorage.getItem('lpk')||''}catch(e){}
function months(){const s=$('#mon'),d=new Date(),th=['ม.ค.','ก.พ.','มี.ค.','เม.ย.','พ.ค.','มิ.ย.','ก.ค.','ส.ค.','ก.ย.','ต.ค.','พ.ย.','ธ.ค.'];let h='';for(let i=0;i<12;i++){const x=new Date(d.getFullYear(),d.getMonth()-i,1);const v=x.getFullYear()+'-'+String(x.getMonth()+1).padStart(2,'0');h+='<option value="'+v+'">'+th[x.getMonth()]+' '+(x.getFullYear()+543)+(i?'':' (เดือนนี้)')+'</option>'}s.innerHTML=h}
async function login(){KEY=($('#key').value||KEY).trim();$('#lmsg').textContent='กำลังตรวจรหัส…';try{await api({action:'ping'});try{localStorage.setItem('lpk',KEY)}catch(e){}$('#login').style.display='none';$('#app').style.display='block';months();load()}catch(e){$('#lmsg').textContent='รหัสไม่ถูกต้อง: '+e.message}}
async function load(){$('#sum').textContent='กำลังโหลด…';try{D=await api({action:'members',month:$('#mon').value});draw()}catch(e){$('#sum').textContent=e.message}}
const P={};
function tk(n){return Math.floor(n/D.pointsPerTicket)}
function fmtN(n){return Number(n).toLocaleString()}
function draw(){if(!D)return;const q=$('#q').value.trim().toLowerCase(),ppt=D.pointsPerTicket,cur=D.month===D.current;
 const L=D.members.filter(m=>!q||[m.nickname,m.department,m.workplace,m.contact].join(' ').toLowerCase().includes(q)).sort((a,b)=>b.points-a.points);
 const tp=D.members.reduce((a,m)=>a+m.points,0),tt=D.members.reduce((a,m)=>a+tk(m.points),0);
 $('#sum').textContent=D.members.length+' คน · '+fmtN(tp)+' คะแนน · '+tt+' สิทธิ์ · '+ppt+' คะแนน = 1 สิทธิ์'+(cur?'':' · ดูย้อนหลังอย่างเดียว');
 $('#list').innerHTML=L.map(m=>{const id=esc(m.id),d=P[m.id]||0,nw=Math.max(0,m.points+d);
  let h='<div class="m'+(open===m.id?' open':'')+'" data-id="'+id+'"><b class="n">'+esc(m.nickname)+'</b><div class="mut">'+esc([m.department,m.workplace,m.contact].filter(Boolean).join(' · '))+'</div><div class="big">'+fmtN(m.points)+' คะแนน = <span class="t">'+tk(m.points)+' สิทธิ์</span></div>';
  if(cur){h+='<div class="row" style="margin-top:8px"><button class="sm add" data-a="t+">+1 สิทธิ์</button><button class="sm sub" data-a="t-">−1 สิทธิ์</button><button class="sm" data-a="tog">'+(open===m.id?'ปิด ▴':'แก้คะแนน ▾')+'</button></div>'+
   '<div class="edit"><div class="lbl add">➕ เพิ่มคะแนน</div><div class="row"><input class="grow" type="number" inputmode="numeric" min="1" data-f="add" placeholder="เช่น 50"><button class="sm add" data-a="add">เพิ่ม</button></div>'+
   '<div class="lbl sub">➖ ลดคะแนน</div><div class="row"><input class="grow" type="number" inputmode="numeric" min="1" data-f="sub" placeholder="เช่น 50"><button class="sm sub" data-a="sub">ลด</button></div>'+
   '<div class="lbl">✏️ ตั้งคะแนนใหม่ทั้งหมด</div><div class="row"><input class="grow" type="number" inputmode="numeric" min="0" data-f="set" placeholder="คะแนนที่ถูกต้อง"><button class="sm" data-a="set">ตั้งค่า</button></div></div>';
   if(d){h+='<div class="pend"><div>จะเปลี่ยนเป็น <b>'+fmtN(nw)+' คะแนน = '+tk(nw)+' สิทธิ์</b> <span class="'+(d>0?'up':'dn')+'">('+(d>0?'+':'')+fmtN(d)+')</span></div><input data-f="note" placeholder="เหตุผล (ไม่บังคับ) เช่น ลืมสแกน"><div class="row"><button class="pri grow" data-a="save">💾 บันทึก</button><button data-a="cancel">ยกเลิก</button></div></div>'}}
  return h+'</div>'}).join('')||'<div class="m mut">ไม่พบสมาชิก</div>'}
function stage(id,delta){const m=D.members.find(x=>x.id===id);let d=(P[id]||0)+delta;if(m.points+d<0)d=-m.points;P[id]=d;if(!d)delete P[id];draw();const el=document.querySelector('.m[data-id="'+CSS.escape(id)+'"] .pend');if(el)el.scrollIntoView({block:'nearest',behavior:'smooth'})}
document.addEventListener('click',async e=>{const b=e.target.closest('[data-a]');if(!b)return;const card=b.closest('.m');const id=card&&card.dataset.id;if(!id)return;const m=D.members.find(x=>x.id===id),ppt=D.pointsPerTicket,a=b.dataset.a;
 const val=f=>Math.floor(+((card.querySelector('[data-f="'+f+'"]')||{}).value||0));
 if(a==='tog'){open=open===id?'':id;draw();return}
 if(a==='t+')return stage(id,ppt);if(a==='t-')return stage(id,-ppt);
 if(a==='add'){const n=val('add');if(n<=0)return toast('ใส่จำนวนที่จะเพิ่ม');return stage(id,n)}
 if(a==='sub'){const n=val('sub');if(n<=0)return toast('ใส่จำนวนที่จะลด');return stage(id,-n)}
 if(a==='set'){const raw=(card.querySelector('[data-f="set"]')||{}).value;if(raw==='')return toast('ใส่คะแนนที่ถูกต้อง');P[id]=0;return stage(id,Math.max(0,Math.floor(+raw))-m.points)}
 if(a==='cancel'){delete P[id];draw();return}
 if(a==='save'){const d=P[id]||0;if(!d)return;const target=Math.max(0,m.points+d);const note=((card.querySelector('[data-f="note"]')||{}).value||'').trim()||((d>0?'เพิ่ม ':'ลด ')+Math.abs(d)+' คะแนน');
  b.disabled=true;b.textContent='กำลังบันทึก…';try{const j=await api({action:'adjust',memberId:id,target:target,note:note});m.points=j.points;delete P[id];open='';draw();toast('บันทึกแล้ว ✓ '+m.nickname+' '+fmtN(j.points)+' คะแนน = '+tk(j.points)+' สิทธิ์')}catch(err){b.disabled=false;b.textContent='💾 บันทึก';toast('ไม่สำเร็จ: '+err.message)}}});
if(KEY){$('#key').value=KEY;login()}
</script></body></html>`;
function doPost(e) { let p = {}; try { p = JSON.parse(e.postData.contents); } catch (x) {} return handle_(p); }
function out_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

/* ---------- Lucky Points helpers ---------- */
function msheet_(name) { return SpreadsheetApp.openById(MEMBER_SHEET_ID).getSheetByName(name); }
function rows_(name) {
  const v = msheet_(name).getDataRange().getValues();
  const h = v.shift().map(String);
  return v.map(r => { const o = {}; h.forEach((k, i) => o[k] = r[i]); return o; });
}
function fmt_(d, f) { return Utilities.formatDate(d, TZ, f); }
function thisMonth_() { return fmt_(new Date(), 'yyyy-MM'); }
function monthOf_(v) { return v instanceof Date ? fmt_(v, 'yyyy-MM') : String(v).slice(0, 7); }
function now_() { return fmt_(new Date(), 'yyyy-MM-dd HH:mm'); }
function settings_() {
  const o = {}; rows_('Settings').forEach(r => o[r.key] = r.value);
  return { shopName: String(o.shopName || ''), pointsPerTicket: Number(o.pointsPerTicket) || 300 };
}
function monthPoints_(month) {
  const m = month || thisMonth_(), pts = {};
  rows_('Purchases').forEach(r => {
    if (String(r.status) === 'approved' && monthOf_(r.month) === m) pts[r.memberId] = (pts[r.memberId] || 0) + (Number(r.amount) || 0);
  });
  return pts;
}
function withLock_(fn) { const l = LockService.getScriptLock(); l.waitLock(20000); try { return fn(); } finally { l.releaseLock(); } }

function handle_(p) {
  const key = props_().getProperty('POS_KEY');
  if (!key || String(p.key) !== key) return out_({ ok: false, error: 'รหัส POS ไม่ถูกต้อง' });
  try {
    switch (p.action) {
      /* ----- common ----- */
      case 'ping': {
        let lucky = null; try { lucky = settings_(); } catch (e) {}
        return out_({ ok: true, shopName: lucky ? lucky.shopName : '', pointsPerTicket: lucky ? lucky.pointsPerTicket : 300, month: thisMonth_(), lucky: !!lucky, sync: true, epoch: props_().getProperty('EPOCH') || '' });
      }
      /* ----- multi-device sync ----- */
      case 'register': return withLock_(() => {
        const n = Number(props_().getProperty('DEV_COUNT') || 0) + 1;
        props_().setProperty('DEV_COUNT', String(n));
        return out_({ ok: true, dev: n });
      });
      case 'seed': return withLock_(() => {
        const ss = data_(), snap = ss.getSheetByName('Snapshot'), ops = ss.getSheetByName('Ops');
        const txt = String(p.snapshot || '');
        if (!txt) return out_({ ok: false, error: 'ไม่มีข้อมูล' });
        snap.clear();
        const rows = [['chunk']];
        for (let i = 0; i < txt.length; i += CHUNK) rows.push([txt.slice(i, i + CHUNK)]);
        snap.getRange(1, 1, rows.length, 1).setNumberFormat('@').setValues(rows);
        ops.clear(); ops.appendRow(['id', 'time', 'device', 'type', 'parts', 'data']);
        const ep = Utilities.getUuid().slice(0, 8);
        props_().setProperty('EPOCH', ep); props_().setProperty('SNAP_CURSOR', '0');
        return out_({ ok: true, epoch: ep, cursor: 0 });
      });
      case 'snapshot': {
        const s = data_().getSheetByName('Snapshot'), n = s.getLastRow();
        const txt = n > 1 ? s.getRange(2, 1, n - 1, 1).getValues().map(r => String(r[0])).join('') : '';
        return out_({ ok: true, epoch: props_().getProperty('EPOCH') || '', cursor: Number(props_().getProperty('SNAP_CURSOR') || 0), data: txt });
      }
      case 'push': {
        const ep = props_().getProperty('EPOCH') || '';
        if (!ep || String(p.epoch) !== ep) return out_({ ok: false, error: 'epoch', epoch: ep });
        const list = Array.isArray(p.ops) ? p.ops : [];
        return withLock_(() => {
          const sh = data_().getSheetByName('Ops'), last = sh.getLastRow();
          const from = Math.max(2, last - 3000);
          const seen = new Set(last >= 2 ? sh.getRange(from, 1, last - from + 1, 1).getValues().map(r => String(r[0])) : []);
          const rows = []; let w = 6;
          list.forEach(o => {
            if (!o || seen.has(String(o.id))) return; seen.add(String(o.id));
            const d = typeof o.data === 'string' ? o.data : JSON.stringify(o.data);
            const parts = []; for (let i = 0; i < d.length; i += CHUNK) parts.push(d.slice(i, i + CHUNK));
            if (!parts.length) parts.push('');
            const r = [String(o.id), String(o.t), String(o.dev), String(o.type), String(parts.length)].concat(parts);
            rows.push(r); w = Math.max(w, r.length);
          });
          if (rows.length) {
            rows.forEach(r => { while (r.length < w) r.push(''); });
            if (sh.getMaxColumns() < w) sh.insertColumnsAfter(sh.getMaxColumns(), w - sh.getMaxColumns());
            const rg = sh.getRange(last + 1, 1, rows.length, w); rg.setNumberFormat('@'); rg.setValues(rows);
          }
          return out_({ ok: true, added: rows.length, total: sh.getLastRow() - 1 });
        });
      }
      case 'pull': {
        const ep = props_().getProperty('EPOCH') || '';
        if (String(p.epoch) !== ep) return out_({ ok: true, reset: true, epoch: ep });
        const sh = data_().getSheetByName('Ops'), total = Math.max(0, sh.getLastRow() - 1);
        const since = Math.max(0, Number(p.since) || 0), n = Math.min(300, total - since);
        if (n <= 0) return out_({ ok: true, ops: [], next: since, total: total });
        const v = sh.getRange(since + 2, 1, n, sh.getLastColumn()).getValues();
        const ops = v.map(r => { const k = Number(r[4]) || 1; return { id: String(r[0]), t: Number(r[1]), dev: String(r[2]), type: String(r[3]), data: r.slice(5, 5 + k).map(String).join('') }; });
        return out_({ ok: true, ops: ops, next: since + n, total: total });
      }
      /* ----- Lucky Points ----- */
      case 'members': {
        const mon = /^\d{4}-\d{2}$/.test(String(p.month || '')) ? String(p.month) : thisMonth_();
        const s = settings_(), pts = monthPoints_(mon);
        const members = rows_('Members').filter(r => String(r.status) === 'active').map(r => ({
          id: String(r.id), nickname: String(r.nickname), department: String(r.department || ''),
          workplace: String(r.workplace || ''), contact: String(r.contact || ''),
          points: Math.floor(pts[String(r.id)] || 0), createdAt: String(r.createdAt || '')
        }));
        return out_({ ok: true, month: mon, current: thisMonth_(), pointsPerTicket: s.pointsPerTicket, members: members });
      }
      case 'purchase': return withLock_(() => {
        const sh = msheet_('Purchases'), last = sh.getLastRow();
        const ids = last > 1 ? sh.getRange(2, 1, last - 1, 1).getValues().map(r => String(r[0])) : [];
        if (ids.indexOf(String(p.id)) >= 0) return out_({ ok: true, duplicate: true });
        const member = rows_('Members').find(r => String(r.id) === String(p.memberId));
        if (!member) return out_({ ok: false, error: 'ไม่พบสมาชิก' });
        const amount = Math.floor(Number(p.amount) || 0), row = last + 1, t = now_();
        sh.getRange(row, 1, 1, 2).setNumberFormat('@'); sh.getRange(row, 4, 1, 6).setNumberFormat('@');
        sh.getRange(row, 1, 1, 9).setValues([[String(p.id), String(p.memberId), amount, String(p.note || ''), 'approved', thisMonth_(), t, t, 'pos']]);
        return out_({ ok: true, points: Math.floor(monthPoints_()[String(p.memberId)] || 0), pointsPerTicket: settings_().pointsPerTicket });
      });
      case 'adjust': return withLock_(() => { // เพิ่ม/ลด/ตั้งคะแนนเดือนนี้ (บันทึกเป็นรายการปรับคะแนน)
        const member = rows_('Members').find(r => String(r.id) === String(p.memberId));
        if (!member) return out_({ ok: false, error: 'ไม่พบสมาชิก' });
        const cur = Math.floor(monthPoints_()[String(p.memberId)] || 0);
        let amount = p.target !== undefined && p.target !== '' ? Math.floor(Number(p.target)) - cur : Math.floor(Number(p.amount) || 0);
        if (!isFinite(amount)) return out_({ ok: false, error: 'จำนวนไม่ถูกต้อง' });
        if (cur + amount < 0) amount = -cur;
        if (amount !== 0) {
          const sh = msheet_('Purchases'), row = sh.getLastRow() + 1, t = now_();
          const id = 'adj' + Utilities.getUuid().replace(/-/g, '').slice(0, 9);
          sh.getRange(row, 1, 1, 2).setNumberFormat('@'); sh.getRange(row, 4, 1, 6).setNumberFormat('@');
          sh.getRange(row, 1, 1, 9).setValues([[id, String(p.memberId), amount, String(p.note || 'ปรับคะแนน'), 'approved', thisMonth_(), t, t, 'pos-adjust']]);
        }
        return out_({ ok: true, points: cur + amount, changed: amount, pointsPerTicket: settings_().pointsPerTicket });
      });
      case 'void': {
        const sh = msheet_('Purchases'), v = sh.getDataRange().getValues();
        for (let i = 1; i < v.length; i++) if (String(v[i][0]) === String(p.id)) {
          sh.getRange(i + 1, 5).setValue('rejected'); sh.getRange(i + 1, 8).setValue(now_()); sh.getRange(i + 1, 9).setValue('pos-refund');
          return out_({ ok: true });
        }
        return out_({ ok: true, notFound: true });
      }
      default: return out_({ ok: false, error: 'unknown action' });
    }
  } catch (err) { return out_({ ok: false, error: String(err) }); }
}
