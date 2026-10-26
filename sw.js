// ชำดี POS — เปิดได้แม้เน็ตหลุด · หน้าแอปดึงเวอร์ชันใหม่จากเว็บก่อนเสมอ (ข้อมูลร้านอยู่ใน IndexedDB ไม่ได้อยู่ในแคชนี้)
const CACHE='chamdee-pos-v7-3';
const CORE=['./','./index.html','./manifest.webmanifest','./icon-192.png','./icon-512.png'];
self.addEventListener('install',e=>{e.waitUntil(caches.open(CACHE).then(c=>Promise.all(CORE.map(u=>fetch(new Request(u,{cache:'reload'})).then(r=>r.ok&&c.put(u,r)).catch(()=>{})))));self.skipWaiting()});
self.addEventListener('activate',e=>{e.waitUntil(caches.keys().then(ks=>Promise.all(ks.filter(k=>k!==CACHE).map(k=>caches.delete(k)))).then(()=>self.clients.claim()))});
self.addEventListener('fetch',e=>{
  const req=e.request;if(req.method!=='GET')return;const u=new URL(req.url);
  if(/script\.google|googleusercontent|peerjs/.test(u.hostname))return; // ระบบสมาชิก/จอลูกค้า ต้องออนไลน์เสมอ
  const isPage=req.mode==='navigate'||(u.origin===location.origin&&/\/(index\.html)?$/.test(u.pathname));
  if(isPage){ // network-first: ได้หน้าแอปล่าสุดทุกครั้งที่มีเน็ต
    e.respondWith(fetch(req,{cache:'no-store'}).then(r=>{if(r&&r.ok){const cp=r.clone();caches.open(CACHE).then(c=>c.put('./index.html',cp))}return r})
      .catch(()=>caches.match('./index.html').then(h=>h||caches.match('./'))));return}
  e.respondWith(caches.match(req).then(hit=>{
    const net=fetch(req).then(r=>{if(r&&(r.ok||r.type==='opaque')){const cp=r.clone();caches.open(CACHE).then(c=>c.put(req,cp))}return r}).catch(()=>hit);
    return hit||net}));
});
