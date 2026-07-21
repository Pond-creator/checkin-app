const CACHE = 'checkin-v3'; // bump ล้างแคชเก่าทั้งหมดทิ้ง

self.addEventListener('install', e => {
  self.skipWaiting(); // ไม่ precache อะไรอีก ตัดปัญหาไฟล์ HTML ค้าง
});

self.addEventListener('activate', e => {
  e.waitUntil(
    caches.keys().then(keys => Promise.all(keys.map(k => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', e => {
  // HTML (การนำทางหน้าเว็บ) → ดึงจากเน็ตสดเสมอ ห้ามใช้แคชเด็ดขาด
  // กันปัญหาผู้ใช้ค้างเวอร์ชันเก่าหลังอัปเดตโค้ด
  if (e.request.mode === 'navigate') {
    e.respondWith(fetch(e.request, { cache: 'no-store' }));
    return;
  }
  // ไฟล์อื่น (css/js/รูป) → network first, fallback แคชได้ถ้าออฟไลน์
  e.respondWith(
    fetch(e.request).catch(() => caches.match(e.request))
  );
});
