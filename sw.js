const CACHE = 'checkin-v2'; // bump เพื่อบังคับล้างแคชเก่า (แก้ปัญหาปุ่มค้างจากไฟล์เก่า)
const ASSETS = [
  '/checkin-app/checkin.html',
  '/checkin-app/index.html'
];

self.addEventListener('install', e => {
  e.waitUntil(caches.open(CACHE).then(c => c.addAll(ASSETS)).then(() => self.skipWaiting()));
});

self.addEventListener('activate', e => {
  e.waitUntil(caches.keys().then(keys =>
    Promise.all(keys.filter(k => k !== CACHE).map(k => caches.delete(k)))
  ).then(() => self.clients.claim()));
});

self.addEventListener('fetch', e => {
  // Network first — always try network, fallback to cache
  e.respondWith(
    fetch(e.request).catch(() => caches.match(e.request))
  );
});
