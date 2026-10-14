/* FIWA Presence - Service Worker
   - Membuat aplikasi bisa dipasang di HP (PWA) dan terbuka lebih cepat.
   - Halaman & konfigurasi: ambil dari jaringan dulu, pakai salinan bila offline.
   - Aset aplikasi (nama ber-hash) dan pustaka CDN: disimpan di cache.
   - Data Supabase TIDAK pernah disimpan di cache (selalu langsung ke server). */
const VERSION = 'fiwa-v1';
const SHELL = ['/', '/manifest.webmanifest', '/icons/icon-192.png', '/icons/icon-512.png'];
const CDN = ['cdn.jsdelivr.net', 'cdnjs.cloudflare.com', 'fonts.googleapis.com', 'fonts.gstatic.com', 'i.ibb.co.com'];

self.addEventListener('install', e => {
  e.waitUntil(caches.open(VERSION).then(c => c.addAll(SHELL)).then(() => self.skipWaiting()));
});

self.addEventListener('activate', e => {
  e.waitUntil(caches.keys()
    .then(keys => Promise.all(keys.filter(k => k !== VERSION).map(k => caches.delete(k))))
    .then(() => self.clients.claim()));
});

const networkFirst = async (req, key) => {
  const cache = await caches.open(VERSION);
  try {
    const res = await fetch(req);
    if (res.ok) cache.put(key || req, res.clone());
    return res;
  } catch (err) {
    const hit = await cache.match(key || req);
    if (hit) return hit;
    throw err;
  }
};

const cacheFirst = async req => {
  const cache = await caches.open(VERSION);
  const hit = await cache.match(req);
  if (hit) return hit;
  const res = await fetch(req);
  if (res.ok || res.type === 'opaque') cache.put(req, res.clone());
  return res;
};

self.addEventListener('fetch', e => {
  const req = e.request;
  if (req.method !== 'GET') return;
  const url = new URL(req.url);

  if (url.origin === self.location.origin) {
    if (req.mode === 'navigate') return e.respondWith(networkFirst(req, '/'));
    if (url.pathname === '/config.js') return e.respondWith(networkFirst(req));
    if (url.pathname.startsWith('/fiwa/app.') || url.pathname.startsWith('/icons/')) return e.respondWith(cacheFirst(req));
    return;
  }
  if (CDN.includes(url.hostname)) e.respondWith(cacheFirst(req));
  // selain itu (termasuk *.supabase.co): biarkan browser mengambil langsung dari server
});
