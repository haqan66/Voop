// Çevrimdışı çalışma için uygulama dosyalarını önbelleğe alır.
const CACHE = 'barkod-kontrol-v7';
const ASSETS = [
  './',
  'index.html',
  'app.css',
  'app.js',
  'manifest.webmanifest',
  'vendor/xlsx.full.min.js',
  'vendor/html5-qrcode.min.js',
  'icons/icon.svg',
  'icons/icon-192.png',
  'icons/icon-512.png',
  'data/urun-listesi.xlsx',
  'data/urun-listesi.js',
];

self.addEventListener('install', (e) => {
  e.waitUntil(caches.open(CACHE).then((c) => c.addAll(ASSETS)).then(() => self.skipWaiting()));
});

self.addEventListener('activate', (e) => {
  e.waitUntil(
    caches.keys()
      .then((keys) => Promise.all(keys.filter((k) => k !== CACHE).map((k) => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});

// Önce ağ, olmazsa önbellek: güncellemeler hemen gelir, internet yoksa uygulama yine açılır.
self.addEventListener('fetch', (e) => {
  const url = new URL(e.request.url);
  // ortak kayıtlar (api.php) hiçbir zaman önbellekten verilmez
  if (e.request.method !== 'GET' || url.origin !== location.origin || url.pathname.endsWith('/api.php')) return;
  e.respondWith(
    fetch(e.request)
      .then((res) => {
        if (res.ok) { const copy = res.clone(); caches.open(CACHE).then((c) => c.put(e.request, copy)); }
        return res;
      })
      .catch(() => caches.match(e.request, { ignoreSearch: true }).then((r) => r || caches.match('index.html')))
  );
});
