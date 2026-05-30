// Service worker PWA cho KBC-HP89
// Network-first cho /static; trang dong khong cache de tranh stale.
const CACHE = 'kbc-hp89-v3';
const PRECACHE = [
  '/static/style.css',
  '/static/icon-192.png',
  '/static/icon-512.png',
  '/static/manifest.json'
];

self.addEventListener('install', (event) => {
  event.waitUntil(
    caches.open(CACHE).then((cache) => cache.addAll(PRECACHE)).catch(() => {})
  );
  self.skipWaiting();
});

self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys().then((keys) =>
      Promise.all(keys.filter((k) => k !== CACHE).map((k) => caches.delete(k)))
    )
  );
  self.clients.claim();
});

self.addEventListener('fetch', (event) => {
  const req = event.request;
  if (req.method !== 'GET') return;

  const url = new URL(req.url);
  // Chi xu ly tai nguyen tinh trong /static (network-first, fallback cache).
  if (url.pathname.startsWith('/static/')) {
    event.respondWith(
      fetch(req)
        .then((res) => {
          const copy = res.clone();
          caches.open(CACHE).then((c) => c.put(req, copy)).catch(() => {});
          return res;
        })
        .catch(() => caches.match(req))
    );
    return;
  }
  // Cac request khac (trang dong, API): de trinh duyet xu ly binh thuong.
});

// ----- Web Push: hien thong bao noi (kem rung) khi nhan duoc push -----
self.addEventListener('push', (event) => {
  let data = {};
  try { data = event.data ? event.data.json() : {}; } catch (e) { data = {}; }
  const title = data.title || 'KBC-HP89';
  const options = {
    body: data.body || '',
    icon: '/static/icon-192.png',
    badge: '/static/icon-192.png',
    vibrate: [200, 100, 200],
    tag: 'kbc-hp89-noti',
    renotify: true,
    data: { url: data.url || '/' }
  };
  event.waitUntil(
    self.registration.showNotification(title, options).then(function () {
      // Bao cho trang dang mo doc "KBC" bang giong noi (chi chay khi app dang mo)
      return self.clients.matchAll({ type: 'window', includeUncontrolled: true });
    }).then(function (wins) {
      wins.forEach(function (w) { w.postMessage({ type: 'noti-speak' }); });
    })
  );
});

// Bam vao thong bao -> mo app/trang lien quan (focus tab cu neu dang mo)
self.addEventListener('notificationclick', (event) => {
  event.notification.close();
  const target = (event.notification.data && event.notification.data.url) || '/';
  event.waitUntil(
    clients.matchAll({ type: 'window', includeUncontrolled: true }).then((wins) => {
      for (const w of wins) {
        if ('focus' in w) { w.navigate(target); return w.focus(); }
      }
      if (clients.openWindow) return clients.openWindow(target);
    })
  );
});
