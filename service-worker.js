const CACHE_NAME = 'cmbank-rsa-v2';
const ASSETS = [
  '/',
  '/index.html',
  '/dashboard.html',
  '/reviewer-dashboard.html',
  '/rsa-dashboard.html',
  '/admin-dashboard.html',
  '/payment-dashboard.html',
  '/super-admin-dashboard.html',
  '/css/index.css',
  '/css/uploader-dashboard.css',
  '/css/reviewer-dashboard.css',
  '/css/admin-dashboard.css',
  '/css/sop-help.css',
  '/css/developer-badge.css',
  '/css/profile-card.css',
  '/css/app-chat.css',
  '/js/auth.js',
  '/js/document-uploader.js',
  '/js/reviewer-dashboard.js',
  '/js/rsa-dashboard.js',
  '/js/admin.js',
  '/js/payment-dashboard.js',
  '/js/super-admin.js',
  '/js/pwa-install.js',
  '/js/app-chat.js',
  '/favicon.svg',
  '/manifest.webmanifest'
];

self.addEventListener('install', (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME).then((cache) => cache.addAll(ASSETS)).catch(() => Promise.resolve())
  );
  self.skipWaiting();
});

self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys().then((keys) => Promise.all(
      keys.filter((k) => k !== CACHE_NAME).map((k) => caches.delete(k))
    ))
  );
  self.clients.claim();
});

self.addEventListener('fetch', (event) => {
  if (event.request.method !== 'GET') return;
  event.respondWith(
    caches.match(event.request).then((cached) => {
      if (cached) return cached;
      return fetch(event.request)
        .then((resp) => {
          const copy = resp.clone();
          caches.open(CACHE_NAME).then((cache) => cache.put(event.request, copy)).catch(() => {});
          return resp;
        })
        .catch(() => cached);
    })
  );
});

self.addEventListener('notificationclick', (event) => {
  event.notification.close();
  const targetUrl = String(
    event.notification?.data?.url ||
    event.notification?.data?.link ||
    '/dashboard.html'
  );

  event.waitUntil((async () => {
    const allClients = await clients.matchAll({ type: 'window', includeUncontrolled: true });
    const sameClient = allClients.find((c) => {
      try { return c.url && c.url.startsWith(self.location.origin); } catch (_) { return false; }
    });
    if (sameClient) {
      await sameClient.focus();
      sameClient.postMessage({ type: 'open-chat-url', url: targetUrl });
      return;
    }
    await clients.openWindow(targetUrl);
  })());
});
