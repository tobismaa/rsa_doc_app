const CACHE_NAME = 'cmbank-rsa-v15';
const BADGE_DB_NAME = 'cmbank-badge-db';
const BADGE_STORE_NAME = 'appState';
const BADGE_COUNT_KEY = 'unreadCount';
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
  '/favicon.png',
  '/favicon.svg',
  '/icons/icon-192.png',
  '/icons/icon-512.png',
  '/manifest.webmanifest'
];

function openBadgeDb() {
  return new Promise((resolve, reject) => {
    const req = indexedDB.open(BADGE_DB_NAME, 1);
    req.onupgradeneeded = () => {
      const db = req.result;
      if (!db.objectStoreNames.contains(BADGE_STORE_NAME)) {
        db.createObjectStore(BADGE_STORE_NAME);
      }
    };
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  });
}

async function getBadgeCount() {
  try {
    const db = await openBadgeDb();
    return await new Promise((resolve) => {
      const tx = db.transaction(BADGE_STORE_NAME, 'readonly');
      const store = tx.objectStore(BADGE_STORE_NAME);
      const req = store.get(BADGE_COUNT_KEY);
      req.onsuccess = () => resolve(Number(req.result || 0) || 0);
      req.onerror = () => resolve(0);
    });
  } catch (_) {
    return 0;
  }
}

async function setBadgeCount(count) {
  const safeCount = Math.max(0, Number(count) || 0);
  try {
    const db = await openBadgeDb();
    await new Promise((resolve) => {
      const tx = db.transaction(BADGE_STORE_NAME, 'readwrite');
      tx.objectStore(BADGE_STORE_NAME).put(safeCount, BADGE_COUNT_KEY);
      tx.oncomplete = () => resolve();
      tx.onerror = () => resolve();
      tx.onabort = () => resolve();
    });
  } catch (_) {}
  try {
    if ('setAppBadge' in self.registration) {
      if (safeCount > 0) {
        await self.registration.setAppBadge(safeCount);
      } else if ('clearAppBadge' in self.registration) {
        await self.registration.clearAppBadge();
      }
    }
  } catch (_) {}
  const allClients = await clients.matchAll({ type: 'window', includeUncontrolled: true });
  allClients.forEach((client) => {
    client.postMessage({ type: 'app-unread-count', count: safeCount });
  });
}

async function incrementBadgeCount(step = 1) {
  const current = await getBadgeCount();
  await setBadgeCount(current + Math.max(1, Number(step) || 1));
}

// FCM background notifications support.
// Keep in service worker so push can display when app is minimized/background.
try {
  importScripts('https://www.gstatic.com/firebasejs/9.22.0/firebase-app-compat.js');
  importScripts('https://www.gstatic.com/firebasejs/9.22.0/firebase-messaging-compat.js');

  firebase.initializeApp({
    apiKey: "AIzaSyABQ-RR3Mlot7Vz2_s06AcFp3AlHb6elmw",
    authDomain: "rsa-doc-app.firebaseapp.com",
    projectId: "rsa-doc-app",
    storageBucket: "rsa-doc-app.firebasestorage.app",
    messagingSenderId: "749343098749",
    appId: "1:749343098749:web:ed78989a0b2c620d156e14",
    measurementId: "G-KQHMRNDZ6X"
  });

  const messaging = firebase.messaging();
  messaging.onBackgroundMessage((payload) => {
    const title = String(payload?.notification?.title || 'New Chat Message');
    const body = String(payload?.notification?.body || 'You have a new message');
    const clickUrl = String(payload?.data?.clickUrl || payload?.fcmOptions?.link || '/dashboard.html');
    incrementBadgeCount(1).catch(() => {});
    self.registration.showNotification(title, {
      body,
      icon: '/favicon.png?v=20260416f',
      badge: '/icons/icon-192.png?v=20260416f',
      requireInteraction: true,
      data: { url: clickUrl }
    });
  });
} catch (_) {}

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
  const requestUrl = new URL(event.request.url);
  if (requestUrl.origin !== self.location.origin) return;

  const isHtmlRequest =
    event.request.mode === 'navigate' ||
    requestUrl.pathname.endsWith('.html') ||
    requestUrl.pathname === '/' ||
    requestUrl.pathname === '';

  if (isHtmlRequest) {
    event.respondWith(
      fetch(event.request)
        .then((resp) => {
          const copy = resp.clone();
          caches.open(CACHE_NAME).then((cache) => cache.put(event.request, copy)).catch(() => {});
          return resp;
        })
        .catch(() => caches.match(event.request).then((cached) => cached || Response.error()))
    );
    return;
  }

  event.respondWith(
    caches.match(event.request).then((cached) => {
      if (cached) return cached;
      return fetch(event.request)
        .then((resp) => {
          const copy = resp.clone();
          caches.open(CACHE_NAME).then((cache) => cache.put(event.request, copy)).catch(() => {});
          return resp;
        })
        .catch(() => cached || Response.error());
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
