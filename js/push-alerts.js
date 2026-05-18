import { getMessaging, getToken, isSupported as isMessagingSupported } from 'https://www.gstatic.com/firebasejs/9.22.0/firebase-messaging.js';
import { arrayUnion } from 'https://www.gstatic.com/firebasejs/9.22.0/firebase-firestore.js';
import { app, db, doc, updateDoc, serverTimestamp } from './firebase-config.js';
import { EMAIL_API_BASE_URL, FCM_WEB_VAPID_KEY } from './email-api-config.js';
import { getSystemSettings } from './shared/system-settings.js?v=20260508a';

function getEmailApiBaseUrl() {
  const runtime = String(window.__EMAIL_API_BASE_URL__ || '').trim();
  const configured = runtime || String(EMAIL_API_BASE_URL || '').trim();
  if (!configured || configured.includes('YOUR-RENDER-URL')) return '';
  return configured.replace(/\/+$/, '');
}

function getServiceWorkerUrl() {
  const cacheBustToken = String(
    localStorage.getItem('cmbank_app_cache_bust_token') ||
    sessionStorage.getItem('cmbank_app_cache_bust_token') ||
    ''
  ).trim();
  return cacheBustToken
    ? `/service-worker.js?v=20260508a&clear=${encodeURIComponent(cacheBustToken)}`
    : '/service-worker.js?v=20260508a';
}

export async function registerPushTokenForCurrentUser(currentUser, profileDocId) {
  try {
    if (!currentUser || !profileDocId) return { ok: false, reason: 'missing-context' };
    const vapidKey = String(window.__FCM_VAPID_KEY__ || FCM_WEB_VAPID_KEY || '').trim();
    if (!vapidKey) return { ok: false, reason: 'missing-vapid-key' };
    if (!('serviceWorker' in navigator) || !('Notification' in window)) return { ok: false, reason: 'unsupported' };
    if (!(await isMessagingSupported())) return { ok: false, reason: 'messaging-unsupported' };
    if (Notification.permission === 'default') {
      try {
        await Notification.requestPermission();
      } catch (_) {}
    }
    if (Notification.permission !== 'granted') return { ok: false, reason: 'permission-not-granted' };

    const reg = await navigator.serviceWorker.register(getServiceWorkerUrl());
    const messaging = getMessaging(app);
    const token = await getToken(messaging, { vapidKey, serviceWorkerRegistration: reg });
    if (!token) return { ok: false, reason: 'empty-token' };

    await updateDoc(doc(db, 'users', profileDocId), {
      fcmTokens: arrayUnion(token),
      fcmLastTokenAt: serverTimestamp(),
      fcmLastTokenPlatform: navigator.userAgent || ''
    });
    return { ok: true, token };
  } catch (error) {
    return { ok: false, reason: String(error?.message || 'registration-failed') };
  }
}

export async function notifyAdminPushEvent({
  currentUser,
  eventType,
  title,
  body,
  clickUrl = '/admin-dashboard.html',
  meta = {}
}) {
  try {
    if (!currentUser || !eventType || !title || !body) return { ok: false, reason: 'missing-fields' };
    const settings = await getSystemSettings(db);
    if (!settings.notificationsPushEnabled) return { ok: false, reason: 'push-disabled' };
    const base = getEmailApiBaseUrl();
    if (!base) return { ok: false, reason: 'missing-api-base' };
    const idToken = await currentUser.getIdToken();
    const response = await fetch(`${base}/api/admin/push-event`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Authorization': `Bearer ${idToken}`
      },
      body: JSON.stringify({
        eventType: String(eventType || '').trim(),
        title: String(title || '').trim(),
        body: String(body || '').trim(),
        clickUrl: String(clickUrl || '/admin-dashboard.html').trim(),
        meta: meta && typeof meta === 'object' ? meta : {}
      })
    });
    return { ok: response.ok };
  } catch (_) {
    return { ok: false, reason: 'request-failed' };
  }
}
