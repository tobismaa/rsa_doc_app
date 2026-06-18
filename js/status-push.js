import { EMAIL_API_BASE_URL } from './email-api-config.js';
import { auth, db } from './firebase-config.js';
import { getSystemSettings } from './shared/system-settings.js?v=20260617a';

function getEmailApiBaseUrl() {
  const runtime = String(window.__EMAIL_API_BASE_URL__ || '').trim();
  const configured = runtime || String(EMAIL_API_BASE_URL || '').trim();
  if (!configured || configured.includes('YOUR-RENDER-URL')) return '';
  return configured.replace(/\/+$/, '');
}

export async function notifyStatusChangePush({
  currentUser,
  submissionId,
  customerName,
  newStatus,
  statusLabel,
  actionLabel,
  message
}) {
  try {
    if (!currentUser || !submissionId || !newStatus) return;
    const settings = await getSystemSettings(db);
    if (!settings.notificationsPushEnabled) return;
    const base = getEmailApiBaseUrl();
    if (!base) return;
    const sourceUser = currentUser || auth.currentUser;
    if (!sourceUser) return;
    const idToken = await sourceUser.getIdToken();
    await fetch(`${base}/api/submission/status-push`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Authorization': `Bearer ${idToken}`
      },
      body: JSON.stringify({
        submissionId: String(submissionId),
        customerName: String(customerName || ''),
        newStatus: String(newStatus || '').toLowerCase(),
        statusLabel: String(statusLabel || ''),
        actionLabel: String(actionLabel || ''),
        message: String(message || '')
      })
    });
  } catch (_) {}
}
