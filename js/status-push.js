import { EMAIL_API_BASE_URL } from './email-api-config.js';

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
  statusLabel
}) {
  try {
    if (!currentUser || !submissionId || !newStatus) return;
    const base = getEmailApiBaseUrl();
    if (!base) return;
    const idToken = await currentUser.getIdToken();
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
        statusLabel: String(statusLabel || '')
      })
    });
  } catch (_) {}
}

