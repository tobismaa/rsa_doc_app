import { app, auth, db } from './firebase-config.js';
import { EMAIL_API_BASE_URL, FCM_WEB_VAPID_KEY } from './email-api-config.js';
import {
  onAuthStateChanged
} from 'https://www.gstatic.com/firebasejs/9.22.0/firebase-auth.js';
import {
  collection,
  doc,
  getDoc,
  getDocs,
  query,
  where,
  orderBy,
  limit,
  onSnapshot,
  addDoc,
  setDoc,
  updateDoc,
  arrayUnion,
  serverTimestamp
} from 'https://www.gstatic.com/firebasejs/9.22.0/firebase-firestore.js';
import {
  getMessaging,
  getToken,
  isSupported as isMessagingSupported
} from 'https://www.gstatic.com/firebasejs/9.22.0/firebase-messaging.js';

let currentUser = null;
let currentProfile = null;
let currentProfileDocId = '';
let currentSubmission = null;
let currentChatId = '';
let unsubMessages = null;
let unsubSubmission = null;
let unsubGlobalChats = null;
const messageSeenMap = new Map();
let globalNoticeTimer = null;
const globalChatLastSeen = new Map();
const FCM_VAPID_KEY = String(window.__FCM_VAPID_KEY__ || FCM_WEB_VAPID_KEY || '').trim();
let pendingChatFromUrl = '';
let pushTokenState = { status: 'idle', detail: 'Push: Not checked' };

function esc(text) {
  const div = document.createElement('div');
  div.textContent = String(text ?? '');
  return div.innerHTML;
}

function fmtDate(v) {
  if (!v) return '-';
  try {
    const d = v.toDate ? v.toDate() : new Date(v);
    if (Number.isNaN(d.getTime())) return '-';
    return d.toLocaleString();
  } catch (_) {
    return '-';
  }
}

function normalizeRole(role) {
  const r = String(role || '').toLowerCase().trim();
  if (r === 'viewer') return 'reviewer';
  return r;
}

function getEmailApiBaseUrl() {
  const runtime = String(window.__EMAIL_API_BASE_URL__ || '').trim();
  const configured = runtime || String(EMAIL_API_BASE_URL || '').trim();
  if (!configured || configured.includes('YOUR-RENDER-URL')) return '';
  return configured.replace(/\/+$/, '');
}

function openChatFromUrl(urlLike) {
  try {
    const u = new URL(String(urlLike || ''), window.location.origin);
    const subId = String(u.searchParams.get('chat') || u.searchParams.get('submissionId') || '').trim();
    if (subId) {
      window.openApplicationChat?.(subId);
    }
  } catch (_) {}
}

function readChatFromCurrentUrl() {
  try {
    const u = new URL(window.location.href);
    return String(u.searchParams.get('chat') || u.searchParams.get('submissionId') || '').trim();
  } catch (_) {
    return '';
  }
}

function canCurrentUserSend(submission) {
  if (!currentUser || !submission) return { ok: false, reason: 'User session unavailable.' };
  const role = normalizeRole(currentProfile?.role || 'uploader');
  const email = String(currentUser.email || '').toLowerCase();
  const status = String(submission.status || '').toLowerCase();

  if (role === 'admin' || role === 'super_admin') return { ok: true, reason: '' };

  if (status === 'pending') {
    return String(submission.assignedTo || '').toLowerCase() === email
      ? { ok: true, reason: '' }
      : { ok: false, reason: 'Chat closed: pending review is assigned to another reviewer.' };
  }
  if (status === 'rejected') {
    return String(submission.uploadedBy || '').toLowerCase() === email
      ? { ok: true, reason: '' }
      : { ok: false, reason: 'Chat closed: correction stage belongs to uploader.' };
  }
  if (status === 'processing_to_pfa' || status === 'approved') {
    return String(submission.assignedToRSA || '').toLowerCase() === email
      ? { ok: true, reason: '' }
      : { ok: false, reason: 'Chat closed: RSA processing moved to another user.' };
  }
  if (status === 'sent_to_pfa' || status === 'rsa_submitted' || status === 'paid') {
    return String(submission.assignedToPayment || '').toLowerCase() === email
      ? { ok: true, reason: '' }
      : { ok: false, reason: 'Chat closed: payment stage moved to another user.' };
  }

  return { ok: false, reason: 'Chat is read-only for this stage.' };
}

function getRoleDisplay() {
  const role = normalizeRole(currentProfile?.role || '');
  if (!role) return 'User';
  return role.charAt(0).toUpperCase() + role.slice(1);
}

function ensureModal() {
  if (document.getElementById('appChatModal')) return;
  const modal = document.createElement('div');
  modal.id = 'appChatModal';
  modal.className = 'app-chat-modal';
  modal.innerHTML = `
    <div class="app-chat-card" role="dialog" aria-modal="true" aria-labelledby="appChatTitle">
      <div class="app-chat-head">
        <div>
          <h3 id="appChatTitle">Application Chat</h3>
          <div class="app-chat-sub" id="appChatSubTitle">-</div>
        </div>
        <div class="app-chat-controls">
          <span id="appChatNotifState" class="app-chat-notif state-unknown">Notifications: Checking...</span>
          <span id="appChatPushState" class="app-chat-notif state-unknown">Push: Not checked</span>
          <button type="button" class="app-chat-btn" id="appChatNotifBtn">Enable Notifications</button>
          <button type="button" class="app-chat-btn danger" id="appChatEscalateBtn">Escalate to Admin</button>
          <button type="button" class="app-chat-btn" id="appChatCloseBtn">Close</button>
        </div>
      </div>
      <div class="app-chat-context" id="appChatContext"></div>
      <div class="app-chat-stage" id="appChatStageNotice"></div>
      <div class="app-chat-body" id="appChatBody">
        <div class="app-chat-empty">Loading chat...</div>
      </div>
      <div class="app-chat-foot">
        <input type="text" id="appChatInput" class="app-chat-input" placeholder="Type a message...">
        <button type="button" class="app-chat-btn primary" id="appChatSendBtn">Send</button>
        <div class="app-chat-notice" id="appChatSendNotice"></div>
      </div>
    </div>
  `;
  document.body.appendChild(modal);

  document.getElementById('appChatCloseBtn')?.addEventListener('click', closeChat);
  modal.addEventListener('click', (e) => {
    if (e.target === modal) closeChat();
  });
  document.getElementById('appChatSendBtn')?.addEventListener('click', sendMessage);
  document.getElementById('appChatInput')?.addEventListener('keypress', (e) => {
    if (e.key === 'Enter') {
      e.preventDefault();
      sendMessage();
    }
  });
  document.getElementById('appChatEscalateBtn')?.addEventListener('click', escalateToAdmin);
  document.getElementById('appChatNotifBtn')?.addEventListener('click', requestNotificationPermissionFromUI);
  updateNotificationStatusUI();
  updatePushTokenStatusUI();
}

function showChatNotice(text, isError = false) {
  const notice = document.getElementById('appChatSendNotice');
  if (!notice) return;
  notice.textContent = String(text || '');
  notice.style.color = isError ? '#ffb3b3' : '';
  if (globalNoticeTimer) clearTimeout(globalNoticeTimer);
  if (text) {
    globalNoticeTimer = setTimeout(() => {
      notice.textContent = '';
      notice.style.color = '';
    }, 4500);
  }
}

function closeChat() {
  const modal = document.getElementById('appChatModal');
  if (modal) modal.classList.remove('active');
  if (typeof unsubMessages === 'function') {
    try { unsubMessages(); } catch (_) {}
    unsubMessages = null;
  }
  if (typeof unsubSubmission === 'function') {
    try { unsubSubmission(); } catch (_) {}
    unsubSubmission = null;
  }
  currentSubmission = null;
  currentChatId = '';
}

function showModal() {
  ensureModal();
  const modal = document.getElementById('appChatModal');
  if (modal) modal.classList.add('active');
  updateNotificationStatusUI();
  updatePushTokenStatusUI();
}

function notificationStatusInfo() {
  if (!('Notification' in window)) {
    return { stateClass: 'state-unsupported', text: 'Notifications: Unsupported', canRequest: false, action: 'none' };
  }
  if (Notification.permission === 'granted') {
    return { stateClass: 'state-granted', text: 'Notifications: Enabled', canRequest: false, action: 'none' };
  }
  if (Notification.permission === 'denied') {
    return { stateClass: 'state-denied', text: 'Notifications: Blocked', canRequest: true, action: 'help' };
  }
  return { stateClass: 'state-default', text: 'Notifications: Not enabled', canRequest: true, action: 'request' };
}

function updateNotificationStatusUI() {
  const statusEl = document.getElementById('appChatNotifState');
  const btn = document.getElementById('appChatNotifBtn');
  if (!statusEl || !btn) return;
  const info = notificationStatusInfo();

  statusEl.textContent = info.text;
  statusEl.classList.remove('state-unknown', 'state-granted', 'state-denied', 'state-default', 'state-unsupported');
  statusEl.classList.add(info.stateClass);
  btn.style.display = info.canRequest ? '' : 'none';
  btn.textContent = info.action === 'help' ? 'How to Enable' : 'Enable Notifications';
  btn.dataset.action = info.action || 'none';
}

async function requestNotificationPermissionFromUI() {
  if (!('Notification' in window)) {
    updateNotificationStatusUI();
    return;
  }
  const btn = document.getElementById('appChatNotifBtn');
  const action = btn?.dataset?.action || 'none';
  if (action === 'help') {
    const host = window.location.origin;
    alert(
      `Notifications are blocked for this app (${host}).\n\n` +
      `To enable:\n` +
      `1. Click the lock/site icon near the address bar.\n` +
      `2. Open Site settings.\n` +
      `3. Set Notifications to Allow.\n` +
      `4. Reload this page.\n\n` +
      `Note: Chrome and Edge keep separate notification permissions.`
    );
    return;
  }
  try {
    await Notification.requestPermission();
  } catch (_) {}
  updateNotificationStatusUI();
  updatePushTokenStatusUI();
}

function setPushTokenState(status, detail) {
  pushTokenState = {
    status: String(status || 'idle'),
    detail: String(detail || 'Push: Not checked')
  };
  updatePushTokenStatusUI();
}

function updatePushTokenStatusUI() {
  const el = document.getElementById('appChatPushState');
  if (!el) return;
  const map = {
    registered: 'state-granted',
    registering: 'state-default',
    blocked: 'state-denied',
    unsupported: 'state-unsupported',
    error: 'state-denied',
    idle: 'state-unknown'
  };
  const css = map[pushTokenState.status] || 'state-unknown';
  el.textContent = pushTokenState.detail;
  el.classList.remove('state-unknown', 'state-granted', 'state-denied', 'state-default', 'state-unsupported');
  el.classList.add(css);
}

function setReadOnlyState(submission) {
  const input = document.getElementById('appChatInput');
  const sendBtn = document.getElementById('appChatSendBtn');
  const notice = document.getElementById('appChatSendNotice');
  const stageNotice = document.getElementById('appChatStageNotice');
  const { ok, reason } = canCurrentUserSend(submission);

  if (input) input.disabled = !ok;
  if (sendBtn) sendBtn.disabled = !ok;
  if (notice) notice.textContent = ok ? '' : reason;
  if (stageNotice) stageNotice.textContent = `Status: ${String(submission?.status || '-')} | Role: ${getRoleDisplay()}`;
}

function maybeNotifyMessage(submission, msg, isOwn) {
  if (isOwn) return;
  if (!msg || !currentUser) return;
  if (!('Notification' in window)) return;
  if (Notification.permission !== 'granted') return;

  const key = `${submission.id}:${msg.id || ''}`;
  if (messageSeenMap.has(key)) return;
  messageSeenMap.set(key, true);

  const title = `New Chat Message - ${submission.customerName || 'Application'}`;
  const body = `${msg.senderName || msg.senderEmail || 'User'}: ${String(msg.text || '').slice(0, 110)}`;
  const n = new Notification(title, { body, tag: `chat-${submission.id}` });
  n.onclick = () => {
    window.focus();
    window.openApplicationChat?.(submission.id);
  };
}

function tsToMillis(v) {
  if (!v) return 0;
  try {
    if (typeof v.toMillis === 'function') return v.toMillis();
    const d = v.toDate ? v.toDate() : new Date(v);
    return Number.isNaN(d.getTime()) ? 0 : d.getTime();
  } catch (_) {
    return 0;
  }
}

function deriveParticipantsFromSubmission(submission) {
  if (!submission) return [];
  const raw = [
    submission.uploadedBy,
    submission.assignedTo,
    submission.reviewedBy,
    submission.assignedToRSA,
    submission.assignedToPayment
  ];
  return Array.from(new Set(raw
    .map((v) => String(v || '').trim().toLowerCase())
    .filter(Boolean)));
}

function maybeNotifyGlobalChat(chatId, chatData) {
  if (!currentUser) return;
  if (!('Notification' in window)) return;
  if (Notification.permission !== 'granted') return;

  const me = String(currentUser.email || '').toLowerCase();
  const by = String(chatData?.lastMessageBy || '').toLowerCase();
  if (!by || by === me) return;

  const when = tsToMillis(chatData?.lastMessageAt || chatData?.updatedAt);
  if (!when) return;
  const key = `global:${chatId}:${when}`;
  if (messageSeenMap.has(key)) return;
  messageSeenMap.set(key, true);

  const title = `New Chat Message - ${chatData?.customerName || 'Application'}`;
  const body = String(chatData?.lastMessage || 'You have a new message').slice(0, 120);
  const n = new Notification(title, { body, tag: `chat-${chatId}` });
  n.onclick = () => {
    window.focus();
    window.openApplicationChat?.(chatId);
  };
}

function startGlobalChatNotifications() {
  if (typeof unsubGlobalChats === 'function') {
    try { unsubGlobalChats(); } catch (_) {}
    unsubGlobalChats = null;
  }
  globalChatLastSeen.clear();
  if (!currentUser) return;
  const myEmail = String(currentUser.email || '').trim().toLowerCase();
  if (!myEmail) return;

  const chatsQuery = query(
    collection(db, 'applicationChats'),
    where('participants', 'array-contains', myEmail),
    limit(120)
  );

  let initialized = false;
  unsubGlobalChats = onSnapshot(
    chatsQuery,
    (snap) => {
      if (!initialized) {
        snap.docs.forEach((d) => {
          const data = d.data() || {};
          const at = tsToMillis(data.lastMessageAt || data.updatedAt);
          if (at) globalChatLastSeen.set(d.id, at);
        });
        initialized = true;
        return;
      }

      snap.docChanges().forEach((change) => {
        if (change.type === 'removed') return;
        const chatId = change.doc.id;
        const data = change.doc.data() || {};
        const at = tsToMillis(data.lastMessageAt || data.updatedAt);
        if (!at) return;

        const prev = globalChatLastSeen.get(chatId) || 0;
        if (at <= prev) return;
        globalChatLastSeen.set(chatId, at);
        maybeNotifyGlobalChat(chatId, data);
      });
    },
    (error) => {
      showChatNotice(`Notification listener unavailable: ${error?.message || 'Permission denied'}`, true);
    }
  );
}

function renderMessages(submission, messages) {
  const body = document.getElementById('appChatBody');
  if (!body) return;
  if (!messages.length) {
    body.innerHTML = '<div class="app-chat-empty">No messages yet.</div>';
    return;
  }
  const me = String(currentUser?.email || '').toLowerCase();
  body.innerHTML = messages.map((m) => {
    const senderEmail = String(m.senderEmail || '').toLowerCase();
    const own = senderEmail && senderEmail === me;
    return `
      <div class="app-chat-msg ${own ? 'me' : ''}">
        <div class="app-chat-meta">
          <span>${esc(m.senderName || m.senderEmail || 'User')} (${esc(m.senderRole || '-')})</span>
          <span>${esc(fmtDate(m.createdAt))}</span>
        </div>
        <div class="app-chat-text">${esc(m.text || '')}</div>
      </div>
    `;
  }).join('');
  body.scrollTop = body.scrollHeight;

  messages.forEach((m) => {
    maybeNotifyMessage(submission, m, String(m.senderEmail || '').toLowerCase() === me);
  });
}

async function loadAndWatchMessages(submissionId, submission) {
  const chatRef = doc(db, 'applicationChats', submissionId);
  const participants = deriveParticipantsFromSubmission(submission);
  const chatSnap = await getDoc(chatRef);
  if (!chatSnap.exists()) {
    await setDoc(chatRef, {
      submissionId,
      customerName: submission.customerName || '',
      uploadedBy: submission.uploadedBy || '',
      participants,
      createdAt: serverTimestamp(),
      escalated: false
    }, { merge: true });
  } else {
    await setDoc(chatRef, {
      submissionId,
      customerName: submission.customerName || '',
      participants
    }, { merge: true });
  }

  if (typeof unsubMessages === 'function') {
    try { unsubMessages(); } catch (_) {}
  }
  const msgQuery = query(
    collection(db, 'applicationChats', submissionId, 'messages'),
    orderBy('createdAt', 'asc')
  );
  unsubMessages = onSnapshot(
    msgQuery,
    (snap) => {
      const msgs = snap.docs.map((d) => ({ id: d.id, ...d.data() }));
      renderMessages(submission, msgs);
    },
    (error) => {
      showChatNotice(`Unable to load messages: ${error?.message || 'Permission denied'}`, true);
    }
  );
}

async function sendMessage() {
  const input = document.getElementById('appChatInput');
  if (!input || !currentSubmission || !currentChatId || !currentUser) return;
  const text = String(input.value || '').trim();
  if (!text) return;

  const access = canCurrentUserSend(currentSubmission);
  if (!access.ok) {
    setReadOnlyState(currentSubmission);
    return;
  }

  try {
    const senderEmail = String(currentUser.email || '').toLowerCase();
    input.value = '';
    await addDoc(collection(db, 'applicationChats', currentChatId, 'messages'), {
      text,
      senderEmail,
      senderName: currentProfile?.fullName || currentUser.email || 'User',
      senderRole: normalizeRole(currentProfile?.role || ''),
      createdAt: serverTimestamp()
    });
    await setDoc(doc(db, 'applicationChats', currentChatId), {
      updatedAt: serverTimestamp(),
      lastMessage: text.slice(0, 300),
      lastMessageAt: serverTimestamp(),
      lastMessageBy: senderEmail,
      participants: deriveParticipantsFromSubmission(currentSubmission)
    }, { merge: true });
    triggerServerPush({
      submissionId: currentChatId,
      customerName: currentSubmission?.customerName || '',
      messageText: text.slice(0, 500)
    }).catch(() => {});
  } catch (error) {
    showChatNotice(`Message failed: ${error?.message || 'Unknown error'}`, true);
    input.value = text;
  }
}

async function escalateToAdmin() {
  if (!currentSubmission || !currentChatId || !currentUser) return;
  const role = normalizeRole(currentProfile?.role || '');
  if (role === 'admin' || role === 'super_admin') return;

  const usersRef = collection(db, 'users');
  const adminSnap = await getDocs(query(usersRef, where('role', '==', 'admin')));
  const superSnap = await getDocs(query(usersRef, where('role', '==', 'super_admin')));
  const adminEmails = [
    ...adminSnap.docs.map((d) => String(d.data()?.email || '').toLowerCase()).filter(Boolean),
    ...superSnap.docs.map((d) => String(d.data()?.email || '').toLowerCase()).filter(Boolean)
  ];
  const unique = Array.from(new Set(adminEmails));

  try {
    await setDoc(doc(db, 'applicationChats', currentChatId), {
      escalated: true,
      escalatedAt: serverTimestamp(),
      escalatedBy: currentUser.email || '',
      adminParticipants: unique,
      participants: Array.from(new Set([
        ...deriveParticipantsFromSubmission(currentSubmission),
        ...unique
      ]))
    }, { merge: true });

    await addDoc(collection(db, 'applicationChats', currentChatId, 'messages'), {
      text: `Escalated to admin by ${currentProfile?.fullName || currentUser.email || 'User'}`,
      senderEmail: currentUser.email || '',
      senderName: currentProfile?.fullName || currentUser.email || 'User',
      senderRole: normalizeRole(currentProfile?.role || ''),
      system: true,
      createdAt: serverTimestamp()
    });
    showChatNotice('Escalated to admin successfully.');
  } catch (error) {
    showChatNotice(`Escalation failed: ${error?.message || 'Unknown error'}`, true);
  }
}

async function fetchUserProfile(email, uid) {
  try {
    const usersRef = collection(db, 'users');
    if (email) {
      const s = await getDocs(query(usersRef, where('email', '==', String(email).toLowerCase())));
      if (!s.empty) return { __docId: s.docs[0].id, ...s.docs[0].data() };
    }
    if (uid) {
      const s = await getDocs(query(usersRef, where('uid', '==', uid)));
      if (!s.empty) return { __docId: s.docs[0].id, ...s.docs[0].data() };
    }
  } catch (_) {}
  return null;
}

async function registerFcmTokenIfPossible() {
  try {
    if (!currentUser || !currentProfileDocId) return;
    if (!FCM_VAPID_KEY) {
      setPushTokenState('error', 'Push: Missing VAPID key');
      return;
    }
    if (!('serviceWorker' in navigator)) {
      setPushTokenState('unsupported', 'Push: Service Worker unsupported');
      return;
    }
    if (!('Notification' in window)) {
      setPushTokenState('unsupported', 'Push: Notifications unsupported');
      return;
    }
    if (Notification.permission !== 'granted') {
      setPushTokenState('blocked', 'Push: Allow notifications first');
      return;
    }
    if (!(await isMessagingSupported())) {
      setPushTokenState('unsupported', 'Push: Messaging not supported');
      return;
    }

    setPushTokenState('registering', 'Push: Registering token...');
    const reg = await navigator.serviceWorker.register('/service-worker.js');
    const messaging = getMessaging(app);
    const token = await getToken(messaging, {
      vapidKey: FCM_VAPID_KEY,
      serviceWorkerRegistration: reg
    });
    if (!token) {
      setPushTokenState('error', 'Push: Token request returned empty');
      return;
    }

    await updateDoc(doc(db, 'users', currentProfileDocId), {
      fcmTokens: arrayUnion(token),
      fcmLastTokenAt: serverTimestamp(),
      fcmLastTokenPlatform: navigator.userAgent || ''
    });
    setPushTokenState('registered', 'Push: Token registered');
  } catch (error) {
    setPushTokenState('error', `Push: ${String(error?.message || 'registration failed')}`);
  }
}

async function triggerServerPush({ submissionId, customerName, messageText }) {
  try {
    const base = getEmailApiBaseUrl();
    if (!base || !submissionId || !currentUser) return;
    const idToken = await currentUser.getIdToken();
    const res = await fetch(`${base}/api/chat/push`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Authorization': `Bearer ${idToken}`
      },
      body: JSON.stringify({
        submissionId,
        customerName: String(customerName || ''),
        messageText: String(messageText || '').slice(0, 500)
      })
    });
    let data = null;
    try { data = await res.json(); } catch (_) {}

    if (!res.ok) {
      const reason = String(data?.error || `HTTP ${res.status}`);
      showChatNotice(`Push server error: ${reason}`, true);
      return;
    }

    const sent = Number(data?.sent || 0);
    if (sent <= 0) {
      const reason = String(data?.reason || 'no-recipient-device-token');
      showChatNotice(`Push not delivered yet: ${reason}`, true);
    }
  } catch (error) {
    const msg = String(error?.message || 'network/cors failure');
    showChatNotice(`Push request failed: ${msg}`, true);
  }
}

window.openApplicationChat = async (submissionId) => {
  try {
    if (!submissionId) return;
    ensureModal();
    showModal();

    const subRef = doc(db, 'submissions', submissionId);
    const subSnap = await getDoc(subRef);
    if (!subSnap.exists()) {
      alert('Application chat unavailable: record not found.');
      closeChat();
      return;
    }
    currentSubmission = { id: submissionId, ...subSnap.data() };
    currentChatId = submissionId;

    const title = document.getElementById('appChatSubTitle');
    const context = document.getElementById('appChatContext');
    if (title) title.textContent = `Application ID: ${submissionId}`;
    if (context) {
      context.innerHTML = `
        <span><strong>Customer:</strong> ${esc(currentSubmission.customerName || '-')}</span>
        <span><strong>Status:</strong> ${esc(currentSubmission.status || '-')}</span>
        <span><strong>Uploader:</strong> ${esc(currentSubmission.uploadedBy || '-')}</span>
        <span><strong>Reviewer:</strong> ${esc(currentSubmission.assignedTo || '-')}</span>
      `;
    }
    setReadOnlyState(currentSubmission);

    if (typeof unsubSubmission === 'function') {
      try { unsubSubmission(); } catch (_) {}
    }
    unsubSubmission = onSnapshot(
      doc(db, 'submissions', submissionId),
      (snap) => {
        if (!snap.exists()) return;
        currentSubmission = { id: submissionId, ...snap.data() };
        setReadOnlyState(currentSubmission);
      },
      (error) => {
        showChatNotice(`Unable to watch submission: ${error?.message || 'Permission denied'}`, true);
      }
    );

    await loadAndWatchMessages(submissionId, currentSubmission);
  } catch (err) {
    showChatNotice(`Unable to open chat: ${err?.message || 'Unknown error'}`, true);
  }
};

function injectChatButtonsFallback() {
  document.querySelectorAll('button[data-chat-submission]').forEach((btn) => {
    if (btn.dataset.chatBound === '1') return;
    btn.dataset.chatBound = '1';
    btn.addEventListener('click', () => {
      const subId = btn.getAttribute('data-chat-submission');
      if (subId) window.openApplicationChat(subId);
    });
  });
}

setInterval(injectChatButtonsFallback, 1500);

// Safety net: if a stale cached row still renders a WhatsApp link in rejected table,
// reroute the click to in-app chat instead of opening WhatsApp web.
document.addEventListener('click', (e) => {
  const waLink = e.target?.closest?.('a[href*="wa.me/"]');
  if (!waLink) return;
  const row = waLink.closest('tr');
  const insideRejectedTable = !!waLink.closest('#rejectedTableBody');
  const looksLikeRejectedRow = row?.querySelector?.('.status-rejected');
  if (!insideRejectedTable && !looksLikeRejectedRow) return;

  const chatBtn = row?.querySelector?.('button[data-chat-submission]');
  if (chatBtn) {
    e.preventDefault();
    chatBtn.click();
    return;
  }

  const submissionId = row?.getAttribute?.('data-submission-id');
  if (submissionId) {
    e.preventDefault();
    window.openApplicationChat?.(submissionId);
  }
});

onAuthStateChanged(auth, async (user) => {
  if (!user) {
    currentUser = null;
    currentProfile = null;
    currentProfileDocId = '';
    if (typeof unsubGlobalChats === 'function') {
      try { unsubGlobalChats(); } catch (_) {}
      unsubGlobalChats = null;
    }
    globalChatLastSeen.clear();
    setPushTokenState('idle', 'Push: Not checked');
    updateNotificationStatusUI();
    return;
  }

  currentUser = user || null;
  currentProfile = null;
  currentProfileDocId = '';
  currentProfile = await fetchUserProfile(user.email || '', user.uid);
  currentProfileDocId = String(currentProfile?.__docId || '').trim();
  if ('Notification' in window && Notification.permission === 'default') {
    try { await Notification.requestPermission(); } catch (_) {}
  }
  await registerFcmTokenIfPossible();
  startGlobalChatNotifications();
  updateNotificationStatusUI();
  const fromUrl = pendingChatFromUrl || readChatFromCurrentUrl();
  if (fromUrl) {
    pendingChatFromUrl = '';
    setTimeout(() => window.openApplicationChat?.(fromUrl), 450);
  }
});

window.addEventListener('focus', updateNotificationStatusUI);
document.addEventListener('visibilitychange', () => {
  if (document.visibilityState === 'visible') updateNotificationStatusUI();
});

if ('serviceWorker' in navigator) {
  navigator.serviceWorker.addEventListener('message', (event) => {
    if (event?.data?.type === 'open-chat-url') {
      pendingChatFromUrl = '';
      openChatFromUrl(event.data.url);
    }
  });
}

pendingChatFromUrl = readChatFromCurrentUrl();
