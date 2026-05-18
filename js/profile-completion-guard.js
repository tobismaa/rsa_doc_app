import { auth, db } from './firebase-config.js';
import { onAuthStateChanged, signOut } from 'https://www.gstatic.com/firebasejs/9.22.0/firebase-auth.js';
import { collection, query, where, getDocs, doc, updateDoc, serverTimestamp, onSnapshot } from 'https://www.gstatic.com/firebasejs/9.22.0/firebase-firestore.js';
import { getMaintenanceSettings, isMaintenanceExemptRole, showMaintenanceOverlay } from './shared/maintenance-mode.js?v=20260507a';
import { getSystemSettings } from './shared/system-settings.js?v=20260508a';
import { registerPushTokenForCurrentUser } from './push-alerts.js';

const REQUIRED_FIELDS = [
  { key: 'fullName', label: 'Full Name' },
  { key: 'location', label: 'Location' },
  { key: 'whatsappCode', label: 'WhatsApp Country Code' },
  { key: 'whatsappLocalNumber', label: 'WhatsApp Number' }
];

const CACHE_CLEAR_HANDLED_KEY = 'cmbank_app_cache_clear_handled_token';
const CACHE_BUST_TOKEN_KEY = 'cmbank_app_cache_bust_token';
let cacheClearWatcherStarted = false;
let cacheClearInProgress = false;
let securityWatchStarted = false;
let sessionTimeoutTimer = null;
let inactivityHandlersBound = false;

const WHATSAPP_COUNTRY_CODES = [
  { code: '+234', label: 'Nigeria', flag: '🇳🇬' },
  { code: '+233', label: 'Ghana', flag: '🇬🇭' },
  { code: '+254', label: 'Kenya', flag: '🇰🇪' },
  { code: '+27', label: 'South Africa', flag: '🇿🇦' },
  { code: '+1', label: 'United States', flag: '🇺🇸' },
  { code: '+44', label: 'United Kingdom', flag: '🇬🇧' },
  { code: '+971', label: 'United Arab Emirates', flag: '🇦🇪' }
];

function normalizeText(value) {
  return String(value || '').trim();
}

function missingFields(userData = {}) {
  return REQUIRED_FIELDS.filter((field) => !normalizeText(userData[field.key]));
}

function buildModal(missing = [], existing = {}) {
  if (document.getElementById('profileGuardModal')) return;

  const style = document.createElement('style');
  style.id = 'profileGuardStyles';
  style.textContent = `
    .profile-guard-backdrop {
      position: fixed;
      inset: 0;
      background: rgba(2, 6, 23, 0.74);
      z-index: 20000;
      display: flex;
      align-items: center;
      justify-content: center;
      padding: 16px;
    }
    .profile-guard-modal {
      width: min(560px, 100%);
      background: #fff;
      border-radius: 14px;
      box-shadow: 0 24px 60px rgba(2, 6, 23, 0.4);
      overflow: hidden;
    }
    .profile-guard-head {
      padding: 16px 18px;
      background: #0f3b67;
      color: #fff;
      font-weight: 700;
    }
    .profile-guard-body {
      padding: 16px 18px;
      color: #334155;
      font-size: 14px;
    }
    .profile-guard-list {
      margin: 0 0 12px;
      padding-left: 18px;
    }
    .profile-guard-grid {
      display: grid;
      grid-template-columns: repeat(2, minmax(0, 1fr));
      gap: 10px;
    }
    .profile-guard-grid label {
      display: block;
      font-size: 12px;
      font-weight: 700;
      color: #334155;
    }
    .profile-guard-grid input {
      width: 100%;
      margin-top: 6px;
      padding: 10px 11px;
      border: 1px solid #d1d5db;
      border-radius: 8px;
      font-size: 14px;
    }
    .profile-guard-grid select {
      width: 100%;
      margin-top: 6px;
      padding: 10px 11px;
      border: 1px solid #d1d5db;
      border-radius: 8px;
      font-size: 14px;
      background: #fff;
    }
    .profile-guard-grid .span-2 {
      grid-column: span 2;
    }
    .profile-guard-foot {
      padding: 12px 18px 16px;
      display: flex;
      justify-content: flex-end;
    }
    .profile-guard-btn {
      border: 0;
      border-radius: 8px;
      background: #0f766e;
      color: #fff;
      font-weight: 700;
      padding: 10px 16px;
      cursor: pointer;
    }
    .profile-guard-btn:disabled {
      opacity: 0.55;
      cursor: not-allowed;
    }
    .profile-guard-error {
      color: #dc2626;
      margin: 8px 0 0;
      font-size: 12px;
      min-height: 16px;
    }
    @media (max-width: 640px) {
      .profile-guard-grid { grid-template-columns: 1fr; }
      .profile-guard-grid .span-2 { grid-column: span 1; }
    }
  `;
  document.head.appendChild(style);

  const backdrop = document.createElement('div');
  backdrop.id = 'profileGuardModal';
  backdrop.className = 'profile-guard-backdrop';
  const missingList = missing.map((f) => `<li>${f.label}</li>`).join('');
  const selectedCode = normalizeText(existing.whatsappCode || '+234');
  const countryCodeOptions = WHATSAPP_COUNTRY_CODES.map((item) => `
    <option value="${item.code}" ${item.code === selectedCode ? 'selected' : ''}>${item.flag} ${item.code} ${item.label}</option>
  `).join('');
  backdrop.innerHTML = `
    <div class="profile-guard-modal" role="dialog" aria-modal="true" aria-labelledby="profileGuardTitle">
      <div id="profileGuardTitle" class="profile-guard-head">Complete Your Profile</div>
      <div class="profile-guard-body">
        <p>Your profile has missing details. Update now to continue.</p>
        <ul class="profile-guard-list">${missingList}</ul>
        <div class="profile-guard-grid">
          <label class="span-2">Full Name
            <input id="pgFullName" type="text" value="${normalizeText(existing.fullName)}" placeholder="Enter full name">
          </label>
          <label class="span-2">Location
            <input id="pgLocation" type="text" value="${normalizeText(existing.location)}" placeholder="Enter location">
          </label>
          <label>WhatsApp Country Code
            <select id="pgWhatsappCode">${countryCodeOptions}</select>
          </label>
          <label>WhatsApp Number
            <input id="pgWhatsappLocal" type="text" value="${normalizeText(existing.whatsappLocalNumber)}" placeholder="10 digits">
          </label>
        </div>
        <p id="profileGuardError" class="profile-guard-error"></p>
      </div>
      <div class="profile-guard-foot">
        <button id="profileGuardSaveBtn" class="profile-guard-btn" type="button">Save and Continue</button>
      </div>
    </div>
  `;
  document.body.appendChild(backdrop);
}

async function clearBrowserAppCaches(cacheClearToken = '') {
  if ('serviceWorker' in navigator) {
    try {
      try {
        navigator.serviceWorker.controller?.postMessage?.({ type: 'clear-app-cache', token: cacheClearToken });
      } catch (_) {}
      const registrations = await navigator.serviceWorker.getRegistrations();
      await Promise.all(registrations.map((registration) => {
        try {
          registration.active?.postMessage?.({ type: 'clear-app-cache', token: cacheClearToken });
          registration.waiting?.postMessage?.({ type: 'clear-app-cache', token: cacheClearToken });
          registration.installing?.postMessage?.({ type: 'clear-app-cache', token: cacheClearToken });
        } catch (_) {}
        return registration.update().catch(() => false);
      }));
    } catch (_) {}
  }

  if ('caches' in window) {
    try {
      const keys = await caches.keys();
      await Promise.all(keys.map((key) => caches.delete(key).catch(() => false)));
    } catch (_) {}
  }
}

function showDashboardAnnouncement(announcement = {}) {
  const existing = document.getElementById('systemAnnouncementBanner');
  const enabled = announcement?.enabled === true;
  const message = String(announcement?.message || '').trim();
  if (!enabled || !message) {
    existing?.remove();
    return;
  }

  const tone = String(announcement?.tone || 'info').trim().toLowerCase();
  const palette = tone === 'warning'
    ? { bg: '#fff7ed', border: '#fdba74', text: '#9a3412' }
    : tone === 'success'
      ? { bg: '#ecfdf5', border: '#6ee7b7', text: '#065f46' }
      : { bg: '#eff6ff', border: '#93c5fd', text: '#1d4ed8' };
  const escapedMessage = normalizeText(message);
  const bannerMarkup = `<div class="system-announcement-shell"><span class="system-announcement-badge">Update</span><div class="system-announcement-track"><span class="system-announcement-text">${escapedMessage}</span></div></div>`;

  if (existing) {
    existing.innerHTML = bannerMarkup;
    existing.style.background = palette.bg;
    existing.style.borderBottom = `1px solid ${palette.border}`;
    existing.style.color = palette.text;
    return;
  }

  if (!document.getElementById('systemAnnouncementBannerStyles')) {
    const style = document.createElement('style');
    style.id = 'systemAnnouncementBannerStyles';
    style.textContent = `
      @keyframes systemAnnouncementScroll {
        0% { transform: translateX(100%); }
        100% { transform: translateX(-100%); }
      }
      .system-announcement-track {
        position: relative;
        overflow: hidden;
        white-space: nowrap;
        flex: 1;
      }
      .system-announcement-text {
        display: inline-block;
        padding-left: 100%;
        white-space: nowrap;
        will-change: transform;
        animation: systemAnnouncementScroll 16s linear infinite;
        font-size: 15px;
        font-weight: 800;
        letter-spacing: 0.01em;
      }
      .system-announcement-shell {
        display: flex;
        align-items: center;
        gap: 14px;
      }
      .system-announcement-badge {
        display: inline-flex;
        align-items: center;
        justify-content: center;
        padding: 6px 12px;
        border-radius: 999px;
        background: rgba(255,255,255,0.75);
        border: 1px solid rgba(255,255,255,0.9);
        box-shadow: 0 8px 18px rgba(15, 23, 42, 0.12);
        font-size: 12px;
        font-weight: 900;
        text-transform: uppercase;
        letter-spacing: 0.08em;
      }
    `;
    document.head.appendChild(style);
  }

  const banner = document.createElement('div');
  banner.id = 'systemAnnouncementBanner';
  banner.innerHTML = bannerMarkup;
  banner.style.cssText = `position:sticky;top:0;z-index:19999;padding:12px 18px;background:${palette.bg};border-bottom:1px solid ${palette.border};color:${palette.text};box-shadow:0 10px 24px rgba(15,23,42,0.08);`;
  document.body.prepend(banner);
}

function startSessionTimeout(minutes = 60) {
  if (sessionTimeoutTimer) clearTimeout(sessionTimeoutTimer);
  const timeoutMs = Math.max(1, Number(minutes || 60)) * 60 * 1000;
  sessionTimeoutTimer = window.setTimeout(async () => {
    try {
      await signOut(auth);
    } finally {
      window.location.href = 'index.html';
    }
  }, timeoutMs);
}

function bindInactivityHandlers(minutes = 60) {
  const restart = () => startSessionTimeout(minutes);
  if (!inactivityHandlersBound) {
    ['click', 'keydown', 'mousemove', 'scroll', 'touchstart'].forEach((eventName) => {
      window.addEventListener(eventName, restart, { passive: true });
    });
    inactivityHandlersBound = true;
  }
  restart();
}

function watchForCacheClearSignal() {
  if (cacheClearWatcherStarted) return;
  cacheClearWatcherStarted = true;

  onSnapshot(doc(db, 'settings', 'system'), async (snap) => {
    const data = snap.exists() ? (snap.data() || {}) : {};
    const token = String(data.cacheClearToken || '').trim();
    const handled = String(localStorage.getItem(CACHE_CLEAR_HANDLED_KEY) || '').trim();
    if (!token || token === handled || cacheClearInProgress) return;

    cacheClearInProgress = true;
    try {
      localStorage.setItem(CACHE_CLEAR_HANDLED_KEY, token);
      localStorage.setItem(CACHE_BUST_TOKEN_KEY, token);
      sessionStorage.setItem(CACHE_BUST_TOKEN_KEY, token);
      await clearBrowserAppCaches(token);
    } catch (_) {
      localStorage.setItem(CACHE_CLEAR_HANDLED_KEY, token);
      localStorage.setItem(CACHE_BUST_TOKEN_KEY, token);
      sessionStorage.setItem(CACHE_BUST_TOKEN_KEY, token);
    }

    const url = new URL(window.location.href);
    url.searchParams.set('_cacheReset', token || Date.now().toString());
    window.location.replace(url.toString());
  }, () => {
    // Fail open if the signal watcher cannot start.
  });
}

function watchForSecuritySignals(userData = {}) {
  if (securityWatchStarted) return;
  securityWatchStarted = true;

  onSnapshot(doc(db, 'settings', 'system'), async (snap) => {
    const systemSettings = snap.exists()
      ? await getSystemSettings(db, { force: true })
      : await getSystemSettings(db, { force: true });

    showDashboardAnnouncement(systemSettings.dashboardAnnouncement);
    bindInactivityHandlers(systemSettings.securityControls.sessionTimeoutMinutes);

    const forceLogoutToken = String(systemSettings.securityControls.forceLogoutToken || '').trim();
    const handled = String(localStorage.getItem('cmbank_force_logout_token') || '').trim();
    if (forceLogoutToken && forceLogoutToken !== handled && !isMaintenanceExemptRole(userData.role)) {
      localStorage.setItem('cmbank_force_logout_token', forceLogoutToken);
      try {
        await signOut(auth);
      } finally {
        window.location.href = 'index.html';
      }
    }
  }, () => {});
}

async function findUserDocByUidOrEmail(uid, email) {
  if (uid) {
    const directDoc = await getDoc(doc(db, 'users', uid));
    if (directDoc.exists()) return directDoc;
  }

  const byUid = query(collection(db, 'users'), where('uid', '==', uid));
  const uidSnap = await getDocs(byUid);
  if (!uidSnap.empty) return uidSnap.docs[0];

  const byEmail = query(collection(db, 'users'), where('email', '==', String(email || '').toLowerCase()));
  const emailSnap = await getDocs(byEmail);
  if (!emailSnap.empty) return emailSnap.docs[0];

  return null;
}

async function enforceProfileCompletion(user) {
  const userDoc = await findUserDocByUidOrEmail(user.uid, user.email);
  if (!userDoc) return;

  const userData = userDoc.data() || {};
  const missing = missingFields(userData);
  if (!missing.length) return;

  buildModal(missing, userData);
  const saveBtn = document.getElementById('profileGuardSaveBtn');
  const errorEl = document.getElementById('profileGuardError');
  const backdrop = document.getElementById('profileGuardModal');
  if (!saveBtn || !backdrop) return;

  document.body.style.overflow = 'hidden';

  const save = async () => {
    const fullName = normalizeText(document.getElementById('pgFullName')?.value);
    const location = normalizeText(document.getElementById('pgLocation')?.value);
    const whatsappCode = normalizeText(document.getElementById('pgWhatsappCode')?.value);
    const whatsappLocalNumber = normalizeText(document.getElementById('pgWhatsappLocal')?.value).replace(/\D/g, '');

    if (!fullName || !location || !whatsappCode || !whatsappLocalNumber) {
      if (errorEl) errorEl.textContent = 'All fields are required.';
      return;
    }
    if (!/^\+?\d{1,4}$/.test(whatsappCode)) {
      if (errorEl) errorEl.textContent = 'Country code is invalid.';
      return;
    }
    if (!/^\d{10}$/.test(whatsappLocalNumber)) {
      if (errorEl) errorEl.textContent = 'WhatsApp number must be exactly 10 digits.';
      return;
    }

    const normalizedCode = whatsappCode.startsWith('+') ? whatsappCode : `+${whatsappCode}`;
    const whatsappNumber = `${normalizedCode}${whatsappLocalNumber}`;

    saveBtn.disabled = true;
    saveBtn.textContent = 'Saving...';
    if (errorEl) errorEl.textContent = '';
    try {
      const waDupSnap = await getDocs(query(collection(db, 'users'), where('whatsappNumber', '==', whatsappNumber)));
      const waDuplicate = waDupSnap.docs.find((d) => d.id !== userDoc.id);
      if (waDuplicate) {
        if (errorEl) errorEl.textContent = 'This WhatsApp number already exists.';
        saveBtn.disabled = false;
        saveBtn.textContent = 'Save and Continue';
        return;
      }

      const phoneDupSnap = await getDocs(query(collection(db, 'users'), where('phone', '==', whatsappNumber)));
      const phoneDuplicate = phoneDupSnap.docs.find((d) => d.id !== userDoc.id);
      if (phoneDuplicate) {
        if (errorEl) errorEl.textContent = 'This WhatsApp number already exists.';
        saveBtn.disabled = false;
        saveBtn.textContent = 'Save and Continue';
        return;
      }

      await updateDoc(doc(db, 'users', userDoc.id), {
        fullName,
        location,
        whatsappCode: normalizedCode,
        whatsappLocalNumber,
        whatsappNumber,
        phone: whatsappNumber,
        updatedAt: serverTimestamp()
      });

      backdrop.remove();
      document.body.style.overflow = '';
      window.location.reload();
    } catch (err) {
      if (errorEl) errorEl.textContent = err?.message || 'Unable to save details.';
      saveBtn.disabled = false;
      saveBtn.textContent = 'Save and Continue';
    }
  };

  saveBtn.addEventListener('click', save);
}

onAuthStateChanged(auth, async (user) => {
  if (!user) return;
  try {
    watchForCacheClearSignal();
    const userDoc = await findUserDocByUidOrEmail(user.uid, user.email);
    if (!userDoc) return;

    const userData = userDoc.data() || {};
    const systemSettings = await getSystemSettings(db, { force: true });
    showDashboardAnnouncement(systemSettings.dashboardAnnouncement);
    bindInactivityHandlers(systemSettings.securityControls.sessionTimeoutMinutes);
    watchForSecuritySignals(userData);
    await registerPushTokenForCurrentUser(user, userDoc.id);
    const maintenanceSettings = await getMaintenanceSettings(db, { force: true });
    if (maintenanceSettings.maintenanceMode && !isMaintenanceExemptRole(userData.role)) {
      showMaintenanceOverlay({
        message: maintenanceSettings.maintenanceMessage || 'Maintenance mode is currently enabled. To protect live workflow data, portal access is temporarily restricted.',
        onSignOut: async () => {
          try {
            await signOut(auth);
          } finally {
            window.location.href = 'index.html';
          }
        }
      });
      return;
    }

    await enforceProfileCompletion(user);
  } catch (_err) {
    // Fail open to avoid blocking legitimate users on transient issues.
  }
});
