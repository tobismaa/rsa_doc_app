import { auth, db } from './firebase-config.js?v=20260625c';
import { onAuthStateChanged, signOut } from 'https://www.gstatic.com/firebasejs/9.22.0/firebase-auth.js';
import { collection, query, where, getDocs, doc, updateDoc, serverTimestamp, onSnapshot } from 'https://www.gstatic.com/firebasejs/9.22.0/firebase-firestore.js';
import { getMaintenanceSettings, isMaintenanceExemptRole, showMaintenanceOverlay } from './shared/maintenance-mode.js?v=20260507a';
import { getDefaultSystemSettings, getSystemSettings } from './shared/system-settings.js?v=20260704c';
import { performAppLogout } from './shared/logout.js?v=20260625b';
import { registerPushTokenForCurrentUser } from './push-alerts.js';

const REQUIRED_FIELDS = [
  { key: 'fullName', label: 'Full Name' },
  { key: 'location', label: 'Location' },
  { key: 'whatsappCode', label: 'WhatsApp Country Code' },
  { key: 'whatsappLocalNumber', label: 'WhatsApp Number' }
];

const CACHE_CLEAR_HANDLED_KEY = 'cmbank_app_cache_clear_handled_token';
const CACHE_BUST_TOKEN_KEY = 'cmbank_app_cache_bust_token';
const FORCE_LOGOUT_ACTIVE_TOKEN_KEY = 'cmbank_force_logout_active_token';
const FORCE_LOGOUT_COMPLETED_TOKEN_KEY = 'cmbank_force_logout_completed_token';
const FORCE_LOGOUT_DEADLINE_KEY = 'cmbank_force_logout_deadline_ms';
const FORCE_LOGOUT_DISMISSED_TOKEN_KEY = 'cmbank_force_logout_dismissed_token';
const TARGETED_LOGOUT_COMPLETED_PREFIX = 'cmbank_targeted_logout_completed_';
const FORCE_LOGOUT_GRACE_MS = 5 * 60 * 1000;
const DASHBOARD_PAGE_TARGETS = {
  'dashboard.html': 'uploader',
  'admin-dashboard.html': 'admin',
  'reviewer-dashboard.html': 'reviewer',
  'rsa-dashboard.html': 'rsa',
  'payment-dashboard.html': 'payment',
  'reports-monitoring-dashboard.html': 'reports_monitoring',
  'super-admin-dashboard.html': 'super_admin'
};
const ANNOUNCEMENT_FONT_FAMILIES = {
  system: 'system-ui, -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif',
  arial: 'Arial, Helvetica, sans-serif',
  trebuchet: '"Trebuchet MS", Arial, sans-serif',
  georgia: 'Georgia, "Times New Roman", serif',
  courier: '"Courier New", Courier, monospace',
  verdana: 'Verdana, Geneva, sans-serif',
  tahoma: 'Tahoma, Geneva, sans-serif'
};
let cacheClearInProgress = false;
let securityWatchStarted = false;
let sessionTimeoutTimer = null;
let inactivityHandlersBound = false;
let forceLogoutNoticeActive = false;
let forceLogoutTimerId = null;
let forceLogoutIntervalId = null;
let currentSecurityUserData = {};
let globalReadOnlyState = {
  enabled: false,
  active: false,
  exempt: false,
  role: '',
  message: ''
};
let forceLogoutLockState = {
  active: false,
  token: '',
  deadlineMs: 0
};
let forceLogoutActionBlockersBound = false;
const targetedLogoutWatchedDocIds = new Set();

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

function ensureForceLogoutNoticeStyles() {
  if (document.getElementById('forceLogoutNoticeStyles')) return;
  const style = document.createElement('style');
  style.id = 'forceLogoutNoticeStyles';
  style.textContent = `
    .force-logout-backdrop {
      position: fixed;
      inset: 0;
      background: rgba(2, 6, 23, 0.82);
      z-index: 25000;
      display: flex;
      align-items: center;
      justify-content: center;
      padding: 16px;
    }
    .force-logout-card {
      width: min(540px, 100%);
      background: #fff;
      border-radius: 16px;
      overflow: hidden;
      box-shadow: 0 24px 60px rgba(2, 6, 23, 0.38);
    }
    .force-logout-head {
      padding: 18px 20px;
      background: linear-gradient(135deg, #b91c1c, #dc2626);
      color: #fff;
      font-size: 18px;
      font-weight: 700;
    }
    .force-logout-body {
      padding: 20px;
      color: #334155;
      font-size: 14px;
      line-height: 1.7;
    }
    .force-logout-countdown {
      margin-top: 14px;
      padding: 12px 14px;
      border-radius: 12px;
      background: #fef2f2;
      border: 1px solid #fecaca;
      color: #991b1b;
      font-weight: 700;
    }
    .force-logout-actions {
      margin-top: 16px;
      display: flex;
      justify-content: flex-end;
      gap: 10px;
      flex-wrap: wrap;
    }
    .force-logout-btn {
      border: 0;
      border-radius: 10px;
      padding: 10px 14px;
      font-weight: 700;
      cursor: pointer;
    }
    .force-logout-btn.secondary {
      background: #e2e8f0;
      color: #0f172a;
    }
    .force-logout-btn.primary {
      background: #0f3b67;
      color: #fff;
    }
    .force-logout-banner {
      position: fixed;
      right: 16px;
      bottom: 16px;
      z-index: 24000;
      width: min(420px, calc(100vw - 32px));
      background: #fff7ed;
      border: 1px solid #fdba74;
      color: #9a3412;
      border-radius: 16px;
      box-shadow: 0 20px 40px rgba(15, 23, 42, 0.18);
      padding: 14px 16px;
    }
    .force-logout-banner strong {
      display: block;
      color: #7c2d12;
      margin-bottom: 6px;
    }
    .force-logout-banner-actions {
      margin-top: 12px;
      display: flex;
      gap: 8px;
      flex-wrap: wrap;
    }
    .global-readonly-banner {
      position: sticky;
      top: 0;
      z-index: 19998;
      padding: 12px 18px;
      background: #fff7ed;
      border-bottom: 1px solid #fdba74;
      color: #9a3412;
      box-shadow: 0 10px 24px rgba(15, 23, 42, 0.08);
    }
    .global-readonly-banner strong {
      display: block;
      margin-bottom: 4px;
      color: #7c2d12;
      font-size: 14px;
    }
    .global-readonly-banner span {
      display: block;
      font-size: 13px;
      line-height: 1.5;
    }
  `;
  document.head.appendChild(style);
}

function getGlobalReadOnlyFallbackMessage() {
  return 'Read-only mode is active. You can view records, but changes are temporarily disabled.';
}

function renderGlobalReadOnlyBanner(state = {}) {
  ensureForceLogoutNoticeStyles();
  const existing = document.getElementById('globalReadOnlyBanner');
  if (!state.enabled) {
    existing?.remove();
    return;
  }

  const message = String(state.message || getGlobalReadOnlyFallbackMessage()).trim();
  const roleLabel = state.exempt ? 'Super admin bypass is active on this browser.' : 'Viewing is still available, but changes are blocked.';
  const markup = `
    <strong>Global Read-Only Mode</strong>
    <span>${message}</span>
    <span style="margin-top:4px;opacity:0.9;">${roleLabel}</span>
  `;

  if (existing) {
    existing.innerHTML = markup;
    return;
  }

  const banner = document.createElement('div');
  banner.id = 'globalReadOnlyBanner';
  banner.className = 'global-readonly-banner';
  banner.innerHTML = markup;
  document.body.prepend(banner);
}

function applyGlobalReadOnlyState({ enabled = false, message = '', role = '' } = {}) {
  const normalizedRole = String(role || '').trim().toLowerCase();
  const exempt = normalizedRole === 'super_admin';
  globalReadOnlyState = {
    enabled: enabled === true,
    active: enabled === true && !exempt,
    exempt,
    role: normalizedRole,
    message: String(message || '').trim() || getGlobalReadOnlyFallbackMessage()
  };
  document.body.dataset.appReadOnly = globalReadOnlyState.active ? 'true' : 'false';
  renderGlobalReadOnlyBanner(globalReadOnlyState);
  window.dispatchEvent(new CustomEvent('app-readonly-changed', { detail: { ...globalReadOnlyState } }));
}

function isForceLogoutLockActive() {
  return forceLogoutLockState.active === true;
}

function applyForceLogoutLockState({ active = false, token = '', deadlineMs = 0 } = {}) {
  forceLogoutLockState = {
    active: active === true,
    token: String(token || '').trim(),
    deadlineMs: Number(deadlineMs || 0)
  };
  document.body.dataset.forceLogoutPending = forceLogoutLockState.active ? 'true' : 'false';
  window.dispatchEvent(new CustomEvent('app-force-logout-changed', { detail: { ...forceLogoutLockState } }));
}

function bindForceLogoutActionBlockers() {
  if (forceLogoutActionBlockersBound) return;
  forceLogoutActionBlockersBound = true;

  document.addEventListener('click', (event) => {
    if (!isForceLogoutLockActive()) return;
    const target = event.target instanceof Element ? event.target : null;
    if (!target) return;
    if (target.closest('#forceLogoutNoticeModal, #forceLogoutNoticeBanner, #globalReadOnlyBanner, #systemAnnouncementBanner')) return;
    if (target.closest('#signOutBtnMobile, #signOutBtnSidebar, .sign-out-btn')) return;
    const blockedTrigger = target.closest('button, [role="button"], input[type="submit"], input[type="button"], .action-btn');
    if (!blockedTrigger) return;
    event.preventDefault();
    event.stopPropagation();
    event.stopImmediatePropagation();
    if (typeof window.showNotification === 'function') {
      window.showNotification('Actions are temporarily disabled while logout is pending.', 'warning');
    }
  }, true);

  document.addEventListener('submit', (event) => {
    if (!isForceLogoutLockActive()) return;
    const target = event.target instanceof Element ? event.target : null;
    if (target?.closest('#forceLogoutNoticeModal')) return;
    event.preventDefault();
    event.stopPropagation();
    event.stopImmediatePropagation();
    if (typeof window.showNotification === 'function') {
      window.showNotification('Form actions are temporarily disabled while logout is pending.', 'warning');
    }
  }, true);
}

window.isAppReadOnlyMode = () => globalReadOnlyState.active === true;
window.getAppReadOnlyState = () => ({ ...globalReadOnlyState });
window.isForceLogoutPending = () => isForceLogoutLockActive();
window.assertAppWritable = (actionLabel = 'This action') => {
  if (window.isForceLogoutPending && window.isForceLogoutPending()) {
    const message = `${String(actionLabel || 'This action').trim()} is unavailable while logout is pending.`;
    if (typeof window.showNotification === 'function') {
      window.showNotification(message, 'warning');
    } else {
      window.alert(message);
    }
    return false;
  }
  if (!window.isAppReadOnlyMode()) return true;
  const message = `${String(actionLabel || 'This action').trim()} is unavailable while read-only mode is active.`;
  if (typeof window.showNotification === 'function') {
    window.showNotification(message, 'warning');
  } else {
    window.alert(message);
  }
  return false;
};

function parseForceLogoutDurationMs(value) {
  const text = String(value || '').trim().toLowerCase();
  const normalized = /^\d+\s*[smh]$/.test(text)
    ? text.replace(/\s+/g, '')
    : (/^\d+(\.\d+)?$/.test(text) ? `${text}m` : '');
  const match = normalized.match(/^(\d+(?:\.\d+)?)([smh])$/);
  if (!match) return 11 * 60 * 1000;
  const amount = Number(match[1]);
  const unit = match[2];
  if (!Number.isFinite(amount) || amount <= 0) return 11 * 60 * 1000;
  if (unit === 's') return amount * 1000;
  if (unit === 'm') return amount * 60 * 1000;
  return amount * 60 * 60 * 1000;
}

function getForceLogoutDeadlineMs(token, durationMsValue) {
  const issuedAtMs = Date.parse(String(token || '').trim());
  const baseMs = Number.isFinite(issuedAtMs) ? issuedAtMs : Date.now();
  const durationMs = Math.max(1000, Number(durationMsValue) || (11 * 60 * 1000));
  return baseMs + durationMs;
}

function clearLocalForceLogoutState(token = '') {
  applyForceLogoutLockState({ active: false, token: '', deadlineMs: 0 });
  if (token) {
    localStorage.setItem(FORCE_LOGOUT_COMPLETED_TOKEN_KEY, String(token || '').trim());
  }
  localStorage.removeItem(FORCE_LOGOUT_ACTIVE_TOKEN_KEY);
  localStorage.removeItem(FORCE_LOGOUT_DEADLINE_KEY);
  localStorage.removeItem(FORCE_LOGOUT_DISMISSED_TOKEN_KEY);
  closeForceLogoutNotice();
  document.getElementById('forceLogoutNoticeBanner')?.remove();
  if (forceLogoutTimerId) {
    window.clearTimeout(forceLogoutTimerId);
    forceLogoutTimerId = null;
  }
  if (forceLogoutIntervalId) {
    window.clearInterval(forceLogoutIntervalId);
    forceLogoutIntervalId = null;
  }
}

function isForceLogoutWindowExpired(token = '', countdownSetting = '11m', deadlineMs = 0) {
  const parsedDeadline = Number(deadlineMs || 0);
  const effectiveDeadline = Number.isFinite(parsedDeadline) && parsedDeadline > 0
    ? parsedDeadline
    : getForceLogoutDeadlineMs(token, parseForceLogoutDurationMs(countdownSetting));
  return Date.now() > effectiveDeadline + FORCE_LOGOUT_GRACE_MS;
}

function formatForceLogoutRemaining(deadlineMs) {
  const remainingMs = Math.max(0, Number(deadlineMs || 0) - Date.now());
  const totalSeconds = Math.ceil(remainingMs / 1000);
  const minutes = Math.floor(totalSeconds / 60);
  const seconds = totalSeconds % 60;
  if (minutes <= 0) return `${seconds} second${seconds === 1 ? '' : 's'}`;
  if (seconds === 0) return `${minutes} minute${minutes === 1 ? '' : 's'}`;
  return `${minutes} minute${minutes === 1 ? '' : 's'} ${seconds} second${seconds === 1 ? '' : 's'}`;
}

function closeForceLogoutNotice() {
  document.getElementById('forceLogoutNoticeModal')?.remove();
  forceLogoutNoticeActive = false;
  if (!document.getElementById('profileGuardModal')) {
    document.body.style.overflow = '';
  }
}

function renderForceLogoutBanner({ token = '', deadlineMs = 0 } = {}) {
  ensureForceLogoutNoticeStyles();
  let banner = document.getElementById('forceLogoutNoticeBanner');
  if (!banner) {
    banner = document.createElement('div');
    banner.id = 'forceLogoutNoticeBanner';
    banner.className = 'force-logout-banner';
    banner.innerHTML = `
      <strong>Forced logout scheduled for system update</strong>
      <div id="forceLogoutBannerMessage"></div>
      <div class="force-logout-countdown" id="forceLogoutBannerCountdown" style="margin-top:10px;"></div>
      <div class="force-logout-banner-actions">
        <button id="forceLogoutReopenBtn" class="force-logout-btn primary" type="button">View Notice</button>
      </div>
    `;
    document.body.appendChild(banner);
  }
  const messageEl = document.getElementById('forceLogoutBannerMessage');
  const countdownEl = document.getElementById('forceLogoutBannerCountdown');
  const reopenBtn = document.getElementById('forceLogoutReopenBtn');
  if (messageEl) {
    messageEl.textContent = 'Please finish your current work and save any new submission as draft so you do not lose it before logout.';
  }
  if (countdownEl) {
    countdownEl.textContent = `Automatic logout in ${formatForceLogoutRemaining(deadlineMs)}.`;
  }
  if (reopenBtn) {
    reopenBtn.onclick = () => {
      localStorage.removeItem(FORCE_LOGOUT_DISMISSED_TOKEN_KEY);
      showForceLogoutNotice({ token, deadlineMs });
      forceLogoutNoticeActive = true;
    };
  }
}

function showForceLogoutNotice({ token = '', deadlineMs = 0 } = {}) {
  ensureForceLogoutNoticeStyles();
  closeForceLogoutNotice();
  forceLogoutNoticeActive = true;

  const backdrop = document.createElement('div');
  backdrop.id = 'forceLogoutNoticeModal';
  backdrop.className = 'force-logout-backdrop';
  backdrop.innerHTML = `
    <div class="force-logout-card" role="dialog" aria-modal="true" aria-labelledby="forceLogoutNoticeTitle">
      <div class="force-logout-head" id="forceLogoutNoticeTitle">System Update Notice</div>
      <div class="force-logout-body">
        <div>A force logout has been scheduled so the latest system update can run properly.</div>
        <div style="margin-top:10px;">You can close this notice and finish what you are doing first.</div>
        <div style="margin-top:10px;">Please save any new submission as draft before the logout time so your work is not lost.</div>
        <div class="force-logout-countdown" id="forceLogoutNoticeCountdown"></div>
        <div class="force-logout-actions">
          <button id="forceLogoutCloseNoticeBtn" class="force-logout-btn secondary" type="button">Close Notice</button>
          <button id="forceLogoutKeepOpenBtn" class="force-logout-btn primary" type="button">Continue Working</button>
        </div>
      </div>
    </div>
  `;
  document.body.appendChild(backdrop);
  document.body.style.overflow = 'hidden';

  const dismiss = () => {
    localStorage.setItem(FORCE_LOGOUT_DISMISSED_TOKEN_KEY, String(token || '').trim());
    closeForceLogoutNotice();
  };
  document.getElementById('forceLogoutCloseNoticeBtn')?.addEventListener('click', dismiss);
  document.getElementById('forceLogoutKeepOpenBtn')?.addEventListener('click', dismiss);
  backdrop.addEventListener('click', (event) => {
    if (event.target === backdrop) dismiss();
  });

  const countdownEl = document.getElementById('forceLogoutNoticeCountdown');
  if (countdownEl) {
    countdownEl.textContent = `Automatic logout in ${formatForceLogoutRemaining(deadlineMs)}.`;
  }
}

async function executeForcedLogout(token) {
  await performAppLogout({
    auth,
    beforeSignOut: async () => {
      clearLocalForceLogoutState(token);
    }
  });
}

function startForceLogoutCountdown({ token = '', deadlineMs = 0 } = {}) {
  if (forceLogoutTimerId) {
    window.clearTimeout(forceLogoutTimerId);
    forceLogoutTimerId = null;
  }
  if (forceLogoutIntervalId) {
    window.clearInterval(forceLogoutIntervalId);
    forceLogoutIntervalId = null;
  }

  const updateCountdownUi = () => {
    const bannerCountdownEl = document.getElementById('forceLogoutBannerCountdown');
    const modalCountdownEl = document.getElementById('forceLogoutNoticeCountdown');
    const message = `Automatic logout in ${formatForceLogoutRemaining(deadlineMs)}.`;
    if (bannerCountdownEl) bannerCountdownEl.textContent = message;
    if (modalCountdownEl) modalCountdownEl.textContent = message;
  };

  updateCountdownUi();
  forceLogoutIntervalId = window.setInterval(updateCountdownUi, 1000);

  const remainingMs = Math.max(0, Number(deadlineMs || 0) - Date.now());
  if (remainingMs <= 0) {
    window.clearInterval(forceLogoutIntervalId);
    forceLogoutIntervalId = null;
    executeForcedLogout(token);
    return;
  }
  forceLogoutTimerId = window.setTimeout(() => {
    if (forceLogoutIntervalId) {
      window.clearInterval(forceLogoutIntervalId);
      forceLogoutIntervalId = null;
    }
    executeForcedLogout(token);
  }, remainingMs);
}

function bootstrapPendingForceLogoutFromStorage() {
  const token = String(localStorage.getItem(FORCE_LOGOUT_ACTIVE_TOKEN_KEY) || '').trim();
  const completedToken = String(localStorage.getItem(FORCE_LOGOUT_COMPLETED_TOKEN_KEY) || '').trim();
  const deadlineMs = Number(localStorage.getItem(FORCE_LOGOUT_DEADLINE_KEY) || 0);
  if (!token || token === completedToken || !Number.isFinite(deadlineMs) || deadlineMs <= 0) {
    clearLocalForceLogoutState();
    return;
  }
  if (Date.now() > deadlineMs + FORCE_LOGOUT_GRACE_MS) {
    clearLocalForceLogoutState(token);
    return;
  }
  applyForceLogoutLockState({ active: true, token, deadlineMs });
  renderForceLogoutBanner({ token, deadlineMs });
  startForceLogoutCountdown({ token, deadlineMs });
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
  const targetDashboards = Array.isArray(announcement?.targetDashboards)
    ? announcement.targetDashboards.map((item) => String(item || '').trim().toLowerCase()).filter(Boolean)
    : [];
  const pageName = String(window.location.pathname || '').split('/').pop().toLowerCase() || 'dashboard.html';
  const currentDashboard = DASHBOARD_PAGE_TARGETS[pageName] || String(currentSecurityUserData?.role || '').trim().toLowerCase();
  if (!enabled || !message || (targetDashboards.length && !targetDashboards.includes(currentDashboard))) {
    existing?.remove();
    return;
  }

  const tone = String(announcement?.tone || 'info').trim().toLowerCase();
  const rawSpeed = Number(announcement?.speed ?? getDefaultSystemSettings().dashboardAnnouncement.speed ?? 30);
  const speedSeconds = Number.isFinite(rawSpeed) ? Math.min(60, Math.max(5, rawSpeed)) : 30;
  const rawFontSize = Number(announcement?.fontSize ?? getDefaultSystemSettings().dashboardAnnouncement.fontSize ?? 15);
  const fontSize = Number.isFinite(rawFontSize) ? Math.min(28, Math.max(12, rawFontSize)) : 15;
  const fontStyleSetting = String(announcement?.fontStyle || getDefaultSystemSettings().dashboardAnnouncement.fontStyle || 'bold').trim().toLowerCase();
  const isItalic = fontStyleSetting === 'italic' || fontStyleSetting === 'bold_italic';
  const isBold = fontStyleSetting === 'bold' || fontStyleSetting === 'bold_italic';
  const fontFamilyKey = String(announcement?.fontFamily || getDefaultSystemSettings().dashboardAnnouncement.fontFamily || 'system').trim().toLowerCase();
  const fontFamily = ANNOUNCEMENT_FONT_FAMILIES[fontFamilyKey] || ANNOUNCEMENT_FONT_FAMILIES.system;
  const palette = tone === 'warning'
    ? { bg: '#fff7ed', border: '#fdba74', text: '#9a3412' }
    : tone === 'success'
      ? { bg: '#ecfdf5', border: '#6ee7b7', text: '#065f46' }
      : { bg: '#eff6ff', border: '#93c5fd', text: '#1d4ed8' };
  const customTextColor = /^#[0-9a-f]{6}$/i.test(String(announcement?.textColor || '').trim())
    ? String(announcement.textColor).trim()
    : '';
  const textColor = customTextColor || palette.text;
  const escapedMessage = normalizeText(message);
  const bannerMarkup = `<div class="system-announcement-shell"><span class="system-announcement-badge">Update</span><div class="system-announcement-track"><span class="system-announcement-text">${escapedMessage}</span></div></div>`;

  if (existing) {
    existing.innerHTML = bannerMarkup;
    existing.style.background = palette.bg;
    existing.style.borderBottom = `1px solid ${palette.border}`;
    existing.style.color = textColor;
    existing.style.setProperty('--system-announcement-duration', `${speedSeconds}s`);
    existing.style.setProperty('--system-announcement-font-size', `${fontSize}px`);
    existing.style.setProperty('--system-announcement-font-style', isItalic ? 'italic' : 'normal');
    existing.style.setProperty('--system-announcement-font-weight', isBold ? '800' : '500');
    existing.style.setProperty('--system-announcement-font-family', fontFamily);
    return;
  }

  if (!document.getElementById('systemAnnouncementBannerStyles')) {
    const style = document.createElement('style');
    style.id = 'systemAnnouncementBannerStyles';
    style.textContent = `
      @keyframes systemAnnouncementScroll {
        0% { transform: translateX(0); }
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
        animation: systemAnnouncementScroll var(--system-announcement-duration, 30s) linear infinite;
        font-size: var(--system-announcement-font-size, 15px);
        font-style: var(--system-announcement-font-style, normal);
        font-weight: var(--system-announcement-font-weight, 800);
        font-family: var(--system-announcement-font-family, system-ui, -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif);
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
      #systemAnnouncementBanner {
        pointer-events: none;
      }
      @media (max-width: 768px) {
        #systemAnnouncementBanner {
          position: relative !important;
          top: auto !important;
          z-index: 10 !important;
          padding: 8px 12px !important;
          box-shadow: 0 4px 12px rgba(15,23,42,0.08) !important;
        }
        .system-announcement-shell {
          gap: 8px;
        }
        .system-announcement-badge {
          display: none;
        }
        .system-announcement-text {
          font-size: min(var(--system-announcement-font-size, 15px), 16px);
        }
      }
    `;
    document.head.appendChild(style);
  }

  const banner = document.createElement('div');
  banner.id = 'systemAnnouncementBanner';
  banner.innerHTML = bannerMarkup;
  banner.style.cssText = `position:sticky;top:0;z-index:900;padding:12px 18px;background:${palette.bg};border-bottom:1px solid ${palette.border};color:${textColor};box-shadow:0 10px 24px rgba(15,23,42,0.08);pointer-events:none;`;
  banner.style.setProperty('--system-announcement-duration', `${speedSeconds}s`);
  banner.style.setProperty('--system-announcement-font-size', `${fontSize}px`);
  banner.style.setProperty('--system-announcement-font-style', isItalic ? 'italic' : 'normal');
  banner.style.setProperty('--system-announcement-font-weight', isBold ? '800' : '500');
  banner.style.setProperty('--system-announcement-font-family', fontFamily);
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

function buildLiveSecuritySettings(source = {}, fallback = {}) {
  const sourceSecurity = source?.securityControls && typeof source.securityControls === 'object'
    ? source.securityControls
    : {};
  const fallbackSecurity = fallback?.securityControls && typeof fallback.securityControls === 'object'
    ? fallback.securityControls
    : {};
  const sessionTimeoutMinutes = Math.max(
    1,
    Number(sourceSecurity.sessionTimeoutMinutes ?? fallbackSecurity.sessionTimeoutMinutes ?? 60) || 60
  );
  const forceLogoutCountdownRaw = sourceSecurity.forceLogoutCountdown
    ?? sourceSecurity.forceLogoutCountdownMinutes
    ?? fallbackSecurity.forceLogoutCountdown
    ?? fallbackSecurity.forceLogoutCountdownMinutes
    ?? '11m';
  const forceLogoutToken = String(sourceSecurity.forceLogoutToken ?? fallbackSecurity.forceLogoutToken ?? '').trim();

  return {
    dashboardAnnouncement: source?.dashboardAnnouncement ?? fallback?.dashboardAnnouncement ?? {},
    globalReadOnlyMode: source?.globalReadOnlyMode ?? fallback?.globalReadOnlyMode ?? false,
    globalReadOnlyMessage: String(source?.globalReadOnlyMessage ?? fallback?.globalReadOnlyMessage ?? '').trim(),
    securityControls: {
      sessionTimeoutMinutes,
      forceLogoutCountdown: /^\d+\s*[smh]$/i.test(String(forceLogoutCountdownRaw || '').trim()) || /^\d+(\.\d+)?$/.test(String(forceLogoutCountdownRaw || '').trim())
        ? String(forceLogoutCountdownRaw).trim()
        : String(fallbackSecurity.forceLogoutCountdown || '11m').trim() || '11m',
      forceLogoutToken
    }
  };
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

async function handleCacheClearSignal(data = {}) {
  const token = String(data.cacheClearToken || '').trim();
  const handled = String(localStorage.getItem(CACHE_CLEAR_HANDLED_KEY) || '').trim();
  if (!token || token === handled || cacheClearInProgress) return;

  cacheClearInProgress = true;
  try {
    const securityControls = data?.securityControls && typeof data.securityControls === 'object'
      ? data.securityControls
      : {};
    const forceLogoutToken = String(securityControls.forceLogoutToken || '').trim();
    const completedToken = String(localStorage.getItem(FORCE_LOGOUT_COMPLETED_TOKEN_KEY) || '').trim();
    const countdownSetting = String(securityControls.forceLogoutCountdown || '11m').trim() || '11m';
    if (forceLogoutToken && forceLogoutToken !== completedToken && !isForceLogoutWindowExpired(forceLogoutToken, countdownSetting)) {
      const currentActiveToken = String(localStorage.getItem(FORCE_LOGOUT_ACTIVE_TOKEN_KEY) || '').trim();
      const existingDeadline = Number(localStorage.getItem(FORCE_LOGOUT_DEADLINE_KEY) || 0);
      const nextDeadline = getForceLogoutDeadlineMs(forceLogoutToken, parseForceLogoutDurationMs(countdownSetting));
      if (currentActiveToken !== forceLogoutToken || !Number.isFinite(existingDeadline) || existingDeadline <= 0) {
        localStorage.setItem(FORCE_LOGOUT_ACTIVE_TOKEN_KEY, forceLogoutToken);
        localStorage.setItem(FORCE_LOGOUT_DEADLINE_KEY, String(nextDeadline));
        localStorage.removeItem(FORCE_LOGOUT_DISMISSED_TOKEN_KEY);
      }
    }
  } catch (_) {}

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
}

function watchForSecuritySignals(userData = {}) {
  currentSecurityUserData = {
    ...currentSecurityUserData,
    ...(userData && typeof userData === 'object' ? userData : {})
  };
  if (securityWatchStarted) return;
  securityWatchStarted = true;

  const evaluateSecurityState = async (systemSettings) => {
    showDashboardAnnouncement(systemSettings.dashboardAnnouncement);
    bindInactivityHandlers(systemSettings.securityControls.sessionTimeoutMinutes);
    applyGlobalReadOnlyState({
      enabled: systemSettings.globalReadOnlyMode === true,
      message: systemSettings.globalReadOnlyMessage || '',
      role: currentSecurityUserData?.role || ''
    });

    const configuredForceLogoutToken = String(systemSettings.securityControls.forceLogoutToken || '').trim();
    const completed = String(localStorage.getItem(FORCE_LOGOUT_COMPLETED_TOKEN_KEY) || '').trim();
    const role = String(currentSecurityUserData?.role || '-').trim() || '-';
    const countdownSetting = String(systemSettings.securityControls.forceLogoutCountdown || '11m');
    const activeTokenBefore = String(localStorage.getItem(FORCE_LOGOUT_ACTIVE_TOKEN_KEY) || '').trim();
    const deadlineBefore = Number(localStorage.getItem(FORCE_LOGOUT_DEADLINE_KEY) || 0);
    const forceLogoutToken = configuredForceLogoutToken;
    const hasPersistedDeadline = Number.isFinite(deadlineBefore) && deadlineBefore > 0;
    if (!forceLogoutToken || isMaintenanceExemptRole(userData.role) || forceLogoutToken === completed) {
      clearLocalForceLogoutState(forceLogoutToken === completed ? forceLogoutToken : '');
      return;
    }

    if (isForceLogoutWindowExpired(forceLogoutToken, countdownSetting, activeTokenBefore === forceLogoutToken ? deadlineBefore : 0)) {
      clearLocalForceLogoutState(forceLogoutToken);
      return;
    }

    const countdownDurationMs = parseForceLogoutDurationMs(countdownSetting);
    const activeToken = String(localStorage.getItem(FORCE_LOGOUT_ACTIVE_TOKEN_KEY) || '').trim();
    let deadlineMs = Number(localStorage.getItem(FORCE_LOGOUT_DEADLINE_KEY) || 0);
    if (activeToken !== forceLogoutToken || !Number.isFinite(deadlineMs) || deadlineMs <= 0) {
      deadlineMs = getForceLogoutDeadlineMs(forceLogoutToken, countdownDurationMs);
      localStorage.setItem(FORCE_LOGOUT_ACTIVE_TOKEN_KEY, forceLogoutToken);
      localStorage.setItem(FORCE_LOGOUT_DEADLINE_KEY, String(deadlineMs));
      localStorage.removeItem(FORCE_LOGOUT_DISMISSED_TOKEN_KEY);
      forceLogoutNoticeActive = false;
    } else if (hasPersistedDeadline) {
      deadlineMs = deadlineBefore;
    }

    applyForceLogoutLockState({ active: true, token: forceLogoutToken, deadlineMs });
    renderForceLogoutBanner({ token: forceLogoutToken, deadlineMs });
    startForceLogoutCountdown({ token: forceLogoutToken, deadlineMs });

    const dismissedToken = String(localStorage.getItem(FORCE_LOGOUT_DISMISSED_TOKEN_KEY) || '').trim();
    if (!forceLogoutNoticeActive && dismissedToken !== forceLogoutToken) {
      forceLogoutNoticeActive = true;
      showForceLogoutNotice({ token: forceLogoutToken, deadlineMs });
    }

    if (dismissedToken === forceLogoutToken) {
      forceLogoutNoticeActive = false;
    }
  };

  onSnapshot(doc(db, 'settings', 'system'), async (snap) => {
    try {
      const rawSettings = snap.exists() ? (snap.data() || {}) : {};
      await handleCacheClearSignal(rawSettings);
      const systemSettings = buildLiveSecuritySettings(rawSettings, getDefaultSystemSettings());
      await evaluateSecurityState(systemSettings);
    } catch (_) {}
  }, () => {});

}

function watchForTargetedLogout(userDocId = '') {
  const normalizedUserDocId = String(userDocId || '').trim();
  if (!normalizedUserDocId || targetedLogoutWatchedDocIds.has(normalizedUserDocId)) return;
  targetedLogoutWatchedDocIds.add(normalizedUserDocId);

  onSnapshot(doc(db, 'users', normalizedUserDocId), async (snap) => {
    try {
      if (!snap.exists()) return;
      const data = snap.data() || {};
      const token = String(data.targetedForceLogoutToken || '').trim();
      if (!token) return;

      const completedOnDoc = String(data.targetedForceLogoutCompletedToken || '').trim();
      const localCompletedKey = `${TARGETED_LOGOUT_COMPLETED_PREFIX}${normalizedUserDocId}`;
      const completedLocally = String(localStorage.getItem(localCompletedKey) || '').trim();
      if (completedOnDoc === token || completedLocally === token) return;

      localStorage.setItem(localCompletedKey, token);
      try {
        await updateDoc(doc(db, 'users', normalizedUserDocId), {
          targetedForceLogoutCompletedToken: token,
          targetedForceLogoutCompletedAt: serverTimestamp(),
          isOnline: false,
          lastLogoutAt: serverTimestamp()
        });
      } catch (_) {}

      await executeForcedLogout(token);
    } catch (_) {}
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

async function findUserDocsByUidOrEmail(uid, email) {
  const docsById = new Map();
  const normalizedUid = String(uid || '').trim();
  const normalizedEmail = String(email || '').trim().toLowerCase();

  if (normalizedUid) {
    const directDoc = await getDoc(doc(db, 'users', normalizedUid)).catch(() => null);
    if (directDoc?.exists?.()) docsById.set(directDoc.id, directDoc);

    const byUidSnap = await getDocs(query(collection(db, 'users'), where('uid', '==', normalizedUid))).catch(() => null);
    byUidSnap?.docs?.forEach((docSnap) => docsById.set(docSnap.id, docSnap));
  }

  if (normalizedEmail) {
    const byEmailSnap = await getDocs(query(collection(db, 'users'), where('email', '==', normalizedEmail))).catch(() => null);
    byEmailSnap?.docs?.forEach((docSnap) => docsById.set(docSnap.id, docSnap));
  }

  return Array.from(docsById.values());
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
  bindForceLogoutActionBlockers();
  bootstrapPendingForceLogoutFromStorage();

  const fallbackUserData = {
    role: '',
    email: String(user.email || '').trim().toLowerCase()
  };

  try {
    const systemSettings = await getSystemSettings(db, { force: true });
    showDashboardAnnouncement(systemSettings.dashboardAnnouncement);
    bindInactivityHandlers(systemSettings.securityControls.sessionTimeoutMinutes);
    applyGlobalReadOnlyState({
      enabled: systemSettings.globalReadOnlyMode === true,
      message: systemSettings.globalReadOnlyMessage || '',
      role: fallbackUserData.role || ''
    });
  } catch (_) {
    bindInactivityHandlers(60);
  }

  watchForSecuritySignals(fallbackUserData);

  try {
    const userDoc = await findUserDocByUidOrEmail(user.uid, user.email);
    const userData = userDoc?.data?.() || fallbackUserData;
    watchForSecuritySignals(userData);

    try {
      const systemSettings = await getSystemSettings(db);
      showDashboardAnnouncement(systemSettings.dashboardAnnouncement);
      bindInactivityHandlers(systemSettings.securityControls.sessionTimeoutMinutes);
      applyGlobalReadOnlyState({
        enabled: systemSettings.globalReadOnlyMode === true,
        message: systemSettings.globalReadOnlyMessage || '',
        role: userData.role || ''
      });
    } catch (_) {}

    try {
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
    } catch (_) {}

    const logoutWatchDocs = await findUserDocsByUidOrEmail(user.uid, user.email).catch(() => []);
    logoutWatchDocs.forEach((docSnap) => {
      if (docSnap?.id) watchForTargetedLogout(docSnap.id);
    });
    if (userDoc?.id && !targetedLogoutWatchedDocIds.has(userDoc.id)) {
      watchForTargetedLogout(userDoc.id);
    }

    if (userDoc?.id) {
      try {
        await registerPushTokenForCurrentUser(user, userDoc.id);
      } catch (_) {}
      try {
        await enforceProfileCompletion(user);
      } catch (_) {}
    }
  } catch (_err) {
    // Fail open to avoid blocking legitimate users on transient issues.
  }
});
