import { EMAIL_API_BASE_URL } from '../email-api-config.js';
import { ADMIN_API_BASE_URL } from '../admin-api-config.js';

export const APP_TIME_ZONE = 'Africa/Lagos';
const DATE_TIME_FORMATTER = new Intl.DateTimeFormat('en-NG', {
  timeZone: APP_TIME_ZONE,
  day: '2-digit',
  month: '2-digit',
  year: 'numeric',
  hour: '2-digit',
  minute: '2-digit',
  hour12: true
});
const DATE_ONLY_FORMATTER = new Intl.DateTimeFormat('en-NG', {
  timeZone: APP_TIME_ZONE,
  day: '2-digit',
  month: 'short',
  year: 'numeric'
});
const LAGOS_DATE_KEY_FORMATTER = new Intl.DateTimeFormat('en-CA', {
  timeZone: APP_TIME_ZONE,
  year: 'numeric',
  month: '2-digit',
  day: '2-digit'
});

let cachedServerEpochMs = 0;
let cachedFetchPerfMs = 0;
let lastFetchAttemptPerfMs = 0;
const SERVER_TIME_TTL_MS = 60 * 1000;

function resolveConfiguredBaseUrl(raw) {
  const value = String(raw || '').trim();
  if (!value || value.includes('YOUR-RENDER-URL')) return '';
  return value.replace(/\/+$/, '');
}

function getTrustedTimeApiBaseUrl() {
  return resolveConfiguredBaseUrl(window.__EMAIL_API_BASE_URL__)
    || resolveConfiguredBaseUrl(window.__ADMIN_API_BASE_URL__)
    || resolveConfiguredBaseUrl(EMAIL_API_BASE_URL)
    || resolveConfiguredBaseUrl(ADMIN_API_BASE_URL);
}

function getPerfNow() {
  if (typeof performance !== 'undefined' && typeof performance.now === 'function') {
    return performance.now();
  }
  return Date.now();
}

function normalizeToDate(value) {
  if (!value) return null;
  if (value instanceof Date) return Number.isNaN(value.getTime()) ? null : value;
  try {
    if (typeof value?.toDate === 'function') {
      const date = value.toDate();
      return Number.isNaN(date?.getTime?.()) ? null : date;
    }
    if (typeof value?.seconds === 'number') {
      const date = new Date(value.seconds * 1000);
      return Number.isNaN(date.getTime()) ? null : date;
    }
    const date = new Date(value);
    return Number.isNaN(date.getTime()) ? null : date;
  } catch (_) {
    return null;
  }
}

export function formatAppDateTime(value, fallback = 'N/A') {
  const date = normalizeToDate(value);
  if (!date) return fallback;
  try {
    return DATE_TIME_FORMATTER.format(date);
  } catch (_) {
    return fallback;
  }
}

export function formatAppDate(value, fallback = 'N/A') {
  const date = normalizeToDate(value);
  if (!date) return fallback;
  try {
    return DATE_ONLY_FORMATTER.format(date);
  } catch (_) {
    return fallback;
  }
}

export async function getTrustedNow({ force = false } = {}) {
  const nowPerfMs = getPerfNow();
  if (!force && cachedServerEpochMs > 0 && (nowPerfMs - cachedFetchPerfMs) <= SERVER_TIME_TTL_MS) {
    return new Date(cachedServerEpochMs + (nowPerfMs - cachedFetchPerfMs));
  }
  if (!force && lastFetchAttemptPerfMs > 0 && (nowPerfMs - lastFetchAttemptPerfMs) <= 3000 && cachedServerEpochMs > 0) {
    return new Date(cachedServerEpochMs + (nowPerfMs - cachedFetchPerfMs));
  }

  const baseUrl = getTrustedTimeApiBaseUrl();
  lastFetchAttemptPerfMs = nowPerfMs;

  if (baseUrl) {
    try {
      const response = await fetch(`${baseUrl}/api/server-time`, {
        method: 'GET',
        headers: { 'Accept': 'application/json' }
      });
      if (response.ok) {
        const payload = await response.json();
        const epochMs = Number(payload?.epochMs || 0);
        if (Number.isFinite(epochMs) && epochMs > 0) {
          cachedServerEpochMs = epochMs;
          cachedFetchPerfMs = getPerfNow();
          return new Date(epochMs);
        }
      }
    } catch (_) {}
  }

  try {
    const response = await fetch('/api/server-time', {
      method: 'GET',
      headers: { 'Accept': 'application/json' }
    });
    if (response.ok) {
      const payload = await response.json();
      const epochMs = Number(payload?.epochMs || 0);
      if (Number.isFinite(epochMs) && epochMs > 0) {
        cachedServerEpochMs = epochMs;
        cachedFetchPerfMs = getPerfNow();
        return new Date(epochMs);
      }
    }
  } catch (_) {}

  const fallback = new Date();
  cachedServerEpochMs = fallback.getTime();
  cachedFetchPerfMs = getPerfNow();
  return fallback;
}

export async function getTrustedDateKey(options = {}) {
  const trustedNow = await getTrustedNow(options);
  return LAGOS_DATE_KEY_FORMATTER.format(trustedNow);
}

export async function getTrustedNowIso(options = {}) {
  const trustedNow = await getTrustedNow(options);
  return trustedNow.toISOString();
}
