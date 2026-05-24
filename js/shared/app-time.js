import { auth, db } from '../firebase-config.js';
import { doc, getDoc, serverTimestamp, setDoc } from 'https://www.gstatic.com/firebasejs/9.22.0/firebase-firestore.js';

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

  lastFetchAttemptPerfMs = nowPerfMs;

  try {
    const currentUid = String(auth.currentUser?.uid || 'public').trim() || 'public';
    const clockRef = doc(db, 'runtimeServerClock', currentUid);
    await setDoc(clockRef, {
      serverNow: serverTimestamp(),
      syncedAt: serverTimestamp()
    }, { merge: true });
    const snap = await getDoc(clockRef);
    const serverNow = snap.exists() ? snap.data()?.serverNow : null;
    const resolvedDate = normalizeToDate(serverNow);
    if (resolvedDate) {
      cachedServerEpochMs = resolvedDate.getTime();
      cachedFetchPerfMs = getPerfNow();
      return resolvedDate;
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
