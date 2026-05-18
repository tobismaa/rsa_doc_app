import {
  doc,
  getDoc
} from "https://www.gstatic.com/firebasejs/9.22.0/firebase-firestore.js";

export const LEGACY_COMMISSION_RATE = 0.02;
export const DEFAULT_COMMISSION_RATE = 0.01;
export const DEFAULT_COMMISSION_EFFECTIVE_FROM_ISO = '2026-05-07T00:00:00+01:00';

let commissionSettingsCache = null;
let commissionSettingsFetchedAt = 0;

function parseRate(value, fallback = DEFAULT_COMMISSION_RATE) {
  const num = Number(value);
  return Number.isFinite(num) && num >= 0 ? num : fallback;
}

function toMillis(value) {
  if (!value) return 0;
  try {
    if (typeof value.toMillis === 'function') return value.toMillis();
    if (typeof value.toDate === 'function') return value.toDate().getTime();
    if (typeof value.seconds === 'number') return value.seconds * 1000;
    const parsed = new Date(value).getTime();
    return Number.isFinite(parsed) ? parsed : 0;
  } catch (_) {
    return 0;
  }
}

function trimTrailingZeros(value) {
  return String(value).replace(/\.0+$/, '').replace(/(\.\d*?)0+$/, '$1');
}

export function formatCommissionRateLabel(rate) {
  const percent = Number(rate || 0) * 100;
  return `${trimTrailingZeros(percent.toFixed(2))}%`;
}

export function getDefaultCommissionSettings() {
  const effectiveFromMs = toMillis(DEFAULT_COMMISSION_EFFECTIVE_FROM_ISO);
  return {
    rate: DEFAULT_COMMISSION_RATE,
    effectiveFromIso: DEFAULT_COMMISSION_EFFECTIVE_FROM_ISO,
    effectiveFromMs,
    rateLabel: formatCommissionRateLabel(DEFAULT_COMMISSION_RATE),
    legacyRate: LEGACY_COMMISSION_RATE
  };
}

export async function getCommissionSettings(db, { force = false } = {}) {
  const now = Date.now();
  if (!force && commissionSettingsCache && (now - commissionSettingsFetchedAt) < 60 * 1000) {
    return commissionSettingsCache;
  }

  const defaults = getDefaultCommissionSettings();
  try {
    const snap = await getDoc(doc(db, 'settings', 'system'));
    const data = snap.exists() ? (snap.data() || {}) : {};
    const rate = parseRate(data.commissionRate, defaults.rate);
    const effectiveFromIso = String(data.commissionRateEffectiveFrom || defaults.effectiveFromIso).trim() || defaults.effectiveFromIso;
    const effectiveFromMs = toMillis(effectiveFromIso) || defaults.effectiveFromMs;
    commissionSettingsCache = {
      rate,
      effectiveFromIso,
      effectiveFromMs,
      rateLabel: formatCommissionRateLabel(rate),
      legacyRate: LEGACY_COMMISSION_RATE
    };
  } catch (_) {
    commissionSettingsCache = defaults;
  }

  commissionSettingsFetchedAt = now;
  return commissionSettingsCache;
}

export function clearCommissionSettingsCache() {
  commissionSettingsCache = null;
  commissionSettingsFetchedAt = 0;
}

export function resolveCommissionRateForTimestamp(timestampMs, settings = null) {
  const config = settings || getDefaultCommissionSettings();
  const effectiveFromMs = Number(config.effectiveFromMs || 0) || getDefaultCommissionSettings().effectiveFromMs;
  const rate = parseRate(config.rate, getDefaultCommissionSettings().rate);
  if (Number.isFinite(timestampMs) && timestampMs > 0 && timestampMs < effectiveFromMs) {
    return LEGACY_COMMISSION_RATE;
  }
  return rate;
}

export function resolveSubmissionCommissionRate(submission = null, settings = null) {
  const storedRate = Number(submission?.commissionRate);
  if (Number.isFinite(storedRate) && storedRate >= 0) return storedRate;

  const submittedMs = Math.max(
    toMillis(submission?.submittedAt),
    toMillis(submission?.uploadedAt),
    toMillis(submission?.createdAt)
  );
  return resolveCommissionRateForTimestamp(submittedMs, settings || getDefaultCommissionSettings());
}

export function buildSubmissionCommissionFields(settings = null, nowMs = Date.now()) {
  const config = settings || getDefaultCommissionSettings();
  const rate = resolveCommissionRateForTimestamp(nowMs, config);
  return {
    commissionRate: rate,
    commissionRatePercent: Number((rate * 100).toFixed(4)),
    commissionRateLabel: formatCommissionRateLabel(rate),
    commissionRateEffectiveFrom: config.effectiveFromIso || getDefaultCommissionSettings().effectiveFromIso,
    commissionRateAssignedAtIso: new Date(nowMs).toISOString()
  };
}

export function getSubmissionCommissionAmount(submission = null, twentyFiveAmount = 0, settings = null) {
  const rate = resolveSubmissionCommissionRate(submission, settings || getDefaultCommissionSettings());
  return Number(twentyFiveAmount || 0) * rate;
}
