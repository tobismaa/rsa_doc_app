import {
  collection,
  getDocs,
  limit,
  query,
  where
} from "https://www.gstatic.com/firebasejs/9.22.0/firebase-firestore.js";

const userFullNameCache = new Map();
const userProfileCache = new Map();

export function normalizeEmail(email) {
  return String(email || '').trim().toLowerCase();
}

function fallbackName(email) {
  return normalizeEmail(email).split('@')[0] || 'Unknown';
}

export async function getUserProfileByEmail(db, email) {
  const normalized = normalizeEmail(email);
  if (!normalized) return null;
  if (userProfileCache.has(normalized)) return userProfileCache.get(normalized);
  try {
    const userQuery = query(collection(db, 'users'), where('email', '==', normalized), limit(1));
    const snapshot = await getDocs(userQuery);
    const profile = snapshot.empty ? null : (snapshot.docs[0].data() || null);
    userProfileCache.set(normalized, profile);
    return profile;
  } catch (_) {
    userProfileCache.set(normalized, null);
    return null;
  }
}

export async function getUserFullName(db, email) {
  const normalized = normalizeEmail(email);
  if (!normalized) return 'Unknown';
  if (userFullNameCache.has(normalized)) return userFullNameCache.get(normalized);
  const profile = await getUserProfileByEmail(db, normalized);
  const fullName = profile?.fullName || profile?.displayName || fallbackName(normalized);
  userFullNameCache.set(normalized, fullName);
  return fullName;
}

export async function ensureUserFullNames(db, emails) {
  if (!Array.isArray(emails) || !emails.length) return;
  await Promise.all(emails.map(async (email) => {
    const normalized = normalizeEmail(email);
    if (!normalized || userFullNameCache.has(normalized)) return;
    const fullName = await getUserFullName(db, normalized);
    userFullNameCache.set(normalized, fullName);
  }));
}

export async function isActiveUserWithRole(db, email, allowedRoles = []) {
  const normalized = normalizeEmail(email);
  if (!normalized) return false;
  try {
    const profile = await getUserProfileByEmail(db, normalized);
    if (!profile) return false;
    const role = String(profile.role || '').toLowerCase();
    const status = String(profile.status || 'active').toLowerCase();
    if (status === 'deactivated') return false;
    return allowedRoles.includes(role);
  } catch (_) {
    return false;
  }
}

export function clearUserDirectoryCaches() {
  userFullNameCache.clear();
  userProfileCache.clear();
}
