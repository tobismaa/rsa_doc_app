import {
  collection,
  doc,
  getDoc,
  getDocs,
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

function getProfileStatusRank(profile = {}) {
  const status = String(profile.status || 'active').trim().toLowerCase();
  if (status === 'active') return 0;
  if (status === 'pending') return 1;
  if (status === 'deactivated') return 2;
  return 3;
}

function getProfileRoleRank(profile = {}) {
  const role = String(profile.role || '').trim().toLowerCase();
  if (role === 'super_admin') return 0;
  if (role === 'admin') return 1;
  if (role === 'reports_monitoring') return 2;
  if (role === 'reviewer') return 3;
  if (role === 'rsa') return 4;
  if (role === 'payment') return 5;
  if (role === 'uploader') return 6;
  return 7;
}

function chooseBestUserDoc(docs = [], preferredUid = '') {
  if (!Array.isArray(docs) || !docs.length) return null;
  const normalizedPreferredUid = String(preferredUid || '').trim();
  if (normalizedPreferredUid) {
    const exactUidDoc = docs.find((snap) => String(snap.data()?.uid || '').trim() === normalizedPreferredUid || String(snap.id || '').trim() === normalizedPreferredUid);
    if (exactUidDoc) return exactUidDoc;
  }

  return docs
    .slice()
    .sort((a, b) => {
      const aData = a.data() || {};
      const bData = b.data() || {};
      const statusDiff = getProfileStatusRank(aData) - getProfileStatusRank(bData);
      if (statusDiff !== 0) return statusDiff;
      const roleDiff = getProfileRoleRank(aData) - getProfileRoleRank(bData);
      if (roleDiff !== 0) return roleDiff;
      return 0;
    })[0] || null;
}

async function findBestUserDoc(db, { email = '', preferredUid = '' } = {}) {
  const normalizedEmail = normalizeEmail(email);
  const normalizedPreferredUid = String(preferredUid || '').trim();

  if (normalizedPreferredUid) {
    try {
      const uidRef = doc(db, 'users', normalizedPreferredUid);
      const uidSnap = await getDoc(uidRef);
      if (uidSnap.exists()) return uidSnap;
    } catch (_) {
      // Continue to query fallbacks.
    }

    try {
      const uidQuery = query(collection(db, 'users'), where('uid', '==', normalizedPreferredUid));
      const uidSnapshot = await getDocs(uidQuery);
      if (!uidSnapshot.empty) {
        return chooseBestUserDoc(uidSnapshot.docs, normalizedPreferredUid);
      }
    } catch (_) {
      // Continue to email and scan fallbacks.
    }
  }

  if (normalizedEmail) {
    try {
      const userQuery = query(collection(db, 'users'), where('email', '==', normalizedEmail));
      const snapshot = await getDocs(userQuery);
      if (!snapshot.empty) {
        return chooseBestUserDoc(snapshot.docs, normalizedPreferredUid);
      }
    } catch (_) {
      // Continue to collection scan fallback.
    }
  }

  if (!normalizedPreferredUid && !normalizedEmail) return null;

  try {
    const snapshot = await getDocs(collection(db, 'users'));
    const matchedDocs = snapshot.docs.filter((snap) => {
      const data = snap.data() || {};
      const docId = String(snap.id || '').trim();
      const dataUid = String(data.uid || '').trim();
      const dataEmail = normalizeEmail(data.email);
      if (normalizedPreferredUid && (docId === normalizedPreferredUid || dataUid === normalizedPreferredUid)) {
        return true;
      }
      if (normalizedEmail && dataEmail === normalizedEmail) {
        return true;
      }
      return false;
    });
    return chooseBestUserDoc(matchedDocs, normalizedPreferredUid);
  } catch (_) {
    return null;
  }
}

export async function getUserProfileByEmail(db, email, options = {}) {
  const normalized = normalizeEmail(email);
  const cacheKey = options?.preferredUid ? `${normalized}::${String(options.preferredUid).trim()}` : normalized;
  if (!normalized) return null;
  if (userProfileCache.has(cacheKey)) return userProfileCache.get(cacheKey);
  try {
    const bestDoc = await findBestUserDoc(db, { email: normalized, preferredUid: options?.preferredUid });
    const profile = bestDoc ? { __docId: bestDoc.id, ...(bestDoc.data() || {}) } : null;
    if (profile) {
      userProfileCache.set(cacheKey, profile);
      userProfileCache.set(normalized, profile);
    }
    return profile;
  } catch (_) {
    return null;
  }
}

export async function getCurrentUserProfile(db, user) {
  if (!user) return null;
  const normalizedEmail = normalizeEmail(user.email);
  const preferredUid = String(user.uid || '').trim();
  const cacheKey = normalizedEmail ? `${normalizedEmail}::${preferredUid}` : preferredUid;
  if (cacheKey && userProfileCache.has(cacheKey)) return userProfileCache.get(cacheKey);

  try {
    const bestDoc = await findBestUserDoc(db, { email: normalizedEmail, preferredUid });
    const profile = bestDoc ? { __docId: bestDoc.id, ...(bestDoc.data() || {}) } : null;
    if (profile) {
      if (cacheKey) userProfileCache.set(cacheKey, profile);
      if (normalizedEmail) userProfileCache.set(normalizedEmail, profile);
    }
    return profile;
  } catch (_) {
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
    userProfileCache.set(normalized, profile);
    const role = String(profile.role || '').toLowerCase();
    const status = String(profile.status || 'active').toLowerCase();
    const leaveStatus = String(profile.leaveStatus || '').toLowerCase();
    if (status === 'deactivated') return false;
    if (leaveStatus === 'on_leave') return false;
    return allowedRoles.includes(role);
  } catch (_) {
    return false;
  }
}

export function clearUserDirectoryCaches() {
  userFullNameCache.clear();
  userProfileCache.clear();
}
