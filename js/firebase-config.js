// Import the functions you need from the SDKs you need
import { getApps, initializeApp } from "https://www.gstatic.com/firebasejs/9.22.0/firebase-app.js";
import { getAuth, onAuthStateChanged } from "https://www.gstatic.com/firebasejs/9.22.0/firebase-auth.js";
import {
  getFirestore,
  collection,
  query,
  where,
  orderBy,
  onSnapshot,
  getDocs,
  getDoc,
  doc,
  updateDoc,
  deleteDoc,
  addDoc,
  serverTimestamp
} from "https://www.gstatic.com/firebasejs/9.22.0/firebase-firestore.js";
import { getAnalytics, isSupported } from "https://www.gstatic.com/firebasejs/9.22.0/firebase-analytics.js";
import { performAppLogout } from './shared/logout.js?v=20260625b';

// Your web app's Firebase configuration
const firebaseConfig = {
  apiKey: "AIzaSyABQ-RR3Mlot7Vz2_s06AcFp3AlHb6elmw",
  authDomain: "rsa-doc-app.firebaseapp.com",
  projectId: "rsa-doc-app",
  storageBucket: "rsa-doc-app.firebasestorage.app",
  messagingSenderId: "749343098749",
  appId: "1:749343098749:web:ed78989a0b2c620d156e14",
  measurementId: "G-KQHMRNDZ6X"
};

// Initialize Firebase
const app = getApps()[0] || initializeApp(firebaseConfig);
const auth = getAuth(app);
const db = getFirestore(app);
let analytics = null;
const targetedLogoutWatchedDocIds = new Set();
const PRESENCE_HEARTBEAT_MS = 60 * 1000;
let presenceHeartbeatTimer = null;
let presenceDocIds = [];

const isLocalDevHost =
  window.location.hostname === "127.0.0.1" ||
  window.location.hostname === "localhost" ||
  window.location.protocol === "file:";

if (!isLocalDevHost) {
  isSupported()
    .then((supported) => {
      if (supported) {
        analytics = getAnalytics(app);
      }
    })
    .catch(() => {
      analytics = null;
    });
}

async function findTargetedLogoutDocs(user) {
  const docsById = new Map();
  const uid = String(user?.uid || '').trim();
  const email = String(user?.email || '').trim().toLowerCase();

  if (uid) {
    const directDoc = await getDoc(doc(db, 'users', uid)).catch(() => null);
    if (directDoc?.exists?.()) docsById.set(directDoc.id, directDoc);

    const uidSnap = await getDocs(query(collection(db, 'users'), where('uid', '==', uid))).catch(() => null);
    uidSnap?.docs?.forEach((docSnap) => docsById.set(docSnap.id, docSnap));
  }

  if (email) {
    const emailSnap = await getDocs(query(collection(db, 'users'), where('email', '==', email))).catch(() => null);
    emailSnap?.docs?.forEach((docSnap) => docsById.set(docSnap.id, docSnap));
  }

  return Array.from(docsById.values());
}

function watchTargetedLogoutDoc(userDocId = '') {
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
      const completedKey = `cmbank_targeted_logout_completed_${normalizedUserDocId}`;
      const completedLocally = String(localStorage.getItem(completedKey) || '').trim();
      if (completedOnDoc === token || completedLocally === token) return;

      localStorage.setItem(completedKey, token);
      await performAppLogout({
        auth,
        beforeSignOut: async () => {
          await updateDoc(doc(db, 'users', normalizedUserDocId), {
            targetedForceLogoutCompletedToken: token,
            targetedForceLogoutCompletedAt: serverTimestamp(),
            isOnline: false,
            lastLogoutAt: serverTimestamp()
          }).catch(() => {});
        }
      });
    } catch (_) {}
  }, () => {});
}

async function writePresence(docIds = [], isOnline = true) {
  const uniqueDocIds = Array.from(new Set(docIds.map((id) => String(id || '').trim()).filter(Boolean)));
  if (uniqueDocIds.length === 0) return;

  await Promise.all(uniqueDocIds.map((userDocId) => updateDoc(doc(db, 'users', userDocId), {
    isOnline,
    lastSeenAt: serverTimestamp(),
    ...(isOnline ? {} : { lastLogoutAt: serverTimestamp() })
  }).catch(() => {})));
}

function startPresenceHeartbeat(docIds = []) {
  presenceDocIds = Array.from(new Set(docIds.map((id) => String(id || '').trim()).filter(Boolean)));
  if (presenceDocIds.length === 0) return;

  writePresence(presenceDocIds, true).catch(() => {});
  if (presenceHeartbeatTimer) window.clearInterval(presenceHeartbeatTimer);
  presenceHeartbeatTimer = window.setInterval(() => {
    if (document.visibilityState === 'hidden') return;
    writePresence(presenceDocIds, true).catch(() => {});
  }, PRESENCE_HEARTBEAT_MS);
}

function bindPresenceEvents() {
  if (window.__cmbankPresenceEventsBound) return;
  window.__cmbankPresenceEventsBound = true;

  window.addEventListener('focus', () => {
    writePresence(presenceDocIds, true).catch(() => {});
  });

  document.addEventListener('visibilitychange', () => {
    if (document.visibilityState === 'visible') {
      writePresence(presenceDocIds, true).catch(() => {});
    }
  });

  window.addEventListener('pagehide', () => {
    writePresence(presenceDocIds, false).catch(() => {});
  });
}

onAuthStateChanged(auth, async (user) => {
  if (!user) {
    if (presenceHeartbeatTimer) {
      window.clearInterval(presenceHeartbeatTimer);
      presenceHeartbeatTimer = null;
    }
    presenceDocIds = [];
    return;
  }
  const docs = await findTargetedLogoutDocs(user).catch(() => []);
  docs.forEach((docSnap) => {
    if (docSnap?.id) watchTargetedLogoutDoc(docSnap.id);
  });
  startPresenceHeartbeat(docs.map((docSnap) => docSnap.id));
  bindPresenceEvents();
});

export {
  app,
  auth,
  db,
  analytics,
  collection,
  query,
  where,
  orderBy,
  onSnapshot,
  getDocs,
  getDoc,
  doc,
  updateDoc,
  deleteDoc,
  addDoc,
  serverTimestamp
};
