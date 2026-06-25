import { signOut } from "https://www.gstatic.com/firebasejs/9.22.0/firebase-auth.js";

const LOGOUT_TIMEOUT_MS = 3500;

function delay(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

async function deleteIndexedDb(name) {
  if (!('indexedDB' in window)) return;
  await new Promise((resolve) => {
    const request = indexedDB.deleteDatabase(name);
    request.onsuccess = () => resolve();
    request.onerror = () => resolve();
    request.onblocked = () => resolve();
  });
}

async function clearIndexedDbStore(dbName, storeName) {
  if (!('indexedDB' in window)) return;
  await new Promise((resolve) => {
    const request = indexedDB.open(dbName);
    request.onerror = () => resolve();
    request.onupgradeneeded = () => {
      try {
        request.transaction?.abort();
      } catch (_) {}
      resolve();
    };
    request.onsuccess = () => {
      const database = request.result;
      try {
        if (!database.objectStoreNames.contains(storeName)) {
          database.close();
          resolve();
          return;
        }
        const transaction = database.transaction(storeName, 'readwrite');
        transaction.objectStore(storeName).clear();
        transaction.oncomplete = () => {
          database.close();
          resolve();
        };
        transaction.onerror = () => {
          database.close();
          resolve();
        };
        transaction.onabort = () => {
          database.close();
          resolve();
        };
      } catch (_) {
        try {
          database.close();
        } catch (_) {}
        resolve();
      }
    };
  });
}

function clearFirebaseAuthStorage() {
  const shouldRemove = (key) => (
    key === 'firebaseLocalStorageDb' ||
    key.startsWith('firebase:authUser:') ||
    key.startsWith('firebase:persistence:')
  );

  try {
    Object.keys(localStorage).forEach((key) => {
      if (shouldRemove(key)) localStorage.removeItem(key);
    });
  } catch (_) {}

  try {
    Object.keys(sessionStorage).forEach((key) => {
      if (shouldRemove(key)) sessionStorage.removeItem(key);
    });
  } catch (_) {}
}

async function clearFirebaseAuthPersistence() {
  clearFirebaseAuthStorage();
  await clearIndexedDbStore('firebaseLocalStorageDb', 'firebaseLocalStorage');
  await deleteIndexedDb('firebaseLocalStorageDb');
}

async function signOutWithTimeout(auth) {
  if (!auth) return;
  await Promise.race([
    signOut(auth),
    delay(LOGOUT_TIMEOUT_MS).then(() => {
      throw new Error('Firebase sign-out timed out');
    })
  ]);
}

export async function performAppLogout({
  auth,
  beforeSignOut,
  redirectTo = 'index.html'
} = {}) {
  if (window.__appLogoutInProgress) return;
  window.__appLogoutInProgress = true;

  try {
    if (typeof beforeSignOut === 'function') {
      await beforeSignOut().catch(() => {});
    }
    await signOutWithTimeout(auth).catch(() => {});
  } finally {
    await clearFirebaseAuthPersistence().catch(() => {});
    window.location.replace(redirectTo);
  }
}
