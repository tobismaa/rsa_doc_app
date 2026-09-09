import { initializeApp } from "https://www.gstatic.com/firebasejs/12.18.0/firebase-app.js";
import { getAuth } from "https://www.gstatic.com/firebasejs/12.18.0/firebase-auth.js";
import { getFirestore } from "https://www.gstatic.com/firebasejs/12.18.0/firebase-firestore.js";

const firebaseConfig = {
  apiKey: "AIzaSyDkij9BPE_66LhKNLaXri07Knqcy0UIApk",
  authDomain: "assetmanager-f2ac7.firebaseapp.com",
  projectId: "assetmanager-f2ac7",
  storageBucket: "assetmanager-f2ac7.firebasestorage.app",
  messagingSenderId: "421067804367",
  appId: "1:421067804367:web:5c29c97e357909818e31ba",
  measurementId: "G-G31DTLTRJB",
};
export const isConfigured = Object.values(firebaseConfig).every(Boolean);
const app = isConfigured ? initializeApp(firebaseConfig, "cmbank-asset-manager") : null;
export const auth = app ? getAuth(app) : null;
export const db = app ? getFirestore(app) : null;
