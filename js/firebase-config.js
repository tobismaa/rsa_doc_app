// Import the functions you need from the SDKs you need
import { initializeApp } from "https://www.gstatic.com/firebasejs/9.22.0/firebase-app.js";
import { getAuth } from "https://www.gstatic.com/firebasejs/9.22.0/firebase-auth.js";
import { getFirestore } from "https://www.gstatic.com/firebasejs/9.22.0/firebase-firestore.js";
import { getAnalytics } from "https://www.gstatic.com/firebasejs/9.22.0/firebase-analytics.js";

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
const app = initializeApp(firebaseConfig);
const analytics = getAnalytics(app);
const auth = getAuth(app);
const db = getFirestore(app);

export { app, auth, db, analytics };
