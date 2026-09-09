// js/config.js

// 🔥 Import Firebase SDK modules (Modular v9+)
import { initializeApp } from "https://www.gstatic.com/firebasejs/9.6.0/firebase-app.js";
import { getAuth } from "https://www.gstatic.com/firebasejs/9.6.0/firebase-auth.js";
import { getFirestore } from "https://www.gstatic.com/firebasejs/9.6.0/firebase-firestore.js";
import { getStorage } from "https://www.gstatic.com/firebasejs/9.6.0/firebase-storage.js";
// Analytics is optional - uncomment if needed
// import { getAnalytics } from "https://www.gstatic.com/firebasejs/9.6.0/firebase-analytics.js";

// 📋 Your Firebase Configuration
const firebaseConfig = {
  apiKey: "AIzaSyCqFDo6SzEBS2vI0rRtoagUegSTU_xNZ8Y",
  authDomain: "evoucher-b42ae.firebaseapp.com",
  projectId: "evoucher-b42ae",
  storageBucket: "evoucher-b42ae.firebasestorage.app",
  messagingSenderId: "108322286006",
  appId: "1:108322286006:web:97c3c3004c356acbcfb917",
  measurementId: "G-82FHTYB6L3"
};

// 🚀 Initialize Firebase
const app = initializeApp(firebaseConfig);

// 🛠️ Export Firebase services
export const auth = getAuth(app);
export const db = getFirestore(app);
export const storage = getStorage(app);
// export const analytics = getAnalytics(app); // Optional

// 🔄 Workflow Configuration
export const WORKFLOW_STEPS = {
  initiated:      { label: "Initiated", next: "operations", previous: null },
  operations:     { label: "Operations Review", next: "fincon", previous: "initiated" },
  fincon:         { label: "FinCon Vetting", next: "iad", previous: "operations" },
  iad:            { label: "Internal Audit", next: "coo", previous: "fincon" },
  coo:            { label: "COO Approval", next: "posting", previous: "iad" },
  posting:        { label: "Operations Posting", next: "settlement", previous: "coo" },
  settlement:     { label: "Settlement", next: "completed", previous: "posting" },
  completed:      { label: "Completed", next: null, previous: "settlement" }
};

// 👥 Role Permissions (Who can do what)
export const ROLE_PERMISSIONS = {
  initiator:  ['create_voucher', 'view_own', 'edit_if_sent_back'],
  operations: ['view_assigned', 'enter_gl_code', 'enter_reference', 'approve_to_fincon', 'send_back'],
  fincon:     ['view_assigned', 'approve', 'send_back'],
  iad:        ['view_assigned', 'approve', 'send_back'],
  coo:        ['view_assigned', 'approve_final', 'redirect_to_posting', 'send_back'],
  admin:      ['all']
};

// 🎨 UI Helpers
export const STATUS_COLORS = {
  pending_operations: '#3b82f6',
  pending_iad: '#8b5cf6',
  pending_fincon: '#ec4899',
  pending_coo: '#f59e0b',
  pending_posting: '#10b981',
  pending_settlement: '#6366f1',
  completed: '#22c55e',
  sent_back: '#ef4444',
  rejected: '#dc2626'
};
