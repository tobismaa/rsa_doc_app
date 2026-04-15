﻿// js/document-uploader.js - COMPLETE FIXED VERSION WITH WORKING FILE SIZE MODALS AND CALCULATIONS
import { auth, db } from './firebase-config.js';
import { BackblazeStorage } from './backblaze-storage.js';
import { queueViewerAssignmentEmail } from './email-alerts.js';
import { notifyStatusChangePush } from './status-push.js';
import {
  collection, query, where, orderBy, onSnapshot, addDoc, updateDoc, doc,
  serverTimestamp, arrayUnion, getDocs, getDoc, setDoc, runTransaction
} from "https://www.gstatic.com/firebasejs/9.22.0/firebase-firestore.js";

// ==================== FIX JSPDF DETECTION - ADD THIS RIGHT AFTER IMPORTS ====================
(function fixJsPDF() {
    // Check all possible locations where jsPDF might be
    if (window.jspdf && window.jspdf.jsPDF) {
        window.jsPDF = window.jspdf.jsPDF;
    }
    else if (window.jsPDF && typeof window.jsPDF === 'function') {
        window.jspdf = window.jspdf || {};
        window.jspdf.jsPDF = window.jsPDF;
    }
    else if (window.jspdf && typeof window.jspdf === 'function') {
        window.jsPDF = window.jspdf;
        window.jspdf = { jsPDF: window.jspdf };
    }
    else {
        // Create a mock that shows a helpful error message
        window.jspdf = {
            jsPDF: function() {
                throw new Error('jsPDF library not loaded. Please add: <script src="https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js"></script>');
            }
        };
    }
})();

// ==================== DOCUMENT TYPES ====================
const DOCUMENT_TYPES = [
  { id: 'birth_certificate', name: 'Birth Certificate / Age Declaration', icon: 'fa-id-card', required: true },
  { id: 'nin', name: 'National Identification Number (NIN)', icon: 'fa-id-card', required: true },
  { id: 'bvn', name: 'BVN', icon: 'fa-id-badge', required: true },
  { id: 'pay_slips', name: '3 Months Pay Slip', icon: 'fa-file-invoice', required: false },
  { id: 'offer_letter', name: 'Offer of Employment Letter', icon: 'fa-file-signature', required: true },
  { id: 'intro_letter', name: 'Introduction Letter', icon: 'fa-file-signature', required: true },
  { id: 'request_letter', name: 'Request Letter', icon: 'fa-file-signature', required: true },
  { id: 'rsa_statement', name: 'RSA Statement', icon: 'fa-file-invoice', required: true },
  { id: 'pfa_form', name: 'PFA Application Form', icon: 'fa-file-pdf', required: true },
  { id: 'consent_letter', name: 'Consent Letter', icon: 'fa-file-signature', required: true },
  { id: 'indemnity_form', name: 'Indemnity Form', icon: 'fa-file-signature', required: true },
  { id: 'mortgage_loan_application', name: 'Mortgage Loan Application Form', icon: 'fa-file-pdf', required: true },
  { id: 'allocation_last_page', name: 'Allocation Last Page', icon: 'fa-file-pdf', required: true },
  { id: 'offer_letter_last_page', name: 'Offer Letter Last Page', icon: 'fa-file-pdf', required: true },
  { id: 'pmi_soa', name: 'PMI SOA', icon: 'fa-file-pdf', required: true },
  { id: 'benefit_application_form', name: 'Benefit Application Form', icon: 'fa-file-alt', required: false },
  { id: 'data_recapture', name: 'Data Recapture', icon: 'fa-file-alt', required: false }
];

// Make document types globally available
const REQUIRED_DOC_TYPES = DOCUMENT_TYPES.filter(d => d.required !== false);
const OPTIONAL_DOC_TYPES = DOCUMENT_TYPES.filter(d => d.required === false);

window.REQUIRED_DOC_TYPES = REQUIRED_DOC_TYPES;
window.OPTIONAL_DOC_TYPES = OPTIONAL_DOC_TYPES;

function formatCurrency(value) {
  const num = Number(value || 0);
  try {
    return num.toLocaleString('en-NG', { style: 'currency', currency: 'NGN', maximumFractionDigits: 2 });
  } catch (e) {
    return '₦' + num.toLocaleString(undefined, { maximumFractionDigits: 2 });
  }
}

function roundDownToNearestThousand(value) {
  const num = Number(value || 0);
  return Math.floor(num / 1000) * 1000;
}

function parseMoney(value) {
  const raw = String(value ?? '').replace(/[^0-9.\-]/g, '');
  const num = Number(raw);
  return Number.isFinite(num) ? num : 0;
}

function getFinancials(submission) {
  const details = submission?.customerDetails || {};
  const rsaBalance = parseMoney(details.rsaBalance || submission?.rsaBalance || 0);
  const computed25 = rsaBalance * 0.25;
  const twentyFive = parseMoney(details.rsa25Percent || submission?.rsa25Percent || computed25);
  const commission2 = twentyFive * 0.02;
  const pfa = String(details.pfa || submission?.pfa || '').trim() || '-';
  return { pfa, twentyFive, commission2 };
}

// ==================== GLOBAL VARIABLES ====================
let currentUser = null;
let currentUserProfile = null;
let allSubmissions = [];
let currentCustomerUploads = {};
let currentEditId = null;
let currentDocType = null;
let currentFile = null;
let approvedAgents = [];
let singlePreviewObjectUrl = null;
let trustedNowCache = { value: null, fetchedAt: 0 };
// Mobile camera photos can be large; compress only when needed.
const MAX_IMAGE_UPLOAD_BYTES = 1024 * 1024; // 1MB
const MAX_PDF_SIZE_BYTES = 1.5 * 1024 * 1024; // 1.5MB
const userFullNames = new Map();
let customerDetailsSaved = false;
const RR_COUNTER_DOC = doc(db, 'counters', 'roundRobin');

// ==================== INJECT MODAL ANIMATIONS (GLOBAL) ====================
function injectModalAnimations() {
  if (document.getElementById('modalAnimations')) return;
  const style = document.createElement('style');
  style.id = 'modalAnimations';
  style.textContent = `
    @keyframes modalSlideIn {
      from { transform: translateY(-30px); opacity: 0; }
      to { transform: translateY(0); opacity: 1; }
    }
  `;
  document.head.appendChild(style);
}
// Call on script load
injectModalAnimations();

// ==================== FILE SIZE WARNING MODAL - SINGLE UPLOAD ====================
function showFileSizeWarningModal(file) {
  // Remove any existing modals first
  const existingModal = document.getElementById('fileSizeWarningModal');
  if (existingModal) {
    existingModal.remove();
  }

  const formatSize = (bytes) => (bytes / (1024 * 1024)).toFixed(2) + ' MB';

  // Create modal element
  const modal = document.createElement('div');
  modal.className = 'modal file-size-warning';
  modal.id = 'fileSizeWarningModal';

  // CRITICAL: Use !important to override CSS conflicts
  modal.style.cssText = `
    display: flex !important;
    position: fixed !important;
    top: 0 !important;
    left: 0 !important;
    width: 100% !important;
    height: 100% !important;
    background: rgba(0, 0, 0, 0.5) !important;
    z-index: 999999 !important;
    align-items: center !important;
    justify-content: center !important;
    opacity: 1 !important;
  `;

  const modalContent = document.createElement('div');
  modalContent.style.cssText = `
    background: white !important;
    border-radius: 12px !important;
    max-width: 500px !important;
    width: 90% !important;
    max-height: 90vh !important;
    overflow-y: auto !important;
    box-shadow: 0 25px 50px -12px rgba(0, 0, 0, 0.25) !important;
    animation: modalSlideIn 0.3s ease !important;
  `;

  modalContent.innerHTML = `
    <div style="padding: 20px; border-bottom: 1px solid #e2e8f0; background: #fee2e2; border-radius: 12px 12px 0 0; display: flex; justify-content: space-between; align-items: center;">
      <h2 style="margin: 0; color: #991b1b; font-size: 20px; display: flex; align-items: center; gap: 10px;">
        <i class="fas fa-exclamation-triangle"></i> File Too Large
      </h2>
      <button id="fileSizeWarningCloseBtn" style="background: transparent; border: none; font-size: 28px; color: #991b1b; cursor: pointer; width: 36px; height: 36px; display: flex; align-items: center; justify-content: center; border-radius: 50%; transition: all 0.2s;">&times;</button>
    </div>
    <div style="padding: 25px;">
      <div style="text-align: center; margin-bottom: 20px;">
        <i class="fas fa-file-pdf" style="font-size: 50px; color: #ef4444;"></i>
      </div>
      <p style="font-size: 16px; color: #1e293b; margin-bottom: 15px; text-align: center;">
        <strong>Maximum file size allowed is 1.5MB</strong>
      </p>
      <div style="background: #f8fafc; border-radius: 8px; padding: 20px; margin-bottom: 20px; text-align: center;">
        <p style="font-size: 16px; font-weight: 600; color: #1e293b; margin-bottom: 10px;">${file.name}</p>
        <p style="font-size: 18px; color: #ef4444; font-weight: 700; background: #fee2e2; display: inline-block; padding: 8px 16px; border-radius: 20px;">
          Size: ${formatSize(file.size)}
        </p>
      </div>
      <div style="background: #f1f5f9; border-radius: 8px; padding: 15px;">
        <p style="font-weight: 600; margin-bottom: 10px; color: #334155;">
          <i class="fas fa-lightbulb" style="color: #f59e0b;"></i> Tips to reduce file size:
        </p>
        <ul style="padding-left: 20px; color: #475569; font-size: 13px; line-height: 1.6;">
          <li>Compress PDF using online tools (ilovepdf.com, smallpdf.com)</li>
          <li>Reduce image resolution before converting to PDF</li>
          <li>Scan documents at 150 DPI instead of 300 DPI</li>
          <li>Save images as JPG instead of PNG</li>
        </ul>
      </div>
    </div>
    <div style="padding: 20px; border-top: 1px solid #e2e8f0; display: flex; justify-content: center; background: #f8fafc; border-radius: 0 0 12px 12px;">
      <button id="understandBtn" style="background: #003366; color: white; border: none; padding: 12px 30px; border-radius: 6px; font-size: 14px; font-weight: 600; cursor: pointer; display: flex; align-items: center; gap: 8px;">
        <i class="fas fa-check"></i> Understood
      </button>
    </div>
  `;

  modal.appendChild(modalContent);

  // SAFER: Append to body or html fallback
  const target = document.body || document.documentElement;
  if (!target) {
    return;
  }
  target.appendChild(modal);

  // Add event listeners
  const closeBtn = document.getElementById('fileSizeWarningCloseBtn');
  const understandBtn = document.getElementById('understandBtn');

  if (closeBtn) closeBtn.addEventListener('click', () => { modal.remove(); });
  if (understandBtn) understandBtn.addEventListener('click', () => { modal.remove(); });

  modal.addEventListener('click', (e) => {
    if (e.target === modal) { modal.remove(); }
  });
}

// ==================== BATCH SIZE WARNING MODAL ====================
function showBatchSizeWarningModal(oversizedFiles, validFiles, originalFiles) {
  const existingModal = document.getElementById('batchSizeWarningModal');
  if (existingModal) existingModal.remove();

  const formatSize = (bytes) => (bytes / (1024 * 1024)).toFixed(2) + ' MB';

  const modal = document.createElement('div');
  modal.className = 'modal batch-warning';
  modal.id = 'batchSizeWarningModal';

  modal.style.cssText = `
    display: flex !important;
    position: fixed !important;
    top: 0 !important;
    left: 0 !important;
    width: 100% !important;
    height: 100% !important;
    background: rgba(0, 0, 0, 0.5) !important;
    z-index: 999999 !important;
    align-items: center !important;
    justify-content: center !important;
    opacity: 1 !important;
  `;

  const modalContent = document.createElement('div');
  modalContent.style.cssText = `
    background: white !important;
    border-radius: 12px !important;
    max-width: 700px !important;
    width: 95% !important;
    max-height: 90vh !important;
    overflow-y: auto !important;
    box-shadow: 0 25px 50px -12px rgba(0, 0, 0, 0.25) !important;
    animation: modalSlideIn 0.3s ease !important;
  `;

  const oversizedListHtml = oversizedFiles.map(file => `
    <div style="display: flex; justify-content: space-between; align-items: center; padding: 12px; background: #fef2f2; border-bottom: 1px solid #fecaca;">
      <div style="display: flex; align-items: center; gap: 10px; flex: 1;">
        <i class="fas fa-exclamation-circle" style="color: #ef4444;"></i>
        <span style="font-weight: 500; color: #1e293b;">${file.name}</span>
      </div>
      <span style="color: #ef4444; font-weight: 600; background: #fee2e2; padding: 4px 8px; border-radius: 4px; white-space: nowrap;">
        ${formatSize(file.size)}
      </span>
    </div>
  `).join('');

  const validListHtml = validFiles.map(file => `
    <div style="display: flex; justify-content: space-between; align-items: center; padding: 12px; background: #f0fdf4; border-bottom: 1px solid #bbf7d0;">
      <div style="display: flex; align-items: center; gap: 10px; flex: 1;">
        <i class="fas fa-check-circle" style="color: #10b981;"></i>
        <span style="font-weight: 500; color: #1e293b;">${file.name}</span>
      </div>
      <span style="color: #10b981; font-weight: 600; background: #dcfce7; padding: 4px 8px; border-radius: 4px; white-space: nowrap;">
        ${formatSize(file.size)}
      </span>
    </div>
  `).join('');

  modalContent.innerHTML = `
    <div style="padding: 20px; border-bottom: 1px solid #e2e8f0; background: #fffbeb; border-radius: 12px 12px 0 0; display: flex; justify-content: space-between; align-items: center;">
      <h2 style="margin: 0; color: #92400e; font-size: 20px; display: flex; align-items: center; gap: 10px;">
        <i class="fas fa-exclamation-triangle"></i> File Size Validation
      </h2>
      <button id="batchSizeWarningCloseBtn" style="background: transparent; border: none; font-size: 28px; color: #92400e; cursor: pointer; width: 36px; height: 36px; display: flex; align-items: center; justify-content: center; border-radius: 50%; transition: all 0.2s;">&times;</button>
    </div>
    <div style="padding: 25px; max-height: 60vh; overflow-y: auto;">
      <p style="font-size: 16px; color: #1e293b; margin-bottom: 20px; text-align: center;">
        <strong>${oversizedFiles.length}</strong> file(s) exceed the 1.5MB limit.
        <strong>${validFiles.length}</strong> file(s) are within limit.
      </p>
      ${oversizedFiles.length > 0 ? `
        <div style="margin-bottom: 25px;">
          <h3 style="font-size: 16px; font-weight: 600; color: #991b1b; margin-bottom: 12px; display: flex; align-items: center; gap: 8px;">
            <i class="fas fa-times-circle"></i> Files Exceeding 1.5MB (will be skipped)
          </h3>
          <div style="border: 1px solid #fecaca; border-radius: 8px; overflow: hidden;">
            ${oversizedListHtml}
          </div>
        </div>
      ` : ''}
      ${validFiles.length > 0 ? `
        <div style="margin-bottom: 20px;">
          <h3 style="font-size: 16px; font-weight: 600; color: #065f46; margin-bottom: 12px; display: flex; align-items: center; gap: 8px;">
            <i class="fas fa-check-circle"></i> Files Within Limit (will be uploaded)
          </h3>
          <div style="border: 1px solid #bbf7d0; border-radius: 8px; overflow: hidden;">
            ${validListHtml}
          </div>
        </div>
      ` : ''}
    </div>
    <div style="padding: 20px; border-top: 1px solid #e2e8f0; display: flex; gap: 10px; justify-content: flex-end; background: #f8fafc; border-radius: 0 0 12px 12px;">
      <button id="cancelBtn" style="background: white; border: 1px solid #cbd5e1; color: #475569; padding: 10px 20px; border-radius: 6px; font-size: 14px; font-weight: 500; cursor: pointer; display: flex; align-items: center; gap: 8px;">
        <i class="fas fa-times"></i> Cancel All
      </button>
      ${validFiles.length > 0 ? `
        <button id="proceedBtn" style="background: #10b981; color: white; border: none; padding: 10px 20px; border-radius: 6px; font-size: 14px; font-weight: 600; cursor: pointer; display: flex; align-items: center; gap: 8px;">
          <i class="fas fa-upload"></i> Upload ${validFiles.length} Valid File(s)
        </button>
      ` : ''}
    </div>
  `;

  modal.appendChild(modalContent);

  const target = document.body || document.documentElement;
  if (!target) {
    return;
  }
  target.appendChild(modal);

  // Event listeners
  const closeBtn = document.getElementById('batchSizeWarningCloseBtn');
  const cancelBtn = document.getElementById('cancelBtn');
  const proceedBtn = document.getElementById('proceedBtn');

  if (closeBtn) closeBtn.addEventListener('click', () => { modal.remove(); });
  if (cancelBtn) cancelBtn.addEventListener('click', () => { modal.remove(); });

  if (proceedBtn) {
    proceedBtn.addEventListener('click', () => {
      modal.remove();
      if (validFiles.length > 0) prepareBatchForMapping(validFiles);
    });
  }

  modal.addEventListener('click', (e) => {
    if (e.target === modal) { modal.remove(); }
  });
}

// ==================== GET VIEWER EMAILS ====================
async function getViewerEmails() {
  const snap = await getDocs(collection(db, 'users'));
  return snap.docs
    .map(d => d.data())
    .filter(data => {
      const role = String(data?.role || '').toLowerCase();
      const status = String(data?.status || 'active').toLowerCase();
      return (role === 'reviewer' || role === 'viewer') && status !== 'deactivated';
    })
    .map(data => data.email)
    .filter(Boolean)
    .sort();
}

function normalizeEmail(email) {
  return String(email || '').trim().toLowerCase();
}

function routingRuleDocId(uploaderEmail) {
  return encodeURIComponent(normalizeEmail(uploaderEmail));
}

async function getUploaderRoutingRule(uploaderEmail) {
  const normalizedUploader = normalizeEmail(uploaderEmail);
  if (!normalizedUploader) return null;
  try {
    const snap = await getDoc(doc(db, 'uploaderRoutingRules', routingRuleDocId(normalizedUploader)));
    if (!snap.exists()) return null;
    const data = snap.data() || {};
    if (data.enabled === false) return null;
    return {
      uploaderEmail: normalizedUploader,
      reviewerEmail: normalizeEmail(data.reviewerEmail),
      rsaEmail: normalizeEmail(data.rsaEmail)
    };
  } catch (_) {
    return null;
  }
}

async function isActiveUserWithRole(email, allowedRoles = []) {
  const normalized = normalizeEmail(email);
  if (!normalized) return false;
  try {
    const userQ = query(collection(db, 'users'), where('email', '==', normalized));
    const snap = await getDocs(userQ);
    if (snap.empty) return false;
    const data = snap.docs[0].data() || {};
    const role = String(data.role || '').toLowerCase();
    const status = String(data.status || 'active').toLowerCase();
    if (status === 'deactivated') return false;
    return allowedRoles.includes(role);
  } catch (_) {
    return false;
  }
}

async function assignRoundRobin(subRef) {
  let uploaderEmail = normalizeEmail(currentUser?.email);
  if (!uploaderEmail) {
    try {
      const subSnap = await getDoc(subRef);
      uploaderEmail = normalizeEmail(subSnap.exists() ? subSnap.data()?.uploadedBy : '');
    } catch (_) { }
  }

  const routingRule = await getUploaderRoutingRule(uploaderEmail);
  const mappedReviewer = routingRule?.reviewerEmail || '';
  if (mappedReviewer && await isActiveUserWithRole(mappedReviewer, ['reviewer', 'viewer'])) {
    await updateDoc(subRef, {
      assignedTo: mappedReviewer,
      assignmentMode: 'uploader_routing',
      assignmentUploader: uploaderEmail || ''
    });
    try {
      const subSnap = await getDoc(subRef);
      if (subSnap.exists()) {
        const subData = subSnap.data();
        await addDoc(collection(db, 'roundRobinAssignments'), {
          submissionId: subRef.id,
          customerName: subData.customerName || 'N/A',
          assignedTo: mappedReviewer,
          assignedBy: currentUser?.email || 'System',
          assignedAt: serverTimestamp(),
          uploadedBy: subData.uploadedBy || uploaderEmail || 'N/A',
          assignmentMethod: 'uploader_routing'
        });
      }
    } catch (_) { }
    return mappedReviewer;
  }

  const viewers = await getViewerEmails();
  if (!viewers.length) return null;

  let assigned = null;
  try {
    await runTransaction(db, async tx => {
      let lastIndex = -1;
      let lastDate = '';
      const today = new Date().toISOString().slice(0, 10);
      const counterSnap = await tx.get(RR_COUNTER_DOC);
      if (counterSnap.exists()) {
        const data = counterSnap.data();
        lastIndex = typeof data.lastIndex === 'number' ? data.lastIndex : -1;
        lastDate = data.lastDate || '';
      }
      if (lastDate !== today) lastIndex = -1;
      const newIndex = (lastIndex + 1) % viewers.length;
      assigned = viewers[newIndex];
      tx.set(RR_COUNTER_DOC, { lastIndex: newIndex, lastDate: today }, { merge: true });
      tx.update(subRef, { assignedTo: assigned, assignmentMode: 'round_robin' });
    });
  } catch (_) {
    // Fallback: if counter transaction fails, still assign to keep workflow moving.
    assigned = viewers[0] || null;
    if (assigned) {
      await updateDoc(subRef, { assignedTo: assigned, assignmentMode: 'round_robin_fallback' });
    }
  }

  if (assigned) {
    try {
      const subSnap = await getDoc(subRef);
      if (subSnap.exists()) {
        const subData = subSnap.data();
        await addDoc(collection(db, 'roundRobinAssignments'), {
          submissionId: subRef.id,
          customerName: subData.customerName || 'N/A',
          assignedTo: assigned,
          assignedBy: currentUser?.email || 'System',
          assignedAt: serverTimestamp(),
          uploadedBy: subData.uploadedBy || 'N/A',
          assignmentMethod: 'round_robin'
        });
      }
    } catch (e) {
      // Silently fail
    }
  }
  return assigned;
}

// ==================== GET APPLICATION STAGE ====================
async function getApplicationStage(submission) {
  if (!submission) return 'Unknown';
  try {
    const status = String(submission.status || '').toLowerCase();
    if (status === 'cleared') return 'Cleared';
    if (status === 'paid') return 'Paid';
    if (status === 'sent_to_pfa' || status === 'rsa_submitted') return 'Sent to PFA';
    if (status === 'processing_to_pfa' || status === 'approved') return 'Processing to PFA';
    if (status === 'rejected') return 'Rejected - Fix Required';
    if (submission.assignedTo) {
      const reviewerName = await getUserFullName(submission.assignedTo);
      return `With Reviewer: ${reviewerName}`;
    }
    const historyQuery = query(
      collection(db, 'roundRobinAssignments'),
      where('submissionId', '==', submission.id),
      orderBy('assignedAt', 'desc')
    );
    const historySnap = await getDocs(historyQuery);
    if (!historySnap.empty) {
      const latest = historySnap.docs[0].data();
      if (latest.assignedTo) {
        const reviewerName = await getUserFullName(latest.assignedTo);
        return `With Reviewer: ${reviewerName}`;
      }
    }
    return 'Pending Assignment';
  } catch (error) {
    return 'Unknown';
  }
}

// ==================== GET USER FULL NAME BY EMAIL ====================
async function getUserFullName(email) {
  if (!email) return 'Unknown';
  if (userFullNames.has(email)) return userFullNames.get(email);
  try {
    const userQuery = query(collection(db, 'users'), where('email', '==', email));
    const userSnapshot = await getDocs(userQuery);
    if (!userSnapshot.empty) {
      const userData = userSnapshot.docs[0].data();
      const fullName = userData.fullName || userData.displayName || email.split('@')[0];
      userFullNames.set(email, fullName);
      return fullName;
    }
  } catch (err) {
    // Silently fail
  }
  const fallbackName = email.split('@')[0];
  userFullNames.set(email, fallbackName);
  return fallbackName;
}

// ==================== DOM ELEMENTS ====================
const userName = document.getElementById('userName');
const userAvatar = document.getElementById('userAvatar');
const newUploadBtn = document.getElementById('newUploadBtn');
const uploadModal = document.getElementById('uploadModal');
const editModal = document.getElementById('editModal');
const singleUploadModal = document.getElementById('singleUploadModal');
const closeModalBtn = document.getElementById('closeModalBtn');
const closeEditModalBtn = document.getElementById('closeEditModalBtn');
const closeSingleModalBtn = document.getElementById('closeSingleModalBtn');
const cancelSingleBtn = document.getElementById('cancelSingleBtn');
const documentGrid = document.getElementById('documentGrid');
const optionalDocumentGrid = document.getElementById('optionalDocumentGrid');
const editDocumentGrid = document.getElementById('editDocumentGrid');
const customerNameInput = document.getElementById('customerName');
const submitCustomerBtn = document.getElementById('submitCustomerBtn');
const submitEditBtn = document.getElementById('submitEditBtn');
const uploadedCountSpan = document.getElementById('uploadedCount');
const totalCountSpan = document.getElementById('totalCount');
const notification = document.getElementById('notification');
const editCustomerName = document.getElementById('editCustomerName');
const rejectionComment = document.getElementById('rejectionComment');
const singleUploadArea = document.getElementById('singleUploadArea');
const singleFileInput = document.getElementById('singleFileInput');
const singleFilePreview = document.getElementById('singleFilePreview');
const uploadModalTitle = document.getElementById('uploadModalTitle');
const uploadDocType = document.getElementById('uploadDocType');
const confirmSingleUpload = document.getElementById('confirmSingleUpload');
const pageTitle = document.getElementById('pageTitle');
const pendingTableBody = document.getElementById('pendingTableBody');
const approvedTableBody = document.getElementById('approvedTableBody');
const rejectedTableBody = document.getElementById('rejectedTableBody');
const activeCommissionTableBody = document.getElementById('activeCommissionTableBody');
const clearedCommissionTableBody = document.getElementById('clearedCommissionTableBody');
const paidCountBadge = document.getElementById('paidCount');
const activeCommissionCountEl = document.getElementById('activeCommissionCount');
const clearedCommissionCountEl = document.getElementById('clearedCommissionCount');
const paidTotal25El = document.getElementById('paidTotal25');
const paidTotal1El = document.getElementById('paidTotal1');
const clearedTotal25El = document.getElementById('clearedTotal25');
const clearedTotal1El = document.getElementById('clearedTotal1');
const activeTotal25Card = document.getElementById('activeTotal25Card');
const activeTotal1Card = document.getElementById('activeTotal1Card');
const clearedTotal25Card = document.getElementById('clearedTotal25Card');
const clearedTotal1Card = document.getElementById('clearedTotal1Card');
const activeCommissionTabBtn = document.getElementById('activeCommissionTabBtn');
const clearedCommissionTabBtn = document.getElementById('clearedCommissionTabBtn');
const activeCommissionSection = document.getElementById('activeCommissionSection');
const clearedCommissionSection = document.getElementById('clearedCommissionSection');
const activeCommissionSearch = document.getElementById('activeCommissionSearch');
const clearedCommissionSearch = document.getElementById('clearedCommissionSearch');
const activeCommissionStartDate = document.getElementById('activeCommissionStartDate');
const activeCommissionEndDate = document.getElementById('activeCommissionEndDate');
const clearedCommissionStartDate = document.getElementById('clearedCommissionStartDate');
const clearedCommissionEndDate = document.getElementById('clearedCommissionEndDate');
const viewerModal = document.getElementById('viewerModal');
const viewerFileName = document.getElementById('viewerFileName');
const documentViewer = document.getElementById('documentViewer');
const helpFab = document.getElementById('helpFab');
const sopHelpModal = document.getElementById('sopHelpModal');
const closeSopHelpModalBtn = document.getElementById('closeSopHelpModalBtn');
const closeSopHelpModalFooterBtn = document.getElementById('closeSopHelpModalFooterBtn');
const batchUploadBtn = document.getElementById('batchUploadBtn');
const batchFileInput = document.getElementById('batchFileInput');
const saveDetailsBtn = document.getElementById('saveDetailsBtn');
const resetDetailsBtn = document.getElementById('resetDetailsBtn');
const batchMappingModal = document.getElementById('batchMappingModal');
const batchMappingList = document.getElementById('batchMappingList');
const closeBatchMappingBtn = document.getElementById('closeBatchMappingBtn');
const cancelBatchMapping = document.getElementById('cancelBatchMapping');
const confirmBatchMapping = document.getElementById('confirmBatchMapping');
const profileNameEl = document.getElementById('profileName');
const profileRegisteredAtEl = document.getElementById('profileRegisteredAt');
const profileEmailEl = document.getElementById('profileEmail');
const profileWhatsappEl = document.getElementById('profileWhatsapp');
const profileLocationEl = document.getElementById('profileLocation');
const profileRoleEl = document.getElementById('profileRole');
const profileStatusEl = document.getElementById('profileStatus');
const registeredAgentsTableBody = document.getElementById('registeredAgentsTableBody');
const customerAgentSelect = document.getElementById('customerAgent');
const agentRegistrationForm = document.getElementById('agentRegistrationForm');
const agentRegistrationModal = document.getElementById('agentRegistrationModal');
const openAgentRegistrationBtn = document.getElementById('openAgentRegistrationBtn');
const closeAgentRegistrationModalBtn = document.getElementById('closeAgentRegistrationModalBtn');
const resetAgentFormBtn = document.getElementById('resetAgentFormBtn');
const submitAgentFormBtn = document.getElementById('submitAgentFormBtn');
let __batchFilesBuffer = [];
let currentCommissionTab = 'active';
let registeredAgents = [];

function renderProfileTab() {
  if (!profileNameEl && !profileEmailEl && !profileRoleEl && !profileStatusEl) return;
  const setProfileField = (el, value) => {
    if (!el) return;
    if ('value' in el) el.value = String(value ?? '');
    else el.textContent = String(value ?? '');
  };
  const fullName = currentUserProfile?.fullName || currentUser?.displayName || currentUser?.email || 'N/A';
  const registeredAt = currentUserProfile?.createdAt ? safeFormatDate(currentUserProfile.createdAt) : '-';
  const email = currentUserProfile?.email || currentUser?.email || 'N/A';
  const whatsapp = currentUserProfile?.whatsappNumber || currentUserProfile?.phone || '-';
  const location = currentUserProfile?.location || '-';
  const role = String(currentUserProfile?.role || 'uploader');
  const status = String(currentUserProfile?.status || 'active');
  setProfileField(profileNameEl, fullName);
  setProfileField(profileRegisteredAtEl, registeredAt);
  setProfileField(profileEmailEl, email);
  setProfileField(profileWhatsappEl, whatsapp);
  setProfileField(profileLocationEl, location);
  setProfileField(profileRoleEl, role.charAt(0).toUpperCase() + role.slice(1));
  setProfileField(profileStatusEl, status.charAt(0).toUpperCase() + status.slice(1));
}

// ==================== INITIALIZATION ====================
document.addEventListener('DOMContentLoaded', () => {
  auth.onAuthStateChanged(async (user) => {
    if (user) {
      currentUser = user;
      try {
        const q = query(collection(db, 'users'), where('email', '==', user.email));
        const snap = await getDocs(q);
        if (!snap.empty) {
          const data = snap.docs[0].data();
          currentUserProfile = data;
          userName.textContent = data.fullName || user.displayName || user.email;
        } else {
          currentUserProfile = { email: user.email, fullName: user.displayName || user.email, role: 'uploader', status: 'active' };
          userName.textContent = user.displayName || user.email;
        }
      } catch (e) {
        currentUserProfile = { email: user.email, fullName: user.displayName || user.email, role: 'uploader', status: 'active' };
        userName.textContent = user.displayName || user.email;
      }
      userAvatar.src = user.photoURL || 'data:image/svg+xml,%3Csvg xmlns=\'http://www.w3.org/2000/svg\' width=\'40\' height=\'40\' viewBox=\'0 0 40 40\'%3E%3Ccircle cx=\'20\' cy=\'20\' r=\'20\' fill=\'%23003366\'/%3E%3Ctext x=\'20\' y=\'25\' text-anchor=\'middle\' fill=\'%23ffffff\' font-size=\'16\'%3E👤%3C/text%3E%3C/svg%3E';
      renderProfileTab();
      totalCountSpan.textContent = DOCUMENT_TYPES.length;
      await loadRegisteredAgents();
      await loadApprovedAgents();
      await loadSubmissions();
    } else {
      window.location.href = 'index.html';
    }
  });
  setupEventListeners();
});

// ==================== PROPERTY RULES ====================
const PROPERTY_RULES = [
  { name: '1 BEDROOM 8 IN 1 FLAT', value: 6000000, fee: 2000, min: 4000000, max: 6499999 },
  { name: '1 BEDROOM 4 IN 1 BUNGALOW', value: 6500000, fee: 2000, min: 6500000, max: 6999999 },
  { name: '1 BEDROOM 2 IN 1 BUNGALOW', value: 7000000, fee: 2000, min: 7000000, max: 9999999 },
  { name: '2 BEDROOM SEMI DETACHED BUNGALOW', value: 13000000, fee: 3000, min: 10000000, max: 14999999 },
  { name: '3 BEDROOM SEMI DETACHED BUNGALOW', value: 15000000, fee: 4000, min: 15000000, max: 19999999 },
  { name: '4 BEDROOM DETACHED BUNGALOW', value: 24000000, fee: 5000, min: 20000000, max: 34999999 },
  { name: '4 BEDROOM DETACHED LUXURY BUNGALOW', value: 35000000, fee: 5000, min: 35000000, max: 59999999 },
  { name: '4 BEDROOM TERRACE DUPLEX', value: 60000000, fee: 10000, min: 60000000, max: 99999999 },
  { name: '5 BEDROOM TERRACE DUPLEX', value: 100000000, fee: 20000, min: 100000000, max: 149999999 },
  { name: '6 BEDROOM TERRACE DUPLEX', value: 150000000, fee: 30000, min: 150000000, max: 199999999 },
  { name: '7 BEDROOM TERRACE DUPLEX', value: 200000000, fee: 40000, min: 200000000, max: 249999999 },
  { name: '8 BEDROOM TERRACE DUPLEX', value: 250000000, fee: 50000, min: 250000000, max: 299999999 }
];

function determinePropertyByRsa(rsaAmount) {
  const n = Number(rsaAmount) || 0;
  for (const r of PROPERTY_RULES) {
    if (n >= r.min && n <= r.max) return r;
  }
  return null;
}

// ==================== UPDATED: SET PROPERTY BY RSA WITH AUTO-FILL ====================
function setPropertyByRsaAndUpdate(rsaAmount) {
  const rule = determinePropertyByRsa(rsaAmount);
  const propEl = document.getElementById('propertyType');
  const propValEl = document.getElementById('propertyValue');
  const propValFmt = document.getElementById('propertyValueFormatted');
  const feeEl = document.getElementById('facilityFee');
  const feeFmt = document.getElementById('facilityFeeFormatted');
  const loanEl = document.getElementById('loanAmount');
  const loanFmt = document.getElementById('loanAmountFormatted');
  const tenorEl = document.getElementById('tenor');

  if (rule) {
    // Set property type
    if (propEl) propEl.value = rule.name;

    // Set property value
    if (propValEl) propValEl.value = rule.value;
    if (propValFmt) propValFmt.textContent = formatCurrency(rule.value);

    // Set facility fee
    if (feeEl) feeEl.value = rule.fee;
    if (feeFmt) feeFmt.textContent = formatCurrency(rule.fee);

    // Calculate exact 25% of RSA (no rounding)
    const rsaBalance = parseFloat(rsaAmount) || 0;
    const rsa25Exact = rsaBalance * 0.25;

    // Calculate loan amount and round DOWN to nearest thousand
    const loanAmount = roundDownToNearestThousand(rule.value - rsa25Exact);

    // Set loan amount
    if (loanEl) loanEl.value = loanAmount;
    if (loanFmt) loanFmt.textContent = formatCurrency(loanAmount);

    // Update tenor if DOB is available
    const dob = document.getElementById('customerDob')?.value;
    if (dob && window.calculateTenorFromDob) {
      const tenor = window.calculateTenorFromDob(dob);
      if (tenorEl) tenorEl.value = tenor;
    }

    // Also update the 25% display
    const rsa25FormattedEl = document.getElementById('rsa25Formatted');
    if (rsa25FormattedEl) {
      rsa25FormattedEl.textContent = formatCurrency(rsa25Exact);
    }

    const rsa25PercentEl = document.getElementById('rsa25Percent');
    if (rsa25PercentEl) {
      rsa25PercentEl.value = rsa25Exact;
    }

  } else {
    // Clear all fields if no rule matches
    if (propEl) propEl.value = '';
    if (propValEl) propValEl.value = '';
    if (propValFmt) propValFmt.textContent = '';
    if (feeEl) feeEl.value = '';
    if (feeFmt) feeFmt.textContent = '';
    if (loanEl) loanEl.value = '';
    if (loanFmt) loanFmt.textContent = '';

    // Also clear 25% display
    const rsa25FormattedEl = document.getElementById('rsa25Formatted');
    if (rsa25FormattedEl) {
      rsa25FormattedEl.textContent = '';
    }

    const rsa25PercentEl = document.getElementById('rsa25Percent');
    if (rsa25PercentEl) {
      rsa25PercentEl.value = '';
    }
  }
}

window.determinePropertyByRsa = determinePropertyByRsa;
window.setPropertyByRsaAndUpdate = setPropertyByRsaAndUpdate;

window.calculateTenorFromDob = function(dobString) {
  if (!dobString) return '';
  const birthDate = new Date(dobString);
  const today = new Date();
  let age = today.getFullYear() - birthDate.getFullYear();
  const monthDiff = today.getMonth() - birthDate.getMonth();
  if (monthDiff < 0 || (monthDiff === 0 && today.getDate() < birthDate.getDate())) {
    age--;
  }
  const tenor = 60 - age;
  return isNaN(tenor) ? '' : Math.max(0, tenor);
};

// ==================== CALCULATE AND DISPLAY CUSTOMER INFO ====================
function calculateAndDisplayCustomerInfo() {
  const dob = document.getElementById('customerDob')?.value;
  const nin = document.getElementById('customerNIN')?.value?.trim() || '';
  const rsaBalance = parseFloat(document.getElementById('rsaBalance')?.value || 0);

  if (!dob) return;

  // Calculate age
  const birthDate = new Date(dob);
  const today = new Date();
  let age = today.getFullYear() - birthDate.getFullYear();
  const monthDiff = today.getMonth() - birthDate.getMonth();
  if (monthDiff < 0 || (monthDiff === 0 && today.getDate() < birthDate.getDate())) {
    age--;
  }

  // Check if adult (age >= 18)
  const isAdult = age >= 18;

  // Check NIN validity (simple check - 11 digits)
  const ninValid = /^\d{11}$/.test(nin);

  // Calculate years to retirement (assuming retirement at 60)
  const yearsToRetirement = Math.max(0, 60 - age);

  // Get exact 25% of RSA (no rounding)
  const rsa25Exact = rsaBalance * 0.25;

  // Get property rule
  const rule = determinePropertyByRsa(rsaBalance);
  const propertyValue = rule ? rule.value : 0;
  const loanAmount = roundDownToNearestThousand(propertyValue - rsa25Exact);

  // Find or create results container
  let resultsContainer = document.getElementById('customerInfoResults');
  if (!resultsContainer) {
    resultsContainer = document.createElement('div');
    resultsContainer.id = 'customerInfoResults';
    resultsContainer.style.cssText = 'margin-top: 20px; background: #f0f9ff; padding: 15px; border-radius: 8px; border-left: 4px solid #003366;';

    // Insert after the form actions
    const formActions = document.querySelector('.form-actions') || document.getElementById('saveDetailsBtn')?.parentNode;
    if (formActions && formActions.parentNode) {
      formActions.parentNode.insertBefore(resultsContainer, formActions.nextSibling);
    } else {
      // Fallback - append to customer details section
      const detailsSection = document.querySelector('.customer-details-section');
      if (detailsSection) detailsSection.appendChild(resultsContainer);
    }
  }

  // Update results
  resultsContainer.innerHTML = `
    <h4 style="margin: 0 0 12px 0; color: #003366; font-size: 16px; display: flex; align-items: center; gap: 8px;">
      <i class="fas fa-chart-line"></i> Customer Analysis
    </h4>
    <div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(180px, 1fr)); gap: 12px;">
      <div style="background: white; padding: 10px; border-radius: 6px;">
        <div style="font-size: 12px; color: #64748b;">Age</div>
        <div style="font-size: 18px; font-weight: 700; color: #003366;">${age} years</div>
      </div>
      <div style="background: white; padding: 10px; border-radius: 6px;">
        <div style="font-size: 12px; color: #64748b;">Adult</div>
        <div style="font-size: 18px; font-weight: 700; color: ${isAdult ? '#10b981' : '#ef4444'};">${isAdult ? 'Yes ✓' : 'No ✗'}</div>
      </div>
      <div style="background: white; padding: 10px; border-radius: 6px;">
        <div style="font-size: 12px; color: #64748b;">NIN Valid</div>
        <div style="font-size: 18px; font-weight: 700; color: ${ninValid ? '#10b981' : '#f59e0b'};">${ninValid ? 'Valid ✓' : 'Check Required'}</div>
      </div>
      <div style="background: white; padding: 10px; border-radius: 6px;">
        <div style="font-size: 12px; color: #64748b;">Years to Retirement</div>
        <div style="font-size: 18px; font-weight: 700; color: #003366;">${yearsToRetirement}</div>
      </div>
      <div style="background: white; padding: 10px; border-radius: 6px;">
        <div style="font-size: 12px; color: #64748b;">25% of RSA</div>
        <div style="font-size: 18px; font-weight: 700; color: #003366;">${formatCurrency(rsa25Exact)}</div>
      </div>
      <div style="background: white; padding: 10px; border-radius: 6px;">
        <div style="font-size: 12px; color: #64748b;">Loan Amount</div>
        <div style="font-size: 18px; font-weight: 700; color: #10b981;">${formatCurrency(loanAmount)}</div>
      </div>
    </div>
  `;

  resultsContainer.style.display = 'block';
}

// ==================== SAVE CUSTOMER DETAILS ====================
function saveCustomerDetails() {
  // Define requiredFields INSIDE the function
  const requiredFields = [
    'customerName', 'customerDob', 'customerEmail', 'customerPhone',
    'customerNIN', 'customerAddress', 'accountNo', 'employer',
    'originatingTP', 'pfa', 'penNo', 'rsaStatementDate', 'rsaBalance'
  ];

  const missingFields = [];
  const invalidFields = [];

  requiredFields.forEach(id => {
    const el = document.getElementById(id);
    if (!el || !el.value.trim()) {
      missingFields.push(id);
      return;
    }
    const value = el.value.trim();
    if (id === 'accountNo' && !/^\d{10}$/.test(value)) invalidFields.push('Account Number (must be exactly 10 digits)');
    if (id === 'customerPhone' && !/^\d{11}$/.test(value)) invalidFields.push('Phone (must be exactly 11 digits)');
    if (id === 'customerNIN' && !/^\d{11}$/.test(value)) invalidFields.push('NIN (must be exactly 11 digits)');
  });

  if (missingFields.length > 0) {
    showNotification('Please fill all required fields before saving', 'error');
    return false;
  }

  if (invalidFields.length > 0) {
    showNotification('Invalid fields: ' + invalidFields.join(', '), 'error');
    return false;
  }

  const rsaBalance = parseFloat(document.getElementById('rsaBalance')?.value || 0);

  // Calculate exact 25% (no rounding)
  const rsa25Exact = rsaBalance * 0.25;

  // Update RSA 25% display with exact value
  const rsa25FormattedEl = document.getElementById('rsa25Formatted');
  if (rsa25FormattedEl) {
    rsa25FormattedEl.textContent = formatCurrency(rsa25Exact);
  }

  // Store exact value in hidden input
  const rsa25PercentEl = document.getElementById('rsa25Percent');
  if (rsa25PercentEl) {
    rsa25PercentEl.value = rsa25Exact;
  }

  // Call property update function
  if (typeof setPropertyByRsaAndUpdate === 'function') {
    setPropertyByRsaAndUpdate(rsaBalance);
  } else if (window.setPropertyByRsaAndUpdate) {
    window.setPropertyByRsaAndUpdate(rsaBalance);
  }

  const dob = document.getElementById('customerDob')?.value;
  if (dob) {
    const tenor = window.calculateTenorFromDob ? window.calculateTenorFromDob(dob) : '';
    const tenorEl = document.getElementById('tenor');
    if (tenorEl) {
      tenorEl.value = tenor;
    }
  }

  // Calculate and display customer info
  if (typeof calculateAndDisplayCustomerInfo === 'function') {
    calculateAndDisplayCustomerInfo();
  }

  customerDetailsSaved = true;
  showNotification('✅ Customer details saved successfully!', 'success');

  if (batchUploadBtn) {
    batchUploadBtn.disabled = false;
  }

  // Enable document grid
  if (documentGrid) {
    documentGrid.style.pointerEvents = 'auto';
    documentGrid.style.opacity = '1';
  }
  if (optionalDocumentGrid) {
    optionalDocumentGrid.style.pointerEvents = 'auto';
    optionalDocumentGrid.style.opacity = '1';
  }

  return true;
}

function resetCustomerDetails() {
  const fields = [
    'customerName', 'customerDob', 'customerEmail', 'customerPhone',
    'customerNIN', 'customerAddress', 'accountNo', 'employer',
    'originatingTP', 'pfa', 'penNo', 'rsaStatementDate', 'rsaBalance',
    'propertyType', 'tenor'
  ];

  // Reset all text fields
  fields.forEach(id => {
    const el = document.getElementById(id);
    if (el) el.value = '';
  });

  // Reset RSA 25% fields with null checks
  const rsa25PercentEl = document.getElementById('rsa25Percent');
  if (rsa25PercentEl) rsa25PercentEl.value = '';

  const rsa25FormattedEl = document.getElementById('rsa25Formatted');
  if (rsa25FormattedEl) rsa25FormattedEl.textContent = '';

  // Reset property value fields
  const propertyValueEl = document.getElementById('propertyValue');
  if (propertyValueEl) propertyValueEl.value = '';

  const propertyValueFormattedEl = document.getElementById('propertyValueFormatted');
  if (propertyValueFormattedEl) propertyValueFormattedEl.textContent = '';

  // Reset facility fee fields
  const facilityFeeEl = document.getElementById('facilityFee');
  if (facilityFeeEl) facilityFeeEl.value = '';

  const facilityFeeFormattedEl = document.getElementById('facilityFeeFormatted');
  if (facilityFeeFormattedEl) facilityFeeFormattedEl.textContent = '';

  // Reset loan amount fields
  const loanAmountEl = document.getElementById('loanAmount');
  if (loanAmountEl) loanAmountEl.value = '';

  const loanAmountFormattedEl = document.getElementById('loanAmountFormatted');
  if (loanAmountFormattedEl) loanAmountFormattedEl.textContent = '';

  // Hide results container
  const resultsContainer = document.getElementById('customerInfoResults');
  if (resultsContainer) {
    resultsContainer.style.display = 'none';
  }

  // Reset state
  customerDetailsSaved = false;

  // Disable batch upload button
  if (batchUploadBtn) {
    batchUploadBtn.disabled = true;
  }

  // Disable document grid
  if (documentGrid) {
    documentGrid.style.pointerEvents = 'none';
    documentGrid.style.opacity = '0.5';
  }
  if (optionalDocumentGrid) {
    optionalDocumentGrid.style.pointerEvents = 'none';
    optionalDocumentGrid.style.opacity = '0.5';
  }

  showNotification('Form reset', 'info');
}

function validateDetailsBeforeBatch() {
  if (!customerDetailsSaved) {
    showNotification('Please save customer details before batch uploading', 'error');
    return false;
  }

  const customerNameInput = document.getElementById('customerName');
  const customerName = customerNameInput ? customerNameInput.value.trim() : '';

  if (!customerName) {
    showNotification('Customer name is required', 'error');
    return false;
  }

  return true;
}

// ==================== EVENT LISTENERS ====================
function setupEventListeners() {
  if (newUploadBtn) newUploadBtn.addEventListener('click', openUploadModal);
  if (helpFab && sopHelpModal) {
    helpFab.addEventListener('click', () => sopHelpModal.classList.add('active'));
  }
  if (closeSopHelpModalBtn && sopHelpModal) {
    closeSopHelpModalBtn.addEventListener('click', () => sopHelpModal.classList.remove('active'));
  }
  if (closeSopHelpModalFooterBtn && sopHelpModal) {
    closeSopHelpModalFooterBtn.addEventListener('click', () => sopHelpModal.classList.remove('active'));
  }
  document.querySelectorAll('.nav-item[data-tab]').forEach(item => {
    item.addEventListener('click', (e) => { e.preventDefault(); switchTab(item.dataset.tab); });
  });
  if (rejectedTableBody) {
    rejectedTableBody.addEventListener('click', (e) => {
      const chatTrigger = e.target.closest('.app-chat-trigger');
      if (chatTrigger) return;
      const waLink = e.target.closest('a[href*="wa.me/"]');
      if (!waLink) return;
      const row = waLink.closest('tr');
      const submissionId = row?.getAttribute('data-submission-id');
      if (!submissionId) return;
      e.preventDefault();
      e.stopPropagation();
      if (typeof window.openApplicationChat === 'function') {
        window.openApplicationChat(submissionId);
      }
    });
  }
  if (closeModalBtn) closeModalBtn.addEventListener('click', () => closeModal(uploadModal));
  if (closeEditModalBtn) closeEditModalBtn.addEventListener('click', () => closeModal(editModal));
  if (closeSingleModalBtn) closeSingleModalBtn.addEventListener('click', () => closeModal(singleUploadModal));
  if (cancelSingleBtn) cancelSingleBtn.addEventListener('click', () => closeModal(singleUploadModal));
  if (submitCustomerBtn) submitCustomerBtn.addEventListener('click', submitCustomer);
  if (submitEditBtn) submitEditBtn.addEventListener('click', submitEdit);
  if (singleUploadArea) {
    setupDragDrop(singleUploadArea, singleFileInput, async (files) => {
      if (files.length > 0) await handleSingleFileSelection(files[0]);
    });
  }
  if (singleFileInput) {
    singleFileInput.addEventListener('change', async (e) => {
      if (e.target.files.length > 0) await handleSingleFileSelection(e.target.files[0]);
    });
  }
  if (batchUploadBtn && batchFileInput) {
    batchUploadBtn.disabled = true;
    batchUploadBtn.addEventListener('click', (e) => {
      e.preventDefault();
      if (validateDetailsBeforeBatch()) batchFileInput.click();
    });
    batchFileInput.addEventListener('change', (e) => {
      const files = Array.from(e.target.files);
      if (files.length > 0) {
        // Only enforce the 1.5MB limit at selection time for PDFs.
        // Images can be larger because we compress them before converting/uploading.
        const oversizedFiles = files.filter((file) => {
          const isPdf = (file?.type === 'application/pdf') || String(file?.name || '').toLowerCase().endsWith('.pdf');
          return isPdf && file.size > MAX_PDF_SIZE_BYTES;
        });
        const validFiles = files.filter((file) => {
          const isPdf = (file?.type === 'application/pdf') || String(file?.name || '').toLowerCase().endsWith('.pdf');
          return !isPdf || file.size <= MAX_PDF_SIZE_BYTES;
        });
        if (oversizedFiles.length > 0) {
          showBatchSizeWarningModal(oversizedFiles, validFiles, files);
        } else {
          prepareBatchForMapping(files);
        }
        batchFileInput.value = '';
      }
    });
  }
  // FIX: Don't exit early if confirmSingleUpload is missing - still show modal
  if (confirmSingleUpload) {
    confirmSingleUpload.addEventListener('click', uploadSingleDocument);
  }
  document.querySelectorAll('.browse-text').forEach(text => {
    text.addEventListener('click', (e) => {
      const input = e.target.closest('.upload-area').querySelector('input[type="file"]');
      if (input) input.click();
    });
  });
  if (customerNameInput) customerNameInput.addEventListener('input', updateSubmitButton);
  if (saveDetailsBtn) saveDetailsBtn.addEventListener('click', saveCustomerDetails);
  if (resetDetailsBtn) resetDetailsBtn.addEventListener('click', resetCustomerDetails);
  if (agentRegistrationForm) {
    agentRegistrationForm.addEventListener('submit', handleAgentRegistration);
  }
  if (openAgentRegistrationBtn) {
    openAgentRegistrationBtn.addEventListener('click', openAgentRegistrationModal);
  }
  if (closeAgentRegistrationModalBtn) {
    closeAgentRegistrationModalBtn.addEventListener('click', closeAgentRegistrationModal);
  }
  if (resetAgentFormBtn) {
    resetAgentFormBtn.addEventListener('click', () => {
      agentRegistrationForm?.reset();
    });
  }

  // ===== UPDATED: RSA Balance input listener with auto-fill =====
  document.getElementById('rsaBalance')?.addEventListener('input', (e) => {
    const rawStr = String(e.target.value).replace(/[^0-9.\-]+/g, '');
    const raw = parseFloat(rawStr) || 0;

    // Calculate exact 25% (no rounding)
    const rsa25Exact = raw * 0.25;

    // Update RSA 25% display immediately
    const rsa25FormattedEl = document.getElementById('rsa25Formatted');
    if (rsa25FormattedEl) {
      rsa25FormattedEl.textContent = formatCurrency(rsa25Exact);
    }

    // Store exact value in hidden input
    const rsa25PercentEl = document.getElementById('rsa25Percent');
    if (rsa25PercentEl) {
      rsa25PercentEl.value = rsa25Exact;
    }

    // Call property update function to auto-fill all property fields
    if (typeof setPropertyByRsaAndUpdate === 'function') {
      setPropertyByRsaAndUpdate(raw);
    } else if (window.setPropertyByRsaAndUpdate) {
      window.setPropertyByRsaAndUpdate(raw);
    }

    // Update customer info display
    if (typeof calculateAndDisplayCustomerInfo === 'function') {
      calculateAndDisplayCustomerInfo();
    }
  });

  // ===== UPDATED: DOB change listener to update tenor =====
  document.getElementById('customerDob')?.addEventListener('change', (e) => {
    const dob = e.target.value;
    const t = window.calculateTenorFromDob ? window.calculateTenorFromDob(dob) : '';
    const tenorEl = document.getElementById('tenor');
    if (tenorEl) tenorEl.value = t;

    // Also recalculate property fields if RSA balance exists
    const rsaBalance = document.getElementById('rsaBalance')?.value;
    if (rsaBalance && rsaBalance.trim() !== '') {
      const rsaVal = parseFloat(rsaBalance) || 0;
      if (typeof setPropertyByRsaAndUpdate === 'function') {
        setPropertyByRsaAndUpdate(rsaVal);
      } else if (window.setPropertyByRsaAndUpdate) {
        window.setPropertyByRsaAndUpdate(rsaVal);
      }
    }
  });

  if (document.getElementById('closeViewer')) {
    document.getElementById('closeViewer').addEventListener('click', closeViewerModal);
  }
  window.addEventListener('click', (e) => {
    if (e.target === uploadModal) closeModal(uploadModal);
    if (e.target === editModal) closeModal(editModal);
    if (e.target === singleUploadModal) closeModal(singleUploadModal);
    if (e.target === viewerModal) closeViewerModal();
    if (e.target === sopHelpModal) sopHelpModal.classList.remove('active');
  });
  if (closeBatchMappingBtn) closeBatchMappingBtn.addEventListener('click', () => batchMappingModal.classList.remove('active'));
  if (cancelBatchMapping) cancelBatchMapping.addEventListener('click', () => batchMappingModal.classList.remove('active'));
  if (confirmBatchMapping) {
    confirmBatchMapping.addEventListener('click', async () => {
      const mapped = __batchFilesBuffer.filter(b => b.selectedType).map(b => ({ file: b.originalFile, mappedType: b.selectedType }));
      if (mapped.length === 0) { showNotification('No files selected for upload', 'error'); return; }
      batchMappingModal.classList.remove('active');
      await handleBatchFiles(mapped);
    });
  }
  try {
    const rsaVal = parseFloat(document.getElementById('rsaBalance')?.value || 0);
    if (rsaVal && window.setPropertyByRsaAndUpdate) window.setPropertyByRsaAndUpdate(rsaVal);
    const dob = document.getElementById('customerDob')?.value;
    if (dob && window.calculateTenorFromDob) {
      const t = window.calculateTenorFromDob(dob);
      const tenorEl = document.getElementById('tenor'); if (tenorEl) tenorEl.value = t;
    }
  } catch (e) { }

  activeCommissionTabBtn?.addEventListener('click', (e) => {
    e.preventDefault();
    switchCommissionTab('active');
  });
  clearedCommissionTabBtn?.addEventListener('click', (e) => {
    e.preventDefault();
    switchCommissionTab('cleared');
  });

  [activeCommissionSearch, activeCommissionStartDate, activeCommissionEndDate].forEach((el) => {
    el?.addEventListener('input', () => renderPaidTable());
    el?.addEventListener('change', () => renderPaidTable());
  });
  [clearedCommissionSearch, clearedCommissionStartDate, clearedCommissionEndDate].forEach((el) => {
    el?.addEventListener('input', () => renderPaidTable());
    el?.addEventListener('change', () => renderPaidTable());
  });
}

// ==================== TAB SWITCHING ====================
window.switchTab = (tabId) => {
  document.querySelectorAll('.nav-item').forEach(item => item.classList.remove('active'));
  document.querySelector(`[data-tab="${tabId}"]`)?.classList.add('active');
  document.querySelectorAll('.tab-content').forEach(content => content.classList.remove('active'));
  document.getElementById(`${tabId}Tab`)?.classList.add('active');
  const titles = { overview: 'Dashboard', pending: 'Pending Documents', approved: 'Approved Documents', rejected: 'Rejected Documents', paid: 'Paid Customers', 'register-agent': 'Register Agent', profile: 'My Profile' };
  if (pageTitle) pageTitle.textContent = titles[tabId] || 'My Documents';
  if (tabId === 'paid') {
    switchCommissionTab(currentCommissionTab || 'active');
    renderPaidTable();
  }
  if (tabId === 'register-agent') {
    loadApprovedAgents();
  }
};

function switchCommissionTab(tab) {
  currentCommissionTab = tab === 'cleared' ? 'cleared' : 'active';
  if (activeCommissionSection) activeCommissionSection.style.display = currentCommissionTab === 'active' ? 'block' : 'none';
  if (clearedCommissionSection) clearedCommissionSection.style.display = currentCommissionTab === 'cleared' ? 'block' : 'none';
  activeCommissionTabBtn?.classList.toggle('active', currentCommissionTab === 'active');
  clearedCommissionTabBtn?.classList.toggle('active', currentCommissionTab === 'cleared');

  const showActive = currentCommissionTab === 'active';
  if (activeTotal25Card) activeTotal25Card.style.display = showActive ? 'block' : 'none';
  if (activeTotal1Card) activeTotal1Card.style.display = showActive ? 'block' : 'none';
  if (clearedTotal25Card) clearedTotal25Card.style.display = showActive ? 'none' : 'block';
  if (clearedTotal1Card) clearedTotal1Card.style.display = showActive ? 'none' : 'block';
}

function getDateFromAny(value) {
  if (!value) return null;
  try {
    if (value.toDate && typeof value.toDate === 'function') return value.toDate();
    if (value.seconds) return new Date(value.seconds * 1000);
    const d = new Date(value);
    return Number.isNaN(d.getTime()) ? null : d;
  } catch (_) {
    return null;
  }
}

function inDateRange(dateObj, startStr, endStr) {
  if (!dateObj) return false;
  if (!startStr && !endStr) return true;

  if (startStr) {
    const start = new Date(startStr);
    start.setHours(0, 0, 0, 0);
    if (dateObj < start) return false;
  }
  if (endStr) {
    const end = new Date(endStr);
    end.setHours(23, 59, 59, 999);
    if (dateObj > end) return false;
  }
  return true;
}

// ==================== DRAG & DROP ====================
function setupDragDrop(area, input, callback) {
  area.addEventListener('dragover', (e) => { e.preventDefault(); area.classList.add('drag-over'); });
  area.addEventListener('dragleave', () => area.classList.remove('drag-over'));
  area.addEventListener('drop', (e) => { e.preventDefault(); area.classList.remove('drag-over'); callback(e.dataTransfer.files); });
}

// ==================== MODAL FUNCTIONS ====================
function openUploadModal() {
  currentEditId = null;
  currentCustomerUploads = {};
  customerNameInput.value = '';
  if (customerAgentSelect) customerAgentSelect.value = '';
  customerDetailsSaved = false;
  if (batchUploadBtn) batchUploadBtn.disabled = true;
  if (window.renderDocumentGridUpload) {
    window.renderDocumentGridUpload(documentGrid, REQUIRED_DOC_TYPES, currentCustomerUploads, 'upload');
    window.renderDocumentGridUpload(optionalDocumentGrid, OPTIONAL_DOC_TYPES, currentCustomerUploads, 'upload');
  }
  // New submission: lock document grids until details are saved.
  if (documentGrid) { documentGrid.style.pointerEvents = 'none'; documentGrid.style.opacity = '0.5'; }
  if (optionalDocumentGrid) { optionalDocumentGrid.style.pointerEvents = 'none'; optionalDocumentGrid.style.opacity = '0.5'; }
  updateSubmitButton();
  uploadModal.classList.add('active');
}

function closeModal(modal) {
  modal.classList.remove('active');
  if (modal === singleUploadModal) {
    currentFile = null;
    clearSingleFilePreview();
    if (confirmSingleUpload) confirmSingleUpload.disabled = true;
    if (singleFileInput) singleFileInput.value = '';
  }
}

function closeViewerModal() {
  if (viewerModal) viewerModal.classList.remove('active');
  if (documentViewer) documentViewer.src = '';
  const viewerHeader = document.querySelector('.viewer-header');
  if (viewerHeader) {
    const existingNav = viewerHeader.querySelector('.viewer-nav');
    if (existingNav) existingNav.remove();
  }
}

// ==================== RENDER DOCUMENT GRID ====================
function renderDocumentGrid(container, documentTypes, uploadedDocs, mode) {
  if (!container) return;
  container.innerHTML = documentTypes.map(doc => {
    const isUploaded = uploadedDocs[doc.id] && uploadedDocs[doc.id].length > 0;
    const statusEmoji = isUploaded ? '✅' : '⬜';
    const statusClass = isUploaded ? 'uploaded' : 'pending';
    return `
      <div class="document-grid-item ${statusClass}" data-doc-type="${doc.id}">
        <div class="doc-status"><span class="status-emoji">${statusEmoji}</span></div>
        <div class="doc-icon"><i class="fas ${doc.icon}"></i></div>
        <div class="doc-name">${doc.name}</div>
        ${isUploaded ? `
          <div class="doc-actions">
            <button class="doc-btn view-btn" onclick="window.viewDocument('${doc.id}')"><i class="fas fa-eye"></i> View</button>
            ${mode === 'edit' ? `
              <button class="doc-btn upload-btn" onclick="window.openSingleUpload('${doc.id}', '${doc.name}')"><i class="fas fa-upload"></i> Re-upload</button>
              <button class="doc-btn remove-btn" onclick="window.removeUploadedDoc('${doc.id}')"><i class="fas fa-trash"></i> Remove</button>
            ` : `<button class="doc-btn remove-btn" onclick="window.removeUploadedDoc('${doc.id}')"><i class="fas fa-trash"></i> Remove</button>`}
          </div>
        ` : `<div class="doc-actions"><button class="doc-btn upload-btn" onclick="window.openSingleUpload('${doc.id}', '${doc.name}')"><i class="fas fa-upload"></i> Upload</button></div>`}
      </div>
    `;
  }).join('');
}
window.renderDocumentGridUpload = renderDocumentGrid;

// ==================== SINGLE UPLOAD ====================
window.openSingleUpload = (docType, docName) => {
  if (!customerDetailsSaved && !currentEditId) { showNotification('Please save customer details first', 'error'); return; }
  currentDocType = docType;
  if (uploadModalTitle) uploadModalTitle.textContent = `Upload ${docName}`;
  if (uploadDocType) uploadDocType.textContent = docName;
  singleUploadModal.classList.add('active');
};

function previewSingleFile(file) {
  const fileType = file.type;
  const isImage = fileType.startsWith('image/');
  if (singlePreviewObjectUrl) { URL.revokeObjectURL(singlePreviewObjectUrl); singlePreviewObjectUrl = null; }
  const previewSrc = isImage ? URL.createObjectURL(file) : '';
  if (isImage) singlePreviewObjectUrl = previewSrc;
  if (singleFilePreview) {
    singleFilePreview.innerHTML = `
      <div class="single-file-preview">
        ${isImage ? `<img src="${previewSrc}" class="preview-image">` : `<i class="fas ${getFileIcon(file.name)} file-icon-large"></i>`}
        <div class="file-details">
          <div class="file-name">${file.name}</div>
          <div class="file-size">${formatFileSize(file.size)}</div>
        </div>
      </div>
    `;
  }
}

function clearSingleFilePreview() {
  if (singlePreviewObjectUrl) { URL.revokeObjectURL(singlePreviewObjectUrl); singlePreviewObjectUrl = null; }
  if (singleFilePreview) singleFilePreview.innerHTML = '';
}

function canvasToJpegBlob(canvas, quality) {
  return new Promise((resolve) => { canvas.toBlob((blob) => resolve(blob), 'image/jpeg', quality); });
}

async function compressImageTo200KB(file, targetBytes = MAX_IMAGE_UPLOAD_BYTES) {
  const dataUrl = await fileToDataURL(file);
  const img = await loadImage(dataUrl);
  const canvas = document.createElement('canvas');
  const ctx = canvas.getContext('2d', { alpha: false });
  if (!ctx) throw new Error('Image processing is not supported on this device.');
  let width = img.width, height = img.height;
  const maxEdge = 1920;
  const longest = Math.max(width, height);
  if (longest > maxEdge) { const ratio = maxEdge / longest; width = Math.max(1, Math.round(width * ratio)); height = Math.max(1, Math.round(height * ratio)); }
  let bestBlob = null;
  for (let pass = 0; pass < 10; pass++) {
    canvas.width = Math.max(1, width); canvas.height = Math.max(1, height);
    ctx.fillStyle = '#ffffff'; ctx.fillRect(0, 0, canvas.width, canvas.height);
    ctx.drawImage(img, 0, 0, canvas.width, canvas.height);
    for (let quality = 0.9; quality >= 0.35; quality -= 0.07) {
      const blob = await canvasToJpegBlob(canvas, quality);
      if (!blob) continue;
      if (!bestBlob || blob.size < bestBlob.size) bestBlob = blob;
      if (blob.size <= targetBytes) {
        const baseName = (file.name || 'photo').replace(/\.[^/.]+$/, '');
        return new File([blob], `${baseName}.jpg`, { type: 'image/jpeg' });
      }
    }
    width = Math.max(360, Math.round(width * 0.85));
    height = Math.max(360, Math.round(height * 0.85));
    if (width <= 360 && height <= 360) break;
  }
  if (bestBlob && bestBlob.size <= targetBytes) {
    const baseName = (file.name || 'photo').replace(/\.[^/.]+$/, '');
    return new File([bestBlob], `${baseName}.jpg`, { type: 'image/jpeg' });
  }
  const targetMb = (targetBytes / (1024 * 1024)).toFixed(0);
  throw new Error(`Could not reduce image below ${targetMb}MB. Please retake at a lower resolution.`);
}

async function handleSingleFileSelection(file) {

  // FIX: Don't exit early if confirmSingleUpload is missing - still show modal
  if (!confirmSingleUpload) {
    // Silently continue
  }

  const isImage = Boolean(file?.type && file.type.startsWith('image/'));
  const isPdf = (file?.type === 'application/pdf') || (String(file?.name || '').toLowerCase().endsWith('.pdf'));

  // PDFs (and unknown files) must obey the 1.5MB cap; images can be larger because we compress them.
  if (!isImage && file.size > MAX_PDF_SIZE_BYTES) {
    showFileSizeWarningModal(file);
    currentFile = null;
    if (confirmSingleUpload) confirmSingleUpload.disabled = true;
    clearSingleFilePreview();
    return;
  }
  if (isPdf && file.size > MAX_PDF_SIZE_BYTES) {
    showFileSizeWarningModal(file);
    currentFile = null;
    if (confirmSingleUpload) confirmSingleUpload.disabled = true;
    clearSingleFilePreview();
    return;
  }

  if (confirmSingleUpload) confirmSingleUpload.disabled = true;
  clearSingleFilePreview();

  try {
    let preparedFile = file;
    if (isImage) {
      // Requirement: on mobile camera snap, compress only when the photo is > 1MB.
      if (file.size > MAX_IMAGE_UPLOAD_BYTES) {
        showNotification('Optimizing photo to 1MB for upload...', 'info');
        preparedFile = await compressImageTo200KB(file, MAX_IMAGE_UPLOAD_BYTES);
        if (preparedFile.size > MAX_IMAGE_UPLOAD_BYTES) throw new Error('Photo must be 1MB or less after compression.');
      }
    }
    currentFile = preparedFile;
    previewSingleFile(preparedFile);
    if (confirmSingleUpload) confirmSingleUpload.disabled = false;
  } catch (error) {
    currentFile = null;
    showNotification(error.message || 'Could not process selected image.', 'error');
  }
}

function fileToDataURL(file) {
  return new Promise((res, rej) => {
    const reader = new FileReader();
    reader.onload = () => res(reader.result);
    reader.onerror = rej;
    reader.readAsDataURL(file);
  });
}

function loadImage(dataUrl) {
  return new Promise((res, rej) => {
    const img = new Image();
    img.onload = () => res(img);
    img.onerror = rej;
    img.src = dataUrl;
  });
}

// ==================== FIXED: IMAGE TO PDF CONVERSION ====================
async function convertImageToPdf(file) {

  try {
    // Check if jsPDF is available
    let jsPDFConstructor = null;

    // Try multiple ways to get jsPDF
    if (window.jspdf && window.jspdf.jsPDF) {
      jsPDFConstructor = window.jspdf.jsPDF;
    } else if (window.jsPDF && typeof window.jsPDF === 'function') {
      jsPDFConstructor = window.jsPDF;
    } else if (typeof jspdf !== 'undefined' && jspdf.jsPDF) {
      jsPDFConstructor = jspdf.jsPDF;
    } else {
      // Show user-friendly error
      showNotification('⚠️ Image to PDF conversion unavailable. Please upload PDF files directly.', 'error');
      throw new Error('jsPDF library not available. Please upload PDF files directly.');
    }

    const dataUrl = await fileToDataURL(file);
    const img = await loadImage(dataUrl);

    // Create PDF
    const pdf = new jsPDFConstructor({
      orientation: 'portrait',
      unit: 'pt',
      format: 'a4'
    });

    const pageWidth = pdf.internal.pageSize.getWidth();
    const pageHeight = pdf.internal.pageSize.getHeight();

    // Calculate image dimensions to fit on page
    let imgWidth = img.width;
    let imgHeight = img.height;

    // Calculate scaling ratio to fit image on page (with 10% margins)
    const widthRatio = (pageWidth * 0.9) / imgWidth;
    const heightRatio = (pageHeight * 0.9) / imgHeight;
    const ratio = Math.min(widthRatio, heightRatio);

    const scaledWidth = imgWidth * ratio;
    const scaledHeight = imgHeight * ratio;

    // Center on page
    const x = (pageWidth - scaledWidth) / 2;
    const y = (pageHeight - scaledHeight) / 2;

    // Add image to PDF
    pdf.addImage(dataUrl, 'JPEG', x, y, scaledWidth, scaledHeight);

    // Get PDF as blob
    const pdfBlob = pdf.output('blob');

    return pdfBlob;

  } catch (error) {
    // Check if it's a jsPDF loading error
    if (error.message && error.message.includes('jsPDF')) {
      showNotification('⚠️ Image to PDF conversion failed: jsPDF library not loaded', 'error');
    } else {
      showNotification('Image conversion failed: ' + error.message, 'error');
    }

    throw error;
  }
}

async function convertImageToPdfUnderLimit(imageFile) {
  const targets = [
    MAX_IMAGE_UPLOAD_BYTES,
    900 * 1024,
    800 * 1024,
    700 * 1024,
    600 * 1024,
    500 * 1024,
    400 * 1024,
    300 * 1024,
    250 * 1024,
    200 * 1024
  ];

  let lastError = null;
  for (const target of targets) {
    try {
      const inputFile = imageFile.size > target
        ? await compressImageTo200KB(imageFile, target)
        : imageFile;
      const pdfBlob = await convertImageToPdf(inputFile);
      if (pdfBlob.size <= MAX_PDF_SIZE_BYTES) {
        return { pdfBlob, inputFile };
      }
      lastError = new Error('Converted PDF still larger than 1.5MB');
    } catch (e) {
      lastError = e;
    }
  }
  throw lastError || new Error('Could not compress image to fit under 1.5MB PDF limit');
}

function findMatchingDocTypeByName(filename, usedTypes) {
  const name = filename.toLowerCase();
  const scores = DOCUMENT_TYPES.map(dt => {
    const idScore = name.includes(dt.id.replace(/_/g, ' ')) ? 10 : 0;
    const words = dt.name.toLowerCase().split(/[^a-z0-9]+/).filter(Boolean);
    let wordScore = 0;
    for (const w of words) if (name.includes(w)) wordScore += 2;
    return { id: dt.id, score: idScore + wordScore };
  }).sort((a,b) => b.score - a.score);
  for (const s of scores) {
    if (s.score <= 0) continue;
    if (usedTypes && usedTypes.includes(s.id)) continue;
    return s.id;
  }
  return null;
}

async function handleBatchFiles(files) {
  if (!files || files.length === 0) return;
  if (!customerDetailsSaved && !currentEditId) { showNotification('Please save customer details before batch uploading', 'error'); return; }
  const customerName = customerNameInput.value.trim();
  if (!customerName) { showNotification('Customer name is required', 'error'); return; }
  if (window.__uploadInProgress) return showNotification('Upload already in progress', 'error');
  window.__uploadInProgress = true;
  const total = files.length;
  if (window.showLoader) window.showLoader('Uploading batch files... 0%');
  try {
    const storage = new BackblazeStorage();
    // Track doc types already occupied (existing uploads + current batch) so two files can't map to one type.
    const usedTypes = new Set(Object.keys(currentCustomerUploads || {}).filter((id) => currentCustomerUploads[id] && currentCustomerUploads[id].length > 0));
    let successCount = 0;
    let failCount = 0;

    for (let idx = 0; idx < files.length; idx++) {
      const item = files[idx];
      const file = item.file || item;
      const mappedType = item.mappedType || item.type || null;

      const isImage = Boolean(file?.type && file.type.startsWith('image/'));
      const isPdf = (file?.type === 'application/pdf') || (String(file?.name || '').toLowerCase().endsWith('.pdf'));
      if (isPdf && file.size > MAX_PDF_SIZE_BYTES) {
        showNotification(`⚠️ ${file.name} exceeds 1.5MB, skipping.`, 'warning');
        failCount++;
        continue;
      }
      // Images can be larger because we compress before converting to PDF.
      if (!isImage && !isPdf && file.size > MAX_PDF_SIZE_BYTES) {
        showNotification(`⚠️ ${file.name} exceeds 1.5MB, skipping.`, 'warning');
        failCount++;
        continue;
      }

      let targetType = mappedType || findMatchingDocTypeByName(file.name, Array.from(usedTypes));
      if (!targetType) {
        const missing = DOCUMENT_TYPES.filter(d => d.required !== false).map(d => d.id).filter(id => !(currentCustomerUploads[id] && currentCustomerUploads[id].length > 0));
        targetType = missing.find((id) => !usedTypes.has(id)) || null;
      }
      if (!targetType) {
        const notUsed = DOCUMENT_TYPES.map(d => d.id).filter(id => !(currentCustomerUploads[id] && currentCustomerUploads[id].length > 0));
        targetType = notUsed.find((id) => !usedTypes.has(id)) || DOCUMENT_TYPES[0].id;
      }
      // Prevent two different files from mapping to the same slot in the same batch.
      if (usedTypes.has(targetType)) {
        showNotification(`⚠️ ${file.name} could not be mapped because "${targetType}" is already selected/filled. Use batch mapping to choose another type.`, 'warning');
        failCount++;
        continue;
      }
      usedTypes.add(targetType);

      let fileToSend = file;
      try {
        if (isImage) {
          // Check jsPDF first
          if (!window.jspdf || (!window.jspdf.jsPDF && !window.jsPDF)) {
            showNotification(`⚠️ Cannot convert ${file.name} to PDF. Please upload PDF files directly.`, 'error');
            failCount++;
            continue;
          }

          const optimizedImage = file.size > MAX_IMAGE_UPLOAD_BYTES
            ? await compressImageTo200KB(file, MAX_IMAGE_UPLOAD_BYTES)
            : file;
          let pdfBlob = await convertImageToPdf(optimizedImage);
          if (pdfBlob.size > MAX_PDF_SIZE_BYTES) {
            try {
              const converted = await convertImageToPdfUnderLimit(file);
              pdfBlob = converted.pdfBlob;
            } catch (e) {
              showNotification(`Converted ${file.name} > 1.5MB, skipped.`, 'error');
              failCount++;
              continue;
            }
          }
          const newName = file.name.replace(/\.[^/.]+$/, '') + '.pdf';
          fileToSend = new File([pdfBlob], newName, { type: 'application/pdf' });
        } else if (file.type === 'application/pdf' || file.name.toLowerCase().endsWith('.pdf')) {
          // Already validated
        } else {
          showNotification(`${file.name} unsupported, only images or PDFs allowed`, 'error');
          failCount++;
          continue;
        }
      } catch (e) {
        showNotification(`Failed to convert ${file.name}`, 'error');
        failCount++;
        continue;
      }

      try {
        const result = await storage.uploadFile(fileToSend, customerName, targetType);
        if (!currentCustomerUploads[targetType]) currentCustomerUploads[targetType] = [];
        currentCustomerUploads[targetType].push({ name: fileToSend.name, fileId: result.fileId, fileUrl: result.fileUrl, uploadedAt: new Date().toISOString(), localAddedAt: Date.now() });
        successCount++;
      } catch (e) {
        showNotification(`Upload failed for ${file.name}`, 'error');
        failCount++;
      }

      try {
        const done = idx + 1;
        const percent = Math.round((done / total) * 100);
        if (window.showLoader) window.showLoader(`Uploading batch files... ${percent}%`);
      } catch (e) { }
    }

    if (window.renderDocumentGridUpload) {
      window.renderDocumentGridUpload(documentGrid, REQUIRED_DOC_TYPES, currentCustomerUploads, 'edit');
      window.renderDocumentGridUpload(optionalDocumentGrid, OPTIONAL_DOC_TYPES, currentCustomerUploads, 'edit');
    }
    updateSubmitButton();
    showNotification(`Batch upload complete: ${successCount} successful, ${failCount} failed`, successCount > 0 ? 'success' : 'error');
  } finally {
    hideLoader();
    window.__uploadInProgress = false;
    __batchFilesBuffer = [];
  }
}

function prepareBatchForMapping(files) {
  if (!files || files.length === 0) return;
  const alreadyUploaded = new Set(Object.keys(currentCustomerUploads || {}).filter(id => currentCustomerUploads[id] && currentCustomerUploads[id].length > 0));
  const usedInBatch = new Set(Array.from(alreadyUploaded));
  __batchFilesBuffer = files.map(f => {
    const g = findMatchingDocTypeByName(f.name, Array.from(usedInBatch));
    const selected = (g && !usedInBatch.has(g)) ? g : null;
    if (selected) usedInBatch.add(selected);
    return { originalFile: f, previewName: f.name, guessedType: g, selectedType: selected };
  });
  renderBatchMappingList();
  if (batchMappingModal) batchMappingModal.classList.add('active');
}

function renderBatchMappingList() {
  if (!batchMappingList) return;
  const alreadyUploaded = new Set(Object.keys(currentCustomerUploads || {}).filter(id => currentCustomerUploads[id] && currentCustomerUploads[id].length > 0));
  batchMappingList.innerHTML = __batchFilesBuffer.map((entry, idx) => {
    const options = DOCUMENT_TYPES.map(d => {
      const disabled = alreadyUploaded.has(d.id) && d.id !== entry.selectedType ? 'disabled' : '';
      const sel = d.id === entry.selectedType ? 'selected' : '';
      return `<option value="${d.id}" ${sel} ${disabled}>${d.name}</option>`;
    }).join('');
    return `
      <div class="batch-item" style="display:flex; gap:12px; align-items:center; padding:8px; border:1px solid #eef2f7; border-radius:8px;">
        <div style="width:40px; text-align:center;"><i class="fas ${getFileIcon(entry.previewName)}"></i></div>
        <div style="flex:1;">
          <div style="font-weight:600;">${entry.previewName}</div>
          <div style="font-size:12px; color:var(--text-muted);">${formatFileSize(entry.originalFile.size)}</div>
        </div>
        <div>
          <select data-idx="${idx}" class="batch-map-select">
            <option value="">-- Skip --</option>
            ${options}
          </select>
        </div>
      </div>
    `;
  }).join('');
  const selects = batchMappingList.querySelectorAll('.batch-map-select');
  selects.forEach(s => {
    const idx = parseInt(s.dataset.idx, 10);
    const sel = __batchFilesBuffer[idx].selectedType;
    if (sel) s.value = sel;
    s.addEventListener('change', (e) => { __batchFilesBuffer[idx].selectedType = e.target.value || null; refreshBatchSelectOptions(); });
  });
  refreshBatchSelectOptions();
}

function refreshBatchSelectOptions() {
  if (!batchMappingList) return;
  // Enforce uniqueness: if a doc type is chosen multiple times, keep the first and clear the rest.
  const seen = new Map(); // docType -> firstIdx
  let hadDuplicates = false;
  __batchFilesBuffer.forEach((b, idx) => {
    const t = b.selectedType;
    if (!t) return;
    if (!seen.has(t)) { seen.set(t, idx); return; }
    hadDuplicates = true;
    b.selectedType = null;
  });

  if (hadDuplicates) {
    // Sync UI to cleared selections.
    const selects = batchMappingList.querySelectorAll('.batch-map-select');
    selects.forEach(s => {
      const idx = parseInt(s.dataset.idx, 10);
      s.value = __batchFilesBuffer[idx].selectedType || '';
    });
    showNotification('Duplicate document mapping removed. Please reselect the cleared ones.', 'warning');
  }

  const chosen = new Set(__batchFilesBuffer.map(b => b.selectedType).filter(Boolean));
  const alreadyUploaded = new Set(Object.keys(currentCustomerUploads || {}).filter(id => currentCustomerUploads[id] && currentCustomerUploads[id].length > 0));
  const selects = batchMappingList.querySelectorAll('.batch-map-select');
  selects.forEach(s => {
    const idx = parseInt(s.dataset.idx, 10);
    const mySelected = __batchFilesBuffer[idx].selectedType;
    Array.from(s.options).forEach(opt => {
      if (!opt.value) { opt.disabled = false; return; }
      if (alreadyUploaded.has(opt.value) && opt.value !== mySelected) { opt.disabled = true; return; }
      opt.disabled = (opt.value !== mySelected && chosen.has(opt.value));
    });
  });
}

async function uploadSingleDocument() {
  if (!currentFile || !currentDocType) return;
  if (!customerDetailsSaved && !currentEditId) { showNotification('Please save customer details first', 'error'); return; }
  if (currentFile.size > MAX_PDF_SIZE_BYTES) { showFileSizeWarningModal(currentFile); return; }
  if (window.__uploadInProgress) { return; }
  window.__uploadInProgress = true;
  if (confirmSingleUpload) {
    confirmSingleUpload.disabled = true;
    confirmSingleUpload.innerHTML = '<i class="fas fa-spinner fa-spin"></i> Uploading...';
  }
  try {
    const storage = new BackblazeStorage();
    const customerName = customerNameInput.value.trim();
    if (!customerName) throw new Error('Please enter customer name first');
    let fileToSend = currentFile;
    if (currentFile.type.startsWith('image/')) {
      try {
        let pdfBlob = null;
        try {
          // Try a normal conversion first (fast path).
          const baseImage = currentFile.size > MAX_IMAGE_UPLOAD_BYTES
            ? await compressImageTo200KB(currentFile, MAX_IMAGE_UPLOAD_BYTES)
            : currentFile;
          pdfBlob = await convertImageToPdf(baseImage);
        } catch (e) {
          throw e;
        }

        if (pdfBlob.size > MAX_PDF_SIZE_BYTES) {
          const ok = confirm('Converted PDF is larger than 1.5MB.\n\nTry compressing the photo further and retry?');
          if (!ok) return;
          const converted = await convertImageToPdfUnderLimit(currentFile);
          pdfBlob = converted.pdfBlob;
        }
        const newName = currentFile.name.replace(/\.[^/.]+$/, '') + '.pdf';
        fileToSend = new File([pdfBlob], newName, { type: 'application/pdf' });
      } catch (e) {
        showNotification('Image conversion failed: ' + e.message, 'error');
        return;
      }
    } else if (currentFile.type === 'application/pdf' || currentFile.name.toLowerCase().endsWith('.pdf')) {
      // Already validated
    } else {
      showNotification('Unsupported file type. Please upload images or PDFs.', 'error');
      return;
    }
    const result = await storage.uploadFile(fileToSend, customerName, currentDocType);
    if (!currentCustomerUploads[currentDocType]) currentCustomerUploads[currentDocType] = [];
    currentCustomerUploads[currentDocType].push({ name: currentFile.name, fileId: result.fileId, fileUrl: result.fileUrl, uploadedAt: new Date().toISOString(), localAddedAt: Date.now() });
    renderDocumentGrid(documentGrid, REQUIRED_DOC_TYPES, currentCustomerUploads, 'upload');
    renderDocumentGrid(optionalDocumentGrid, OPTIONAL_DOC_TYPES, currentCustomerUploads, 'upload');
    updateSubmitButton();
    showNotification('✅ Document uploaded successfully!', 'success');
    closeModal(singleUploadModal);
  } catch (error) {
    showNotification('Upload failed: ' + error.message, 'error');
  } finally {
    if (confirmSingleUpload) {
      confirmSingleUpload.disabled = false;
      confirmSingleUpload.innerHTML = 'Upload Document';
    }
    window.__uploadInProgress = false;
  }
}

// ==================== VIEW DOCUMENT ====================
window.viewDocument = (docType) => {
  const docs = currentCustomerUploads[docType];
  if (!docs || docs.length === 0) return;
  const lastDoc = docs[docs.length - 1];
  window.open(lastDoc.fileUrl, '_blank');
};

// ==================== VIEW SUBMISSION DOCUMENTS ====================
window.viewSubmissionDocs = (submissionId) => {
  const sub = allSubmissions.find(s => s.id === submissionId);
  if (!sub || !sub.documents || sub.documents.length === 0) { showNotification('No documents available', 'error'); return; }
  const firstDoc = sub.documents[0];
  const docTypeLabel = DOCUMENT_TYPES.find(t => t.id === firstDoc.documentType)?.name || firstDoc.documentType || 'Document';
  if (viewerModal && viewerFileName && documentViewer) {
    viewerFileName.textContent = `${sub.customerName} - ${docTypeLabel}`;
    documentViewer.src = firstDoc.fileUrl?.trim();
    viewerModal.classList.add('active');
  }
  if (sub.documents.length > 1) {
    let currentIndex = 0;
    const showDoc = (index) => {
      const doc = sub.documents[index];
      const docTypeLabel = DOCUMENT_TYPES.find(t => t.id === doc.documentType)?.name || doc.documentType || 'Document';
      viewerFileName.textContent = `${sub.customerName} - ${docTypeLabel} (${index + 1}/${sub.documents.length})`;
      documentViewer.src = doc.fileUrl?.trim();
    };
    const addViewerNav = () => {
      const viewerHeader = viewerModal ? viewerModal.querySelector('.viewer-header') : document.querySelector('.viewer-header');
      if (!viewerHeader) return;
      const existingNav = viewerHeader.querySelector('.viewer-nav');
      if (existingNav) existingNav.remove();
      const nav = document.createElement('div');
      nav.className = 'viewer-nav';
      nav.style.cssText = 'display: flex; gap: 10px; align-items: center;';
      nav.innerHTML = `
        <button id="prevDoc" class="action-btn" ${currentIndex === 0 ? 'disabled' : ''}><i class="fas fa-chevron-left"></i> Prev</button>
        <span id="docCounter" style="font-size: 14px; color: #666;">${currentIndex + 1}/${sub.documents.length}</span>
        <button id="nextDoc" class="action-btn" ${currentIndex === sub.documents.length - 1 ? 'disabled' : ''}><i class="fas fa-chevron-right"></i> Next</button>
      `;
      const closeBtn = viewerHeader.querySelector('#closeViewer') || document.getElementById('closeViewer');
      viewerHeader.insertBefore(nav, closeBtn);
      document.getElementById('prevDoc').onclick = () => { if (currentIndex > 0) { currentIndex--; showDoc(currentIndex); addViewerNav(); } };
      document.getElementById('nextDoc').onclick = () => { if (currentIndex < sub.documents.length - 1) { currentIndex++; showDoc(currentIndex); addViewerNav(); } };
    };
    addViewerNav();
  }
};

// ==================== SHOW APPLICATION TRACKING ====================
window.showApplicationTrack = async (submissionId) => {
  const sub = allSubmissions.find(s => s.id === submissionId);
  if (!sub) return;
  try {
    const stage = await getApplicationStage(sub);
    const trackModal = document.createElement('div');
    trackModal.className = 'modal';
    trackModal.id = 'trackModal';
    trackModal.style.cssText = 'display: flex !important; position: fixed; top: 0; left: 0; width: 100%; height: 100%; background: rgba(0,0,0,0.5); z-index: 999999; align-items: center; justify-content: center;';
    trackModal.innerHTML = `
      <div class="modal-content" style="max-width: 500px;">
        <div class="modal-header">
          <h2><i class="fas fa-map-marker-alt"></i> Application Tracking</h2>
          <button class="close-btn" onclick="this.closest('.modal').remove()">&times;</button>
        </div>
        <div class="modal-body" style="padding: 30px;">
          <div style="text-align: center; margin-bottom: 30px;">
            <i class="fas fa-file-alt" style="font-size: 60px; color: var(--cm-primary);"></i>
            <h3 style="margin: 15px 0 5px;">${sub.customerName}</h3>
            <p style="color: var(--gray-500);">Application ID: ${sub.id.substring(0, 8)}...</p>
          </div>
          <div style="background: #f8fafc; border-radius: 12px; padding: 20px;">
            <div style="display: flex; align-items: center; gap: 15px; margin-bottom: 15px;">
              <div style="width: 40px; height: 40px; background: ${getStageColor(stage)}; border-radius: 50%; display: flex; align-items: center; justify-content: center;">
                <i class="fas ${getStageIcon(stage)}" style="color: white;"></i>
              </div>
              <div>
                <div style="font-size: 12px; color: var(--gray-500);">Current Stage</div>
                <div style="font-size: 18px; font-weight: 600; color: var(--gray-800);">${stage}</div>
              </div>
            </div>
            <div style="border-top: 1px solid var(--gray-200); padding-top: 15px;">
              <div style="display: flex; justify-content: space-between; margin-bottom: 10px;">
                <span style="color: var(--gray-600);">Status:</span>
                <span class="status-badge status-${sub.status}">${sub.status}</span>
              </div>
              <div style="display: flex; justify-content: space-between; margin-bottom: 10px;">
                <span style="color: var(--gray-600);">Uploaded:</span>
                <span>${safeFormatDate(sub.uploadedAt)}</span>
              </div>
              ${sub.assignedTo ? `<div style="display: flex; justify-content: space-between; margin-bottom: 10px;"><span style="color: var(--gray-600);">Assigned To:</span><span>${sub.assignedTo}</span></div>` : ''}
              ${sub.reviewedAt ? `<div style="display: flex; justify-content: space-between; margin-bottom: 10px;"><span style="color: var(--gray-600);">Reviewed:</span><span>${safeFormatDate(sub.reviewedAt)}</span></div>` : ''}
              ${sub.comment ? `<div style="margin-top: 15px;"><span style="color: var(--gray-600);">Comment:</span><p style="background: white; padding: 10px; border-radius: 8px; margin-top: 5px;">${sub.comment}</p></div>` : ''}
            </div>
          </div>
          <div style="margin-top: 25px; text-align: center;">
            <button class="action-btn" onclick="this.closest('.modal').remove()" style="padding: 10px 30px;">Close</button>
          </div>
        </div>
      </div>
    `;
    const target = document.body || document.documentElement;
    if (target) target.appendChild(trackModal);
    trackModal.addEventListener('click', (e) => { if (e.target === trackModal) trackModal.remove(); });
  } catch (error) {
    showNotification('Could not load tracking information', 'error');
  }
};

function getStageColor(stage) {
  if (stage.includes('Reviewer')) return '#3b82f6';
  if (stage.includes('Approved')) return '#10b981';
  if (stage.includes('Rejected')) return '#ef4444';
  if (stage.includes('RSA')) return '#8b5cf6';
  return '#f59e0b';
}
function getStageIcon(stage) {
  if (stage.includes('Reviewer')) return 'fa-user-check';
  if (stage.includes('Approved')) return 'fa-check-circle';
  if (stage.includes('Rejected')) return 'fa-times-circle';
  if (stage.includes('RSA')) return 'fa-building';
  return 'fa-clock';
}

// ==================== SUBMIT CUSTOMER ====================
function updateSubmitButton() {
  if (!submitCustomerBtn) return;
  const customerName = customerNameInput.value.trim();
  const hasAnyDoc = Object.keys(currentCustomerUploads || {}).some(id => currentCustomerUploads[id] && currentCustomerUploads[id].length > 0);
  if (currentEditId) { submitCustomerBtn.disabled = !(customerName && hasAnyDoc); return; }
  const requiredDocs = DOCUMENT_TYPES.filter(d => d.required !== false).map(d => d.id);
  const uploadedRequired = requiredDocs.filter(id => currentCustomerUploads[id] && currentCustomerUploads[id].length > 0).length;
  uploadedCountSpan.textContent = Object.keys(currentCustomerUploads || {}).filter((id) => currentCustomerUploads[id] && currentCustomerUploads[id].length > 0).length;
  submitCustomerBtn.disabled = !(customerName && uploadedRequired === requiredDocs.length && customerDetailsSaved);
}

function getMissingRequiredSubmissionFields() {
  const requiredFields = [
    { id: 'customerName', label: 'Customer Name' }, { id: 'customerDob', label: 'Date of Birth' },
    { id: 'customerEmail', label: 'Email' }, { id: 'customerPhone', label: 'Phone' },
    { id: 'customerNIN', label: 'NIN' }, { id: 'customerAddress', label: 'Address' },
    { id: 'accountNo', label: 'Account Number' }, { id: 'employer', label: 'Employer' },
    { id: 'originatingTP', label: 'Originating Transfer Pin' }, { id: 'pfa', label: 'PFA' },
    { id: 'penNo', label: 'PEN Number' }, { id: 'rsaStatementDate', label: 'RSA Statement Date' },
    { id: 'rsaBalance', label: 'RSA Balance' }, { id: 'propertyType', label: 'Property Type' },
    { id: 'propertyValue', label: 'Property Value' }, { id: 'loanAmount', label: 'Loan Amount' },
  ];
  const missingLabels = [], invalidLabels = [];
  requiredFields.forEach(field => {
    const el = document.getElementById(field.id);
    if (!el) return;
    const value = String(el.value || '').trim();
    if (!value) { missingLabels.push(field.label); return; }
    if (field.id === 'accountNo' && !/^\d{10}$/.test(value)) invalidLabels.push('Account Number (must be exactly 10 digits)');
    if (field.id === 'customerPhone' && !/^\d{11}$/.test(value)) invalidLabels.push('Phone (must be exactly 11 digits)');
    if (field.id === 'customerNIN' && !/^\d{11}$/.test(value)) invalidLabels.push('NIN (must be exactly 11 digits)');
  });
  return invalidLabels.length > 0 ? [...missingLabels, ...invalidLabels] : missingLabels;
}

async function submitCustomer() {
  const customerName = customerNameInput.value.trim();
  if (!customerName) return;
  if (!customerDetailsSaved) { showNotification('Please save customer details before submitting', 'error'); return; }
  submitCustomerBtn.disabled = true;
  submitCustomerBtn.innerHTML = '<i class="fas fa-spinner fa-spin"></i> Submitting...';
  try {
    const missingFields = getMissingRequiredSubmissionFields();
    if (missingFields.length > 0) {
      submitCustomerBtn.disabled = false;
      submitCustomerBtn.innerHTML = '<i class="fas fa-check-circle"></i> Submit Customer Documents';
      return showNotification('All fields are compulsory. Missing or invalid: ' + missingFields.join(', '), 'error');
    }
    if (currentEditId) {
      const hasAnyDoc = Object.keys(currentCustomerUploads || {}).some(id => currentCustomerUploads[id] && currentCustomerUploads[id].length > 0);
      if (!hasAnyDoc) {
        submitCustomerBtn.disabled = false;
        submitCustomerBtn.innerHTML = '<i class="fas fa-check-circle"></i> Submit Customer Documents';
        return showNotification('Please upload at least one document before submitting fix.', 'error');
      }
      const documents = [];
      Object.entries(currentCustomerUploads).forEach(([type, files]) => {
        files.forEach(file => { documents.push({ documentType: type, ...file }); });
      });
      const customerDetails = {
        name: customerName,
        dob: document.getElementById('customerDob')?.value || '',
        email: document.getElementById('customerEmail')?.value?.trim() || '',
        phone: document.getElementById('customerPhone')?.value?.trim() || '',
        nin: document.getElementById('customerNIN')?.value?.trim() || '',
        address: document.getElementById('customerAddress')?.value?.trim() || '',
        accountNo: document.getElementById('accountNo')?.value?.trim() || '',
        employer: document.getElementById('employer')?.value?.trim() || '',
        originatingTP: document.getElementById('originatingTP')?.value?.trim() || '',
        pfa: document.getElementById('pfa')?.value?.trim() || '',
        penNo: document.getElementById('penNo')?.value?.trim() || '',
        rsaStatementDate: document.getElementById('rsaStatementDate')?.value || '',
        rsaBalance: document.getElementById('rsaBalance')?.value || '',
        rsa25: document.getElementById('rsa25Percent')?.value || '',
        propertyType: document.getElementById('propertyType')?.value?.trim() || '',
        tenor: document.getElementById('tenor')?.value || '',
        propertyValue: document.getElementById('propertyValue')?.value || '',
        facilityFee: document.getElementById('facilityFee')?.value || '',
        loanAmount: document.getElementById('loanAmount')?.value || ''
      };
      const submissionRef = doc(db, 'submissions', currentEditId);
      const existingSub = allSubmissions.find((s) => s.id === currentEditId) || {};
      const reviewerToReassign = String(existingSub.reviewedBy || existingSub.assignedTo || '').trim();
      await updateDoc(submissionRef, {
        customerName, customerDetails, status: 'pending', documents,
        documentTypes: Object.keys(currentCustomerUploads), reuploadedAt: serverTimestamp(),
        fixSubmitted: true, fixLocked: true, fixSubmittedAt: serverTimestamp(),
        assignedTo: reviewerToReassign || existingSub.assignedTo || '',
        reviewedAt: null,
        comment: ''
      });
      notifyStatusChangePush({
        currentUser,
        submissionId: currentEditId,
        customerName,
        newStatus: 'pending',
        statusLabel: 'Pending Review',
        actionLabel: 'Application Re-Submitted',
        message: `Application for ${customerName} was re-submitted and is back in pending review.`
      }).catch(() => {});
      showNotification(reviewerToReassign
        ? '✅ Fix submitted and reassigned to the same reviewer for another review.'
        : '✅ Fix submitted successfully!', 'success');
      closeModal(uploadModal);
      currentEditId = null;
      updateSubmitButton();
      return;
    }
    const requiredDocs = DOCUMENT_TYPES.filter(d => d.required !== false).map(d => d.id);
    const missing = requiredDocs.filter(id => !(currentCustomerUploads[id] && currentCustomerUploads[id].length > 0));
    if (missing.length > 0) {
      submitCustomerBtn.disabled = false;
      submitCustomerBtn.innerHTML = '<i class="fas fa-check-circle"></i> Submit Customer Documents';
      return showNotification('Please upload required documents: ' + missing.join(', '), 'error');
    }
    const documents = [];
    Object.entries(currentCustomerUploads).forEach(([type, files]) => {
      files.forEach(file => { documents.push({ documentType: type, ...file }); });
    });
    const customerDetails = {
      name: customerName,
      dob: document.getElementById('customerDob')?.value || '',
      email: document.getElementById('customerEmail')?.value?.trim() || '',
      phone: document.getElementById('customerPhone')?.value?.trim() || '',
      nin: document.getElementById('customerNIN')?.value?.trim() || '',
      address: document.getElementById('customerAddress')?.value?.trim() || '',
      accountNo: document.getElementById('accountNo')?.value?.trim() || '',
      employer: document.getElementById('employer')?.value?.trim() || '',
      originatingTP: document.getElementById('originatingTP')?.value?.trim() || '',
      pfa: document.getElementById('pfa')?.value?.trim() || '',
      penNo: document.getElementById('penNo')?.value?.trim() || '',
      rsaStatementDate: document.getElementById('rsaStatementDate')?.value || '',
      rsaBalance: document.getElementById('rsaBalance')?.value || '',
      rsa25: document.getElementById('rsa25Percent')?.value || '',
      propertyType: document.getElementById('propertyType')?.value?.trim() || '',
      tenor: document.getElementById('tenor')?.value || '',
      propertyValue: document.getElementById('propertyValue')?.value || '',
      facilityFee: document.getElementById('facilityFee')?.value || '',
      loanAmount: document.getElementById('loanAmount')?.value || ''
    };
    const selectedAgentId = String(customerAgentSelect?.value || '').trim();
    const selectedAgent = approvedAgents.find((a) => a.id === selectedAgentId) || null;
    const agentSummary = selectedAgent
      ? `Selected agent:\n- Name: ${selectedAgent.fullName || '-'}\n- Phone: ${selectedAgent.contactNumber || '-'}`
      : 'Selected agent: No Agent';
    const proceed = confirm(
      `Please confirm submission details:\n\nCustomer: ${customerName || '-'}\n${agentSummary}\n\nProceed with submission?`
    );
    if (!proceed) {
      submitCustomerBtn.disabled = false;
      submitCustomerBtn.innerHTML = '<i class="fas fa-check-circle"></i> Submit Customer Documents';
      return;
    }
    const uploaderEmail = normalizeEmail(currentUser?.email);
    const subRef = await addDoc(collection(db, 'submissions'), {
      customerName, customerDetails, uploadedBy: uploaderEmail, uploadedAt: serverTimestamp(),
      status: 'pending', comment: '', documents, documentTypes: Object.keys(currentCustomerUploads),
      agentId: selectedAgent?.id || '',
      agentName: selectedAgent?.fullName || '',
      agentContactNumber: selectedAgent?.contactNumber || '',
      agentAccountNumber: selectedAgent?.accountNumber || '',
      agentAccountBank: selectedAgent?.accountBank || ''
    });
    const assignedEmail = await assignRoundRobin(subRef).catch(err => { return null; });
    let assignmentEmailFailed = false;
    if (assignedEmail) {
      const emailResult = await queueViewerAssignmentEmail({
        submissionId: subRef.id,
        viewerEmail: assignedEmail,
        customerName,
        uploaderEmail: currentUser?.email || ''
      }).catch((emailErr) => {
        console.warn('viewer assignment email failed:', emailErr);
        return { queued: true, sent: false, reason: 'send-failed' };
      });

      if (emailResult?.sent === false) {
        assignmentEmailFailed = true;
      }
    }
    if (assignmentEmailFailed) {
      showNotification(`✅ Submitted and assigned to ${assignedEmail}, but email alert failed.`, 'warning');
    } else {
      showNotification(assignedEmail ? `✅ Submitted – assigned to ${assignedEmail}` : '✅ Customer documents submitted successfully!', 'success');
    }
    notifyStatusChangePush({
      currentUser,
      submissionId: subRef.id,
      customerName,
      newStatus: 'pending',
      statusLabel: 'Pending Review',
      actionLabel: 'New Submission',
      message: `A new application for ${customerName} has been submitted for review.`
    }).catch(() => {});
    closeModal(uploadModal);
  } catch (error) {
    showNotification('Submission failed: ' + error.message, 'error');
  } finally {
    submitCustomerBtn.innerHTML = '<i class="fas fa-check-circle"></i> Submit Customer Documents';
  }
}

// ==================== EDIT FUNCTIONS ====================
window.openEditModal = async (id) => {
  const sub = allSubmissions.find(s => s.id === id);
  if (!sub) return;
  if (sub.fixSubmitted || sub.fixLocked || sub.fixSubmittedAt) { showNotification('This rejected application has already been fixed and submitted.', 'warning'); return; }
  currentEditId = id;
  const details = sub.customerDetails || {};
  try {
    const pick = (keys, fallback = '') => {
      for (const key of keys) {
        const dVal = details?.[key];
        if (dVal !== undefined && dVal !== null && String(dVal).trim() !== '') return String(dVal);
        const sVal = sub?.[key];
        if (sVal !== undefined && sVal !== null && String(sVal).trim() !== '') return String(sVal);
      }
      return fallback;
    };
    const fieldMap = [
      { id: 'customerName', keys: ['name', 'customerName'], fallback: sub.customerName || '', label: 'Customer Name' },
      { id: 'customerDob', keys: ['dob', 'dateOfBirth', 'customerDob'], fallback: '', label: 'Date of Birth' },
      { id: 'customerEmail', keys: ['email', 'customerEmail'], fallback: '', label: 'Email' },
      { id: 'customerPhone', keys: ['phone', 'customerPhone'], fallback: '', label: 'Phone' },
      { id: 'customerNIN', keys: ['nin', 'customerNIN'], fallback: '', label: 'NIN' },
      { id: 'customerAddress', keys: ['address', 'customerAddress'], fallback: '', label: 'Address' },
      { id: 'accountNo', keys: ['accountNo'], fallback: '', label: 'Account Number' },
      { id: 'employer', keys: ['employer'], fallback: '', label: 'Employer' },
      { id: 'originatingTP', keys: ['originatingTP'], fallback: '', label: 'Originating Transfer Pin' },
      { id: 'pfa', keys: ['pfa', 'pfaName'], fallback: '', label: 'PFA' },
      { id: 'penNo', keys: ['penNo'], fallback: '', label: 'PEN Number' },
      { id: 'rsaStatementDate', keys: ['rsaStatementDate'], fallback: '', label: 'RSA Statement Date' },
      { id: 'rsaBalance', keys: ['rsaBalance'], fallback: '', label: 'RSA Balance' },
      { id: 'propertyType', keys: ['propertyType'], fallback: '', label: 'Property Type' },
      { id: 'propertyValue', keys: ['propertyValue'], fallback: '', label: 'Property Value' },
      { id: 'loanAmount', keys: ['loanAmount'], fallback: '', label: 'Loan Amount' },
    ];
    const missingFields = [];
    fieldMap.forEach((field) => {
      const el = document.getElementById(field.id);
      if (!el) return;
      const value = pick(field.keys, field.fallback);
      el.value = value;
      if (!String(value || '').trim()) missingFields.push(field.label);
    });
    if (missingFields.length > 0) showNotification('Some saved fields are missing: ' + missingFields.join(', '), 'warning');
    customerDetailsSaved = true;
  } catch (e) { }
  const existingUploads = {};
  (sub.documents || []).forEach(doc => {
    if (!existingUploads[doc.documentType]) existingUploads[doc.documentType] = [];
    existingUploads[doc.documentType].push(doc);
  });
  currentCustomerUploads = existingUploads;
  if (batchUploadBtn) batchUploadBtn.disabled = false;
  if (window.renderDocumentGridUpload) {
    window.renderDocumentGridUpload(documentGrid, REQUIRED_DOC_TYPES, currentCustomerUploads, 'edit');
    window.renderDocumentGridUpload(optionalDocumentGrid, OPTIONAL_DOC_TYPES, currentCustomerUploads, 'edit');
  }
  updateSubmitButton();
  if (customerNameInput) customerNameInput.focus();
  uploadModal.classList.add('active');
};

async function submitEdit() {
  if (!currentEditId) return;
  submitEditBtn.disabled = true;
  submitEditBtn.innerHTML = '<i class="fas fa-spinner fa-spin"></i> Updating...';
  try {
    const newDocuments = [];
    for (const [type, files] of Object.entries(currentCustomerUploads)) {
      for (const file of files) {
        if (file.localAddedAt && file.localAddedAt > Date.now() - 60000) {
          newDocuments.push({ documentType: type, name: file.name, fileId: file.fileId, fileUrl: file.fileUrl, uploadedAt: new Date().toISOString() });
        }
      }
    }
    if (newDocuments.length > 0) {
      const submissionRef = doc(db, 'submissions', currentEditId);
      const existingSub = allSubmissions.find((s) => s.id === currentEditId) || {};
      const reviewerToReassign = String(existingSub.reviewedBy || existingSub.assignedTo || '').trim();
      await updateDoc(submissionRef, {
        status: 'pending',
        documents: arrayUnion(...newDocuments),
        reuploadedAt: serverTimestamp(),
        fixSubmitted: true,
        fixLocked: true,
        fixSubmittedAt: serverTimestamp(),
        assignedTo: reviewerToReassign || existingSub.assignedTo || '',
        reviewedAt: null,
        comment: ''
      });
      notifyStatusChangePush({
        currentUser,
        submissionId: currentEditId,
        customerName: existingSub.customerName || '',
        newStatus: 'pending',
        statusLabel: 'Pending Review',
        actionLabel: 'Application Re-Submitted',
        message: `Application for ${existingSub.customerName || 'this customer'} was re-submitted and is back in pending review.`
      }).catch(() => {});
    }
    showNotification('✅ Documents re-uploaded and sent back for reviewer action.', 'success');
    closeModal(editModal);
  } catch (error) {
    showNotification('Update failed: ' + error.message, 'error');
  } finally {
    submitEditBtn.innerHTML = '<i class="fas fa-upload"></i> Re-upload Documents';
  }
}

// ==================== LOAD SUBMISSIONS ====================
async function loadSubmissions() {
  const uploaderEmail = normalizeEmail(currentUser?.email);
  if (!uploaderEmail) return;
  const q = query(collection(db, 'submissions'), where('uploadedBy', '==', uploaderEmail), orderBy('uploadedAt', 'desc'));
  onSnapshot(q, async (snapshot) => {
    allSubmissions = [];
    const emails = new Set();
    snapshot.forEach((doc) => {
      const data = doc.data();
      allSubmissions.push({ id: doc.id, ...data });
      if (data.uploadedBy) emails.add(data.uploadedBy);
      if (data.reviewedBy) emails.add(data.reviewedBy);
      if (data.assignedTo) emails.add(data.assignedTo);
    });
    try { await ensureUserFullNames(Array.from(emails)); } catch (e) { }
    renderPendingTable(); renderApprovedTable(); renderRejectedTable(); renderPaidTable(); updateDashboardCards();
  }, (error) => { showNotification('Error loading submissions', 'error'); });
}

async function ensureUserFullNames(emails) {
  if (!emails || emails.length === 0) return;
  for (const email of emails) {
    if (!email || userFullNames.has(email)) continue;
    try {
      const q = query(collection(db, 'users'), where('email', '==', email));
      const snap = await getDocs(q);
      if (!snap.empty) { const d = snap.docs[0].data(); userFullNames.set(email, d.fullName || d.displayName || email.split('@')[0]); }
      else { userFullNames.set(email, email.split('@')[0]); }
    } catch (e) { userFullNames.set(email, email.split('@')[0]); }
  }
}

function safeFormatDate(dateValue) {
  if (!dateValue) return 'N/A';
  try {
    let date;
    if (dateValue.toDate && typeof dateValue.toDate === 'function') date = dateValue.toDate();
    else if (typeof dateValue === 'string') date = new Date(dateValue);
    else if (dateValue.seconds) date = new Date(dateValue.seconds * 1000);
    else if (dateValue instanceof Date) date = dateValue;
    else return 'N/A';
    return date.toLocaleDateString('en-NG', { day: 'numeric', month: 'short', year: 'numeric', hour: '2-digit', minute: '2-digit' });
  } catch (error) { return 'Invalid date'; }
}

function normalizeWhatsAppPhone(raw) {
  const digits = String(raw || '').replace(/\D/g, '');
  if (!digits) return '';
  if (digits.startsWith('0') && digits.length === 11) return `234${digits.slice(1)}`;
  if (digits.length === 10) return `234${digits}`;
  if (digits.startsWith('234')) return digits;
  return digits;
}

function renderWhatsAppLink(raw) {
  const display = String(raw || '').trim();
  const normalized = normalizeWhatsAppPhone(display);
  if (!normalized) return '-';
  return `<a href="https://wa.me/${normalized}" target="_blank" rel="noopener noreferrer">${display}</a>`;
}

function updateDashboardCards() {
  const pending = allSubmissions.filter(s => s.status === 'pending').length;
  const approved = allSubmissions.filter(s => {
    const st = String(s.status || '').toLowerCase();
    return st === 'processing_to_pfa' || st === 'approved';
  }).length;
  const rejected = allSubmissions.filter(s => s.status === 'rejected').length;
  document.getElementById('cardPendingCount').textContent = pending;
  document.getElementById('cardApprovedCount').textContent = approved;
  document.getElementById('cardRejectedCount').textContent = rejected;
}

async function renderPendingTable() {
  if (!pendingTableBody) { return; }
  const pending = allSubmissions.filter(s => s.status === 'pending');
  if (pending.length === 0) { pendingTableBody.innerHTML = '<tr><td colspan="8" class="no-data">No pending documents</td></tr>'; return; }
  let html = '';
  for (const sub of pending) {
    const date = safeFormatDate(sub.uploadedAt);
    const assignedName = sub.assignedTo ? await getUserFullName(sub.assignedTo) : 'Not assigned';
    const whatsapp = renderWhatsAppLink(sub.customerDetails?.phone || sub.customerPhone || '');
    html += `<tr><td><strong>${sub.customerName}</strong></td><td>${whatsapp}</td><td>${assignedName}</td><td>${date}</td><td><span class="status-badge status-pending">Pending</span></td><td>${sub.comment || '-'}</td><td><button class="action-btn view-btn-small" onclick="window.viewSubmissionDocs('${sub.id}')"><i class="fas fa-eye"></i> View</button> <button class="action-btn" onclick="window.openApplicationChat('${sub.id}')" title="Application Chat"><i class="fas fa-comments"></i> Chat</button></td><td><button class="action-btn track-btn" onclick="window.showApplicationTrack('${sub.id}')"><i class="fas fa-map-marker-alt"></i> Track</button></td></tr>`;
  }
  pendingTableBody.innerHTML = html;
}

async function renderApprovedTable() {
  if (!approvedTableBody) { return; }
  const approved = allSubmissions.filter(s => {
    const st = String(s.status || '').toLowerCase();
    return st === 'processing_to_pfa' || st === 'approved';
  });
  if (approved.length === 0) { approvedTableBody.innerHTML = '<tr><td colspan="9" class="no-data">No approved documents</td></tr>'; return; }
  let html = '';
  for (const sub of approved) {
    const uploadDate = safeFormatDate(sub.uploadedAt);
    const approvedDate = safeFormatDate(sub.reviewedAt);
    const approvedBy = (sub.reviewedBy && userFullNames.get(sub.reviewedBy)) ? userFullNames.get(sub.reviewedBy) : (sub.reviewedBy || '-');
    const assignedName = sub.assignedTo ? await getUserFullName(sub.assignedTo) : 'Not assigned';
    const whatsapp = renderWhatsAppLink(sub.customerDetails?.phone || sub.customerPhone || '');
    html += `<tr><td><strong>${sub.customerName}</strong></td><td>${whatsapp}</td><td>${assignedName}</td><td>${uploadDate}</td><td><span class="status-badge status-approved">Processing to PFA</span></td><td>${approvedBy}</td><td>${approvedDate}</td><td><button class="action-btn view-btn-small" onclick="window.viewSubmissionDocs('${sub.id}')"><i class="fas fa-eye"></i> View</button> <button class="action-btn" onclick="window.openApplicationChat('${sub.id}')" title="Application Chat"><i class="fas fa-comments"></i> Chat</button></td><td><button class="action-btn track-btn" onclick="window.showApplicationTrack('${sub.id}')"><i class="fas fa-map-marker-alt"></i> Track</button></td></tr>`;
  }
  approvedTableBody.innerHTML = html;
}

async function renderRejectedTable() {
  if (!rejectedTableBody) { return; }
  const rejected = allSubmissions.filter(s => s.status === 'rejected');
  if (rejected.length === 0) { rejectedTableBody.innerHTML = '<tr><td colspan="11" class="no-data">No rejected documents</td></tr>'; return; }
  let html = '';
  for (const sub of rejected) {
    const isFixLocked = !!(sub.fixSubmitted || sub.fixLocked || sub.fixSubmittedAt);
    const date = safeFormatDate(sub.uploadedAt);
    const rejectedDate = safeFormatDate(sub.reviewedAt);
    const rejectedBy = (sub.reviewedBy && userFullNames.get(sub.reviewedBy)) ? userFullNames.get(sub.reviewedBy) : (sub.reviewedBy || '-');
    const assignedName = sub.assignedTo ? await getUserFullName(sub.assignedTo) : 'Not assigned';
    const chatBtn = `<button class="action-btn app-chat-trigger" data-chat-submission="${sub.id}" onclick="window.openApplicationChat('${sub.id}')" title="Application Chat"><i class="fas fa-comments"></i> Chat</button>`;
    html += `<tr data-submission-id="${sub.id}"><td><strong>${sub.customerName}</strong></td><td>${chatBtn}</td><td>${assignedName}</td><td>${date}</td><td><span class="status-badge status-rejected">Rejected</span></td><td>${sub.comment || 'No reason provided'}</td><td>${rejectedBy}</td><td>${rejectedDate}</td><td><button class="action-btn edit-btn" onclick="window.openEditModal('${sub.id}')" ${isFixLocked ? 'disabled style="opacity:.6;cursor:not-allowed;" title="Already fixed and submitted"' : ''}><i class="fas fa-edit"></i> ${isFixLocked ? 'Fixed' : 'Re-upload'}</button></td><td><button class="action-btn view-btn-small" onclick="window.viewSubmissionDocs('${sub.id}')"><i class="fas fa-eye"></i> View</button></td><td><button class="action-btn track-btn" onclick="window.showApplicationTrack('${sub.id}')"><i class="fas fa-map-marker-alt"></i> Track</button></td></tr>`;
  }
  rejectedTableBody.innerHTML = html;
}

function renderPaidTable() {
  if (!activeCommissionTableBody || !clearedCommissionTableBody) return;

  const activeAll = allSubmissions.filter(s => String(s.status || '').toLowerCase() === 'paid');
  const clearedAll = allSubmissions.filter(s => String(s.status || '').toLowerCase() === 'cleared');
  const paidTabTotal = activeAll.length + clearedAll.length;
  if (paidCountBadge) {
    paidCountBadge.textContent = String(paidTabTotal);
    paidCountBadge.style.display = paidTabTotal > 0 ? 'inline-flex' : 'none';
  }
  if (activeCommissionCountEl) activeCommissionCountEl.textContent = String(activeAll.length);
  if (clearedCommissionCountEl) clearedCommissionCountEl.textContent = String(clearedAll.length);

  // Totals are based on ACTIVE commission only.
  let total25 = 0;
  let total2 = 0;
  activeAll.forEach((sub) => {
    const { twentyFive, commission2 } = getFinancials(sub);
    total25 += twentyFive;
    total2 += commission2;
  });
  if (paidTotal25El) paidTotal25El.textContent = formatCurrency(total25);
  if (paidTotal1El) paidTotal1El.textContent = formatCurrency(total2);

  // Separate totals for CLEARED commission.
  let clearedTotal25 = 0;
  let clearedTotal2 = 0;
  clearedAll.forEach((sub) => {
    const { twentyFive, commission2 } = getFinancials(sub);
    clearedTotal25 += twentyFive;
    clearedTotal2 += commission2;
  });
  if (clearedTotal25El) clearedTotal25El.textContent = formatCurrency(clearedTotal25);
  if (clearedTotal1El) clearedTotal1El.textContent = formatCurrency(clearedTotal2);

  const activeSearch = String(activeCommissionSearch?.value || '').trim().toLowerCase();
  const activeStart = String(activeCommissionStartDate?.value || '').trim();
  const activeEnd = String(activeCommissionEndDate?.value || '').trim();
  const activeFiltered = activeAll.filter((sub) => {
    const text = `${sub.customerName || ''} ${sub.pfa || ''} ${sub.customerDetails?.pfa || ''}`.toLowerCase();
    const paidDate = getDateFromAny(sub.paidAt || sub.updatedAt);
    const textOk = !activeSearch || text.includes(activeSearch);
    return textOk && inDateRange(paidDate, activeStart, activeEnd);
  });

  if (!activeFiltered.length) {
    activeCommissionTableBody.innerHTML = '<tr><td colspan="7" class="no-data">No active commission records</td></tr>';
  } else {
    activeCommissionTableBody.innerHTML = activeFiltered.map(sub => {
      const { pfa, twentyFive, commission2 } = getFinancials(sub);
      const paidDate = safeFormatDate(sub.paidAt || sub.updatedAt);
      const agentName = sub.agentName || '-';

      return `
        <tr>
          <td><strong>${sub.customerName || '-'}</strong></td>
          <td>${pfa}</td>
          <td>${agentName}</td>
          <td>${formatCurrency(twentyFive)}</td>
          <td>${formatCurrency(commission2)}</td>
          <td><span class="status-badge status-approved">Paid</span></td>
          <td>${paidDate}</td>
        </tr>
      `;
    }).join('');
  }

  const clearedSearch = String(clearedCommissionSearch?.value || '').trim().toLowerCase();
  const clearedStart = String(clearedCommissionStartDate?.value || '').trim();
  const clearedEnd = String(clearedCommissionEndDate?.value || '').trim();
  const clearedFiltered = clearedAll.filter((sub) => {
    const text = `${sub.customerName || ''} ${sub.pfa || ''} ${sub.customerDetails?.pfa || ''}`.toLowerCase();
    const clearedDate = getDateFromAny(sub.clearedAt || sub.updatedAt);
    const textOk = !clearedSearch || text.includes(clearedSearch);
    return textOk && inDateRange(clearedDate, clearedStart, clearedEnd);
  });

  if (!clearedFiltered.length) {
    clearedCommissionTableBody.innerHTML = '<tr><td colspan="7" class="no-data">No cleared commission records</td></tr>';
  } else {
    clearedCommissionTableBody.innerHTML = clearedFiltered.map(sub => {
      const { pfa, twentyFive, commission2 } = getFinancials(sub);
      const clearedDate = safeFormatDate(sub.clearedAt || sub.updatedAt);
      const agentName = sub.agentName || '-';

      return `
        <tr>
          <td><strong>${sub.customerName || '-'}</strong></td>
          <td>${pfa}</td>
          <td>${agentName}</td>
          <td>${formatCurrency(twentyFive)}</td>
          <td>${formatCurrency(commission2)}</td>
          <td><span class="status-badge status-pending">Cleared</span></td>
          <td>${clearedDate}</td>
        </tr>
      `;
    }).join('');
  }
}

function showNotification(message, type = 'info') {
  if (!notification) return;
  notification.textContent = message;
  notification.className = `notification ${type}`;
  notification.style.display = 'block';
  setTimeout(() => { notification.style.display = 'none'; }, 3000);
}

async function loadApprovedAgents() {
  try {
    const currentUid = String(currentUser?.uid || '').trim();
    const currentEmail = String(currentUser?.email || '').trim().toLowerCase();
    let rows = [];

    if (currentUid) {
      const ownedSnap = await getDocs(query(
        collection(db, 'agents'),
        where('status', '==', 'approved'),
        where('createdByUid', '==', currentUid)
      ));
      rows = ownedSnap.docs.map((d) => ({ id: d.id, ...(d.data() || {}) }));
    } else {
      const snap = await getDocs(query(collection(db, 'agents'), where('status', '==', 'approved')));
      rows = snap.docs.map((d) => ({ id: d.id, ...(d.data() || {}) }));
    }

    approvedAgents = rows
      .filter((agent) => {
        if (currentUid && String(agent.createdByUid || '') === currentUid) return true;
        const createdByEmail = String(agent.createdBy || '').trim().toLowerCase();
        return !!currentEmail && createdByEmail === currentEmail;
      })
      .sort((a, b) => String(a.fullName || '').localeCompare(String(b.fullName || '')));
    populateApprovedAgentSelect();
  } catch (_) {
    approvedAgents = [];
    populateApprovedAgentSelect();
  }
}

function openAgentRegistrationModal() {
  if (!agentRegistrationModal) return;
  agentRegistrationModal.classList.add('active');
}

function closeAgentRegistrationModal() {
  if (!agentRegistrationModal) return;
  agentRegistrationModal.classList.remove('active');
}

function normalizeAgentDate(value) {
  if (!value) return '-';
  if (value.toDate && typeof value.toDate === 'function') return safeFormatDate(value);
  if (value.seconds) return safeFormatDate(value);
  const asDate = new Date(value);
  if (Number.isNaN(asDate.getTime())) return '-';
  return safeFormatDate(asDate);
}

function statusBadgeForAgent(status) {
  const key = String(status || 'pending').toLowerCase();
  if (key === 'approved') return '<span class="status-badge status-approved">Approved</span>';
  if (key === 'rejected') return '<span class="status-badge status-rejected">Rejected</span>';
  return '<span class="status-badge status-pending">Pending</span>';
}

function escAttr(value) {
  return String(value ?? '')
    .replace(/&/g, '&amp;')
    .replace(/"/g, '&quot;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');
}

window.viewAgentDetails = (agentId) => {
  const agent = registeredAgents.find((row) => row.id === agentId);
  if (!agent) {
    showNotification('Agent record not found', 'error');
    return;
  }
  const detailsModal = document.createElement('div');
  detailsModal.className = 'modal active';
  detailsModal.innerHTML = `
    <div class="modal-content">
      <div class="modal-header">
        <h2><i class="fas fa-user-tie"></i> Agent Details</h2>
        <button class="close-btn" type="button">&times;</button>
      </div>
      <div class="modal-body">
        <div class="customer-input-grid">
          <div><label>Full Name</label><input type="text" readonly value="${escAttr(agent.fullName || '-')}"></div>
          <div><label>Contact Number</label><input type="text" readonly value="${escAttr(agent.contactNumber || '-')}"></div>
          <div><label>Account Number</label><input type="text" readonly value="${escAttr(agent.accountNumber || '-')}"></div>
          <div><label>Account Bank</label><input type="text" readonly value="${escAttr(agent.accountBank || '-')}"></div>
          <div><label>Status</label><input type="text" readonly value="${escAttr(String(agent.status || 'pending'))}"></div>
          <div><label>Date Registered</label><input type="text" readonly value="${escAttr(normalizeAgentDate(agent.createdAt))}"></div>
        </div>
      </div>
      <div class="modal-footer">
        <button class="cancel-btn" type="button">Close</button>
      </div>
    </div>
  `;
  document.body.appendChild(detailsModal);
  const close = () => detailsModal.remove();
  detailsModal.querySelectorAll('button').forEach((btn) => btn.addEventListener('click', close));
  detailsModal.addEventListener('click', (e) => {
    if (e.target === detailsModal) close();
  });
};

async function loadRegisteredAgents() {
  if (!registeredAgentsTableBody) return;
  registeredAgentsTableBody.innerHTML = '<tr><td colspan="7" class="loading-row"><div class="loading-spinner"></div> Loading agents...</td></tr>';
  try {
    const currentUid = String(currentUser?.uid || '').trim();
    const currentEmail = String(currentUser?.email || '').trim().toLowerCase();
    const byId = new Map();

    if (currentUid) {
      const ownSnap = await getDocs(query(collection(db, 'agents'), where('createdByUid', '==', currentUid)));
      ownSnap.docs.forEach((row) => {
        byId.set(row.id, { id: row.id, ...(row.data() || {}) });
      });
    }
    if (currentEmail) {
      const emailSnap = await getDocs(query(collection(db, 'agents'), where('createdBy', '==', currentEmail)));
      emailSnap.docs.forEach((row) => {
        if (!byId.has(row.id)) byId.set(row.id, { id: row.id, ...(row.data() || {}) });
      });
    }

    registeredAgents = Array.from(byId.values()).sort((a, b) => {
      const aTime = getDateFromAny(a.createdAt)?.getTime() || 0;
      const bTime = getDateFromAny(b.createdAt)?.getTime() || 0;
      return bTime - aTime;
    });

    if (!registeredAgents.length) {
      registeredAgentsTableBody.innerHTML = '<tr><td colspan="7" class="no-data">No registered agents yet</td></tr>';
      return;
    }

    registeredAgentsTableBody.innerHTML = registeredAgents.map((agent) => `
      <tr>
        <td><strong>${agent.fullName || '-'}</strong></td>
        <td>${agent.contactNumber || '-'}</td>
        <td>${agent.accountNumber || '-'}</td>
        <td>${agent.accountBank || '-'}</td>
        <td>${statusBadgeForAgent(agent.status)}</td>
        <td>${normalizeAgentDate(agent.createdAt)}</td>
        <td><button class="action-btn view-btn-small" onclick="window.viewAgentDetails('${agent.id}')"><i class="fas fa-eye"></i> View</button></td>
      </tr>
    `).join('');
  } catch (error) {
    registeredAgents = [];
    registeredAgentsTableBody.innerHTML = '<tr><td colspan="7" class="no-data">Unable to load agents</td></tr>';
  }
}

function populateApprovedAgentSelect() {
  if (!customerAgentSelect) return;
  const current = String(customerAgentSelect.value || '');
  customerAgentSelect.innerHTML = '<option value="">No Agent</option>' + approvedAgents.map((agent) => (
    `<option value="${agent.id}">${agent.fullName || 'Unnamed'} - ${agent.contactNumber || '-'}</option>`
  )).join('');
  if (current && approvedAgents.some((a) => a.id === current)) {
    customerAgentSelect.value = current;
  }
}

async function handleAgentRegistration(e) {
  e.preventDefault();
  const fullName = String(document.getElementById('agentFullName')?.value || '').trim();
  const contactNumber = String(document.getElementById('agentContactNumber')?.value || '').trim();
  const accountNumber = String(document.getElementById('agentAccountNumber')?.value || '').trim();
  const accountBank = String(document.getElementById('agentAccountBank')?.value || '').trim();

  if (!fullName || !contactNumber || !accountNumber || !accountBank) {
    showNotification('Please complete all agent fields', 'error');
    return;
  }
  if (!/^\d{10,11}$/.test(contactNumber.replace(/\D/g, ''))) {
    showNotification('Agent contact number must be 10 or 11 digits', 'error');
    return;
  }
  if (!/^\d{10}$/.test(accountNumber.replace(/\D/g, ''))) {
    showNotification('Agent account number must be exactly 10 digits', 'error');
    return;
  }

  try {
    if (submitAgentFormBtn) {
      submitAgentFormBtn.disabled = true;
      submitAgentFormBtn.innerHTML = '<i class="fas fa-spinner fa-spin"></i> Submitting...';
    }
    await addDoc(collection(db, 'agents'), {
      fullName,
      contactNumber,
      accountNumber,
      accountBank,
      status: 'pending',
      createdBy: currentUser?.email || '',
      createdByUid: currentUser?.uid || '',
      createdAt: serverTimestamp()
    });
    showNotification('Agent registration submitted for admin approval', 'success');
    agentRegistrationForm?.reset();
    closeAgentRegistrationModal();
    await loadRegisteredAgents();
  } catch (err) {
    showNotification('Failed to submit agent registration', 'error');
  } finally {
    if (submitAgentFormBtn) {
      submitAgentFormBtn.disabled = false;
      submitAgentFormBtn.innerHTML = '<i class="fas fa-paper-plane"></i> Submit for Approval';
    }
  }
}

function getFileIcon(filename) {
  const ext = filename?.split('.').pop().toLowerCase();
  const icons = { 'pdf': 'fa-file-pdf', 'doc': 'fa-file-word', 'docx': 'fa-file-word', 'jpg': 'fa-file-image', 'jpeg': 'fa-file-image', 'png': 'fa-file-image' };
  return icons[ext] || 'fa-file-alt';
}

function formatFileSize(bytes) {
  if (!bytes) return '0 Bytes';
  const k = 1024;
  const sizes = ['Bytes', 'KB', 'MB'];
  const i = Math.floor(Math.log(bytes) / Math.log(k));
  return parseFloat((bytes / Math.pow(k, i)).toFixed(1)) + ' ' + sizes[i];
}

function hideLoader() {
  const loader = document.getElementById('globalLoader');
  if (loader) loader.style.display = 'none';
}
window.hideLoader = hideLoader;
window.showLoader = function(message = 'Processing...') {
  const loader = document.getElementById('globalLoader');
  const loaderText = document.getElementById('loaderText');
  if (loader) loader.style.display = 'flex';
  if (loaderText) loaderText.textContent = message;
};

// ==================== MAKE FUNCTIONS GLOBAL ====================
if (typeof window.openSingleUpload !== 'function') window.openSingleUpload = openSingleUpload;
if (typeof window.viewDocument !== 'function') window.viewDocument = viewDocument;
if (typeof window.viewSubmissionDocs !== 'function') window.viewSubmissionDocs = viewSubmissionDocs;
if (typeof window.openEditModal !== 'function') window.openEditModal = openEditModal;
if (typeof window.switchTab !== 'function') window.switchTab = switchTab;
if (typeof window.showApplicationTrack !== 'function') window.showApplicationTrack = showApplicationTrack;
if (typeof window.saveCustomerDetails !== 'function') window.saveCustomerDetails = saveCustomerDetails;
if (typeof window.resetCustomerDetails !== 'function') window.resetCustomerDetails = resetCustomerDetails;

window.removeUploadedDoc = (docType) => {
  if (!currentCustomerUploads[docType] || currentCustomerUploads[docType].length === 0) return;
  currentCustomerUploads[docType].pop();
  renderDocumentGrid(documentGrid, REQUIRED_DOC_TYPES, currentCustomerUploads, 'upload');
  renderDocumentGrid(optionalDocumentGrid, OPTIONAL_DOC_TYPES, currentCustomerUploads, 'upload');
  updateSubmitButton();
};

// ==================== INITIALIZE RECENT SEARCH ====================
document.addEventListener('DOMContentLoaded', function() {
  const recentSearch = document.getElementById('recentSearch');
  if (recentSearch) {
    recentSearch.addEventListener('input', function(e) {
      const searchTerm = e.target.value.toLowerCase();
      const rows = document.querySelectorAll('#recentTableBody tr');
      rows.forEach(row => { const text = row.textContent.toLowerCase(); row.style.display = text.includes(searchTerm) ? '' : 'none'; });
    });
  }
  ['pending', 'approved', 'rejected'].forEach(tab => {
    const searchInput = document.getElementById(`${tab}Search`);
    if (searchInput) {
      searchInput.addEventListener('input', function(e) {
        const searchTerm = e.target.value.toLowerCase();
        const rows = document.querySelectorAll(`#${tab}TableBody tr`);
        rows.forEach(row => { const text = row.textContent.toLowerCase(); row.style.display = text.includes(searchTerm) ? '' : 'none'; });
      });
    }
  });
});
