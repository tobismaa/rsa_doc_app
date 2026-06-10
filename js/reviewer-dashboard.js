// js/reviewer-dashboard.js - FIXED VERSION WITH DIRECT SAVE ONLY
import { auth, db } from './firebase-config.js';
import { queueUploaderApprovedEmail, queueUploaderRejectedEmail, queueRsaApprovalEmail } from './email-alerts.js';
import { notifyStatusChangePush } from './status-push.js';
import { getSystemSettings } from './shared/system-settings.js?v=20260508a';
import { formatAppDate, formatAppDateTime, getTrustedDateKey, getTrustedNowIso } from './shared/app-time.js';
import {
    getCurrentUserProfile as getCurrentUserProfileShared,
    getUserProfileByEmail as getUserProfileByEmailShared,
    getUserFullName as getUserFullNameShared,
    normalizeEmail as normalizeEmailShared
} from './shared/user-directory.js?v=20260518a';
import { getUploaderRoutingRule as getUploaderRoutingRuleShared, routingRuleDocId as routingRuleDocIdShared } from './shared/uploader-routing.js?v=20260427e';
import {
    getTimestampMillis as getStageTimestampMillis,
    getSubmissionReviewEntryAt,
    getSubmissionApprovalEntryAt,
    getSubmissionRejectionEntryAt
} from './shared/submission-stage.js?v=20260609a';
import {
    buildDashboardStageReport,
    renderDashboardStageReport,
    exportDashboardStageReportExcel
} from './shared/dashboard-stage-report.js?v=20260610a';
import { signOut } from "https://www.gstatic.com/firebasejs/9.22.0/firebase-auth.js";
import {
    collection,
    query,
    where,
    orderBy,
    onSnapshot,
    updateDoc,
    doc,
    getDocs,
    getDoc,
    serverTimestamp,
    runTransaction,
    addDoc
} from "https://www.gstatic.com/firebasejs/9.22.0/firebase-firestore.js";

// ==================== DOCUMENT TYPES MAPPING ====================
const DOCUMENT_TYPES = {
    'birth_certificate': 'Birth Certificate',
    'nin': 'National ID (NIN)',
    'pay_slips': 'Pay Slips',
    'offer_letter': 'Offer Letter',
    'intro_letter': 'Introduction Letter',
    'request_letter': 'Request Letter',
    'rsa_statement': 'RSA Statement',
    'pfa_form': 'PFA Form',
    'consent_letter': 'Consent Letter',
    'indemnity_form': 'Indemnity Form',
    'utility_bill': 'Utility Bill',
    'benefit_application_form': 'Benefit Application Form',
    'credit_life': 'Credit Life'
};

// ==================== GLOBAL VARIABLES ====================
let currentUser = null;
let currentUserData = null;
let allSubmissions = [];
let currentSubmissionId = null;
let currentViewerSubmission = null;
let currentViewerIndex = 0;
let downloadInProgress = false;
let submissionsLoadVersion = 0;
let currentReviewerStageReport = null;

function getTimestampMsSafe(value) {
    if (!value) return 0;
    try {
        if (typeof value?.toMillis === 'function') return value.toMillis();
        if (typeof value?.toDate === 'function') return value.toDate().getTime();
        const parsed = new Date(value).getTime();
        return Number.isFinite(parsed) ? parsed : 0;
    } catch (_) {
        return 0;
    }
}

function getEffectiveSubmissionDocuments(submission = null) {
    const rawDocs = Array.isArray(submission?.documents) ? submission.documents : [];
    if (rawDocs.length <= 1) return rawDocs;

    const latestByType = new Map();
    rawDocs.forEach((docItem, index) => {
        const type = String(docItem?.documentType || '').trim();
        const key = type || `__index_${index}`;
        const previous = latestByType.get(key);
        const currentScore = Math.max(
            getTimestampMsSafe(docItem?.uploadedAt),
            Number(docItem?.localAddedAt || 0)
        );
        const previousScore = previous
            ? Math.max(
                getTimestampMsSafe(previous?.uploadedAt),
                Number(previous?.localAddedAt || 0)
            )
            : -1;

        if (!previous || currentScore >= previousScore) {
            latestByType.set(key, docItem);
        }
    });

    const orderedTypes = Array.isArray(submission?.documentTypes) ? submission.documentTypes : [];
    const normalizedOrder = orderedTypes
        .map((type) => latestByType.get(String(type || '').trim()))
        .filter(Boolean);

    const alreadyIncluded = new Set(normalizedOrder);
    const remainder = Array.from(latestByType.values()).filter((docItem) => !alreadyIncluded.has(docItem));
    return [...normalizedOrder, ...remainder];
}

async function getLatestSubmissionById(submissionId) {
    const normalizedId = String(submissionId || '').trim();
    if (!normalizedId) return null;

    const cached = allSubmissions.find((s) => s.id === normalizedId) || null;

    try {
        const snap = await getDoc(doc(db, 'submissions', normalizedId));
        if (!snap.exists()) return cached;

        const fresh = { id: snap.id, ...snap.data() };
        const existingIndex = allSubmissions.findIndex((s) => s.id === normalizedId);
        if (existingIndex >= 0) {
            allSubmissions[existingIndex] = fresh;
        } else {
            allSubmissions.push(fresh);
        }
        return fresh;
    } catch (_) {
        return cached;
    }
}

function buildViewerDocumentUrl(fileUrl, submission = null, docIndex = 0) {
    const cleanUrl = String(fileUrl || '').trim();
    if (!cleanUrl) return '';

    try {
        const url = new URL(cleanUrl, window.location.origin);
        const versionSeed = Math.max(
            getTimestampMsSafe(submission?.reuploadedAt),
            getTimestampMsSafe(submission?.finalSubmittedAt),
            getTimestampMsSafe(submission?.rsaSubmittedAt),
            getTimestampMsSafe(submission?.fixSubmittedAt),
            getTimestampMsSafe(submission?.uploadedAt)
        ) || Date.now();
        url.searchParams.set('_v', `${versionSeed}-${docIndex}`);
        return url.toString();
    } catch (_) {
        const joiner = cleanUrl.includes('?') ? '&' : '?';
        return `${cleanUrl}${joiner}_v=${Date.now()}-${docIndex}`;
    }
}

function isRsaProcessingStatus(status) {
    const normalized = String(status || '').toLowerCase();
    return normalized === 'processing_to_pfa' || normalized === 'approved';
}

function getReviewerWorkflowState(submission = {}) {
    const normalizedStatus = String(submission?.status || '').toLowerCase().trim();
    if (!String(submission?.reviewedBy || '').trim()) return '';
    if (normalizedStatus === 'rejected' || normalizedStatus === 'rejected_by_rsa') return 'rejected';
    if (
        normalizedStatus === 'processing_to_pfa' ||
        normalizedStatus === 'approved' ||
        normalizedStatus === 'sent_to_pfa' ||
        normalizedStatus === 'rsa_submitted' ||
        normalizedStatus === 'paid' ||
        normalizedStatus === 'cleared'
    ) {
        return 'approved';
    }
    return '';
}

function getReviewerDecisionState(submission = {}) {
    const decision = String(submission?.reviewerDecision || '').toLowerCase().trim();
    if (decision === 'approved' || decision === 'rejected') return decision;

    const normalizedStatus = String(submission?.status || '').toLowerCase().trim();
    if (normalizedStatus === 'rejected') return 'rejected';
    if (
        normalizedStatus === 'processing_to_pfa' ||
        normalizedStatus === 'approved' ||
        normalizedStatus === 'sent_to_pfa' ||
        normalizedStatus === 'rsa_submitted' ||
        normalizedStatus === 'paid' ||
        normalizedStatus === 'cleared'
    ) {
        return 'approved';
    }
    return '';
}

function isAssignedToCurrentReviewer(submission = {}) {
    return normalizeEmail(submission?.assignedTo) === normalizeEmail(currentUser?.email);
}

function isReviewedByCurrentReviewer(submission = {}) {
    return normalizeEmail(submission?.reviewedBy) === normalizeEmail(currentUser?.email);
}

function isPendingForCurrentReviewer(submission = {}) {
    return String(submission?.status || '').toLowerCase().trim() === 'pending' && isAssignedToCurrentReviewer(submission);
}

function isApprovedByCurrentReviewer(submission = {}) {
    return isReviewedByCurrentReviewer(submission) && getReviewerDecisionState(submission) === 'approved';
}

function isRejectedByCurrentReviewer(submission = {}) {
    return isReviewedByCurrentReviewer(submission) && getReviewerDecisionState(submission) === 'rejected';
}

function formatReviewerStatusLabel(status) {
    const normalized = String(status || '').toLowerCase();
    if (normalized === 'processing_to_pfa' || normalized === 'approved') return 'Processing to PFA';
    if (normalized === 'rejected_by_rsa') return 'Rejected by RSA';
    if (!normalized) return '-';
    return normalized.charAt(0).toUpperCase() + normalized.slice(1);
}

function escapeHtml(value) {
    return String(value ?? '')
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&#39;');
}

function getRejectionHistoryEntries(submission) {
    const rawHistory = Array.isArray(submission?.rejectionHistory) ? submission.rejectionHistory : [];
    const normalizedHistory = rawHistory
        .map((entry) => {
            if (typeof entry === 'string') {
                const reason = entry.trim();
                return reason ? { reason, rejectedAt: null, rejectedBy: '' } : null;
            }
            const reason = String(entry?.reason || '').trim();
            if (!reason) return null;
            return {
                reason,
                rejectedAt: entry?.rejectedAt || null,
                rejectedBy: String(entry?.rejectedBy || entry?.reviewerEmail || '').trim()
            };
        })
        .filter(Boolean);

    if (normalizedHistory.length > 0) return normalizedHistory;

    const fallback = String(
        submission?.latestRejectionReason ||
        submission?.previousRejectionReason ||
        submission?.comment ||
        ''
    ).trim();
    return fallback ? [{
        reason: fallback,
        rejectedAt: submission?.latestRejectedAt || submission?.previousRejectedAt || submission?.reviewedAt || null,
        rejectedBy: String(submission?.latestRejectedBy || submission?.previousRejectedBy || submission?.reviewedBy || '').trim()
    }] : [];
}

function getRejectionHistoryReasons(submission) {
    return getRejectionHistoryEntries(submission).map((entry) => entry.reason).filter(Boolean);
}

function hasReviewContext(submission) {
    return getRejectionHistoryReasons(submission).length > 0;
}

function getPreviousRejectedBy(submission) {
    return String(
        submission?.latestRejectedBy ||
        submission?.previousRejectedBy ||
        submission?.reviewedBy ||
        ''
    ).trim();
}

function getPreviousRejectedAt(submission) {
    return submission?.latestRejectedAt || submission?.previousRejectedAt || submission?.reviewedAt || null;
}

function getTimestampMs(value) {
    if (!value) return 0;
    if (typeof value?.toMillis === 'function') return value.toMillis();
    if (typeof value?.seconds === 'number') return value.seconds * 1000;
    const parsed = new Date(value).getTime();
    return Number.isFinite(parsed) ? parsed : 0;
}

function isResubmittedSubmission(submission) {
    return String(submission?.status || '').toLowerCase() === 'pending' && (
        submission?.resubmittedAfterRejection === true ||
        Number(submission?.fixCount || 0) > 0 ||
        Boolean(submission?.reuploadedAt)
    );
}

function renderReviewContext(submission, { multiline = false } = {}) {
    const reasons = getRejectionHistoryReasons(submission);
    if (!reasons.length) return '';

    const lines = reasons.map((reason, index) => `Rejection ${index + 1}: ${escapeHtml(reason)}`);
    const html = lines.join('<br>');
    return multiline ? html : `<div class="review-context-snippet">${html}</div>`;
}

function renderReviewContextButton(submissionId, submission) {
    if (!hasReviewContext(submission)) return '-';
    return `
        <button class="action-btn review-context-trigger" onclick="window.openReviewContextModal('${submissionId}')">
            <i class="fas fa-eye"></i> View
        </button>
    `;
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
    return `<a href="https://wa.me/${normalized}" target="_blank" rel="noopener noreferrer" aria-label="Open WhatsApp chat" title="Open WhatsApp chat"><i class="fab fa-whatsapp"></i></a>`;
}

// Cache for user full names
const userFullNameCache = new Map();
let reviewerLeaveHistoryLoaded = false;
let reviewerMyLeaveHistory = [];
let reviewerReliefLeaveHistory = [];

// ==================== DOM ELEMENTS ====================
const userName = document.getElementById('userName');
const userAvatar = document.getElementById('userAvatar');
const pendingTableBody = document.getElementById('pendingTableBody');
const approvedTableBody = document.getElementById('approvedTableBody');
const rejectedTableBody = document.getElementById('rejectedTableBody');
const commentModal = document.getElementById('commentModal');
const reviewContextModal = document.getElementById('reviewContextModal');
const viewerModal = document.getElementById('viewerModal');
const closeCommentModal = document.getElementById('closeCommentModal');
const closeReviewContextModal = document.getElementById('closeReviewContextModal');
const closeViewer = document.getElementById('closeViewer');
const cancelComment = document.getElementById('cancelComment');
const closeReviewContextBtn = document.getElementById('closeReviewContextBtn');
const approveBtn = document.getElementById('approveDocument');
const rejectBtn = document.getElementById('rejectDocument');
const commentText = document.getElementById('commentText');
const modalCustomerName = document.getElementById('modalCustomerName');
const modalDocumentType = document.getElementById('modalDocumentType');
const previousRejectionBox = document.getElementById('previousRejectionBox');
const reviewContextCustomerName = document.getElementById('reviewContextCustomerName');
const reviewContextContact = document.getElementById('reviewContextContact');
const reviewContextHistory = document.getElementById('reviewContextHistory');
const viewerFileName = document.getElementById('viewerFileName');
const documentViewer = document.getElementById('documentViewer');
const viewerDownloadBtn = document.getElementById('viewerDownloadBtn');
const viewerDownloadSection = document.getElementById('viewerDownloadSection');
const viewerNav = document.getElementById('viewerNav');
const viewerContextBox = document.getElementById('viewerContextBox');
const notification = document.getElementById('notification');
const pageTitle = document.getElementById('pageTitle');
const pendingCountBadge = document.getElementById('pendingCount');
const approvedCountBadge = document.getElementById('approvedCount');
const rejectedCountBadge = document.getElementById('rejectedCount');
const zoomInBtn = document.getElementById('zoomInBtn');
const zoomOutBtn = document.getElementById('zoomOutBtn');
const zoomResetBtn = document.getElementById('zoomResetBtn');
const zoomLevel = document.getElementById('zoomLevel');
const selectedCount = document.getElementById('selectedCount');
const profileNameEl = document.getElementById('profileName');
const profileRegisteredAtEl = document.getElementById('profileRegisteredAt');
const profileEmailEl = document.getElementById('profileEmail');
const profileWhatsappEl = document.getElementById('profileWhatsapp');
const profileLocationEl = document.getElementById('profileLocation');
const profileRoleEl = document.getElementById('profileRole');
const profileStatusEl = document.getElementById('profileStatus');
const reviewerReportMeta = document.getElementById('reviewerReportMeta');
const reviewerReportStartDate = document.getElementById('reviewerReportStartDate');
const reviewerReportEndDate = document.getElementById('reviewerReportEndDate');
const reviewerReportSummaryBody = document.getElementById('reviewerReportSummaryBody');
const reviewerReportDetailsBody = document.getElementById('reviewerReportDetailsBody');
const generateReviewerReportBtn = document.getElementById('generateReviewerReportBtn');
const exportReviewerReportBtn = document.getElementById('exportReviewerReportBtn');

function initializeReviewerReportDates() {
    const today = new Date();
    const sixDaysAgo = new Date(today.getTime() - (6 * 24 * 60 * 60 * 1000));
    const toInputValue = (date) => `${date.getFullYear()}-${`${date.getMonth() + 1}`.padStart(2, '0')}-${`${date.getDate()}`.padStart(2, '0')}`;
    if (reviewerReportStartDate && !reviewerReportStartDate.value) reviewerReportStartDate.value = toInputValue(sixDaysAgo);
    if (reviewerReportEndDate && !reviewerReportEndDate.value) reviewerReportEndDate.value = toInputValue(today);
}

function resolveReviewerKnownName(email) {
    const normalized = normalizeEmail(email);
    if (!normalized) return 'Unassigned';
    if (normalizeEmail(currentUserData?.email) === normalized) {
        return currentUserData?.fullName || currentUserData?.displayName || normalized;
    }
    const matchedSubmission = allSubmissions.find((sub) => (
        normalizeEmail(sub.uploadedBy) === normalized
        || normalizeEmail(sub.reviewedBy) === normalized
        || normalizeEmail(sub.assignedTo) === normalized
    ));
    return matchedSubmission?.uploadedByName || matchedSubmission?.reviewedByName || normalized;
}

function buildReviewerStageReport() {
    const startDate = String(reviewerReportStartDate?.value || '').trim();
    const endDate = String(reviewerReportEndDate?.value || '').trim();
    if (!startDate || !endDate) throw new Error('Choose both start date and end date.');
    if (startDate > endDate) throw new Error('Start date cannot be after end date.');
    const sourceRecords = allSubmissions.filter((sub) => normalizeEmail(sub.assignedTo));
    return buildDashboardStageReport({
        stageId: 'reviewer',
        records: sourceRecords,
        rangeStart: startDate,
        rangeEnd: endDate,
        resolveName: resolveReviewerKnownName
    });
}

function renderProfileTab() {
    if (!profileNameEl && !profileEmailEl && !profileRoleEl && !profileStatusEl) return;
    const fullName = currentUserData?.fullName || currentUserData?.displayName || currentUser?.displayName || currentUser?.email || 'N/A';
    const registeredAt = currentUserData?.createdAt ? formatTimestamp(currentUserData.createdAt) : '-';
    const email = currentUserData?.email || currentUser?.email || 'N/A';
    const whatsapp = currentUserData?.whatsappNumber || currentUserData?.phone || '-';
    const location = currentUserData?.location || '-';
    const role = String(currentUserData?.role || 'reviewer');
    const normalizedRole = role;
    const status = String(currentUserData?.status || 'active');
    if (profileNameEl) profileNameEl.textContent = fullName;
    if (profileRegisteredAtEl) profileRegisteredAtEl.textContent = registeredAt;
    if (profileEmailEl) profileEmailEl.textContent = email;
    if (profileWhatsappEl) profileWhatsappEl.textContent = whatsapp;
    if (profileLocationEl) profileLocationEl.textContent = location;
    if (profileRoleEl) profileRoleEl.textContent = normalizedRole.charAt(0).toUpperCase() + normalizedRole.slice(1);
    if (profileStatusEl) profileStatusEl.textContent = status.charAt(0).toUpperCase() + status.slice(1);
}

// ==================== ZOOM LEVEL ====================
let currentZoom = 1.0;

// ==================== DOCUMENT FETCH HELPER ====================
async function fetchWithCorsFallback(url) {
    const cleanUrl = url?.toString().trim().replace(/[\s\n\r\t]+/g, '');
    if (!cleanUrl) throw new Error('Invalid URL');

    try {
        const response = await fetch(cleanUrl, {
            mode: 'cors',
            credentials: 'omit',
            headers: {
                'Accept': 'application/pdf, image/*, */*'
            }
        });
        if (!response.ok) {
            throw new Error(`Document fetch failed: ${response.status}`);
        }
        return response;
    } catch (error) {
        const proxyUrl = getBackblazeDownloadProxyUrl(cleanUrl);
        if (!proxyUrl) {
            error.corsBlocked = true;
            throw error;
        }

        const proxyResponse = await fetch(proxyUrl, { credentials: 'same-origin' });
        if (!proxyResponse.ok) {
            throw new Error(`Document proxy failed: ${proxyResponse.status}`);
        }
        return proxyResponse;
    }
}

function getBackblazeDownloadProxyUrl(cleanUrl) {
    try {
        const parsed = new URL(cleanUrl);
        const isBackblaze = parsed.protocol === 'https:' && /\.backblazeb2\.com$/i.test(parsed.hostname);
        if (!isBackblaze || !parsed.pathname.startsWith('/file/cmbank-rsa-documents/')) return '';
        return `/api/backblaze-download.php?url=${encodeURIComponent(cleanUrl)}`;
    } catch (error) {
        return '';
    }
}

// ==================== GET USER FULL NAME BY EMAIL ====================
async function getUserFullName(email) {
    const normalized = normalizeEmailShared(email);
    if (!normalized) return 'Unknown';
    if (userFullNameCache.has(normalized)) {
        return userFullNameCache.get(normalized);
    }
    const fullName = await getUserFullNameShared(db, normalized);
    userFullNameCache.set(normalized, fullName);
    return fullName;
}

// ==================== CHECK IF FILE IS PDF ====================
function isPdfFile(fileUrl, fileName) {
    const lowerName = (fileName || '').toLowerCase();
    const lowerUrl = (fileUrl || '').toLowerCase();
    return lowerName.endsWith('.pdf') || lowerUrl.includes('.pdf') || lowerUrl.includes('application/pdf');
}

// ==================== CHECK IF FILE IS IMAGE ====================
function isImageFile(fileName) {
    const lowerName = fileName.toLowerCase();
    return lowerName.endsWith('.jpg') || 
           lowerName.endsWith('.jpeg') || 
           lowerName.endsWith('.png') || 
           lowerName.endsWith('.gif') || 
           lowerName.endsWith('.bmp') || 
           lowerName.endsWith('.webp');
}

// ==================== CONVERT IMAGE TO PDF WITH BETTER FORMAT SUPPORT ====================
async function imageToPdf(imageBlob, imageName) {
    try {
        const { PDFDocument } = PDFLib;
        const pdfDoc = await PDFDocument.create();
        const page = pdfDoc.addPage([595, 842]); // A4 size
        
        const imageBytes = await imageBlob.arrayBuffer();
        let image;
        const lowerName = imageName.toLowerCase();
        
        // Try to detect image type from blob if name doesn't indicate clearly
        const imageType = imageBlob.type || '';
        
        try {
            if (lowerName.endsWith('.jpg') || lowerName.endsWith('.jpeg') || imageType.includes('jpeg')) {
                image = await pdfDoc.embedJpg(imageBytes);
            } else if (lowerName.endsWith('.png') || imageType.includes('png')) {
                image = await pdfDoc.embedPng(imageBytes);
            } else {
                // Try to detect format from magic numbers
                const arr = new Uint8Array(imageBytes.slice(0, 4));
                if (arr[0] === 0xFF && arr[1] === 0xD8) {
                    // JPEG magic number
                    image = await pdfDoc.embedJpg(imageBytes);
                } else if (arr[0] === 0x89 && arr[1] === 0x50 && arr[2] === 0x4E && arr[3] === 0x47) {
                    // PNG magic number
                    image = await pdfDoc.embedPng(imageBytes);
                } else {
                    throw new Error(`Unsupported image format: ${imageName}`);
                }
            }
        } catch (embedError) {
            throw new Error(`Failed to embed image: ${embedError.message}`);
        }
        
        const pageWidth = page.getWidth();
        const pageHeight = page.getHeight();
        const imgWidth = image.width;
        const imgHeight = image.height;
        const scale = Math.min(pageWidth / imgWidth, pageHeight / imgHeight) * 0.9;
        
        page.drawImage(image, {
            x: (pageWidth - imgWidth * scale) / 2,
            y: (pageHeight - imgHeight * scale) / 2,
            width: imgWidth * scale,
            height: imgHeight * scale,
        });
        
        return await pdfDoc.save();
    } catch (error) {
        throw new Error(`Failed to convert image to PDF: ${error.message}`);
    }
}

// ==================== SAVE FILE WITH LOCATION PICKER (fallback) ====================
async function saveFileWithLocationPicker(blob, defaultFileName) {
    if (!('showSaveFilePicker' in window)) {
        // Fallback for browsers that don't support showSaveFilePicker
        showNotification('Save picker not supported. Using direct download...', 'info');
        downloadBlobAsFile(blob, defaultFileName);
        return true;
    }

    try {
        const fileHandle = await window.showSaveFilePicker({
            suggestedName: defaultFileName,
            types: [{
                description: 'PDF Document',
                accept: { 'application/pdf': ['.pdf'] }
            }]
        });

        const writable = await fileHandle.createWritable();
        await writable.write(blob);
        await writable.close();

        showNotification(`ÃƒÂ¢Ã…â€œÃ¢â‚¬Â¦ Saved: ${defaultFileName}`, 'success');
        return true;
    } catch (error) {
        if (error.name === 'AbortError') {
            showNotification('Save cancelled', 'info');
        } else {
            showNotification('Save failed: ' + error.message, 'error');
            
            // Fallback to direct download
            downloadBlobAsFile(blob, defaultFileName);
        }
        return false;
    }
}

// ==================== SAVE BLOB TO FOLDER WITH CUSTOMER SUBDIRECTORY ====================
async function saveBlobToFolderPicker(blob, defaultFileName, customerName = 'Customer') {
    // If the directory picker isn't available, fall back to the old behavior
    if (!('showDirectoryPicker' in window)) {
        showNotification('Folder picker not supported. Falling back to save dialog...', 'info');
        return saveFileWithLocationPicker(blob, defaultFileName);
    }

    try {
        showNotification('ÃƒÂ°Ã…Â¸Ã¢â‚¬Å“Ã‚Â Please select a destination folder...', 'info');
        const dirHandle = await window.showDirectoryPicker({
            mode: 'readwrite',
            startIn: 'downloads'
        });

        const safeCustomerName = customerName.replace(/[^a-zA-Z0-9\s_-]/g, '_').trim() || 'Customer';
        const customerFolder = await dirHandle.getDirectoryHandle(safeCustomerName, { create: true });
        const fileHandle = await customerFolder.getFileHandle(defaultFileName, { create: true });
        const writable = await fileHandle.createWritable();
        await writable.write(blob);
        await writable.close();

        showNotification(`ÃƒÂ¢Ã…â€œÃ¢â‚¬Â¦ Saved to ${safeCustomerName}/${defaultFileName}`, 'success');
        return true;
    } catch (error) {
        if (error.name === 'AbortError') {
            showNotification('Save cancelled', 'info');
        } else {
            showNotification('Save failed: ' + error.message, 'error');
            // As a last resort try the save-file picker
            await saveFileWithLocationPicker(blob, defaultFileName);
        }
        return false;
    }
}

// ==================== DOWNLOAD BLOB AS FILE (FALLBACK) ====================
function downloadBlobAsFile(blob, fileName) {
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = fileName;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(url);
    showNotification('ÃƒÂ¢Ã…â€œÃ¢â‚¬Â¦ Download started', 'success');
}

function openDirectDocumentDownload(fileUrl, fileName = 'document.pdf') {
    const cleanUrl = fileUrl?.toString().trim().replace(/[\s\n\r\t]+/g, '');
    if (!cleanUrl) return false;

    const link = document.createElement('a');
    link.href = cleanUrl;
    link.target = '_blank';
    link.rel = 'noopener noreferrer';
    link.download = (fileName || 'document.pdf').replace(/[\\/:*?"<>|]/g, '_');
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    return true;
}

// ==================== DOWNLOAD ALL TO FOLDER (folder picker) ====================
window.downloadAll = async (submissionId) => {
    const sub = allSubmissions.find(s => s.id === submissionId);
    if (!sub) return;

    if (downloadInProgress) {
        showNotification('Download already in progress', 'warning');
        return;
    }

    // Browser security requires explicit user action to choose a directory.
    // It's not possible for a web page to unilaterally write into the user's
    // Documents or Downloads folder without a picker. The File System Access
    // API forces the user to select a location to prevent malware from
    // arbitrarily saving files on the machine.
    if (!('showDirectoryPicker' in window)) {
        showNotification('Folder picker not supported. Use Chrome/Edge.', 'error');
        return;
    }

    try {
        downloadInProgress = true;
        
        if (!sub.documents || sub.documents.length === 0) {
            showNotification('No documents to download', 'error');
            return;
        }

        // let the loader display while we ask for folder and begin downloading
        showLoader('Preparing download...');
        showNotification('ÃƒÂ°Ã…Â¸Ã¢â‚¬Å“Ã‚Â Please select a folder for download...', 'info');
        const dirHandle = await window.showDirectoryPicker({
            mode: 'readwrite',
            startIn: 'downloads'
        });
        const safeCustomerName = sub.customerName.replace(/[^a-zA-Z0-9\s_-]/g, '_').trim();
        const customerFolder = await dirHandle.getDirectoryHandle(safeCustomerName, { create: true });
        const totalDocs = sub.documents.length;
        let savedCount = 0;
        let directOpenedCount = 0;
        let failedCount = 0;

        for (const [index, doc] of sub.documents.entries()) {
            const docTypeLabel = DOCUMENT_TYPES[doc.documentType] || doc.documentType || 'Document';
            const fileUrl = doc.fileUrl?.trim();
            const fileName = doc.name || 'document';
            const fileExt = fileName.split('.').pop()?.toLowerCase() || 'pdf';
            const outputFileName = `${safeCustomerName}_${docTypeLabel}.${fileExt}`;

            const percent = Math.round(((index + 1) / totalDocs) * 100);
            showLoader(`Downloading (${percent}%) ÃƒÂ¢Ã¢â€šÂ¬Ã¢â‚¬Å“ ${docTypeLabel}`);

            try {
                const response = await fetchWithCorsFallback(fileUrl);
                const blob = await response.blob();
                const fileHandle = await customerFolder.getFileHandle(outputFileName, { create: true });
                const writable = await fileHandle.createWritable();
                await writable.write(blob);
                await writable.close();
                savedCount++;
            } catch (err) {
                if (openDirectDocumentDownload(fileUrl, outputFileName)) {
                    directOpenedCount++;
                    showNotification(`Storage blocked secure save for ${docTypeLabel}; opened directly.`, 'warning');
                    continue;
                }
                failedCount++;
                showNotification(`ÃƒÂ¢Ã…Â¡Ã‚Â ÃƒÂ¯Ã‚Â¸Ã‚Â Failed: ${docTypeLabel}`, 'warning');
            }
        }

        showLoader('Finalizing...');
        if (failedCount > 0 || directOpenedCount > 0) {
            showNotification(`Saved ${savedCount}, opened ${directOpenedCount} directly, ${failedCount} failed`, 'warning');
            return;
        }
        showNotification(`ÃƒÂ¢Ã…â€œÃ¢â‚¬Â¦ All documents saved to: ${safeCustomerName}/`, 'success');
    } catch (error) {
        if (error.name === 'AbortError') {
            showNotification('Download cancelled', 'info');
        } else {
            showNotification('Download failed: ' + error.message, 'error');
        }
    } finally {
        downloadInProgress = false;
        hideLoader();
    }
};

// ==================== FORMAT TIMESTAMP ====================
function formatTimestamp(timestamp) {
    return formatAppDateTime(timestamp, 'N/A');
}

function normalizeDocumentType(documentType = '') {
    return String(documentType).toLowerCase().trim().replace(/[\s-]+/g, '_');
}

function getCanonicalDocumentType(doc = {}, fallbackType = '') {
    const rawType = normalizeDocumentType(doc.documentType || doc.type || doc.docType || fallbackType || '');
    const rawName = normalizeDocumentType(doc.name || doc.fileName || '');
    const source = `${rawType} ${rawName}`;

    if (source.includes('birth') && source.includes('certificate')) return 'birth_certificate';
    if (source.includes('nin') || source.includes('national_id')) return 'nin';
    if (source.includes('rsa') && source.includes('statement')) return 'rsa_statement';
    if (source.includes('pfa') && (source.includes('form') || source.includes('application'))) return 'pfa_form';
    if (source.includes('bvn')) return 'bvn';
    return rawType;
}

function formatSimpleDate(value) {
    if (!value) return '';
    return formatAppDate(value, String(value));
}

function formatNaira(value) {
    if (value === null || value === undefined || value === '') return '';
    const parsed = typeof value === 'number' ? value : Number(String(value).replace(/,/g, ''));
    if (Number.isNaN(parsed)) return String(value);
    return parsed.toLocaleString('en-NG', {
        style: 'currency',
        currency: 'NGN',
        minimumFractionDigits: 2,
        maximumFractionDigits: 2
    });
}

function getSubmissionField(submission, ...keys) {
    const details = submission?.customerDetails || {};
    const detailEntries = Object.entries(details);
    const rootEntries = Object.entries(submission || {});

    for (const key of keys) {
        const normalizedKey = normalizeDocumentType(key);

        if (details[key] !== undefined && details[key] !== null && String(details[key]).trim() !== '') return details[key];
        if (submission?.[key] !== undefined && submission[key] !== null && String(submission[key]).trim() !== '') return submission[key];

        const detailMatch = detailEntries.find(([k, v]) => normalizeDocumentType(k) === normalizedKey && v !== null && v !== undefined && String(v).trim() !== '');
        if (detailMatch) return detailMatch[1];

        const rootMatch = rootEntries.find(([k, v]) => normalizeDocumentType(k) === normalizedKey && v !== null && v !== undefined && String(v).trim() !== '');
        if (rootMatch) return rootMatch[1];
    }
    return '';
}

function renderViewerContextForType(canonicalType) {
    if (!viewerContextBox) return;

    let html = '';
    if (canonicalType === 'birth_certificate') {
        const dob = getSubmissionField(currentViewerSubmission, 'dob', 'dateOfBirth', 'customerDob');
        if (dob) html = `<strong>Date of Birth:</strong> ${formatSimpleDate(dob)}`;
    } else if (canonicalType === 'nin') {
        const nin = getSubmissionField(currentViewerSubmission, 'nin', 'customerNIN', 'nationalId', 'nationalIdentificationNumber');
        if (nin) html = `<strong>NIN:</strong> ${nin}`;
    } else if (canonicalType === 'rsa_statement') {
        const rsaBalance = getSubmissionField(currentViewerSubmission, 'rsaBalance', 'rsa_balance');
        const rsaStatementDate = getSubmissionField(currentViewerSubmission, 'rsaStatementDate', 'rsa_statement_date');
        const balanceText = formatNaira(rsaBalance);
        const dateText = formatSimpleDate(rsaStatementDate);
        const parts = [];
        if (balanceText) parts.push(`<strong>RSA Balance:</strong> ${balanceText}`);
        if (dateText) parts.push(`<strong>RSA Statement Date:</strong> ${dateText}`);
        html = parts.join(' | ');
    } else if (canonicalType === 'pfa_form' || canonicalType === 'pfa_application_form') {
        const pfaName = getSubmissionField(currentViewerSubmission, 'pfa', 'pfaName', 'pensionFundAdministrator');
        if (pfaName) html = `<strong>PFA Name:</strong> ${pfaName}`;
    }

    if (html) {
        viewerContextBox.innerHTML = html;
        viewerContextBox.style.display = 'inline-block';
    } else {
        viewerContextBox.innerHTML = '';
        viewerContextBox.style.display = 'none';
    }
}

function clearViewerContext() {
    if (!viewerContextBox) return;
    viewerContextBox.innerHTML = '';
    viewerContextBox.style.display = 'none';
}

// ==================== INITIALIZATION ====================
document.addEventListener('DOMContentLoaded', () => {
    auth.onAuthStateChanged(async (user) => {
        if (user) {
            currentUser = user;
            const normalizedCurrentEmail = normalizeEmail(user.email);
            
            try {
                const resolvedProfile = await getCurrentUserProfileShared(db, user);

                if (resolvedProfile) {
                    currentUserData = resolvedProfile;
                    userName.textContent = currentUserData.fullName || currentUserData.displayName || user.email.split('@')[0];
                } else {
                    currentUserData = { email: user.email, fullName: user.displayName || user.email, role: 'reviewer', status: 'active' };
                    userName.textContent = user.displayName || user.email.split('@')[0];
                }
            } catch (err) {
                currentUserData = { email: user.email, fullName: user.displayName || user.email, role: 'reviewer', status: 'active' };
                userName.textContent = user.displayName || user.email.split('@')[0];
            }

            userAvatar.src = user.photoURL || 'data:image/svg+xml,%3Csvg xmlns=\'http://www.w3.org/2000/svg\' width=\'40\' height=\'40\' viewBox=\'0 0 40 40\'%3E%3Ccircle cx=\'20\' cy=\'20\' r=\'20\' fill=\'%23003366\'/%3E%3Ctext x=\'20\' y=\'25\' text-anchor=\'middle\' fill=\'%23ffffff\' font-size=\'16\'%3EÃƒÂ°Ã…Â¸Ã¢â‚¬ËœÃ‚Â¤%3C/text%3E%3C/svg%3E';

            // Replace broken-encoded avatar fallback with a stable initial.
            if (!user.photoURL && userAvatar) {
                userAvatar.src = 'data:image/svg+xml,%3Csvg xmlns=%27http://www.w3.org/2000/svg%27 width=%2740%27 height=%2740%27 viewBox=%270 0 40 40%27%3E%3Ccircle cx=%2720%27 cy=%2720%27 r=%2720%27 fill=%27%23003366%27/%3E%3Ctext x=%2720%27 y=%2725%27 text-anchor=%27middle%27 fill=%27%23ffffff%27 font-size=%2716%27 font-family=%27Arial%27%3ER%3C/text%3E%3C/svg%3E';
            }

            const role = String(currentUserData?.role || '').toLowerCase();
            if (role === 'reviewer') {
                renderProfileTab();
                loadSubmissions();
            } else {
                showNotification('Access denied. Reviewer privileges required.', 'error');
                setTimeout(() => { window.location.href = 'index.html'; }, 2000);
            }
        } else {
            window.location.href = 'index.html';
        }
    });
    setupEventListeners();
    setupZoomControls();
    setupForceRefreshButtons();
    setupIdleLogout();
});

window.signOutUser = async () => {
    try {
        const userId = currentUserData?.id || currentUser?.uid || '';
        if (userId) {
            await updateDoc(doc(db, 'users', userId), {
                isOnline: false,
                lastSeenAt: serverTimestamp(),
                lastLogoutAt: serverTimestamp()
            }).catch(() => {});
        }
        await signOut(auth);
    } catch (e) { }
    window.location.href = 'index.html';
};

const IDLE_TIMEOUT_MS = 30 * 60 * 1000;
let idleLastActivity = Date.now();
let idleIntervalHandle = null;

function setupIdleLogout() {
    if (idleIntervalHandle) return;
    const bump = () => { idleLastActivity = Date.now(); };
    ['mousemove', 'mousedown', 'keydown', 'touchstart', 'scroll'].forEach((ev) => {
        window.addEventListener(ev, bump, { passive: true });
    });
    window.addEventListener('focus', bump);
    document.addEventListener('visibilitychange', () => { if (!document.hidden) bump(); });

    idleIntervalHandle = setInterval(() => {
        if (!auth.currentUser) return;
        if (Date.now() - idleLastActivity >= IDLE_TIMEOUT_MS) {
            window.signOutUser();
        }
    }, 60 * 1000);
}

function forceHardRefresh() {
    const url = new URL(window.location.href);
    url.searchParams.set('_', Date.now().toString());
    window.location.replace(url.toString());
}

function setupForceRefreshButtons() {
    document.getElementById('forceRefreshBtn')?.addEventListener('click', (e) => {
        e.preventDefault();
        forceHardRefresh();
    });
    document.getElementById('forceRefreshBtnMobile')?.addEventListener('click', (e) => {
        e.preventDefault();
        forceHardRefresh();
    });
}

// ==================== SETUP ZOOM CONTROLS ====================
function setupZoomControls() {
    if (zoomInBtn) {
        zoomInBtn.addEventListener('click', () => {
            currentZoom = Math.min(currentZoom + 0.25, 3.0);
            applyZoom();
        });
    }
    
    if (zoomOutBtn) {
        zoomOutBtn.addEventListener('click', () => {
            currentZoom = Math.max(currentZoom - 0.25, 0.5);
            applyZoom();
        });
    }
    
    if (zoomResetBtn) {
        zoomResetBtn.addEventListener('click', () => {
            currentZoom = 1.0;
            applyZoom();
        });
    }
}

function applyZoom() {
    if (documentViewer) {
        documentViewer.style.transform = `scale(${currentZoom})`;
        documentViewer.style.transformOrigin = 'center top';
        documentViewer.style.transition = 'transform 0.3s ease';
    }
    if (zoomLevel) {
        zoomLevel.textContent = `${Math.round(currentZoom * 100)}%`;
    }
}

// ==================== EVENT LISTENERS ====================
function setupEventListeners() {
    if (closeCommentModal) closeCommentModal.addEventListener('click', () => closeModal(commentModal));
    if (closeReviewContextModal) closeReviewContextModal.addEventListener('click', () => closeModal(reviewContextModal));
    if (closeViewer) closeViewer.addEventListener('click', closeViewerModal);
    if (cancelComment) cancelComment.addEventListener('click', () => closeModal(commentModal));
    if (closeReviewContextBtn) closeReviewContextBtn.addEventListener('click', () => closeModal(reviewContextModal));

    if (approveBtn) approveBtn.addEventListener('click', () => confirmApprove());
    if (rejectBtn) rejectBtn.addEventListener('click', () => confirmReject());
    document.getElementById('reviewerLeaveMineBtn')?.addEventListener('click', () => {
        document.getElementById('reviewerLeaveMineSection')?.style.setProperty('display', '');
        document.getElementById('reviewerLeaveReliefSection')?.style.setProperty('display', 'none');
    });
    document.getElementById('reviewerLeaveReliefBtn')?.addEventListener('click', () => {
        document.getElementById('reviewerLeaveMineSection')?.style.setProperty('display', 'none');
        document.getElementById('reviewerLeaveReliefSection')?.style.setProperty('display', '');
    });
    document.getElementById('closeReviewerLeaveApplicationsBtn')?.addEventListener('click', () => {
        document.getElementById('reviewerLeaveApplicationsModal')?.classList.remove('active');
    });
    
    if (viewerDownloadBtn) viewerDownloadBtn.addEventListener('click', saveCurrentDocumentWithPicker);

    window.addEventListener('click', (e) => {
        if (e.target === commentModal) closeModal(commentModal);
        if (e.target === reviewContextModal) closeModal(reviewContextModal);
        if (e.target === viewerModal) closeViewerModal();
        if (e.target === document.getElementById('reviewerLeaveApplicationsModal')) {
            document.getElementById('reviewerLeaveApplicationsModal')?.classList.remove('active');
        }
    });
}

// ==================== LOAD SUBMISSIONS ====================
async function loadSubmissions() {
    const reviewerEmail = normalizeEmail(currentUser?.email);
    if (!reviewerEmail) return;
    const loadVersion = ++submissionsLoadVersion;

    // Load broadly, then filter locally so previously reviewed items remain visible
    // after they are moved onward to RSA/PFA and no longer stay in assignedTo.
    const q = query(
        collection(db, 'submissions'),
        orderBy('uploadedAt', 'desc')
    );

    const processSnapshot = async (snapshot) => {
        const relevantDocs = snapshot.docs.filter((docSnap) => {
            const data = docSnap.data() || {};
            return normalizeEmail(data.assignedTo) === reviewerEmail
                || normalizeEmail(data.reviewedBy) === reviewerEmail;
        });
        const uploaderEmails = [...new Set(relevantDocs.map(doc => doc.data().uploadedBy))];
        const reviewerEmails = [...new Set(relevantDocs.map(doc => doc.data().reviewedBy))];
        const allEmails = [...new Set([...uploaderEmails, ...reviewerEmails].filter(Boolean))];
        await Promise.all(allEmails.map(email => getUserFullName(email)));

        const nextSubmissions = await Promise.all(relevantDocs.map(async (docSnap) => {
            const data = docSnap.data();
            const uploaderName = await getUserFullName(data.uploadedBy);
            const reviewerName = data.reviewedBy ? await getUserFullName(data.reviewedBy) : null;
            return {
                id: docSnap.id,
                ...data,
                uploadedByName: uploaderName,
                reviewedByName: reviewerName
            };
        }));
        if (loadVersion !== submissionsLoadVersion) return;
        allSubmissions = nextSubmissions;

        renderAllTables();
        updatePendingCount();
        if (typeof updateDashboardCards === 'function') updateDashboardCards();
        if (typeof renderRecentReviews === 'function') renderRecentReviews();
    };

    try {
        onSnapshot(q, processSnapshot, async (error) => {
            await loadSubmissionsFallback();
        });
    } catch (error) {
        await loadSubmissionsFallback();
    }
}

async function loadSubmissionsFallback() {
    try {
        const reviewerEmail = normalizeEmail(currentUser?.email);
        if (!reviewerEmail) return;
        const loadVersion = ++submissionsLoadVersion;

        const fallbackQuery = query(
            collection(db, 'submissions'),
            orderBy('uploadedAt', 'desc')
        );
        const snapshot = await getDocs(fallbackQuery);
        const docsSorted = snapshot.docs.filter((docSnap) => {
            const data = docSnap.data() || {};
            return normalizeEmail(data.assignedTo) === reviewerEmail
                || normalizeEmail(data.reviewedBy) === reviewerEmail;
        });

        const uploaderEmails = [...new Set(docsSorted.map(doc => doc.data().uploadedBy))];
        const reviewerEmails = [...new Set(docsSorted.map(doc => doc.data().reviewedBy))];
        const allEmails = [...new Set([...uploaderEmails, ...reviewerEmails].filter(Boolean))];
        await Promise.all(allEmails.map(email => getUserFullName(email)));

        const nextSubmissions = await Promise.all(docsSorted.map(async (docSnap) => {
            const data = docSnap.data();
            const uploaderName = await getUserFullName(data.uploadedBy);
            const reviewerName = data.reviewedBy ? await getUserFullName(data.reviewedBy) : null;
            return {
                id: docSnap.id,
                ...data,
                uploadedByName: uploaderName,
                reviewedByName: reviewerName
            };
        }));
        if (loadVersion !== submissionsLoadVersion) return;
        allSubmissions = nextSubmissions;

        renderAllTables();
        updatePendingCount();
        if (typeof updateDashboardCards === 'function') updateDashboardCards();
        if (typeof renderRecentReviews === 'function') renderRecentReviews();
        showNotification('Live updates unavailable. Showing latest data.', 'info');
    } catch (error) {
        showNotification('Could not load submissions: ' + error.message, 'error');
    }
}

// ==================== UPDATE PENDING COUNT ====================
function updatePendingCount() {
    updateNavCounts();
}

function updateNavCounts() {
    const pendingSubmissions = allSubmissions.filter(isPendingForCurrentReviewer).length;
    const approvedSubmissions = allSubmissions.filter(isApprovedByCurrentReviewer).length;
    const rejectedSubmissions = allSubmissions.filter(isRejectedByCurrentReviewer).length;
    [
        [pendingCountBadge, pendingSubmissions],
        [approvedCountBadge, approvedSubmissions],
        [rejectedCountBadge, rejectedSubmissions]
    ].forEach(([badge, count]) => {
        if (badge) {
            badge.textContent = String(count);
            badge.style.display = 'inline-block';
        }
    });
}

// --- DASHBOARD HELPERS ---
function updateDashboardCards() {
    const approved = allSubmissions.filter(isApprovedByCurrentReviewer).length;
    const pending = allSubmissions.filter(isPendingForCurrentReviewer).length;
    const rejected = allSubmissions.filter(isRejectedByCurrentReviewer).length;
    document.getElementById('vCardApprovedCount') && (document.getElementById('vCardApprovedCount').textContent = approved);
    document.getElementById('vCardPendingCount') && (document.getElementById('vCardPendingCount').textContent = pending);
    document.getElementById('vCardRejectedCount') && (document.getElementById('vCardRejectedCount').textContent = rejected);
    updateNavCounts();
}

function renderRecentReviews() {
    const tbody = document.getElementById('vRecentTableBody');
    if (!tbody) return;
    const q = (document.getElementById('vRecentSearch')?.value || '').toLowerCase();
    const items = allSubmissions.filter(s => {
        if (!isReviewedByCurrentReviewer(s)) return false;
        const decision = getReviewerDecisionState(s);
        return decision === 'approved' || decision === 'rejected';
    })
        .slice().sort((a,b) => {
            const ta = a.reviewedAt?.seconds ? a.reviewedAt.seconds*1000 : (a.reviewedAt ? new Date(a.reviewedAt).getTime() : 0);
            const tb = b.reviewedAt?.seconds ? b.reviewedAt.seconds*1000 : (b.reviewedAt ? new Date(b.reviewedAt).getTime() : 0);
            return tb - ta;
        }).slice(0,10);

    tbody.innerHTML = items.filter(sub => {
        if (!q) return true;
        const hay = `${sub.customerName || ''} ${sub.uploadedByName || ''}`.toLowerCase();
        return hay.includes(q);
    }).map(sub => {
        const dt = sub.reviewedAt ? formatTimestamp(sub.reviewedAt) : 'N/A';
        return `
            <tr>
                <td><strong>${sub.customerName || 'Unknown'}</strong></td>
                <td>${escapeHtml(formatReviewerStatusLabel(sub.status))}</td>
                <td>${dt}</td>
                <td>${sub.uploadedByName || 'N/A'}</td>
                <td><button class="action-btn view-btn-small" onclick="window.viewSubmission('${sub.id}')"><i class="fas fa-eye"></i></button></td>
            </tr>
        `;
    }).join('');
}

// ==================== RENDER TABLES ====================
function renderAllTables() {
    const pendingSubs = allSubmissions
        .filter(isPendingForCurrentReviewer)
        .slice()
        .sort((a, b) => {
            const aTime = getStageTimestampMillis(getSubmissionReviewEntryAt(a));
            const bTime = getStageTimestampMillis(getSubmissionReviewEntryAt(b));
            return bTime - aTime;
        });
    renderPendingTable(pendingSubs);

    // Viewer should only see approved/rejected items they reviewed
    const approvedSubs = allSubmissions
        .filter(isApprovedByCurrentReviewer)
        .slice()
        .sort((a, b) => getStageTimestampMillis(getSubmissionApprovalEntryAt(b)) - getStageTimestampMillis(getSubmissionApprovalEntryAt(a)));
    renderApprovedTable(approvedSubs);

    const rejectedSubs = allSubmissions
        .filter(isRejectedByCurrentReviewer)
        .slice()
        .sort((a, b) => getStageTimestampMillis(getSubmissionRejectionEntryAt(b)) - getStageTimestampMillis(getSubmissionRejectionEntryAt(a)));
    renderRejectedTable(rejectedSubs);
}

function renderPendingTable(submissions) {
    if (!pendingTableBody) return;

    if (submissions.length === 0) {
        pendingTableBody.innerHTML = '<tr><td colspan="9" class="no-data">No pending documents found</td></tr>';
        return;
    }

    pendingTableBody.innerHTML = submissions.map(sub => {
        const isResubmitted = isResubmittedSubmission(sub);
        let date = 'N/A';
        if (isResubmitted && sub.reuploadedAt) {
            date = `${formatTimestamp(sub.reuploadedAt)} (Re-uploaded)`;
        } else if (getSubmissionReviewEntryAt(sub)) {
            date = formatTimestamp(getSubmissionReviewEntryAt(sub));
        }

        const docTypes = sub.documentTypes?.map(type => DOCUMENT_TYPES[type] || type).join(', ') || 'N/A';
        const docCount = getEffectiveSubmissionDocuments(sub).length;
        const whatsapp = renderWhatsAppLink(sub.customerDetails?.phone || sub.customerPhone || '');
        const statusBadges = `
            <span class="status-badge status-pending">Pending</span>
            ${isResubmitted ? '<span class="status-badge status-resubmitted">Resubmitted</span>' : ''}
        `;
        const reviewContext = renderReviewContextButton(sub.id, sub);

        return `
            <tr class="${isResubmitted ? 'submission-row-resubmitted' : ''}">
                <td><strong>${sub.customerName || 'Unknown'}</strong></td>
                <td>${sub.agentName || 'No Agent'}</td>
                <td>${date}</td>
                <td>${docTypes} <small class="text-muted">(${docCount})</small></td>
                <td>${whatsapp}</td>
                <td>${sub.uploadedByName || 'N/A'}</td>
                <td>${statusBadges}</td>
                <td>${reviewContext}</td>
                <td>
                    <div class="action-buttons">
                        <button class="action-btn view-btn-small" onclick="window.viewSubmission('${sub.id}')">
                            <i class="fas fa-eye"></i> View
                        </button>
                        <button class="action-btn download-all-btn" onclick="window.downloadAll('${sub.id}')" ${downloadInProgress ? 'disabled' : ''}>
                            ${downloadInProgress ? '<i class="fas fa-spinner fa-spin"></i>' : '<i class="fas fa-download"></i>'} Download
                        </button>
                        <button class="action-btn" onclick="window.openApplicationChat('${sub.id}')">
                            <i class="fas fa-comments"></i> Chat
                        </button>
                        <button class="action-btn review-btn" onclick="window.openReviewModal('${sub.id}')">
                            <i class="fas fa-check-circle"></i> Review
                        </button>
                    </div>
                </td>
            </tr>
        `;
    }).join('');
}

function renderApprovedTable(submissions) {
    if (!approvedTableBody) return;

    if (submissions.length === 0) {
        approvedTableBody.innerHTML = '<tr><td colspan="8" class="no-data">No approved documents found</td></tr>';
        return;
    }

    approvedTableBody.innerHTML = submissions.map(sub => {
        let uploadDate = 'N/A';
        if (sub.uploadedAt) {
            uploadDate = formatTimestamp(sub.uploadedAt);
        }
        
        const approvedDate = formatTimestamp(getSubmissionApprovalEntryAt(sub)) || 'N/A';
        const whatsapp = renderWhatsAppLink(sub.customerDetails?.phone || sub.customerPhone || '');

        return `
            <tr>
                <td><strong>${sub.customerName || 'Unknown'}</strong></td>
                <td>${sub.agentName || 'No Agent'}</td>
                <td>${uploadDate}</td>
                <td>${whatsapp}</td>
                <td>${sub.uploadedByName || 'N/A'}</td>
                <td>${sub.reviewedByName || 'N/A'}</td>
                <td>${approvedDate}</td>
                <td>
                    <div class="action-buttons">
                        <button class="action-btn view-btn-small" onclick="window.viewSubmission('${sub.id}')">
                            <i class="fas fa-eye"></i> View
                        </button>
                        <button class="action-btn download-all-btn" onclick="window.downloadAll('${sub.id}')" ${downloadInProgress ? 'disabled' : ''}>
                            ${downloadInProgress ? '<i class="fas fa-spinner fa-spin"></i>' : '<i class="fas fa-download"></i>'} Download
                        </button>
                        <button class="action-btn" onclick="window.openApplicationChat('${sub.id}')">
                            <i class="fas fa-comments"></i> Chat
                        </button>
                    </div>
                </td>
            </tr>
        `;
    }).join('');
}

function renderRejectedTable(submissions) {
    if (!rejectedTableBody) return;

    if (submissions.length === 0) {
        rejectedTableBody.innerHTML = '<tr><td colspan="8" class="no-data">No rejected documents found</td></tr>';
        return;
    }

    rejectedTableBody.innerHTML = submissions.map(sub => {
        const date = formatTimestamp(getSubmissionRejectionEntryAt(sub)) || 'N/A';
        
        const docTypes = sub.documentTypes?.map(type => DOCUMENT_TYPES[type] || type).join(', ') || 'N/A';
        const docCount = getEffectiveSubmissionDocuments(sub).length;
        const chatBtn = `<button class="action-btn app-chat-trigger" data-chat-submission="${sub.id}" onclick="window.openApplicationChat('${sub.id}')"><i class="fas fa-comments"></i> Chat</button>`;
        const rejectionLabel = String(sub.status || '').toLowerCase() === 'rejected_by_rsa' ? 'Rejected by RSA' : 'Rejected';
        const rejectionReason = String(sub.latestRejectionReason || sub.previousRejectionReason || sub.comment || '').trim() || 'No reason provided';

        return `
            <tr>
                <td><strong>${sub.customerName || 'Unknown'}</strong></td>
                <td>${sub.agentName || 'No Agent'}</td>
                <td>${date}</td>
                <td>${docTypes} <small class="text-muted">(${docCount})</small></td>
                <td>${chatBtn}</td>
                <td>${sub.uploadedByName || 'N/A'}</td>
                <td><strong>${rejectionLabel}:</strong> ${rejectionReason}</td>
                <td>
                    <div class="action-buttons">
                        <button class="action-btn view-btn-small" onclick="window.viewSubmission('${sub.id}')">
                            <i class="fas fa-eye"></i> View
                        </button>
                        <button class="action-btn download-all-btn" onclick="window.downloadAll('${sub.id}')" ${downloadInProgress ? 'disabled' : ''}>
                            ${downloadInProgress ? '<i class="fas fa-spinner fa-spin"></i>' : '<i class="fas fa-download"></i>'} Download
                        </button>
                    </div>
                </td>
            </tr>
        `;
    }).join('');
}

// ==================== VIEW SUBMISSION ====================
window.viewSubmission = async (submissionId) => {
    const sub = await getLatestSubmissionById(submissionId);
    const docs = getEffectiveSubmissionDocuments(sub);
    if (!sub || docs.length === 0) {
        showNotification('No documents available', 'error');
        return;
    }

    // Store current submission and reset index
    currentViewerSubmission = { ...sub, documents: docs };
    currentViewerIndex = 0;

    if (viewerDownloadBtn) {
        viewerDownloadBtn.style.display = 'inline-flex';
    }
    
    if (viewerDownloadSection) {
        viewerDownloadSection.style.display = 'none';
    }

    // Show the first document
    showDocumentAtIndex(0);
};

function showDocumentAtIndex(index) {
    if (!currentViewerSubmission || !currentViewerSubmission.documents || 
        index < 0 || index >= currentViewerSubmission.documents.length) {
        return;
    }

    const doc = currentViewerSubmission.documents[index];
    const fallbackDocType = currentViewerSubmission?.documentTypes?.[index] || '';
    const canonicalType = getCanonicalDocumentType(doc, fallbackDocType);
    const docTypeLabel = DOCUMENT_TYPES[canonicalType] || DOCUMENT_TYPES[doc.documentType] || doc.documentType || 'Document';
    
    // Update filename with counter if multiple documents
    if (currentViewerSubmission.documents.length > 1) {
        viewerFileName.textContent = `${docTypeLabel} (${index + 1}/${currentViewerSubmission.documents.length})`;
    } else {
        viewerFileName.textContent = docTypeLabel;
    }
    
    // Set document source
    documentViewer.src = buildViewerDocumentUrl(doc.fileUrl, currentViewerSubmission, index);
    
    // Update download button data
    if (viewerDownloadBtn) {
        viewerDownloadBtn.dataset.currentUrl = doc.fileUrl?.trim();
        viewerDownloadBtn.dataset.currentName = doc.name || 'document.pdf';
        viewerDownloadBtn.dataset.customerName = currentViewerSubmission.customerName || 'Customer';
        viewerDownloadBtn.dataset.docType = docTypeLabel;
    }
    
    // Update navigation buttons
    updateViewerNavigation(index);
    renderViewerContextForType(canonicalType);
    
    // Show modal and reset zoom
    viewerModal.classList.add('active');
    currentZoom = 1.0;
    applyZoom();
}

function updateViewerNavigation(currentIndex) {
    if (!viewerNav || !currentViewerSubmission) return;
    
    const totalDocs = currentViewerSubmission.documents.length;
    
    if (totalDocs <= 1) {
        viewerNav.innerHTML = '';
        return;
    }
    
    viewerNav.innerHTML = `
        <button class="nav-btn nav-btn-text" id="prevDocBtn" ${currentIndex === 0 ? 'disabled' : ''}>
            &#8249; Prev
        </button>
        <span class="nav-counter">${currentIndex + 1}/${totalDocs}</span>
        <button class="nav-btn nav-btn-text" id="nextDocBtn" ${currentIndex === totalDocs - 1 ? 'disabled' : ''}>
            Next &#8250;
        </button>
    `;
    
    const prevBtn = document.getElementById('prevDocBtn');
    const nextBtn = document.getElementById('nextDocBtn');
    
    if (prevBtn) {
        prevBtn.addEventListener('click', () => {
            if (currentViewerIndex > 0) {
                currentViewerIndex--;
                showDocumentAtIndex(currentViewerIndex);
            }
        });
    }
    
    if (nextBtn) {
        nextBtn.addEventListener('click', () => {
            if (currentViewerIndex < currentViewerSubmission.documents.length - 1) {
                currentViewerIndex++;
                showDocumentAtIndex(currentViewerIndex);
            }
        });
    }
}

// ==================== RSA ROUND-ROBIN ASSIGNMENT ====================
const RSA_COUNTER_DOC = doc(db, 'counters', 'roundRobinRSA');

function normalizeEmail(email) {
    return normalizeEmailShared(email);
}

function routingRuleDocId(uploaderEmail) {
    return routingRuleDocIdShared(uploaderEmail);
}

async function getUploaderRoutingRule(uploaderEmail) {
    return getUploaderRoutingRuleShared(db, uploaderEmail);
}

async function isActiveRSAUser(email) {
    const normalized = normalizeEmail(email);
    if (!normalized) return false;
    try {
        const data = await getUserProfileByEmailShared(db, normalized);
        if (!data) return false;
        const role = String(data.role || '').trim().toLowerCase();
        const status = String(data.status || 'active').toLowerCase();
        const leaveStatus = String(data.leaveStatus || '').toLowerCase();
        return role === 'rsa'
            && status !== 'deactivated'
            && leaveStatus !== 'on_leave'
            && data?.skipRsaRoundRobin !== true;
    } catch (_) {
        return false;
    }
}

async function getRSAEmails() {
    const snap = await getDocs(collection(db, 'users'));
    return snap.docs
        .map(d => d.data() || {})
        .filter((u) => String(u.role || '').trim().toLowerCase() === 'rsa')
        .filter((u) => String(u.status || 'active').toLowerCase() !== 'deactivated' && String(u.leaveStatus || '').toLowerCase() !== 'on_leave' && u?.skipRsaRoundRobin !== true)
        .map(u => normalizeEmail(u.email))
        .filter(Boolean)
        .sort(); // alphabetical so order is deterministic
}

async function assignRoundRobinRSA(submissionRef) {
    let uploaderEmail = '';
    try {
        const subSnap = await getDoc(submissionRef);
        if (subSnap.exists()) uploaderEmail = normalizeEmail(subSnap.data()?.uploadedBy);
    } catch (_) { }

    const routingRule = await getUploaderRoutingRule(uploaderEmail);
    const mappedRsa = routingRule?.rsaEmail || '';
    if (mappedRsa && await isActiveRSAUser(mappedRsa)) {
        await updateDoc(submissionRef, { assignedToRSA: mappedRsa, rsaAssignedAt: serverTimestamp(), rsaAssignmentMode: 'uploader_routing' });
        try {
            const subSnap = await getDoc(submissionRef);
            if (subSnap.exists()) {
                const subData = subSnap.data();
                await addDoc(collection(db, 'roundRobinAssignmentsRSA'), {
                    submissionId: submissionRef.id,
                    customerName: subData.customerName || 'N/A',
                    assignedToRSA: mappedRsa,
                    assignedBy: currentUser?.email || 'System',
                    assignedAt: serverTimestamp(),
                    reviewedBy: subData.reviewedBy || 'N/A',
                    assignmentMethod: 'uploader_routing'
                });
            }
        } catch (_) { }
        return mappedRsa;
    }

    const rsaUsers = await getRSAEmails();
    if (!rsaUsers.length) return null;
    const systemSettings = await getSystemSettings(db);
    if (!systemSettings.rsaRoundRobinEnabled) {
        await updateDoc(submissionRef, { assignedToRSA: '', rsaAssignmentMode: 'round_robin_disabled' });
        return null;
    }
    
    let assigned = null;
    let assignmentMethod = 'round_robin';
    const trustedDateKey = await getTrustedDateKey();
    try {
        await runTransaction(db, async tx => {
            let lastIndex = -1;
            const counterSnap = await tx.get(RSA_COUNTER_DOC);
            
            if (counterSnap.exists()) {
                const data = counterSnap.data();
                lastIndex = typeof data.lastIndex === 'number' ? data.lastIndex : -1;
            }
            const newIndex = (lastIndex + 1) % rsaUsers.length;
            assigned = rsaUsers[newIndex];
            
            tx.set(RSA_COUNTER_DOC, { lastIndex: newIndex, lastDate: trustedDateKey }, { merge: true });
            tx.update(submissionRef, { assignedToRSA: assigned, rsaAssignedAt: serverTimestamp(), rsaAssignmentMode: 'round_robin' });
        });
    } catch (error) {
        assigned = rsaUsers[0] || null;
        if (assigned) {
            await updateDoc(submissionRef, { assignedToRSA: assigned, rsaAssignedAt: serverTimestamp(), rsaAssignmentMode: 'round_robin_fallback' });
            assignmentMethod = 'round_robin_fallback';
        }
        console.error('RSA round-robin transaction failed; using fallback assignment.', {
            submissionId: submissionRef?.id || '',
            fallbackAssignedTo: assigned,
            rsaUsers,
            error
        });
    }
    
    // Track assignment in history collection (after transaction)
    if (assigned) {
        try {
            const subSnap = await getDoc(submissionRef);
            if (subSnap.exists()) {
                const subData = subSnap.data();
                await addDoc(collection(db, 'roundRobinAssignmentsRSA'), {
                    submissionId: submissionRef.id,
                    customerName: subData.customerName || 'N/A',
                    assignedToRSA: assigned,
                    assignedBy: currentUser?.email || 'System',
                    assignedAt: serverTimestamp(),
                    reviewedBy: subData.reviewedBy || 'N/A',
                    assignmentMethod
                });
            }
        } catch (e) {
        }
    }
    
    return assigned;
}

// ==================== OPEN REVIEW MODAL ====================
window.openReviewModal = (submissionId) => {
    const sub = allSubmissions.find(s => s.id === submissionId);
    if (!sub) return;

    currentSubmissionId = submissionId;
    modalCustomerName.textContent = sub.customerName;
    modalDocumentType.textContent = `${sub.documents?.length || 0} documents uploaded`;
    commentText.value = '';
    if (previousRejectionBox) {
        previousRejectionBox.innerHTML = '';
        previousRejectionBox.style.display = 'none';
    }
    commentModal.classList.add('active');
};

window.openReviewContextModal = (submissionId) => {
    const sub = allSubmissions.find(s => s.id === submissionId);
    if (!sub || !reviewContextModal) return;

    const entries = getRejectionHistoryEntries(sub);
    const contactValue = String(sub.customerDetails?.phone || sub.customerPhone || '-').trim() || '-';

    if (reviewContextCustomerName) {
        reviewContextCustomerName.textContent = sub.customerName || 'Unknown';
    }
    if (reviewContextContact) {
        reviewContextContact.textContent = `Contact: ${contactValue}`;
    }
    if (reviewContextHistory) {
        if (entries.length) {
            reviewContextHistory.innerHTML = `
                <ol class="review-context-list">
                    ${entries.map((entry, index) => {
                        const timeText = entry.rejectedAt ? formatTimestamp(entry.rejectedAt) : 'Time not available';
                        return `<li><strong>Rejection ${index + 1}:</strong> ${escapeHtml(entry.reason)}<span class="review-context-time">${escapeHtml(timeText)}</span></li>`;
                    }).join('')}
                </ol>
            `;
            reviewContextHistory.style.display = 'block';
        } else {
            reviewContextHistory.innerHTML = 'No rejection context available.';
            reviewContextHistory.style.display = 'block';
        }
    }

    reviewContextModal.classList.add('active');
};

// ==================== CONFIRM ACTIONS ====================
function confirmApprove() {
    const sub = allSubmissions.find(s => s.id === currentSubmissionId);
    if (!sub) return;
    
    const confirmed = confirm(`Are you sure you want to APPROVE this submission for ${sub.customerName}?\n\nThis action will assign the approved documents to an RSA user for processing.`);
    if (confirmed) {
        reviewDocument('approved');
    }
}

function confirmReject() {
    const sub = allSubmissions.find(s => s.id === currentSubmissionId);
    if (!sub) return;
    
    const confirmed = confirm(`Are you sure you want to REJECT this submission for ${sub.customerName}?\n\nMake sure to provide a detailed rejection reason in the comment field.`);
    if (confirmed) {
        reviewDocument('rejected');
    }
}

// ==================== REVIEW DOCUMENT ====================
async function reviewDocument(action) {
    if (!currentSubmissionId) return;

    const comment = commentText.value.trim();
    const currentSub = allSubmissions.find(s => s.id === currentSubmissionId);
    const customerName = currentSub?.customerName || 'Customer';
    const uploaderEmail = currentSub?.uploadedBy || '';
    const systemSettings = await getSystemSettings(db, { force: true });
    const rejectionRules = systemSettings.rejectionRules || {};
    const rolePermissions = systemSettings.rolePermissions || {};
    const minRejectLength = Number(rejectionRules.minLength || 0);
    if (action === 'approved' && rolePermissions.reviewerCanApprove === false) {
        showNotification('Reviewer approvals are currently disabled by Super Admin.', 'error');
        return;
    }
    if (action === 'rejected' && rolePermissions.reviewerCanReject === false) {
        showNotification('Reviewer rejections are currently disabled by Super Admin.', 'error');
        return;
    }

    if (action === 'rejected' && rejectionRules.reviewerRequired !== false && !comment) {
        showNotification('Please provide a reason for rejection', 'error');
        return;
    }
    if (action === 'rejected' && comment && minRejectLength > 0 && comment.length < minRejectLength) {
        showNotification(`Rejection reason must be at least ${minRejectLength} characters.`, 'error');
        return;
    }

    try {
        const submissionRef = doc(db, 'submissions', currentSubmissionId);
        
        // If approving, assign to RSA user using round-robin
        if (action === 'approved') {
            const rsaAssigned = await assignRoundRobinRSA(submissionRef);
            
            await updateDoc(submissionRef, {
                status: 'processing_to_pfa',
                comment: comment || '',
                reviewedBy: currentUser.email,
                reviewedAt: serverTimestamp(),
                reviewerDecision: 'approved',
                reviewerDecisionBy: currentUser.email,
                reviewerDecisionAt: serverTimestamp(),
                rsaReady: true
            });

            if (rsaAssigned) {
                queueRsaApprovalEmail({
                    submissionId: currentSubmissionId,
                    rsaEmail: rsaAssigned,
                    customerName,
                    reviewerEmail: currentUser?.email || '',
                    uploaderEmail
                }).catch((emailError) => {
                });
            }

            if (uploaderEmail) {
                queueUploaderApprovedEmail({
                    submissionId: currentSubmissionId,
                    uploaderEmail,
                    customerName,
                    reviewerEmail: currentUser?.email || '',
                    rsaEmail: rsaAssigned || ''
                }).catch((emailError) => {
                });
            }
            notifyStatusChangePush({
                currentUser,
                submissionId: currentSubmissionId,
                customerName,
                newStatus: 'processing_to_pfa',
                statusLabel: 'Processing to PFA',
                actionLabel: 'Application Approved',
                message: `Application for ${customerName || 'this customer'} was approved and moved to Processing to PFA.`
            }).catch(() => {});
            
            showNotification(`Document moved to Processing to PFA and assigned to RSA: ${rsaAssigned || 'pending'}`, 'success');
        } else {
            // If rejecting, don't assign to RSA
            const currentHistory = getRejectionHistoryEntries(currentSub);
            const nextHistory = [
                ...currentHistory,
                {
                    reason: comment,
                    rejectedAt: await getTrustedNowIso(),
                    rejectedBy: currentUser.email
                }
            ].filter((entry) => String(entry?.reason || '').trim());
            await updateDoc(submissionRef, {
                status: action,
                comment: comment || '',
                rejectionHistory: nextHistory,
                latestRejectionReason: comment || '',
                latestRejectedBy: currentUser.email,
                latestRejectedAt: serverTimestamp(),
                previousRejectionReason: comment || '',
                previousRejectedBy: currentUser.email,
                previousRejectedAt: serverTimestamp(),
                resubmittedAfterRejection: false,
                reviewedBy: currentUser.email,
                reviewedAt: serverTimestamp(),
                reviewerDecision: 'rejected',
                reviewerDecisionBy: currentUser.email,
                reviewerDecisionAt: serverTimestamp(),
                rsaReady: true
            });

            if (uploaderEmail) {
                queueUploaderRejectedEmail({
                    submissionId: currentSubmissionId,
                    uploaderEmail,
                    customerName,
                    reviewerEmail: currentUser?.email || '',
                    rejectionReason: comment || ''
                }).catch((emailError) => {
                });
            }
            notifyStatusChangePush({
                currentUser,
                submissionId: currentSubmissionId,
                customerName,
                newStatus: 'rejected',
                statusLabel: 'Rejected',
                actionLabel: 'Application Rejected',
                message: `Application for ${customerName || 'this customer'} was rejected and needs correction.`
            }).catch(() => {});
            
            showNotification('Document rejected successfully!', 'success');
        }
        
        closeModal(commentModal);

    } catch (error) {
        showNotification('Failed to update status: ' + error.message, 'error');
    }
}

// ==================== VIEW DOCUMENT ====================
window.viewDocument = (fileUrl, fileName, originalName = '') => {
    viewerFileName.textContent = fileName;
    documentViewer.src = fileUrl?.trim();
    
    if (viewerDownloadBtn) {
        viewerDownloadBtn.style.display = 'inline-flex';
        viewerDownloadBtn.dataset.currentUrl = fileUrl?.trim();
        viewerDownloadBtn.dataset.currentName = originalName || 'document.pdf';
    }
    
    if (viewerDownloadSection) {
        viewerDownloadSection.style.display = 'none';
    }
    
    // Clear navigation for single document view
    if (viewerNav) {
        viewerNav.innerHTML = '';
    }
    clearViewerContext();
    
    viewerModal.classList.add('active');
    currentZoom = 1.0;
    applyZoom();
};

function closeViewerModal() {
    viewerModal.classList.remove('active');
    documentViewer.src = '';
    currentZoom = 1.0;
    currentViewerSubmission = null;
    currentViewerIndex = 0;
    
    if (viewerDownloadBtn) {
        viewerDownloadBtn.style.display = 'none';
    }
    
    if (viewerDownloadSection) {
        viewerDownloadSection.style.display = 'none';
    }
    
    if (viewerNav) {
        viewerNav.innerHTML = '';
    }
    clearViewerContext();
}

// ==================== SAVE CURRENT DOCUMENT WITH LOCATION PICKER ====================
async function saveCurrentDocumentWithPicker() {
    if (!viewerDownloadBtn || !viewerDownloadBtn.dataset.currentUrl) {
        showNotification('No document to save', 'error');
        return;
    }

    const fileUrl = viewerDownloadBtn.dataset.currentUrl;
    const originalName = viewerDownloadBtn.dataset.currentName || 'document.pdf';
    const customerName = viewerDownloadBtn.dataset.customerName || 'Customer';
    const docType = viewerDownloadBtn.dataset.docType || 'Document';

    showLoader('Downloading document...');

    try {
        const response = await fetchWithCorsFallback(fileUrl);
        const blob = await response.blob();
        
        // Generate a clean filename
        const safeCustomerName = customerName.replace(/[^a-zA-Z0-9\s_-]/g, '_').trim();
        const cleanFileName = `${safeCustomerName}_${docType}.pdf`;
        
        // prompt folder and save inside customer subfolder
        await saveBlobToFolderPicker(blob, cleanFileName, customerName);
    } catch (error) {
        const safeCustomerName = customerName.replace(/[^a-zA-Z0-9\s_-]/g, '_').trim() || 'Customer';
        const cleanFileName = `${safeCustomerName}_${docType}.pdf`.replace(/[\\/:*?"<>|]/g, '_');
        if (openDirectDocumentDownload(fileUrl, cleanFileName || originalName)) {
            showNotification('Storage blocked secure save, so the document opened directly.', 'warning');
        } else {
            showNotification('Save failed: ' + error.message, 'error');
        }
    } finally {
        hideLoader();
    }
}

// ==================== ADDITIONAL FILES UPLOAD ====================
function setupAdditionalFileUpload() {
    const uploadArea = document.getElementById('additionalUploadArea');
    const fileInput = document.getElementById('additionalFileInput');
    
    if (!uploadArea || !fileInput) return;
    
    // Remove existing listeners by cloning
    const newUploadArea = uploadArea.cloneNode(true);
    uploadArea.parentNode.replaceChild(newUploadArea, uploadArea);
    
    const newFileInput = fileInput.cloneNode(true);
    fileInput.parentNode.replaceChild(newFileInput, fileInput);
    
    const updatedUploadArea = document.getElementById('additionalUploadArea');
    const updatedFileInput = document.getElementById('additionalFileInput');
    
    updatedUploadArea.addEventListener('click', () => {
        updatedFileInput.click();
    });
    
    updatedUploadArea.addEventListener('dragover', (e) => {
        e.preventDefault();
        updatedUploadArea.style.borderColor = 'var(--cm-primary)';
        updatedUploadArea.style.background = 'var(--cm-light)';
    });
    
    updatedUploadArea.addEventListener('dragleave', () => {
        updatedUploadArea.style.borderColor = '#cbd5e1';
        updatedUploadArea.style.background = 'transparent';
    });
    
    updatedUploadArea.addEventListener('drop', (e) => {
        e.preventDefault();
        updatedUploadArea.style.borderColor = '#cbd5e1';
        updatedUploadArea.style.background = 'transparent';
        
        const files = Array.from(e.dataTransfer.files);
        handleAdditionalFiles(files);
    });
    
    updatedFileInput.addEventListener('change', (e) => {
        const files = Array.from(e.target.files);
        handleAdditionalFiles(files);
    });
}

function handleAdditionalFiles(files) {
    if (files.length + additionalFiles.length > 10) {
        showNotification('Maximum 10 files total allowed', 'error');
        return;
    }
    
    // Support more image formats
    const allowedTypes = ['.pdf', '.jpg', '.jpeg', '.png', '.gif', '.bmp', '.webp'];
    const validFiles = files.filter(file => {
        const ext = '.' + file.name.split('.').pop().toLowerCase();
        return allowedTypes.includes(ext);
    });
    
    if (validFiles.length !== files.length) {
        showNotification('Some files were skipped (only PDF and common image formats allowed)', 'warning');
    }
    
    validFiles.forEach(file => {
        const fileId = `additional_${additionalFileCounter++}`;
        additionalFiles.push({
            id: fileId,
            file: file,
            name: file.name,
            type: file.type,
            size: file.size,
            isAdditional: true
        });
    });
    
    updateFilePreview();
    updateSelectedCount();
}

function updateFilePreview() {
    const previewContainer = document.getElementById('additionalFilesPreview');
    if (!previewContainer) return;
    
    if (additionalFiles.length === 0) {
        previewContainer.innerHTML = '';
        return;
    }
    
    let html = `
        <div class="preview-header">
            <span><i class="fas fa-paperclip"></i> Added Files (${additionalFiles.length})</span>
            <button class="clear-files-btn" onclick="window.clearAdditionalFiles()">
                <i class="fas fa-trash"></i> Clear All
            </button>
        </div>
    `;
    
    additionalFiles.forEach(file => {
        const fileExt = '.' + file.name.split('.').pop().toLowerCase();
        const isPdf = fileExt === '.pdf';
        const fileSize = (file.size / 1024).toFixed(1) + ' KB';
        
        html += `
            <div class="preview-item">
                <i class="fas ${isPdf ? 'fa-file-pdf' : 'fa-file-image'}" style="color: ${isPdf ? '#dc2626' : '#10b981'};"></i>
                <span class="file-name" title="${file.name}">${file.name.substring(0, 25)}${file.name.length > 25 ? '...' : ''}</span>
                <span class="file-size">${fileSize}</span>
                ${!isPdf ? '<span class="convert-badge">ÃƒÂ¢Ã¢â‚¬Â Ã¢â‚¬â„¢PDF</span>' : ''}
                <button class="remove-file-btn" onclick="window.removeAdditionalFile('${file.id}')" title="Remove">
                    <i class="fas fa-times-circle"></i>
                </button>
            </div>
        `;
    });
    
    previewContainer.innerHTML = html;
}

window.removeAdditionalFile = (fileId) => {
    additionalFiles = additionalFiles.filter(f => {
        if (f.id === fileId && f._localUrl) {
            URL.revokeObjectURL(f._localUrl);
        }
        return f.id !== fileId;
    });
    updateFilePreview();
    renumberBadges();
    updateSelectedCount();
};

window.clearAdditionalFiles = () => {
    additionalFiles.forEach(f => {
        if (f._localUrl) URL.revokeObjectURL(f._localUrl);
    });
    additionalFiles = [];
    updateFilePreview();
    const fileInput = document.getElementById('additionalFileInput');
    if (fileInput) fileInput.value = '';
    renumberBadges();
    updateSelectedCount();
};
    // Update select all checkbox
window.switchTab = (tabId, triggerEl = null) => {
    document.querySelectorAll('.nav-item').forEach(item => item.classList.remove('active'));
    const activeItem = triggerEl || document.querySelector(`.nav-item[data-tab="${tabId}"]`);
    if (activeItem) activeItem.classList.add('active');

    document.querySelectorAll('.tab-content').forEach(content => content.classList.remove('active'));
    document.getElementById(tabId + 'Tab')?.classList.add('active');

    const titles = {
        'pending': 'Pending Documents Review',
        'approved': 'Approved Documents',
        'rejected': 'Rejected Documents',
        'report': 'Reviewer Report',
        'leave': 'Leave History',
        'profile': 'My Profile',
        'help': 'Help & SOP'
    };
    if (pageTitle) pageTitle.textContent = titles[tabId] || 'Dashboard';
    if (tabId === 'report') {
        initializeReviewerReportDates();
    }
    if (tabId === 'leave') {
        renderReviewerLeaveHistory().catch(() => {});
    }
};

function getLeaveTimestampMillis(value) {
    if (!value) return 0;
    try {
        if (typeof value.toMillis === 'function') return value.toMillis();
        if (typeof value.toDate === 'function') return value.toDate().getTime();
        const d = new Date(value);
        return Number.isNaN(d.getTime()) ? 0 : d.getTime();
    } catch (_) {
        return 0;
    }
}

function buildReviewerLeaveHistoryRecords(audits, mode = 'mine') {
    const currentEmail = normalizeEmail(currentUser?.email);
    const activated = audits
        .filter((entry) => entry.action === 'user_leave_activated')
        .filter((entry) => mode === 'mine' ? normalizeEmail(entry.userEmail) === currentEmail : normalizeEmail(entry.relieverEmail) === currentEmail)
        .sort((a, b) => getLeaveTimestampMillis(b.timestamp) - getLeaveTimestampMillis(a.timestamp));
    const resumed = audits.filter((entry) => entry.action === 'user_leave_resumed');
    return activated.map((startEntry) => {
        const startMs = getLeaveTimestampMillis(startEntry.timestamp);
        const matchingResume = resumed
            .filter((entry) => normalizeEmail(entry.userEmail) === normalizeEmail(startEntry.userEmail) && String(entry.stage || '') === String(startEntry.stage || ''))
            .filter((entry) => getLeaveTimestampMillis(entry.timestamp) >= startMs)
            .sort((a, b) => getLeaveTimestampMillis(a.timestamp) - getLeaveTimestampMillis(b.timestamp))[0] || null;
        return {
            id: `${startEntry.id || startMs}-${mode}`,
            originalUserEmail: normalizeEmail(startEntry.userEmail),
            relieverEmail: normalizeEmail(startEntry.relieverEmail),
            stage: startEntry.stage || '',
            startAt: startEntry.timestamp || null,
            endAt: matchingResume?.timestamp || null,
            startAtMs: startMs,
            endAtMs: getLeaveTimestampMillis(matchingResume?.timestamp),
            movedCount: Number(startEntry.movedCount || 0),
            returnedCount: Number(matchingResume?.returnedCount || 0),
            finalizedCount: Number(matchingResume?.finalizedCount || 0),
            status: matchingResume ? 'Completed' : 'Active'
        };
    });
}

async function loadReviewerLeaveHistory() {
    const auditSnap = await getDocs(query(collection(db, 'audit'), orderBy('timestamp', 'desc')));
    const audits = auditSnap.docs.map((docSnap) => ({ id: docSnap.id, ...(docSnap.data() || {}) }));
    reviewerMyLeaveHistory = buildReviewerLeaveHistoryRecords(audits, 'mine');
    reviewerReliefLeaveHistory = buildReviewerLeaveHistoryRecords(audits, 'relief');
    reviewerLeaveHistoryLoaded = true;
}

function renderReviewerLeaveRows(records, bodyId, includeOriginalUser = false) {
    const body = document.getElementById(bodyId);
    if (!body) return;
    if (!records.length) {
        body.innerHTML = `<tr><td colspan="${includeOriginalUser ? 8 : 7}" class="no-data">No leave records found</td></tr>`;
        return;
    }
    body.innerHTML = records.map((record) => `
        <tr>
            ${includeOriginalUser ? `<td>${record.originalUserEmail || '-'}</td>` : ''}
            <td>${formatTimestamp(record.startAt)}</td>
            <td>${record.endAt ? formatTimestamp(record.endAt) : '-'}</td>
            <td>${record.status}</td>
            <td>${record.relieverEmail || '-'}</td>
            <td>${record.movedCount}</td>
            <td>${record.returnedCount}/${record.finalizedCount}</td>
            <td><button class="action-btn view-btn-small" onclick="window.openReviewerLeaveApplications('${record.id}')"><i class="fas fa-eye"></i> View</button></td>
        </tr>
    `).join('');
}

async function renderReviewerLeaveHistory() {
    if (!reviewerLeaveHistoryLoaded) {
        await loadReviewerLeaveHistory();
    }
    renderReviewerLeaveRows(reviewerMyLeaveHistory, 'reviewerMyLeaveTableBody', false);
    renderReviewerLeaveRows(reviewerReliefLeaveHistory, 'reviewerReliefLeaveTableBody', true);
}

window.openReviewerLeaveApplications = async (recordId) => {
    const record = [...reviewerMyLeaveHistory, ...reviewerReliefLeaveHistory].find((item) => item.id === recordId);
    if (!record) return;
    const submissionsSnap = await getDocs(collection(db, 'submissions'));
    const endMs = record.endAtMs || Number.MAX_SAFE_INTEGER;
    const rows = submissionsSnap.docs
        .map((docSnap) => ({ id: docSnap.id, ...(docSnap.data() || {}) }))
        .filter((sub) => normalizeEmail(sub.leaveCoverOriginalEmail) === record.originalUserEmail)
        .filter((sub) => normalizeEmail(sub.leaveCoverRelieverEmail) === record.relieverEmail || !record.relieverEmail)
        .filter((sub) => String(sub.leaveCoverStage || '') === String(record.stage || ''))
        .filter((sub) => {
            const movedMs = getLeaveTimestampMillis(sub.leaveCoverStartedAt);
            return movedMs >= record.startAtMs && movedMs <= endMs + (24 * 60 * 60 * 1000);
        });
    const body = document.getElementById('reviewerLeaveApplicationsBody');
    const title = document.getElementById('reviewerLeaveApplicationsTitle');
    if (title) title.textContent = `Leave Applications - ${record.originalUserEmail || 'User'}`;
    if (body) {
        body.innerHTML = rows.length ? rows.map((sub) => `
            <tr>
                <td>${sub.customerName || 'Unknown'}</td>
                <td>${sub.status || '-'}</td>
                <td>${formatTimestamp(sub.leaveCoverStartedAt)}</td>
                <td>${formatTimestamp(sub.leaveCoverReturnedAt)}</td>
                <td>${formatTimestamp(sub.leaveCoverFinalizedAt)}</td>
            </tr>
        `).join('') : '<tr><td colspan="5" class="no-data">No applications found for this leave record</td></tr>';
    }
    document.getElementById('reviewerLeaveApplicationsModal')?.classList.add('active');
};

// ==================== MODAL UTILITIES ====================
function closeModal(modal) {
    if (!modal) return;
    modal.classList.remove('active');
    if (modal === commentModal) {
        currentSubmissionId = null;
        commentText.value = '';
        if (previousRejectionBox) {
            previousRejectionBox.innerHTML = '';
            previousRejectionBox.style.display = 'none';
        }
    } else if (modal === reviewContextModal) {
        if (reviewContextCustomerName) reviewContextCustomerName.textContent = '-';
        if (reviewContextContact) reviewContextContact.textContent = 'Contact: -';
        if (reviewContextHistory) {
            reviewContextHistory.innerHTML = '';
            reviewContextHistory.style.display = 'none';
        }
    }
}

// ----------------- loader helpers -----------------
// functions used by both viewer and uploader dashboards
function showLoader(msg) {
    const loader = document.getElementById('globalLoader');
    const text = document.getElementById('loaderText');
    if (loader && text) {
        text.textContent = msg || 'Processing...';
        loader.style.display = 'flex';
        // small timeout for CSS transition
        setTimeout(() => loader.classList.add('active'), 10);
    }
}
function hideLoader() {
    const loader = document.getElementById('globalLoader');
    if (loader) {
        loader.classList.remove('active');
        setTimeout(() => loader.style.display = 'none', 300);
    }
}

// ==================== NOTIFICATION ====================
function showNotification(message, type = 'info') {
    if (!notification) return;

    notification.textContent = message;
    notification.className = `notification ${type}`;
    notification.style.display = 'block';
    notification.classList.add('show');

    setTimeout(() => {
        notification.classList.remove('show');
        setTimeout(() => {
            notification.style.display = 'none';
        }, 300);
    }, 3000);
}

// ==================== SIGN OUT ====================
window.signOutUser = () => {
    window.location.href = 'index.html';
};
// Merge feature removed; no merged-PDF globals.

initializeReviewerReportDates();

generateReviewerReportBtn?.addEventListener('click', async () => {
    const originalHtml = generateReviewerReportBtn.innerHTML;
    generateReviewerReportBtn.disabled = true;
    generateReviewerReportBtn.innerHTML = '<i class="fas fa-spinner fa-spin"></i> Generating...';
    try {
        currentReviewerStageReport = buildReviewerStageReport();
        renderDashboardStageReport(currentReviewerStageReport, {
            metaEl: reviewerReportMeta,
            summaryBodyEl: reviewerReportSummaryBody,
            detailsBodyEl: reviewerReportDetailsBody
        });
        showNotification('Reviewer report generated successfully.', 'success');
    } catch (error) {
        showNotification(error?.message || 'Failed to generate reviewer report.', 'error');
    } finally {
        generateReviewerReportBtn.disabled = false;
        generateReviewerReportBtn.innerHTML = originalHtml;
    }
});

exportReviewerReportBtn?.addEventListener('click', async () => {
    if (!currentReviewerStageReport) {
        showNotification('Generate a report first.', 'warning');
        return;
    }
    const originalHtml = exportReviewerReportBtn.innerHTML;
    exportReviewerReportBtn.disabled = true;
    exportReviewerReportBtn.innerHTML = '<i class="fas fa-spinner fa-spin"></i> Exporting...';
    try {
        await exportDashboardStageReportExcel(currentReviewerStageReport, 'CMBank RSA Reviewer Dashboard');
        showNotification('Reviewer report Excel downloaded.', 'success');
    } catch (error) {
        showNotification(error?.message || 'Failed to export reviewer report.', 'error');
    } finally {
        exportReviewerReportBtn.disabled = false;
        exportReviewerReportBtn.innerHTML = originalHtml;
    }
});
