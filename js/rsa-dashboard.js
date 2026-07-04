import { auth, db } from './firebase-config.js?v=20260625c';
import { getSystemSettings } from './shared/system-settings.js?v=20260617a';
import { performAppLogout } from './shared/logout.js?v=20260625b';
import {
    collection,
    addDoc,
    doc,
    getDoc,
    query,
    where,
    orderBy,
    onSnapshot,
    getDocs,
    updateDoc,
    runTransaction,
    limit,
    serverTimestamp
} from "https://www.gstatic.com/firebasejs/9.22.0/firebase-firestore.js";
import { queueUploaderFinalSubmissionEmail } from './email-alerts.js';
import { notifyStatusChangePush } from './status-push.js';
import { formatAppDate, formatAppDateTime, getTrustedDateKey } from './shared/app-time.js';
import {
    getUserProfileByEmail as getUserProfileByEmailShared,
    getCurrentUserProfile as getCurrentUserProfileShared,
    getUserFullName as getUserFullNameShared,
    normalizeEmail as normalizeEmailShared
} from './shared/user-directory.js?v=20260518a';
import {
    getTimestampMillis as getStageTimestampMillis,
    getSubmissionRsaEntryAt,
    getSubmissionRejectionEntryAt,
    getSubmissionFinalSubmissionEntryAt
} from './shared/submission-stage.js?v=20260609a';
import {
    buildDashboardStageReport,
    renderDashboardStageReport,
    exportDashboardStageReportExcel
} from './shared/dashboard-stage-report.js?v=20260610a';
import {
    getUploaderRoutingRule as getUploaderRoutingRuleShared,
    routingRuleDocId as routingRuleDocIdShared
} from './shared/uploader-routing.js?v=20260427e';

// ==================== DOCUMENT TYPES MAPPING ====================
const DOCUMENT_TYPES = {
    'birth_certificate': 'Birth Certificate / Age Declaration',
    'nin': 'National Identification Number (NIN)',
    'bvn': 'BVN',
    'pay_slips': '3 Months Pay Slip',
    'offer_letter': 'Offer of Employment Letter',
    'intro_letter': 'Introduction Letter',
    'request_letter': 'Request Letter',
    'rsa_statement': 'RSA Statement',
    'pfa_form': 'PFA Application Form',
    'consent_letter': 'Consent Letter',
    'indemnity_form': 'Indemnity Form',
    'mortgage_loan_application': 'Mortgage Loan Application Form',
    'allocation_last_page': 'Allocation Last Page',
    'offer_letter_last_page': 'Offer Letter Last Page',
    'pmi_soa': 'PMI SOA',
    'benefit_application_form': 'Benefit Application Form',
    'data_recapture': 'Data Recapture',
    'credit_life': 'Credit Life'
};

let currentUser = null;
let currentRsaProfileData = null;
let currentTab = 'approved';
let allSubmissions = [];
let currentRsaDisplayedSubmissions = [];
let currentRsaExcelResults = [];
let currentRsaExcelFilteredResults = [];
let currentRsaExcelRange = { start: '', end: '' };
let currentRsaExcelSelectedIds = new Set();
let rsaLeaveHistoryLoaded = false;
let rsaMyLeaveHistory = [];
let rsaReliefLeaveHistory = [];
const userFullNameCache = new Map();
let unsubscribeQueue = null;
let queueLoadSeq = 0;
let currentRsaStageReport = null;
const RSA_DASHBOARD_TABS = ['approved', 'rejected', 'finally-submitted', 'report', 'leave', 'profile', 'help'];

function getInitialRsaTab() {
    const hashTab = decodeURIComponent(String(window.location.hash || '').replace(/^#/, '')).trim();
    return RSA_DASHBOARD_TABS.includes(hashTab) ? hashTab : 'approved';
}

function rememberRsaTab(tabId) {
    if (!RSA_DASHBOARD_TABS.includes(tabId)) return;
    if (window.location.hash === `#${tabId}`) return;
    history.replaceState(null, '', `#${tabId}`);
}

const IDLE_TIMEOUT_MS = 30 * 60 * 1000;
let idleLastActivity = Date.now();
let idleIntervalHandle = null;

const rsaTableBody = document.getElementById('rsaTableBody');
const rsaSearch = document.getElementById('rsaSearch');
const rsaStartDate = document.getElementById('rsaStartDate');
const rsaEndDate = document.getElementById('rsaEndDate');
const rsaFilterBtn = document.getElementById('rsaFilterBtn');
const pageTitle = document.getElementById('pageTitle');
const approvedCountBadge = document.getElementById('approvedCount');
const rejectedCountBadge = document.getElementById('rejectedCount');
const userNameEl = document.getElementById('userName');
const userRoleEl = document.getElementById('userRole');
const finalSubmitBtn = document.getElementById('finalSubmitBtn');
const rsaRejectModal = document.getElementById('rsaRejectModal');
const closeRsaRejectModalBtn = document.getElementById('closeRsaRejectModal');
const cancelRsaRejectModalBtn = document.getElementById('cancelRsaRejectModal');
const confirmRsaRejectBtn = document.getElementById('confirmRsaRejectBtn');
const rsaRejectCustomerNameEl = document.getElementById('rsaRejectCustomerName');
const rsaRejectReasonInput = document.getElementById('rsaRejectReasonInput');
const profileNameEl = document.getElementById('profileName');
const profileRegisteredAtEl = document.getElementById('profileRegisteredAt');
const profileEmailEl = document.getElementById('profileEmail');
const profileWhatsappEl = document.getElementById('profileWhatsapp');
const profileLocationEl = document.getElementById('profileLocation');
const profileRoleEl = document.getElementById('profileRole');
const profileStatusEl = document.getElementById('profileStatus');
const rsaSelectAllRows = document.getElementById('rsaSelectAllRows');
const rsaExcelExportNavItem = document.getElementById('rsaExcelExportNavItem');
const approvedExcelExportBar = document.querySelector('.excel-export-bar');
const openRsaExcelDateModalBtn = document.getElementById('openRsaExcelDateModalBtn');
const rsaExcelDateModal = document.getElementById('rsaExcelDateModal');
const closeRsaExcelDateModalBtn = document.getElementById('closeRsaExcelDateModal');
const cancelRsaExcelDateModalBtn = document.getElementById('cancelRsaExcelDateModal');
const submitRsaExcelDateRangeBtn = document.getElementById('submitRsaExcelDateRangeBtn');
const rsaExcelStartDateInput = document.getElementById('rsaExcelStartDate');
const rsaExcelEndDateInput = document.getElementById('rsaExcelEndDate');
const rsaExcelResultsModal = document.getElementById('rsaExcelResultsModal');
const closeRsaExcelResultsModalBtn = document.getElementById('closeRsaExcelResultsModal');
const closeRsaExcelResultsFooterBtn = document.getElementById('closeRsaExcelResultsFooterBtn');
const rsaExcelRangeSummary = document.getElementById('rsaExcelRangeSummary');
const rsaExcelResultsCount = document.getElementById('rsaExcelResultsCount');
const rsaExcelResultsSearch = document.getElementById('rsaExcelResultsSearch');
const rsaExcelResultsTableBody = document.getElementById('rsaExcelResultsTableBody');
const rsaExcelResultsSelectAll = document.getElementById('rsaExcelResultsSelectAll');
const exportRsaLevel2AllBtn = document.getElementById('exportRsaLevel2AllBtn');
const exportRsaLevel2SelectedBtn = document.getElementById('exportRsaLevel2SelectedBtn');
const rsaReportMeta = document.getElementById('rsaReportMeta');
const rsaReportStartDate = document.getElementById('rsaReportStartDate');
const rsaReportEndDate = document.getElementById('rsaReportEndDate');
const rsaReportSummaryBody = document.getElementById('rsaReportSummaryBody');
const rsaReportDetailsBody = document.getElementById('rsaReportDetailsBody');
const generateRsaStageReportBtn = document.getElementById('generateRsaStageReportBtn');
const exportRsaStageReportBtn = document.getElementById('exportRsaStageReportBtn');

let currentDetailsSubmissionId = null;
let currentRejectSubmissionId = null;

function renderProfileTab() {
    if (!profileNameEl && !profileEmailEl && !profileRoleEl && !profileStatusEl) return;
    const fullName = currentRsaProfileData?.fullName || currentRsaProfileData?.displayName || currentUser?.displayName || currentUser?.email || 'N/A';
    const registeredAt = currentRsaProfileData?.createdAt ? formatTimestamp(currentRsaProfileData.createdAt) : '-';
    const email = currentRsaProfileData?.email || currentUser?.email || 'N/A';
    const whatsapp = currentRsaProfileData?.whatsappNumber || currentRsaProfileData?.phone || '-';
    const location = currentRsaProfileData?.location || '-';
    const role = String(currentRsaProfileData?.role || 'rsa');
    const status = String(currentRsaProfileData?.status || 'active');
    if (profileNameEl) profileNameEl.textContent = fullName;
    if (profileRegisteredAtEl) profileRegisteredAtEl.textContent = registeredAt;
    if (profileEmailEl) profileEmailEl.textContent = email;
    if (profileWhatsappEl) profileWhatsappEl.textContent = whatsapp;
    if (profileLocationEl) profileLocationEl.textContent = location;
    if (profileRoleEl) profileRoleEl.textContent = role.toUpperCase();
    if (profileStatusEl) profileStatusEl.textContent = status.charAt(0).toUpperCase() + status.slice(1);
}

// format time helper
function formatTimestamp(ts) {
    return formatAppDateTime(ts, 'N/A');
}

function normalizeWhatsAppPhone(raw) {
    const digits = String(raw || '').replace(/\D/g, '');
    if (!digits) return '';
    if (digits.startsWith('0') && digits.length === 11) return `234${digits.slice(1)}`;
    if (digits.length === 10) return `234${digits}`;
    if (digits.startsWith('234')) return digits;
    return digits;
}

function setWhatsAppContact(containerEl, rawPhone) {
    if (!containerEl) return;
    const display = String(rawPhone || '').trim();
    const normalized = normalizeWhatsAppPhone(display);
    if (!normalized) {
        containerEl.textContent = '-';
        return;
    }
    containerEl.innerHTML = `<a href="https://wa.me/${normalized}" target="_blank" rel="noopener noreferrer">${display}</a>`;
}

function getApprovedTimestamp(sub) {
    return getSubmissionRsaEntryAt(sub);
}

function normalizeEmail(email) {
    return normalizeEmailShared(email);
}

function routingRuleDocId(uploaderEmail) {
    return routingRuleDocIdShared(uploaderEmail);
}

async function getUploaderRoutingRule(uploaderEmail) {
    return getUploaderRoutingRuleShared(db, uploaderEmail);
}

async function isActivePaymentUser(email) {
    const normalized = normalizeEmail(email);
    if (!normalized) return false;
    try {
        const data = await getUserProfileByEmailShared(db, normalized);
        if (!data) return false;
        const role = String(data.role || '').toLowerCase();
        const status = String(data.status || 'active').toLowerCase();
        const leaveStatus = String(data.leaveStatus || '').toLowerCase();
        return role === 'payment' && status !== 'deactivated' && leaveStatus !== 'on_leave';
    } catch (_) {
        return false;
    }
}

async function getActivePaymentUsers() {
    const paymentQuery = query(collection(db, 'users'), where('role', '==', 'payment'));
    const snap = await getDocs(paymentQuery);
    return snap.docs
        .map((d) => d.data() || {})
        .filter((u) => String(u.status || 'active').toLowerCase() !== 'deactivated' && String(u.leaveStatus || '').toLowerCase() !== 'on_leave')
        .map((u) => normalizeEmail(u.email))
        .filter(Boolean)
        .filter((email, idx, arr) => arr.indexOf(email) === idx)
        .sort();
}

async function assignPaymentRoundRobin(submissionId) {
    const submissionRef = doc(db, 'submissions', submissionId);
    let uploaderEmail = '';
    try {
        const subSnap = await getDoc(submissionRef);
        if (subSnap.exists()) uploaderEmail = normalizeEmail(subSnap.data()?.uploadedBy);
    } catch (_) { }

    const routingRule = await getUploaderRoutingRule(uploaderEmail);
    const mappedPayment = routingRule?.paymentEmail || '';
    if (mappedPayment && await isActivePaymentUser(mappedPayment)) {
        await updateDoc(submissionRef, {
            assignedToPayment: mappedPayment,
            paymentAssignedAt: serverTimestamp(),
            paymentAssignmentMethod: 'uploader_routing'
        });
        return mappedPayment;
    }

    const paymentUsers = await getActivePaymentUsers();
    const systemSettings = await getSystemSettings(db);

    if (!paymentUsers.length) {
        await updateDoc(submissionRef, {
            assignedToPayment: '',
            paymentAssignedAt: serverTimestamp(),
            paymentAssignmentMethod: 'unassigned'
        });
        return '';
    }

    if (!systemSettings.paymentRoundRobinEnabled) {
        await updateDoc(submissionRef, {
            assignedToPayment: '',
            paymentAssignedAt: serverTimestamp(),
            paymentAssignmentMethod: 'round_robin_disabled'
        });
        return '';
    }

    const counterRef = doc(db, 'counters', 'roundRobinPayment');
    let assignedPayment = '';
    const trustedDateKey = await getTrustedDateKey();

    try {
        await runTransaction(db, async (tx) => {
            const counterSnap = await tx.get(counterRef);
            let lastIndex = -1;

            if (counterSnap.exists()) {
                const data = counterSnap.data() || {};
                lastIndex = typeof data.lastIndex === 'number' ? data.lastIndex : -1;
            }

            const nextIndex = (lastIndex + 1) % paymentUsers.length;
            assignedPayment = paymentUsers[nextIndex];

            tx.set(counterRef, {
                lastIndex: nextIndex,
                lastDate: trustedDateKey,
                lastAssignedTo: assignedPayment,
                updatedAt: serverTimestamp()
            }, { merge: true });

            tx.update(submissionRef, {
                assignedToPayment: assignedPayment,
                paymentAssignedAt: serverTimestamp(),
                paymentAssignmentMethod: 'round_robin'
            });
        });
    } catch (_) {
        assignedPayment = paymentUsers[0] || '';
        if (assignedPayment) {
            await updateDoc(submissionRef, {
                assignedToPayment: assignedPayment,
                paymentAssignedAt: serverTimestamp(),
                paymentAssignmentMethod: 'round_robin_fallback'
            });
        }
    }

    return assignedPayment;
}

async function getUserFullName(email) {
    const normalizedEmail = normalizeEmailShared(email);
    if (!normalizedEmail) return '';
    if (userFullNameCache.has(normalizedEmail)) return userFullNameCache.get(normalizedEmail);
    const fullName = await getUserFullNameShared(db, normalizedEmail);
    userFullNameCache.set(normalizedEmail, fullName);
    return fullName;
}

async function loadCurrentRsaProfile(user) {
    if (!user) return false;

    const email = String(user.email || '').trim().toLowerCase();
    let profileData = null;

    try {
        profileData = await getCurrentUserProfileShared(db, user);
        if (profileData) {
            currentRsaProfileData = profileData;
        }
    } catch (error) {
    }

    const displayName = profileData?.fullName || profileData?.displayName || user.displayName || email.split('@')[0] || 'RSA User';
    if (userNameEl) userNameEl.textContent = displayName;
    if (userRoleEl) userRoleEl.textContent = 'RSA';
    if (!currentRsaProfileData) {
        currentRsaProfileData = { email, fullName: displayName, role: 'rsa', status: 'active' };
    }
    renderProfileTab();
    updateRsaExcelAccessUI();

    const role = String(profileData?.role || '').toLowerCase();
    if (role && role !== 'rsa') {
        showNotification('Access denied. RSA privileges required.', 'error');
        setTimeout(() => { window.location.href = 'index.html'; }, 1500);
        return false;
    }
    return true;
}

function getCurrentRsaRoleLevel() {
    const rawLevel = currentRsaProfileData?.roleLevel ?? currentRsaProfileData?.accessLevel ?? 1;
    const normalized = String(rawLevel || '').trim().toLowerCase();
    if (Number(rawLevel) === 2 || normalized === '2' || normalized === 'level 2' || normalized === 'level2') {
        return 2;
    }
    return 1;
}

function isCurrentRsaLevelTwo() {
    return getCurrentRsaRoleLevel() >= 2;
}

function initializeRsaReportDates() {
    const today = new Date();
    const sixDaysAgo = new Date(today.getTime() - (6 * 24 * 60 * 60 * 1000));
    const toInputValue = (date) => `${date.getFullYear()}-${`${date.getMonth() + 1}`.padStart(2, '0')}-${`${date.getDate()}`.padStart(2, '0')}`;
    if (rsaReportStartDate && !rsaReportStartDate.value) rsaReportStartDate.value = toInputValue(sixDaysAgo);
    if (rsaReportEndDate && !rsaReportEndDate.value) rsaReportEndDate.value = toInputValue(today);
}

function resolveRsaKnownName(email) {
    const normalized = normalizeEmail(email);
    if (!normalized) return 'Unassigned';
    if (normalizeEmail(currentRsaProfileData?.email) === normalized) {
        return currentRsaProfileData?.fullName || currentRsaProfileData?.displayName || normalized;
    }
    return userFullNameCache.get(normalized) || normalized;
}

async function fetchRsaStageReportSourceRecords() {
    if (isCurrentRsaLevelTwo()) {
        const snapshot = await getDocs(query(collection(db, 'submissions'), orderBy('uploadedAt', 'desc')));
        const records = snapshot.docs
            .map((docSnap) => ({ id: docSnap.id, ...docSnap.data() }))
            .filter((sub) => normalizeEmail(sub.assignedToRSA));
        const emails = [...new Set(records.flatMap((sub) => [sub.assignedToRSA, sub.assignedTo, sub.reviewedBy, sub.uploadedBy]).filter(Boolean))];
        await Promise.all(emails.map((email) => getUserFullName(email)));
        return records;
    }
    return allSubmissions.filter((sub) => normalizeEmail(sub.assignedToRSA));
}

async function buildRsaStageReport() {
    const startDate = String(rsaReportStartDate?.value || '').trim();
    const endDate = String(rsaReportEndDate?.value || '').trim();
    if (!startDate || !endDate) throw new Error('Choose both start date and end date.');
    if (startDate > endDate) throw new Error('Start date cannot be after end date.');
    const sourceRecords = await fetchRsaStageReportSourceRecords();
    return buildDashboardStageReport({
        stageId: 'rsa',
        records: sourceRecords,
        rangeStart: startDate,
        rangeEnd: endDate,
        resolveName: resolveRsaKnownName
    });
}

function updateRsaExcelAccessUI() {
    const isLevelTwo = isCurrentRsaLevelTwo();
    if (rsaExcelExportNavItem) {
        rsaExcelExportNavItem.style.display = isLevelTwo ? '' : 'none';
    }
    if (approvedExcelExportBar) {
        approvedExcelExportBar.style.display = 'flex';
    }
    if (!isLevelTwo && currentTab === 'report') {
        switchTab('approved');
    }
}

// ------- utility helpers (borrowed from viewer dashboard) -------
async function fetchWithCorsFallback(url) {
    const cleanUrl = url?.toString().trim().replace(/[\s\n\r\t]+/g, '');
    if (!cleanUrl) throw new Error('Invalid URL');

    const proxyUrl = getBackblazeDownloadProxyUrl(cleanUrl);
    if (proxyUrl) {
        const proxyResponse = await fetch(proxyUrl, { credentials: 'same-origin' });
        if (!proxyResponse.ok) throw new Error(`Document proxy failed: ${proxyResponse.status}`);
        return proxyResponse;
    }

    try {
        const response = await fetch(cleanUrl, { mode: 'cors', credentials: 'omit' });
        if (!response.ok) throw new Error(`Document fetch failed: ${response.status}`);
        return response;
    } catch (error) {
        error.corsBlocked = true;
        throw error;
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

function openDirectDocumentDownload(fileUrl, fileName = 'document.pdf') {
    const cleanUrl = fileUrl?.toString().trim().replace(/[\s\n\r\t]+/g, '');
    if (!cleanUrl) return false;

    const link = document.createElement('a');
    link.href = getBackblazeDownloadProxyUrl(cleanUrl) || cleanUrl;
    link.target = '_blank';
    link.rel = 'noopener noreferrer';
    link.download = (fileName || 'document.pdf').replace(/[\\/:*?"<>|]/g, '_');
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    return true;
}

// PDF/image helpers
function isPdfFile(fileUrl, fileName) {
    const lowerName = (fileName||'').toLowerCase();
    const lowerUrl = (fileUrl||'').toLowerCase();
    return lowerName.endsWith('.pdf') || lowerUrl.includes('.pdf') || lowerUrl.includes('application/pdf');
}
function isImageFile(fileName) {
    const lowerName = fileName.toLowerCase();
    return lowerName.endsWith('.jpg') || lowerName.endsWith('.jpeg') || lowerName.endsWith('.png') || lowerName.endsWith('.gif') || lowerName.endsWith('.bmp') || lowerName.endsWith('.webp');
}
async function imageToPdf(imageBlob, imageName) {
    try {
        const { PDFDocument } = PDFLib;
        const pdfDoc = await PDFDocument.create();
        const page = pdfDoc.addPage([595, 842]);
        const imageBytes = await imageBlob.arrayBuffer();
        let image;
        const lowerName = imageName.toLowerCase();
        const imageType = imageBlob.type || '';
        try {
            if (lowerName.endsWith('.jpg') || lowerName.endsWith('.jpeg') || imageType.includes('jpeg')) {
                image = await pdfDoc.embedJpg(imageBytes);
            } else if (lowerName.endsWith('.png') || imageType.includes('png')) {
                image = await pdfDoc.embedPng(imageBytes);
            } else {
                const arr = new Uint8Array(imageBytes.slice(0,4));
                if (arr[0]===0xFF && arr[1]===0xD8) {
                    image = await pdfDoc.embedJpg(imageBytes);
                } else if (arr[0]===0x89 && arr[1]===0x50 && arr[2]===0x4E && arr[3]===0x47) {
                    image = await pdfDoc.embedPng(imageBytes);
                } else {
                    throw new Error(`Unsupported image format: ${imageName}`);
                }
            }
        } catch (embedError) {
            throw embedError;
        }
        const imgDims = image.scale(0.5);
        page.drawImage(image, { x: 50, y: page.getHeight() - imgDims.height - 50, width: imgDims.width, height: imgDims.height });
        const pdfBytes = await pdfDoc.save();
        return pdfBytes;
    } catch (e) {
        throw e;
    }
}

function showNotification(message, type='info') {
    const notification = document.getElementById('notification');
    if (!notification) return;
    notification.textContent = message;
    notification.className = `notification ${type}`;
    notification.style.display = 'block';
    setTimeout(()=>{ notification.style.display='none'; }, 3000);
}

function escapeHtml(value) {
    if (!value) return '';
    const div = document.createElement('div');
    div.textContent = String(value);
    return div.innerHTML;
}

function showLoader(msg) {
    const loader = document.getElementById('globalLoader');
    const text = document.getElementById('loaderText');
    if (loader && text) {
        text.textContent = msg || 'Processing...';
        loader.style.display = 'flex';
        setTimeout(()=>loader.classList.add('active'),10);
    }
}
function hideLoader() {
    const loader = document.getElementById('globalLoader');
    if (loader) {
        loader.classList.remove('active');
        setTimeout(()=>loader.style.display='none',300);
    }
}

async function saveBlobToFolderPicker(blob, defaultFileName, customerName='Customer') {
    if (!('showDirectoryPicker' in window)) {
        showNotification('Folder picker not supported. Falling back to save dialog...', 'info');
        return saveFileWithLocationPicker(blob, defaultFileName);
    }
    try {
        showNotification('📁 Please select a destination folder...', 'info');
        const dirHandle = await window.showDirectoryPicker({ mode:'readwrite', startIn:'downloads' });
        const safeCustomerName = customerName.replace(/[^a-zA-Z0-9\s_-]/g,'_').trim() || 'Customer';
        const customerFolder = await dirHandle.getDirectoryHandle(safeCustomerName, { create:true });
        const fileHandle = await customerFolder.getFileHandle(defaultFileName,{ create:true });
        const writable = await fileHandle.createWritable();
        await writable.write(blob);
        await writable.close();
        showNotification(`✅ Saved to ${safeCustomerName}/${defaultFileName}`,'success');
        return true;
    } catch (error) {
        if (error.name === 'AbortError') {
            showNotification('Save cancelled','info');
        } else {
            showNotification('Save failed: '+error.message,'error');
            await saveFileWithLocationPicker(blob, defaultFileName);
        }
        return false;
    }
}

async function saveFileWithLocationPicker(blob, defaultFileName) {
    if (!('showSaveFilePicker' in window)) {
        triggerDirectDownload(blob, defaultFileName);
        return true;
    }

    try {
        const extension = String(defaultFileName || '').includes('.')
            ? `.${String(defaultFileName).split('.').pop().toLowerCase()}`
            : '.bin';
        const fileHandle = await window.showSaveFilePicker({
            suggestedName: defaultFileName,
            types: [{
                description: 'Download',
                accept: {
                    [blob.type || 'application/octet-stream']: [extension]
                }
            }]
        });
        const writable = await fileHandle.createWritable();
        await writable.write(blob);
        await writable.close();
        showNotification(`Saved: ${defaultFileName}`, 'success');
        return true;
    } catch (error) {
        if (error.name === 'AbortError') {
            showNotification('Save cancelled', 'info');
        } else {
            showNotification('Save failed: ' + error.message, 'error');
            triggerDirectDownload(blob, defaultFileName);
        }
        return false;
    }
}

function triggerDirectDownload(blob, fileName) {
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = fileName;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(url);
    showNotification('✅ Download started','success');
}

async function downloadBlobAsFile(blob, fileName) {
    return saveFileWithLocationPicker(blob, fileName);
}

function parseMoneyValue(value) {
    const raw = String(value ?? '').replace(/[^0-9.\-]/g, '');
    const num = Number(raw);
    return Number.isFinite(num) ? num : '';
}

function formatMoneyForExcel(value) {
    const num = Number(value);
    return Number.isFinite(num) ? num.toFixed(2) : '';
}

function roundDownToThousand(value) {
    const num = Number(value);
    if (!Number.isFinite(num)) return '';
    return Math.floor(num / 1000) * 1000;
}

function roundUpToThousand(value) {
    const num = Number(value);
    if (!Number.isFinite(num)) return '';
    return Math.ceil(num / 1000) * 1000;
}

function toTitleCaseWords(value) {
    const text = String(value ?? '').trim().toLowerCase();
    if (!text) return '';
    return text.replace(/\b([a-z])([a-z']*)/g, (_, first, rest) => `${first.toUpperCase()}${rest}`);
}

function toPenCodeCase(value) {
    const text = String(value ?? '').trim().toLowerCase();
    if (!text) return '';
    return text.replace(/[a-z]/, (match) => match.toUpperCase());
}

function getYearOfBirth(dob) {
    const value = String(dob || '').trim();
    if (!value) return '';
    const date = new Date(value);
    if (!Number.isNaN(date.getTime())) return String(date.getFullYear());
    const match = value.match(/\b(19|20)\d{2}\b/);
    return match ? match[0] : '';
}

function getRsaExcelRows(submissions) {
    return submissions.map((sub, index) => {
        const details = sub.customerDetails || {};
        const propertyValue = parseMoneyValue(details.propertyValue);
        const loanAmount = roundUpToThousand(parseMoneyValue(details.loanAmount));
        const rsaBalance = parseMoneyValue(details.rsaBalance);
        const expected25 = Number.isFinite(Number(rsaBalance)) ? Number(rsaBalance) * 0.25 : '';
        const equityValue = expected25 === '' ? '' : roundDownToThousand(expected25);
        const houseNumber = details.houseNumber || sub.houseNumber || '';
        const tenor = details.tenor || sub.tenor || '';

        return [
            index + 1,
            details.accountNo || '',
            toTitleCaseWords(sub.customerName || details.name || ''),
            toTitleCaseWords(details.originatingTP || ''),
            toTitleCaseWords(details.pfa || ''),
            toPenCodeCase(details.penNo || ''),
            details.rsaStatementDate || '',
            rsaBalance === '' ? '' : Number(formatMoneyForExcel(rsaBalance)),
            expected25 === '' ? '' : Number(formatMoneyForExcel(expected25)),
            details.propertyType || '',
            toTitleCaseWords(details.originatingTP || ''),
            houseNumber,
            tenor,
            propertyValue === '' ? '' : Number(formatMoneyForExcel(propertyValue)),
            equityValue === '' ? '' : Number(formatMoneyForExcel(equityValue)),
            loanAmount === '' ? '' : Number(formatMoneyForExcel(loanAmount)),
            toTitleCaseWords(details.address || ''),
            getYearOfBirth(details.dob),
            details.mortgageLoanApplicationFormDate || ''
        ];
    });
}

async function downloadRsaExcel(submissions, scopeLabel = 'selected') {
    if (!submissions.length) {
        showNotification('No Processing to PFA applications available for Excel export', 'warning');
        return;
    }
    if (!window.ExcelJS) {
        showNotification('Excel export library is not available. Please refresh and try again.', 'error');
        return;
    }

    const headers = [
        'S/N',
        'ACCOUNT NUMBER',
        'Full Name (Surname First)',
        'Originating Contact Centre',
        'Name of PFA',
        'PEN CODE NUMBER',
        'RSA STATEMENT DATE',
        'RSA BALANCE (NGN)',
        'EXPECTED 25 %RSA (NGN)',
        'HOUSE TYPE',
        'LOCATION',
        'HOUSE NUMBER',
        'TENOR',
        'PROPERTY VALUE',
        'EQUITY VALUE',
        'LOAN AMOUNT',
        'CUSTOMER ADDRESS',
        'YEAR OF BIRTH',
        'MORTGAGE LOAN APPLICATION FORM DATE'
    ];
    const rows = getRsaExcelRows(submissions);
    const workbook = new window.ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('RSA Export');
    worksheet.addRow(headers);
    rows.forEach((row) => worksheet.addRow(row));

    worksheet.columns = [
        { width: 8 },
        { width: 18 },
        { width: 30 },
        { width: 28 },
        { width: 24 },
        { width: 20 },
        { width: 18 },
        { width: 18 },
        { width: 20 },
        { width: 22 },
        { width: 24 },
        { width: 16 },
        { width: 12 },
        { width: 18 },
        { width: 18 },
        { width: 18 },
        { width: 36 },
        { width: 14 },
        { width: 28 }
    ];

    worksheet.eachRow((row, rowNumber) => {
        row.eachCell((cell, colNumber) => {
            cell.font = { name: 'Calibri', size: 11, bold: rowNumber === 1, color: { argb: 'FF000000' } };
            cell.alignment = { vertical: 'middle', wrapText: true };
            cell.border = {
                top: { style: 'thin', color: { argb: 'FF999999' } },
                left: { style: 'thin', color: { argb: 'FF999999' } },
                bottom: { style: 'thin', color: { argb: 'FF999999' } },
                right: { style: 'thin', color: { argb: 'FF999999' } }
            };
            if (rowNumber === 1) {
                cell.fill = {
                    type: 'pattern',
                    pattern: 'solid',
                    fgColor: { argb: 'FFD9EAD3' }
                };
            }
            const numberColumns = new Set([1, 8, 9, 14, 15, 16, 18]);
            if (rowNumber > 1 && numberColumns.has(colNumber) && typeof cell.value === 'number') {
                cell.numFmt = colNumber === 1 || colNumber === 18 ? '0' : '0.00';
            }
        });
    });

    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], {
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    });
    const stamp = await getTrustedDateKey();
    await saveFileWithLocationPicker(blob, `rsa_processing_to_pfa_${scopeLabel}_${stamp}.xlsx`);
}

async function getAllRsaExcelSubmissions() {
    const snap = await getDocs(query(collection(db, 'submissions'), orderBy('uploadedAt', 'desc')));
    const next = snap.docs.map((d) => ({ id: d.id, ...(d.data() || {}) }));
    const filtered = next.filter((sub) => {
        const status = String(sub.status || '').toLowerCase();
        return (status === 'processing_to_pfa' || status === 'approved') && !sub.finalSubmitted && !sub.rsaSubmitted;
    });
    const reviewerEmails = [...new Set(filtered.map((s) => s.reviewedBy).filter(Boolean))];
    const uploaderEmails = [...new Set(filtered.map((s) => s.uploadedBy).filter(Boolean))];
    await Promise.all([...reviewerEmails, ...uploaderEmails].map((email) => getUserFullName(email)));
    return filtered.map((sub) => ({
        ...sub,
        reviewedByName: sub.reviewedBy ? (userFullNameCache.get(String(sub.reviewedBy).toLowerCase()) || sub.reviewedBy) : '-',
        uploadedByName: sub.uploadedBy ? (userFullNameCache.get(String(sub.uploadedBy).toLowerCase()) || sub.uploadedBy) : '-'
    }));
}

function getSubmissionRsaEntryDateMillis(submission) {
    return getStageTimestampMillis(getSubmissionRsaEntryAt(submission));
}

function formatDateRangeLabel(startValue, endValue) {
    const toLabel = (value) => {
        if (!value) return 'Any date';
        return formatAppDate(`${value}T00:00:00`, value);
    };
    return `${toLabel(startValue)} to ${toLabel(endValue)}`;
}

function closeRsaExcelDateModal() {
    rsaExcelDateModal?.classList.remove('active');
}

function openRsaExcelDateModal() {
    if (!isCurrentRsaLevelTwo()) {
        showNotification('This export flow is available only to RSA Level 2.', 'warning');
        return;
    }
    if (rsaExcelStartDateInput && rsaExcelEndDateInput && (!rsaExcelStartDateInput.value || !rsaExcelEndDateInput.value)) {
        const end = new Date();
        const start = new Date();
        start.setDate(end.getDate() - 30);
        rsaExcelEndDateInput.value = end.toISOString().slice(0, 10);
        rsaExcelStartDateInput.value = start.toISOString().slice(0, 10);
    }
    rsaExcelDateModal?.classList.add('active');
}

function closeRsaExcelResultsModal() {
    rsaExcelResultsModal?.classList.remove('active');
    currentRsaExcelResults = [];
    currentRsaExcelFilteredResults = [];
    currentRsaExcelRange = { start: '', end: '' };
    currentRsaExcelSelectedIds = new Set();
    if (rsaExcelResultsSearch) rsaExcelResultsSearch.value = '';
    if (rsaExcelResultsSelectAll) {
        rsaExcelResultsSelectAll.checked = false;
        rsaExcelResultsSelectAll.indeterminate = false;
    }
}

function getSelectedRsaExcelResults() {
    const selectedIds = Array.from(currentRsaExcelSelectedIds);
    return currentRsaExcelResults.filter((sub) => selectedIds.includes(sub.id));
}

function updateRsaExcelResultsSelectAllState() {
    if (!rsaExcelResultsSelectAll) return;
    const visibleRows = currentRsaExcelFilteredResults;
    const checkedVisibleCount = visibleRows.filter((sub) => currentRsaExcelSelectedIds.has(sub.id)).length;
    rsaExcelResultsSelectAll.checked = visibleRows.length > 0 && checkedVisibleCount === visibleRows.length;
    rsaExcelResultsSelectAll.indeterminate = checkedVisibleCount > 0 && checkedVisibleCount < visibleRows.length;
}

function getRsaExcelResultsSearchText(sub) {
    const status = String(sub.status || '').toLowerCase();
    const statusLabel = (status === 'processing_to_pfa' || status === 'approved') ? 'processing to pfa' : String(sub.status || '');
    return [
        sub.customerName,
        sub.agentName,
        sub.uploadedByName,
        sub.uploadedBy,
        sub.reviewedByName,
        sub.reviewedBy,
        statusLabel
    ].map((value) => String(value || '').toLowerCase()).join(' ');
}

function syncRsaExcelFilteredResults() {
    const queryText = String(rsaExcelResultsSearch?.value || '').trim().toLowerCase();
    currentRsaExcelFilteredResults = !queryText
        ? [...currentRsaExcelResults]
        : currentRsaExcelResults.filter((sub) => getRsaExcelResultsSearchText(sub).includes(queryText));
}

function renderRsaExcelResultsModal() {
    if (!rsaExcelResultsTableBody) return;
    syncRsaExcelFilteredResults();
    if (rsaExcelRangeSummary) {
        rsaExcelRangeSummary.textContent = formatDateRangeLabel(currentRsaExcelRange.start, currentRsaExcelRange.end);
    }
    if (rsaExcelResultsCount) {
        const total = currentRsaExcelResults.length;
        const visible = currentRsaExcelFilteredResults.length;
        rsaExcelResultsCount.textContent = visible === total
            ? `${total} application${total === 1 ? '' : 's'} found`
            : `${visible} of ${total} application${total === 1 ? '' : 's'} shown`;
    }
    if (rsaExcelResultsSelectAll) {
        rsaExcelResultsSelectAll.checked = false;
        rsaExcelResultsSelectAll.indeterminate = false;
    }
    if (!currentRsaExcelFilteredResults.length) {
        rsaExcelResultsTableBody.innerHTML = '<tr><td colspan="8" style="text-align:center;padding:36px;color:#666">No applications found for this date range</td></tr>';
        return;
    }
    rsaExcelResultsTableBody.innerHTML = currentRsaExcelFilteredResults.map((sub) => {
        const reviewer = sub.reviewedByName || sub.reviewedBy || '-';
        const uploader = sub.uploadedByName || sub.uploadedBy || '-';
        const status = String(sub.status || '').toLowerCase();
        const statusLabel = (status === 'processing_to_pfa' || status === 'approved') ? 'Processing to PFA' : (sub.status || '-');
        return `
            <tr>
                <td><input type="checkbox" class="rsa-export-results-select" value="${sub.id}" ${currentRsaExcelSelectedIds.has(sub.id) ? 'checked' : ''} title="Select ${sub.customerName || 'customer'}"></td>
                <td><strong>${sub.customerName || '-'}</strong></td>
                <td>${sub.agentName || 'No Agent'}</td>
                <td>${uploader}</td>
                <td>${reviewer}</td>
                <td>${formatTimestamp(sub.uploadedAt)}</td>
                <td>${formatTimestamp(getApprovedTimestamp(sub))}</td>
                <td>${statusLabel}</td>
            </tr>
        `;
    }).join('');
    updateRsaExcelResultsSelectAllState();
}

async function handleRsaLevelTwoDateRangeSubmit() {
    if (!isCurrentRsaLevelTwo()) {
        showNotification('This export flow is available only to RSA Level 2.', 'warning');
        return;
    }
    const startValue = rsaExcelStartDateInput?.value || '';
    const endValue = rsaExcelEndDateInput?.value || '';
    if (!startValue || !endValue) {
        showNotification('Please choose both start and end dates.', 'warning');
        return;
    }
    if (new Date(`${startValue}T00:00:00`).getTime() > new Date(`${endValue}T23:59:59`).getTime()) {
        showNotification('Start date cannot be after end date.', 'warning');
        return;
    }

    try {
        showLoader('Loading RSA applications for export...');
        const submissions = await getAllRsaExcelSubmissions();
        const startMs = new Date(`${startValue}T00:00:00`).getTime();
        const endMs = new Date(`${endValue}T23:59:59`).getTime();
        currentRsaExcelResults = submissions.filter((sub) => {
            const entryAtMs = getSubmissionRsaEntryDateMillis(sub);
            return entryAtMs >= startMs && entryAtMs <= endMs;
        }).sort((a, b) => getSubmissionRsaEntryDateMillis(b) - getSubmissionRsaEntryDateMillis(a));
        currentRsaExcelSelectedIds = new Set();
        if (rsaExcelResultsSearch) rsaExcelResultsSearch.value = '';
        currentRsaExcelRange = { start: startValue, end: endValue };
        renderRsaExcelResultsModal();
        closeRsaExcelDateModal();
        rsaExcelResultsModal?.classList.add('active');
    } catch (error) {
        showNotification('Failed to load RSA export applications.', 'error');
    } finally {
        hideLoader();
    }
}

function getSelectedRsaSubmissions() {
    const selectedIds = Array.from(document.querySelectorAll('.rsa-export-select:checked'))
        .map((input) => input.value)
        .filter(Boolean);
    return currentRsaDisplayedSubmissions.filter((sub) => selectedIds.includes(sub.id));
}

function updateRsaSelectAllState() {
    if (!rsaSelectAllRows) return;
    const rowChecks = Array.from(document.querySelectorAll('.rsa-export-select'));
    const checked = rowChecks.filter((input) => input.checked);
    rsaSelectAllRows.checked = rowChecks.length > 0 && checked.length === rowChecks.length;
    rsaSelectAllRows.indeterminate = checked.length > 0 && checked.length < rowChecks.length;
}

// ==================== MERGE FUNCTIONS COPY START ====================
let currentMergeSubmission = null;
let mergeSelectionOrder = [];
let additionalFiles = [];
let additionalFileCounter = 0;
let mergedPdfBlob = null;
let mergedPdfUrl = null;
let mergeInProgress = false;

async function getLatestSubmissionById(submissionId) {
    const normalizedId = String(submissionId || '').trim();
    if (!normalizedId) return null;

    const cached = allSubmissions.find((s) => s.id === normalizedId) || null;

    try {
        const snap = await getDoc(doc(db, 'submissions', normalizedId));
        if (!snap.exists()) return cached;

        const fresh = { id: snap.id, ...snap.data() };
        const index = allSubmissions.findIndex((s) => s.id === normalizedId);
        if (index >= 0) {
            allSubmissions[index] = fresh;
        } else {
            allSubmissions.push(fresh);
        }
        return fresh;
    } catch (_) {
        return cached;
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

window.openMergeModal = async (submissionId) => {
    const sub = await getLatestSubmissionById(submissionId);
    if (!sub) return;
    currentMergeSubmission = sub;
    mergeSelectionOrder = [];
    additionalFiles = [];
    additionalFileCounter = 0;
    mergedPdfBlob = null;
    if (mergedPdfUrl) { URL.revokeObjectURL(mergedPdfUrl); mergedPdfUrl = null; }
    const mergeDownloadSection = document.getElementById('mergeDownloadSection');
    if (mergeDownloadSection) mergeDownloadSection.style.display='none';
    document.getElementById('mergeCustomerName').textContent = sub.customerName;
    const effectiveDocs = getEffectiveSubmissionDocuments(sub);
    if (!effectiveDocs.length) {
        showNotification('No documents available to merge','error');
        return;
    }
    refreshMergeList();
    const previewContainer = document.getElementById('additionalFilesPreview');
    if (previewContainer) previewContainer.innerHTML='';
    const fileInput = document.getElementById('additionalFileInput');
    if (fileInput) fileInput.value='';
    setTimeout(setupAdditionalFileUpload,100);
    document.getElementById('mergeModal').classList.add('active');
    updateSelectedCount();
};

window.handleMergeCheckbox = (checkbox,itemId) => {
    if (checkbox.checked) {
        if (!mergeSelectionOrder.includes(itemId)) mergeSelectionOrder.push(itemId);
    } else {
        mergeSelectionOrder = mergeSelectionOrder.filter(id=>id!==itemId);
    }
    renumberBadges();
    updateSelectedCount();
    const selectAllCheckbox = document.getElementById('selectAllDocuments');
    if (selectAllCheckbox) {
        const totalItems = document.querySelectorAll('.merge-item input[type="checkbox"]').length;
        const checkedItems = mergeSelectionOrder.length;
        selectAllCheckbox.checked = checkedItems === totalItems && totalItems > 0;
        selectAllCheckbox.indeterminate = checkedItems>0 && checkedItems<totalItems;
    }
};

window.selectAllDocuments = (checkbox) => {
    const checkboxes = document.querySelectorAll('.merge-item input[type="checkbox"]');
    mergeSelectionOrder = [];
    if (checkbox.checked) {
        checkboxes.forEach(cb=>{ cb.checked=true; mergeSelectionOrder.push(cb.value); });
    } else {
        checkboxes.forEach(cb=>{ cb.checked=false; });
    }
    renumberBadges();
    updateSelectedCount();
};

function updateSelectedCount() {
    const selectedCount = document.getElementById('selectedCount');
    if (selectedCount) selectedCount.textContent = mergeSelectionOrder.length;
}

function renumberBadges() {
    mergeSelectionOrder.forEach((itemId,order)=>{
        const badge = document.getElementById(`order_${itemId}`);
        if (badge) { badge.textContent = `#${order+1}`; badge.style.display='inline-block'; }
    });
    document.querySelectorAll('.merge-order-badge').forEach(badge=>{
        if (!mergeSelectionOrder.includes(badge.id.replace('order_',''))) {
            badge.textContent=''; badge.style.display='none';
        }
    });
}

function refreshMergeList() {
    if (!currentMergeSubmission) return;
    const originalDocs = getEffectiveSubmissionDocuments(currentMergeSubmission);
    let combinedHtml = `
        <div class="merge-list-header">
            <div class="select-all">
                <input type="checkbox" id="selectAllDocuments" onchange="window.selectAllDocuments(this)">
                <label for="selectAllDocuments">Select All</label>
            </div>
            <div class="merge-instructions">
                <i class="fas fa-info-circle"></i>
                <span>Check documents to merge. Order will be preserved (#1, #2, #3...)</span>
            </div>
        </div>
        <div class="merge-items-container">
    `;
    originalDocs.forEach((doc,idx)=>{
        const itemId = `orig_${idx}`;
        const isChecked = mergeSelectionOrder.includes(itemId);
        const isPdf = isPdfFile(doc.fileUrl, doc.name);
        combinedHtml += `
            <div class="merge-item" data-url="${doc.fileUrl}" data-name="${doc.name}" id="item_${itemId}">
                <input type="checkbox" value="${itemId}" ${isChecked?'checked':''} onchange="window.handleMergeCheckbox(this,'${itemId}')">
                <span class="merge-order-badge" id="order_${itemId}" style="display: ${isChecked?'inline-block':'none'};">${isChecked?`#${mergeSelectionOrder.indexOf(itemId)+1}`:''}</span>
                <div class="merge-thumb"></div>
                <div class="merge-info">
                    <div>${DOCUMENT_TYPES[doc.documentType]||doc.documentType||'Document'}</div>
                    <div class="doc-filename">${doc.name}</div>
                </div>
            </div>
        `;
    });
    additionalFiles.forEach(f=>{
        const itemId = f.id;
        const isChecked = mergeSelectionOrder.includes(itemId);
        const isPdf = isPdfFile('', f.name);
        combinedHtml += `
            <div class="merge-item additional" data-url="" data-name="${f.name}" id="item_${itemId}">
                <input type="checkbox" value="${itemId}" ${isChecked?'checked':''} onchange="window.handleMergeCheckbox(this,'${itemId}')">
                <span class="merge-order-badge" id="order_${itemId}" style="display: ${isChecked?'inline-block':'none'};">${isChecked?`#${mergeSelectionOrder.indexOf(itemId)+1}`:''}</span>
                <div class="merge-thumb"></div>
                <div class="merge-info">
                    <div>${f.name}</div>
                    <div class="doc-filename">(additional)</div>
                </div>
                <button class="remove-item-btn" onclick="window.removeAdditionalFile('${itemId}')" title="Remove">
                    <i class="fas fa-times-circle"></i>
                </button>
            </div>
        `;
    });
    combinedHtml += '</div>';
    const mergeListEl = document.getElementById('mergeList');
    if (mergeListEl) mergeListEl.innerHTML = combinedHtml;
    const selectAllCheckbox = document.getElementById('selectAllDocuments');
    if (selectAllCheckbox) {
        const totalItems = originalDocs.length + additionalFiles.length;
        const checkedItems = mergeSelectionOrder.length;
        selectAllCheckbox.checked = checkedItems===totalItems && totalItems>0;
        selectAllCheckbox.indeterminate = checkedItems>0 && checkedItems<totalItems;
    }
    renderMergeThumbnails();
}

async function renderMergeThumbnails() {
    // Thumbnails removed (pdfThumbnail was not included in RSA dashboard build).
    // Keep placeholder empty to avoid runtime errors.
    return;
}

function setupAdditionalFileUpload() {
    const uploadArea = document.getElementById('additionalUploadArea');
    const fileInput = document.getElementById('additionalFileInput');
    if (!uploadArea || !fileInput) return;
    const newUploadArea = uploadArea.cloneNode(true);
    uploadArea.parentNode.replaceChild(newUploadArea, uploadArea);
    const newFileInput = fileInput.cloneNode(true);
    fileInput.parentNode.replaceChild(newFileInput, fileInput);
    const updatedUploadArea = document.getElementById('additionalUploadArea');
    const updatedFileInput = document.getElementById('additionalFileInput');
    updatedUploadArea.addEventListener('click', () => { updatedFileInput.click(); });
    updatedUploadArea.addEventListener('dragover', (e)=>{
        e.preventDefault();
        updatedUploadArea.style.borderColor='var(--cm-primary)';
        updatedUploadArea.style.background='var(--cm-light)';
    });
    updatedUploadArea.addEventListener('dragleave', ()=>{
        updatedUploadArea.style.borderColor='#cbd5e1';
        updatedUploadArea.style.background='transparent';
    });
    updatedUploadArea.addEventListener('drop', (e)=>{
        e.preventDefault();
        updatedUploadArea.style.borderColor='#cbd5e1';
        updatedUploadArea.style.background='transparent';
        const files = Array.from(e.dataTransfer.files);
        handleAdditionalFiles(files);
    });
    updatedFileInput.addEventListener('change', (e)=>{
        const files = Array.from(e.target.files);
        handleAdditionalFiles(files);
    });
}

function handleAdditionalFiles(files) {
    if (files.length + additionalFiles.length > 10) {
        showNotification('Maximum 10 files total allowed','error');
        return;
    }
    const allowedTypes = ['.pdf','.jpg','.jpeg','.png','.gif','.bmp','.webp'];
    const validFiles = files.filter(file => {
        const ext = '.' + file.name.split('.').pop().toLowerCase();
        return allowedTypes.includes(ext);
    });
    if (validFiles.length !== files.length) {
        showNotification('Some files were skipped (only PDF and common image formats allowed)','warning');
    }
    validFiles.forEach(file=>{
        const fileId = `additional_${additionalFileCounter++}`;
        additionalFiles.push({ id:fileId, file:file, name:file.name, type:file.type, size:file.size, isAdditional:true });
    });
    refreshMergeList();
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
    previewContainer.innerHTML = additionalFiles.map(f=>{
        return `<div class="file-preview" id="preview_${f.id}">
            <span>${f.name}</span>
            <button class="remove-file" onclick="window.removeAdditionalFile('${f.id}')"><i class="fas fa-times-circle"></i></button>
        </div>`;
    }).join('');
}

window.removeAdditionalFile = (id) => {
    additionalFiles = additionalFiles.filter(f=>f.id!==id);
    refreshMergeList();
    updateFilePreview();
    mergeSelectionOrder = mergeSelectionOrder.filter(i=>i!==id);
    updateSelectedCount();
};

async function handleMergeDocuments() {
    const originalDocs = getEffectiveSubmissionDocuments(currentMergeSubmission);
    if (mergeSelectionOrder.length === 0) {
        showNotification('Please select documents to merge','error');
        return;
    }
    mergeInProgress = true;
    const mergeDocumentsBtn = document.getElementById('mergeDocumentsBtn');
    if (mergeDocumentsBtn) {
        mergeDocumentsBtn.disabled = true;
        mergeDocumentsBtn.innerHTML = '<i class="fas fa-spinner fa-spin"></i> Merging...';
    }
    const mergeDownloadSection = document.getElementById('mergeDownloadSection');
    if (mergeDownloadSection) mergeDownloadSection.style.display='none';
    try {
        const { PDFDocument } = PDFLib;
        const mergedPdf = await PDFDocument.create();
        const totalDocs = mergeSelectionOrder.length;
        let processedCount = 0;
        let successCount = 0;
        for (const itemId of mergeSelectionOrder) {
            processedCount++;
            try {
                if (itemId.startsWith('orig_')) {
                    const index = parseInt(itemId.replace('orig_',''));
                    const doc = originalDocs[index];
                    if (!doc || !doc.fileUrl) {
                        showNotification(`⚠️ Skipping document ${processedCount}: Invalid document`,'warning');
                        continue;
                    }
                    showNotification(`📄 Processing ${processedCount}/${totalDocs}: ${DOCUMENT_TYPES[doc.documentType]||'Document'}...`,'info');
                    let response;
                    try {
                        response = await fetchWithCorsFallback(doc.fileUrl);
                    } catch (fetchError) {
                        showNotification(`⚠️ Could not fetch document ${processedCount}, skipping...`,'warning');
                        continue;
                    }
                    if (!response.ok) {
                        showNotification(`⚠️ Failed to fetch document ${processedCount} (${response.status}), skipping...`,'warning');
                        continue;
                    }
                    const blob = await response.blob();
                    const fileName = doc.name || 'file.pdf';
                    if (blob.size === 0) {
                        showNotification(`⚠️ Document ${processedCount} is empty, skipping...`,'warning');
                        continue;
                    }
                    const arrayBuffer = await blob.arrayBuffer();
                    const isPdf = blob.type.includes('pdf') || fileName.toLowerCase().endsWith('.pdf') || (arrayBuffer.byteLength>4 && new Uint8Array(arrayBuffer.slice(0,4))[0]===0x25);
                    if (isPdf) {
                        try {
                            const pdf = await PDFDocument.load(arrayBuffer,{ignoreEncryption:true,throwOnInvalidObject:false});
                            const copiedPages = await mergedPdf.copyPages(pdf,pdf.getPageIndices());
                            copiedPages.forEach(page=>mergedPdf.addPage(page));
                            successCount++;
                        } catch (pdfError) {
                            if (isImageFile(fileName)) {
                                showNotification(`🔄 Attempting to convert as image...`,'info');
                                try {
                                    const pdfBytes = await imageToPdf(blob,fileName);
                                    const pdf = await PDFDocument.load(pdfBytes);
                                    const copiedPages = await mergedPdf.copyPages(pdf,pdf.getPageIndices());
                                    copiedPages.forEach(page=>mergedPdf.addPage(page));
                                    successCount++;
                                } catch (convertError) {
                                    throw new Error(`Failed to convert: ${convertError.message}`);
                                }
                            } else {
                                throw new Error(`Invalid PDF file: ${doc.name||'document'}`);
                            }
                        }
                    } else if (isImageFile(fileName)) {
                        showNotification(`🔄 Converting image to PDF...`,'info');
                        try {
                            const pdfBytes = await imageToPdf(blob,fileName);
                            const pdf = await PDFDocument.load(pdfBytes);
                            const copiedPages = await mergedPdf.copyPages(pdf,pdf.getPageIndices());
                            copiedPages.forEach(page=>mergedPdf.addPage(page));
                            successCount++;
                        } catch (convertError) {
                            throw new Error(`Failed to convert image: ${convertError.message}`);
                        }
                    } else {
                        showNotification(`⚠️ Unsupported file format: ${fileName}, skipping...`,'warning');
                    }
                } else {
                    const fileObj = additionalFiles.find(f=>f.id===itemId);
                    if (!fileObj) {
                        showNotification(`⚠️ Skipping file ${processedCount}: File not found`,'warning');
                        continue;
                    }
                    const file = fileObj.file;
                    showNotification(`📄 Processing ${processedCount}/${totalDocs}: ${file.name}...`,'info');
                    const arrayBuffer = await file.arrayBuffer();
                    const fileName = file.name.toLowerCase();
                    const isPdf = file.type.includes('pdf') || fileName.endsWith('.pdf') || (arrayBuffer.byteLength>4 && new Uint8Array(arrayBuffer.slice(0,4))[0]===0x25);
                    if (isPdf) {
                        try {
                            const pdf = await PDFDocument.load(arrayBuffer,{ignoreEncryption:true,throwOnInvalidObject:false});
                            const copiedPages = await mergedPdf.copyPages(pdf,pdf.getPageIndices());
                            copiedPages.forEach(page=>mergedPdf.addPage(page));
                            successCount++;
                        } catch (pdfError) {
                            throw new Error(`Invalid PDF file: ${file.name}`);
                        }
                    } else if (fileName.match(/\.(jpg|jpeg|png|gif|bmp|webp)$/)) {
                        showNotification(`🔄 Converting image to PDF: ${file.name}...`,'info');
                        try {
                            const pdfBytes = await imageToPdf(file,file.name);
                            const pdf = await PDFDocument.load(pdfBytes);
                            const copiedPages = await mergedPdf.copyPages(pdf,pdf.getPageIndices());
                            copiedPages.forEach(page=>mergedPdf.addPage(page));
                            successCount++;
                        } catch (convertError) {
                            throw new Error(`Failed to convert image: ${convertError.message}`);
                        }
                    } else {
                        showNotification(`⚠️ Unsupported file format: ${file.name}, skipping...`,'warning');
                    }
                }
            } catch (itemError) {
                showNotification(`⚠️ Error processing item ${processedCount}: ${itemError.message}`,'warning');
            }
        }
        if (mergedPdf.getPageCount() === 0) {
            throw new Error('No valid pages were merged. Please check your files.');
        }
        showNotification(`✅ Successfully merged ${successCount} out of ${totalDocs} documents (${mergedPdf.getPageCount()} pages)`,'success');
        showNotification('⏳ Finalizing merged document...','info');
        const mergedPdfBytes = await mergedPdf.save();
        mergedPdfBlob = new Blob([mergedPdfBytes], { type: 'application/pdf' });
        if (mergedPdfUrl) URL.revokeObjectURL(mergedPdfUrl);
        mergedPdfUrl = URL.createObjectURL(mergedPdfBlob);
        if (mergeDownloadSection) {
            mergeDownloadSection.style.display='block';
            mergeDownloadSection.innerHTML = `
                <div class="merge-success-message">
                    <i class="fas fa-check-circle"></i>
                    <span>PDF merged successfully! (${mergedPdf.getPageCount()} pages from ${successCount} documents)</span>
                </div>
                <div class="merge-action-buttons">
                    <button class="action-btn view-merged-btn" id="viewMergedPdf">
                        <i class="fas fa-eye"></i> View Merged PDF
                    </button>
                    <button class="action-btn download-merged-btn" id="saveMergedPdf">
                        <i class="fas fa-save"></i> Save PDF
                    </button>
                </div>
            `;
            document.getElementById('viewMergedPdf')?.addEventListener('click', viewMergedPDF);
            document.getElementById('saveMergedPdf')?.addEventListener('click', saveMergedPDFWithPicker);
        }
    } catch (error) {
        showNotification('Failed to merge: ' + error.message,'error');
    } finally {
        mergeInProgress = false;
        if (mergeDocumentsBtn) {
            mergeDocumentsBtn.disabled = false;
            mergeDocumentsBtn.innerHTML = '<i class="fas fa-compress-alt"></i> Merge Selected';
        }
    }
}

async function saveMergedPDFWithPicker() {
    if (!mergedPdfBlob) {
        showNotification('No merged PDF available','error');
        return;
    }
    const customerName = currentMergeSubmission?.customerName || 'Customer';
    const timestamp = new Date().toISOString().slice(0,10).replace(/-/g,'');
    const fileName = `${customerName.replace(/[^a-zA-Z0-9]/g,'_')}_Merged_${timestamp}.pdf`;
    showLoader('Preparing merged PDF...');
    try { await saveBlobToFolderPicker(mergedPdfBlob, fileName, customerName); }
    finally { hideLoader(); }
}

function viewMergedPDF() {
    if (!mergedPdfBlob || !mergedPdfUrl) {
        showNotification('No merged PDF available','error');
        return;
    }
    const customerName = currentMergeSubmission?.customerName || 'Customer';
    const viewerFileName = document.getElementById('viewerFileName');
    const documentViewer = document.getElementById('documentViewer');
    if (viewerFileName) viewerFileName.textContent = `${customerName} - Merged Documents`;
    if (documentViewer) documentViewer.src = mergedPdfUrl;
    const viewerModal = document.getElementById('viewerModal');
    if (viewerModal) viewerModal.classList.add('active');
    const viewerDownloadBtn = document.getElementById('viewerDownloadBtn');
    if (viewerDownloadBtn) viewerDownloadBtn.style.display='none';
    // Keep merge modal open when viewing merged PDF
    // closeMergeModalFunc();
}

function closeMergeModalFunc() {
    const mergeModal = document.getElementById('mergeModal');
    if (mergeModal) mergeModal.classList.remove('active');
    currentMergeSubmission = null;
    mergeSelectionOrder=[];
    additionalFiles=[];
    additionalFileCounter=0;
    mergedPdfBlob=null;
    if (mergedPdfUrl) { URL.revokeObjectURL(mergedPdfUrl); mergedPdfUrl=null; }
    const mergeDownloadSection = document.getElementById('mergeDownloadSection');
    if (mergeDownloadSection) mergeDownloadSection.style.display='none';
    updateSelectedCount();
}

// ==================== MERGE FUNCTIONS COPY END ====================

function renderRows(submissions) {
    if (!rsaTableBody) return;
    if (submissions.length === 0) {
        rsaTableBody.innerHTML = '<tr><td colspan="5" style="text-align:center;padding:40px;color:#666">No records</td></tr>';
        return;
    }
    rsaTableBody.innerHTML = submissions.map(sub=>{
        const docs = (sub.documents||[]).length;
        const viewer = sub.reviewedBy || '-';
        const status = sub.status || '';
        return `<tr>
            <td>${sub.customerName || '-'}</td>
            <td>${viewer}</td>
            <td>${formatTimestamp(sub.uploadedAt)}</td>
            <td>${status.charAt(0).toUpperCase()+status.slice(1)}</td>
            <td>${docs}</td>
        </tr>`;
    }).join('');
}

function loadQueue() {
    // We need to fetch both rsaReady==true records AND rsaSubmitted==true records (old structure)
    // Firestore doesn't support OR in where clauses, so we'll fetch all assigned records and filter in memory
    let q = query(
        collection(db,'submissions'),
        where('assignedToRSA','==', currentUser?.email || ''),
        orderBy('uploadedAt','desc')
    );

    if (typeof unsubscribeQueue === 'function') {
        try { unsubscribeQueue(); } catch (e) { /* ignore */ }
    }

    unsubscribeQueue = onSnapshot(q, async (snap) => {
        const seq = ++queueLoadSeq;
        // Filter for records that are either rsaReady=true OR rsaSubmitted=true (old structure)
        const next = snap.docs
            .map(d => ({ id: d.id, ...d.data() }))
            .filter(s => s.rsaReady === true || s.rsaSubmitted === true || String(s.status || '').toLowerCase() === 'rejected_by_rsa');

        const reviewerEmails = [...new Set(next.map(s => s.reviewedBy).filter(Boolean))];
        const uploaderEmails = [...new Set(next.map(s => s.uploadedBy).filter(Boolean))];
        await Promise.all([...reviewerEmails, ...uploaderEmails].map(email => getUserFullName(email)));
        if (seq !== queueLoadSeq) return;

        allSubmissions = next.map(sub => ({
            ...sub,
            reviewedByName: sub.reviewedBy ? (userFullNameCache.get(String(sub.reviewedBy).toLowerCase()) || sub.reviewedBy) : '-',
            uploadedByName: sub.uploadedBy ? (userFullNameCache.get(String(sub.uploadedBy).toLowerCase()) || sub.uploadedBy) : '-'
        }));
        updateApprovedCount();
        renderCurrentTab();
    }, (err) => {
        showNotification('Could not load approved applications: ' + err.message, 'error');
    });
}

function updateApprovedCount() {
    // Exclude both old and new finally submitted records from approved count
    const cnt = allSubmissions.filter(s => (String(s.status || '').toLowerCase() === 'processing_to_pfa' || String(s.status || '').toLowerCase() === 'approved') && !s.finalSubmitted && !s.rsaSubmitted).length;
    if (approvedCountBadge) {
        approvedCountBadge.textContent = cnt;
        approvedCountBadge.style.display = 'inline-block';
    }
    const rejectedCnt = allSubmissions.filter((s) => String(s.status || '').toLowerCase() === 'rejected_by_rsa').length;
    if (rejectedCountBadge) {
        rejectedCountBadge.textContent = rejectedCnt;
        rejectedCountBadge.style.display = 'inline-block';
    }
    updateFinallySubmittedCount();
}

function switchTab(tabId) {
    tabId = RSA_DASHBOARD_TABS.includes(tabId) ? tabId : 'approved';
    currentTab = tabId;
    rememberRsaTab(tabId);
    document.querySelectorAll('.nav-item').forEach(nav=>nav.classList.remove('active'));
    document.querySelector(`[data-tab="${tabId}"]`)?.classList.add('active');
    document.querySelectorAll('.tab-content').forEach(tab=>tab.classList.remove('active'));
    const targetTab = document.getElementById(`${tabId}Tab`);
    targetTab?.classList.add('active');
    const titles = {
        approved: 'Processing to PFA',
        rejected: 'Rejected by RSA',
        'finally-submitted': 'Finally Submitted Applications',
        report: 'RSA Report',
        profile: 'My Profile',
        help: 'Help & SOP'
    };
    if (pageTitle) pageTitle.textContent = titles[tabId] || 'RSA Queue';
    if (tabId === 'report') {
        initializeRsaReportDates();
    }
    renderCurrentTab();
}

function renderCurrentTab() {
    if (currentTab === 'approved') {
        const start = rsaStartDate.value? new Date(rsaStartDate.value) : null;
        const end = rsaEndDate.value? new Date(rsaEndDate.value) : null;
        if (end) end.setHours(23,59,59,999);
        let list = allSubmissions.filter(s=>{
            const entryAt = getSubmissionRsaEntryDateMillis(s);
            if (start && entryAt < start.getTime()) return false;
            if (end && entryAt > end.getTime()) return false;
            return true;
        });
        if (rsaSearch && rsaSearch.value.trim()) {
            const qstr = rsaSearch.value.trim().toLowerCase();
            list = list.filter(s => {
                return (s.customerName||'').toLowerCase().includes(qstr) ||
                       (s.reviewedBy||'').toLowerCase().includes(qstr) ||
                       (s.reviewedByName||'').toLowerCase().includes(qstr) ||
                       (s.uploadedBy||'').toLowerCase().includes(qstr) ||
                       (s.uploadedByName||'').toLowerCase().includes(qstr);
            });
        }
        // Exclude both old and new finally submitted records
        list = list.filter(s => {
            const status = String(s.status || '').toLowerCase();
            return (status === 'processing_to_pfa' || status === 'approved') && !s.finalSubmitted && !s.rsaSubmitted;
        });
        list = list.slice().sort((a, b) => getSubmissionRsaEntryDateMillis(b) - getSubmissionRsaEntryDateMillis(a));
        currentRsaDisplayedSubmissions = list;
        renderRsaRows(list);
    } else if (currentTab === 'finally-submitted') {
        renderFinallySubmittedTab();
    } else if (currentTab === 'rejected') {
        renderRejectedRsaTab();
    } else if (currentTab === 'leave') {
        renderRsaLeaveHistory();
    }
}

function getTimestampMillis(value) {
    if (!value) return 0;
    try {
        if (typeof value.toMillis === 'function') return value.toMillis();
        if (typeof value.toDate === 'function') return value.toDate().getTime();
        const date = new Date(value);
        return Number.isNaN(date.getTime()) ? 0 : date.getTime();
    } catch (_) {
        return 0;
    }
}

function buildLeaveHistoryRecords(audits, mode = 'mine') {
    const currentEmail = normalizeEmail(currentUser?.email);
    const activated = audits
        .filter((entry) => entry.action === 'user_leave_activated')
        .filter((entry) => mode === 'mine'
            ? normalizeEmail(entry.userEmail) === currentEmail
            : normalizeEmail(entry.relieverEmail) === currentEmail)
        .sort((a, b) => getTimestampMillis(b.timestamp) - getTimestampMillis(a.timestamp));

    const resumed = audits.filter((entry) => entry.action === 'user_leave_resumed');
    return activated.map((startEntry) => {
        const startMs = getTimestampMillis(startEntry.timestamp);
        const matchingResume = resumed
            .filter((entry) => normalizeEmail(entry.userEmail) === normalizeEmail(startEntry.userEmail) && String(entry.stage || '') === String(startEntry.stage || ''))
            .filter((entry) => getTimestampMillis(entry.timestamp) >= startMs)
            .sort((a, b) => getTimestampMillis(a.timestamp) - getTimestampMillis(b.timestamp))[0] || null;
        return {
            id: `${startEntry.id || startMs}-${mode}`,
            originalUserEmail: normalizeEmail(startEntry.userEmail),
            relieverEmail: normalizeEmail(startEntry.relieverEmail),
            stage: startEntry.stage || '',
            startAt: startEntry.timestamp || null,
            endAt: matchingResume?.timestamp || null,
            startAtMs: startMs,
            endAtMs: getTimestampMillis(matchingResume?.timestamp),
            movedCount: Number(startEntry.movedCount || 0),
            returnedCount: Number(matchingResume?.returnedCount || 0),
            finalizedCount: Number(matchingResume?.finalizedCount || 0),
            status: matchingResume ? 'Completed' : 'Active',
            activatedBy: startEntry.performedBy || '',
            resumedBy: matchingResume?.performedBy || ''
        };
    });
}

async function loadRsaLeaveHistory() {
    const [auditSnap] = await Promise.all([
        getDocs(query(collection(db, 'audit'), orderBy('timestamp', 'desc'), limit(500)))
    ]);
    const audits = auditSnap.docs.map((docSnap) => ({ id: docSnap.id, ...(docSnap.data() || {}) }));
    rsaMyLeaveHistory = buildLeaveHistoryRecords(audits, 'mine');
    rsaReliefLeaveHistory = buildLeaveHistoryRecords(audits, 'relief');
    rsaLeaveHistoryLoaded = true;
}

function renderRsaLeaveRows(records, bodyId, includeOriginalUser = false) {
    const body = document.getElementById(bodyId);
    if (!body) return;
    if (!records.length) {
        body.innerHTML = `<tr><td colspan="${includeOriginalUser ? 8 : 7}" style="text-align:center;padding:32px;color:#666">No leave records found</td></tr>`;
        return;
    }
    body.innerHTML = records.map((record) => `
        <tr>
            ${includeOriginalUser ? `<td>${escapeHtml(userFullNameCache.get(record.originalUserEmail) || record.originalUserEmail || '-')}</td>` : ''}
            <td>${formatTimestamp(record.startAt)}</td>
            <td>${record.endAt ? formatTimestamp(record.endAt) : '-'}</td>
            <td>${escapeHtml(record.status)}</td>
            <td>${escapeHtml(userFullNameCache.get(record.relieverEmail) || record.relieverEmail || '-')}</td>
            <td>${record.movedCount}</td>
            <td>${record.returnedCount}/${record.finalizedCount}</td>
            <td><button class="action-btn" onclick="window.openRsaLeaveApplications('${record.id}')"><i class="fas fa-eye"></i> View</button></td>
        </tr>
    `).join('');
}

async function renderRsaLeaveHistory() {
    if (!rsaLeaveHistoryLoaded) {
        await loadRsaLeaveHistory();
    }
    renderRsaLeaveRows(rsaMyLeaveHistory, 'rsaMyLeaveTableBody', false);
    renderRsaLeaveRows(rsaReliefLeaveHistory, 'rsaReliefLeaveTableBody', true);
}

window.openRsaLeaveApplications = async (recordId) => {
    const record = [...rsaMyLeaveHistory, ...rsaReliefLeaveHistory].find((item) => item.id === recordId);
    if (!record) return;
    const submissionsSnap = await getDocs(query(collection(db, 'submissions'), where('leaveCoverOriginalEmail', '==', record.originalUserEmail)));
    const startMs = record.startAtMs || 0;
    const endMs = record.endAtMs || Number.MAX_SAFE_INTEGER;
    const matches = submissionsSnap.docs
        .map((docSnap) => ({ id: docSnap.id, ...(docSnap.data() || {}) }))
        .filter((sub) => normalizeEmail(sub.leaveCoverOriginalEmail) === record.originalUserEmail)
        .filter((sub) => normalizeEmail(sub.leaveCoverRelieverEmail) === record.relieverEmail || !record.relieverEmail)
        .filter((sub) => String(sub.leaveCoverStage || '') === String(record.stage || ''))
        .filter((sub) => {
            const movedMs = getTimestampMillis(sub.leaveCoverStartedAt);
            return movedMs >= startMs && movedMs <= endMs + (24 * 60 * 60 * 1000);
        });

    const body = document.getElementById('rsaLeaveApplicationsBody');
    const title = document.getElementById('rsaLeaveApplicationsTitle');
    if (title) title.textContent = `Leave Applications - ${userFullNameCache.get(record.originalUserEmail) || record.originalUserEmail || 'User'}`;
    if (body) {
        body.innerHTML = matches.length
            ? matches.map((sub) => `
                <tr>
                    <td>${escapeHtml(sub.customerName || 'Unknown')}</td>
                    <td>${escapeHtml(String(sub.status || '-'))}</td>
                    <td>${formatTimestamp(sub.leaveCoverStartedAt)}</td>
                    <td>${formatTimestamp(sub.leaveCoverReturnedAt)}</td>
                    <td>${formatTimestamp(sub.leaveCoverFinalizedAt)}</td>
                </tr>
            `).join('')
            : '<tr><td colspan="5" style="text-align:center;padding:24px;color:#666">No applications found for this leave record</td></tr>';
    }
    document.getElementById('rsaLeaveApplicationsModal')?.classList.add('active');
};

function renderRejectedRsaTab() {
    const body = document.getElementById('rsaRejectedTableBody');
    if (!body) return;

    let list = allSubmissions
        .filter((s) => String(s.status || '').toLowerCase() === 'rejected_by_rsa')
        .slice()
        .sort((a, b) => getStageTimestampMillis(getSubmissionRejectionEntryAt(b)) - getStageTimestampMillis(getSubmissionRejectionEntryAt(a)));
    if (rsaSearch && rsaSearch.value.trim()) {
        const qstr = rsaSearch.value.trim().toLowerCase();
        list = list.filter((s) =>
            (s.customerName || '').toLowerCase().includes(qstr) ||
            (s.uploadedBy || '').toLowerCase().includes(qstr) ||
            (s.uploadedByName || '').toLowerCase().includes(qstr)
        );
    }

    if (!list.length) {
        body.innerHTML = '<tr><td colspan="9" style="text-align:center;padding:40px;color:#666">No RSA-rejected applications</td></tr>';
        return;
    }

    body.innerHTML = list.map((sub) => `
        <tr>
            <td><strong>${sub.customerName || '-'}</strong></td>
            <td>${sub.agentName || 'No Agent'}</td>
            <td>${sub.uploadedByName || sub.uploadedBy || '-'}</td>
            <td>${formatTimestamp(sub.uploadedAt)}</td>
            <td>${formatTimestamp(sub.latestRejectedAt || sub.reviewedAt)}</td>
            <td>${sub.latestRejectedBy || currentUser?.email || '-'}</td>
            <td>${sub.latestRejectionReason || sub.comment || '-'}</td>
            <td><span class="status-badge status-rejected">Rejected by RSA</span></td>
            <td>
                <button class="action-btn" onclick="window.openCustomerDetails('${sub.id}')"><i class="fas fa-eye"></i> Details</button>
                <button class="action-btn" onclick="window.downloadAllRsa('${sub.id}')"><i class="fas fa-download"></i> Download All</button>
                <button class="action-btn" onclick="window.openApplicationChat('${sub.id}')"><i class="fas fa-comments"></i> Chat</button>
            </td>
        </tr>
    `).join('');
}

function renderFinallySubmittedTab() {
    const finallySubmittedTableBody = document.getElementById('finallySubmittedTableBody');
    if (!finallySubmittedTableBody) {
        return;
    }

    // Support both old structure (rsaSubmitted=true) and new structure (finalSubmitted=true)
    let list = allSubmissions.filter(s => {
        return s.finalSubmitted === true || s.rsaSubmitted === true;
    });
    list = list.slice().sort((a, b) => getStageTimestampMillis(getSubmissionFinalSubmissionEntryAt(b)) - getStageTimestampMillis(getSubmissionFinalSubmissionEntryAt(a)));

    const finalStartDate = document.getElementById('finalStartDate')?.value || '';
    const finalEndDate = document.getElementById('finalEndDate')?.value || '';
    const startMs = finalStartDate ? new Date(`${finalStartDate}T00:00:00`).getTime() : 0;
    const endMs = finalEndDate ? new Date(`${finalEndDate}T23:59:59`).getTime() : 0;
    if (startMs || endMs) {
        list = list.filter((sub) => {
            const entryMs = getStageTimestampMillis(getSubmissionFinalSubmissionEntryAt(sub));
            if (!entryMs) return false;
            if (startMs && entryMs < startMs) return false;
            if (endMs && entryMs > endMs) return false;
            return true;
        });
    }

    // Apply filters
    const finalSearch = document.getElementById('finalSearch');
    if (finalSearch && finalSearch.value.trim()) {
        const qstr = finalSearch.value.trim().toLowerCase();
        list = list.filter(s => {
            return (s.customerName||'').toLowerCase().includes(qstr) ||
                   (s.reviewedBy||'').toLowerCase().includes(qstr) ||
                   (s.reviewedByName||'').toLowerCase().includes(qstr) ||
                   (s.uploadedBy||'').toLowerCase().includes(qstr) ||
                   (s.uploadedByName||'').toLowerCase().includes(qstr);
        });
    }

    if (list.length === 0) {
        finallySubmittedTableBody.innerHTML = '<tr><td colspan="7" style="text-align:center;padding:40px;color:#666">No finally submitted applications</td></tr>';
        return;
    }

    const htmlRows = list.map(sub => {
        const reviewer = sub.reviewedByName || sub.reviewedBy || '-';
        const uploader = sub.uploadedByName || sub.uploadedBy || '-';
        // Support both old (rsaSubmittedBy) and new (finalSubmittedBy) structure
        const rsaOfficer = sub.finalSubmittedBy || sub.rsaSubmittedBy || currentUser?.email || '-';
        const uploadedAt = formatTimestamp(sub.uploadedAt);
        const approvedAt = formatTimestamp(sub.reviewedAt);
        const submittedAt = formatTimestamp(sub.finalSubmittedAt || sub.rsaSubmittedAt || sub.uploadedAt);

        return `<tr>
            <td><strong>${sub.customerName || 'Unknown'}</strong></td>
            <td>${sub.agentName || 'No Agent'}</td>
            <td>${uploader}</td>
            <td>${reviewer}</td>
            <td>${rsaOfficer}</td>
            <td>${uploadedAt}</td>
            <td>${approvedAt}</td>
            <td>${submittedAt}</td>
            <td><span class="status-badge status-approved">Sent to PFA</span></td>
            <td>
                <button class="action-btn" onclick="window.openCustomerDetails('${sub.id}')"><i class="fas fa-eye"></i> Details</button>
                <button class="action-btn" onclick="window.downloadAllRsa('${sub.id}')"><i class="fas fa-download"></i> Download All</button>
                <button class="action-btn merge-btn" onclick="window.openMergeModal('${sub.id}')"><i class="fas fa-compress-alt"></i> Merge</button>
                <button class="action-btn" onclick="window.openApplicationChat('${sub.id}')"><i class="fas fa-comments"></i> Chat</button>
            </td>
        </tr>`;
    });

    const joinedHtml = htmlRows.join('');
    finallySubmittedTableBody.innerHTML = joinedHtml;

    updateFinallySubmittedCount();
}

function updateFinallySubmittedCount() {
    // Support both old structure (rsaSubmitted=true) and new structure (finalSubmitted=true)
    const cnt = allSubmissions.filter(s => s.finalSubmitted === true || s.rsaSubmitted === true).length;
    const badge = document.getElementById('finallySubmittedCount');
    if (badge) {
        badge.textContent = cnt;
        badge.style.display = 'inline-block';
    }
}

function renderRsaRows(submissions) {
    if (!rsaTableBody) return;
    currentRsaDisplayedSubmissions = submissions;
    if (rsaSelectAllRows) {
        rsaSelectAllRows.checked = false;
        rsaSelectAllRows.indeterminate = false;
    }
    if (submissions.length === 0) {
        rsaTableBody.innerHTML = '<tr><td colspan="10" style="text-align:center;padding:40px;color:#666">No records</td></tr>';
        return;
    }
    rsaTableBody.innerHTML = submissions.map(sub=>{
        const docs = (sub.documents||[]).length;
        const reviewer = sub.reviewedByName || sub.reviewedBy || '-';
        const uploader = sub.uploadedByName || sub.uploadedBy || '-';
        const approvedAt = formatTimestamp(getApprovedTimestamp(sub));
        const status = String(sub.status || '');
        const statusLabel = (status.toLowerCase() === 'processing_to_pfa' || status.toLowerCase() === 'approved')
            ? 'Processing to PFA'
            : (status.charAt(0).toUpperCase() + status.slice(1));
        const actions = `
            <button class="action-btn" onclick="window.openCustomerDetails('${sub.id}')"><i class="fas fa-eye"></i> Details</button>
            <button class="action-btn" onclick="window.downloadAllRsa('${sub.id}')"><i class="fas fa-download"></i> Download All</button>
            <button class="action-btn merge-btn" onclick="window.openMergeModal('${sub.id}')"><i class="fas fa-compress-alt"></i> Merge</button>
            <button class="action-btn" onclick="window.openApplicationChat('${sub.id}')"><i class="fas fa-comments"></i> Chat</button>
            <button class="action-btn" style="background:#dc2626;color:#fff;border:none;" onclick="window.rejectRsaSubmission('${sub.id}')"><i class="fas fa-times-circle"></i> Reject</button>
            <button class="action-btn" style="background:#10b981;color:#fff;border:none;" onclick="window.finalSubmitRsa('${sub.id}')"><i class="fas fa-paper-plane"></i> Final Submit</button>
        `;
        return `<tr>
            <td><input type="checkbox" class="rsa-export-select" value="${sub.id}" title="Select ${sub.customerName || 'customer'} for Excel"></td>
            <td>${sub.customerName||'-'}</td>
            <td>${sub.agentName || 'No Agent'}</td>
            <td>${uploader}</td>
            <td>${reviewer}</td>
            <td>${formatTimestamp(sub.uploadedAt)}</td>
            <td>${approvedAt}</td>
            <td>${statusLabel}</td>
            <td>${docs}</td>
            <td>${actions}</td>
        </tr>`;
    }).join('');
}

window.finalSubmitRsa = async (submissionId) => {
    if (typeof window.assertAppWritable === 'function' && !window.assertAppWritable('RSA final submission')) return;
    const sub = allSubmissions.find(s => s.id === submissionId);
    if (!sub) return;
    const rolePermissions = (await getSystemSettings(db, { force: true })).rolePermissions || {};
    if (rolePermissions.rsaCanApprove === false) {
        showNotification('RSA final submissions are currently disabled by Super Admin.', 'error');
        return;
    }

    const currentStatus = String(sub.status || '').toLowerCase();
    if (!(currentStatus === 'processing_to_pfa' || currentStatus === 'approved')) {
        showNotification('Only applications in Processing to PFA can be finally submitted.', 'warning');
        return;
    }

    const customerName = sub.customerName || 'this customer';
    const confirmed = confirm(`Final submit application for ${customerName}?\n\nThis will remove it from your RSA queue.`);
    if (!confirmed) return;

    try {
        showLoader('Submitting final application...');
        await updateDoc(doc(db, 'submissions', submissionId), {
            status: 'sent_to_pfa',
            rsaReady: true,
            rsaSubmitted: true,
            rsaSubmittedAt: serverTimestamp(),
            rsaSubmittedBy: currentUser?.email || '',
            finalSubmitted: true,
            finalSubmittedAt: serverTimestamp(),
            finalSubmittedBy: currentUser?.email || ''
        });
        const assignedPayment = await assignPaymentRoundRobin(submissionId);
        if (assignedPayment) {
            try {
                await addDoc(collection(db, 'roundRobinAssignmentsPayment'), {
                    submissionId,
                    customerName: sub.customerName || '',
                    assignedToPayment: assignedPayment,
                    assignedBy: currentUser?.email || '',
                    assignedAt: serverTimestamp()
                });
            } catch (_) {}
        }

        if (sub.uploadedBy) {
            queueUploaderFinalSubmissionEmail({
                submissionId,
                uploaderEmail: sub.uploadedBy,
                customerName: sub.customerName || '',
                rsaEmail: currentUser?.email || ''
            }).catch(() => {});
        }
        notifyStatusChangePush({
            currentUser,
            submissionId,
            customerName: sub.customerName || '',
            newStatus: 'sent_to_pfa',
            statusLabel: 'Sent to PFA',
            actionLabel: 'Application Sent to PFA',
            message: `Application for ${sub.customerName || 'this customer'} was finally submitted and sent to PFA.`
        }).catch(() => {});

        // Update the submission in allSubmissions array to reflect the change
        const subIndex = allSubmissions.findIndex(s => s.id === submissionId);
        if (subIndex >= 0) {
            allSubmissions[subIndex] = {
                ...allSubmissions[subIndex],
                status: 'sent_to_pfa',
                rsaSubmitted: true,
                rsaSubmittedAt: new Date(),
                rsaSubmittedBy: currentUser?.email || '',
                finalSubmitted: true,
                finalSubmittedAt: new Date(),
                finalSubmittedBy: currentUser?.email || '',
                assignedToPayment: assignedPayment || ''
            };
        }

        updateApprovedCount();
        updateFinallySubmittedCount();
        renderCurrentTab();
        closeCustomerDetailsModal();

        showNotification(`Final submission recorded. Status set to Sent to PFA${assignedPayment ? ` and assigned to ${assignedPayment}` : ''}.`, 'success');
    } catch (error) {
        showNotification('Final submission failed: ' + (error?.message || 'Unknown error'), 'error');
    } finally {
        hideLoader();
    }
};


rsaFilterBtn?.addEventListener('click', renderCurrentTab);

// search triggers
rsaSearch && rsaSearch.addEventListener('input', renderCurrentTab);
rsaSelectAllRows?.addEventListener('change', () => {
    document.querySelectorAll('.rsa-export-select').forEach((input) => {
        input.checked = rsaSelectAllRows.checked;
    });
    updateRsaSelectAllState();
});
rsaTableBody?.addEventListener('change', (e) => {
    if (e.target?.classList?.contains('rsa-export-select')) updateRsaSelectAllState();
});
document.querySelectorAll('.export-rsa-all-excel').forEach((btn) => btn.addEventListener('click', async () => {
    try {
        const submissions = isCurrentRsaLevelTwo()
            ? await getAllRsaExcelSubmissions()
            : currentRsaDisplayedSubmissions;
        if (!submissions.length) {
            showNotification('No RSA records available for Excel export', 'info');
            return;
        }
        downloadRsaExcel(submissions, isCurrentRsaLevelTwo() ? 'all_rsa_users' : 'all');
    } catch (error) {
        showNotification('Failed to prepare RSA Excel export', 'error');
    }
}));
document.querySelectorAll('.export-rsa-selected-excel').forEach((btn) => btn.addEventListener('click', () => {
    const selected = getSelectedRsaSubmissions();
    if (!selected.length) {
        showNotification('Select at least one customer first', 'warning');
        return;
    }
    downloadRsaExcel(selected, 'selected');
}));
openRsaExcelDateModalBtn?.addEventListener('click', openRsaExcelDateModal);
closeRsaExcelDateModalBtn?.addEventListener('click', closeRsaExcelDateModal);
cancelRsaExcelDateModalBtn?.addEventListener('click', closeRsaExcelDateModal);
submitRsaExcelDateRangeBtn?.addEventListener('click', handleRsaLevelTwoDateRangeSubmit);
closeRsaExcelResultsModalBtn?.addEventListener('click', closeRsaExcelResultsModal);
closeRsaExcelResultsFooterBtn?.addEventListener('click', closeRsaExcelResultsModal);
rsaExcelResultsSelectAll?.addEventListener('change', () => {
    currentRsaExcelFilteredResults.forEach((sub) => {
        if (rsaExcelResultsSelectAll.checked) {
            currentRsaExcelSelectedIds.add(sub.id);
        } else {
            currentRsaExcelSelectedIds.delete(sub.id);
        }
    });
    document.querySelectorAll('.rsa-export-results-select').forEach((input) => {
        input.checked = currentRsaExcelSelectedIds.has(input.value);
    });
    updateRsaExcelResultsSelectAllState();
});
rsaExcelResultsTableBody?.addEventListener('change', (e) => {
    if (!e.target?.classList?.contains('rsa-export-results-select')) return;
    if (e.target.checked) {
        currentRsaExcelSelectedIds.add(e.target.value);
    } else {
        currentRsaExcelSelectedIds.delete(e.target.value);
    }
    updateRsaExcelResultsSelectAllState();
});
let rsaExcelResultsSearchDebounce = null;
rsaExcelResultsSearch?.addEventListener('input', () => {
    if (rsaExcelResultsSearchDebounce) window.clearTimeout(rsaExcelResultsSearchDebounce);
    rsaExcelResultsSearchDebounce = window.setTimeout(() => {
        renderRsaExcelResultsModal();
    }, 180);
});
exportRsaLevel2AllBtn?.addEventListener('click', () => {
    if (!currentRsaExcelResults.length) {
        showNotification('No applications available in this export window.', 'warning');
        return;
    }
    downloadRsaExcel(currentRsaExcelResults, 'level2_date_range_all');
});
exportRsaLevel2SelectedBtn?.addEventListener('click', () => {
    const selected = getSelectedRsaExcelResults();
    if (!selected.length) {
        showNotification('Select at least one application from the export list.', 'warning');
        return;
    }
    downloadRsaExcel(selected, 'level2_date_range_selected');
});

// Finally Submitted tab filters
const finalFilterBtn = document.getElementById('finalFilterBtn');
const finalSearch = document.getElementById('finalSearch');
finalFilterBtn?.addEventListener('click', renderCurrentTab);
finalSearch?.addEventListener('input', renderCurrentTab);
document.getElementById('rsaLeaveMineBtn')?.addEventListener('click', () => {
    document.getElementById('rsaLeaveMineSection')?.style.setProperty('display', '');
    document.getElementById('rsaLeaveReliefSection')?.style.setProperty('display', 'none');
    const mineBtn = document.getElementById('rsaLeaveMineBtn');
    const reliefBtn = document.getElementById('rsaLeaveReliefBtn');
    if (mineBtn) { mineBtn.style.background = '#003366'; mineBtn.style.color = '#fff'; mineBtn.style.border = 'none'; }
    if (reliefBtn) { reliefBtn.style.background = ''; reliefBtn.style.color = ''; reliefBtn.style.border = ''; }
});
document.getElementById('rsaLeaveReliefBtn')?.addEventListener('click', () => {
    document.getElementById('rsaLeaveMineSection')?.style.setProperty('display', 'none');
    document.getElementById('rsaLeaveReliefSection')?.style.setProperty('display', '');
    const mineBtn = document.getElementById('rsaLeaveMineBtn');
    const reliefBtn = document.getElementById('rsaLeaveReliefBtn');
    if (reliefBtn) { reliefBtn.style.background = '#003366'; reliefBtn.style.color = '#fff'; reliefBtn.style.border = 'none'; }
    if (mineBtn) { mineBtn.style.background = ''; mineBtn.style.color = ''; mineBtn.style.border = ''; }
});

// tab navigation
document.querySelectorAll('.nav-item[data-tab]').forEach(item => {
    item.addEventListener('click', (e) => {
        e.preventDefault();
        const tab = item.dataset.tab;
        switchTab(tab);
    });
});

// authentication
auth.onAuthStateChanged(user=>{
    if (!user) return window.location.href = 'index.html';
    currentUser = user;
    loadCurrentRsaProfile(user).then((allowed) => {
        if (!allowed) return;
        loadQueue();
        switchTab(getInitialRsaTab());
    });
});

window.signOutUser = async () => {
    await performAppLogout({
        auth,
        beforeSignOut: async () => {
            const userId = currentRsaProfileData?.id || currentUser?.uid || '';
            if (userId) {
                await updateDoc(doc(db, 'users', userId), {
                    isOnline: false,
                    lastSeenAt: serverTimestamp(),
                    lastLogoutAt: serverTimestamp()
                }).catch(() => {});
            }
        }
    });
};

// merge modal controls
document.getElementById('closeMergeModal')?.addEventListener('click', closeMergeModalFunc);
document.getElementById('cancelMerge')?.addEventListener('click', closeMergeModalFunc);
document.getElementById('mergeDocumentsBtn')?.addEventListener('click', handleMergeDocuments);

// Customer details modal controls
document.getElementById('closeCustomerDetails')?.addEventListener('click', closeCustomerDetailsModal);
document.getElementById('closeCustomerDetailsFooter')?.addEventListener('click', closeCustomerDetailsModal);
document.getElementById('closeViewer')?.addEventListener('click', closeViewerModal);
document.getElementById('closeRsaLeaveApplicationsBtn')?.addEventListener('click', () => {
    document.getElementById('rsaLeaveApplicationsModal')?.classList.remove('active');
});
closeRsaRejectModalBtn?.addEventListener('click', closeRsaRejectModal);
cancelRsaRejectModalBtn?.addEventListener('click', closeRsaRejectModal);
confirmRsaRejectBtn?.addEventListener('click', submitRsaRejection);
rsaRejectReasonInput?.addEventListener('input', () => {
    rsaRejectReasonInput.style.borderColor = '';
    rsaRejectReasonInput.style.boxShadow = '';
});
finalSubmitBtn?.addEventListener('click', (e) => {
    e.preventDefault();
    if (currentDetailsSubmissionId) window.finalSubmitRsa(currentDetailsSubmissionId);
});
setupForceRefreshButtons();
setupIdleLogout();

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

// Close modals when clicking outside
window.addEventListener('click', (e) => {
    const customerDetailsModal = document.getElementById('customerDetailsModal');
    const viewerModal = document.getElementById('viewerModal');
    const mergeModal = document.getElementById('mergeModal');
    
    if (e.target === customerDetailsModal) closeCustomerDetailsModal();
    if (e.target === viewerModal) closeViewerModal();
    if (e.target === mergeModal) closeMergeModalFunc();
    if (e.target === rsaRejectModal) closeRsaRejectModal();
    if (e.target === rsaExcelDateModal) closeRsaExcelDateModal();
    if (e.target === rsaExcelResultsModal) closeRsaExcelResultsModal();
    if (e.target === document.getElementById('rsaLeaveApplicationsModal')) {
        document.getElementById('rsaLeaveApplicationsModal')?.classList.remove('active');
    }
});

// ==================== CUSTOMER DETAILS MODAL ====================
window.openCustomerDetails = async (submissionId) => {
    const sub = await getLatestSubmissionById(submissionId);
    if (!sub) return;

    currentDetailsSubmissionId = submissionId;
    
    const details = sub.customerDetails || {};
    const doc = document;
    
    // Populate personal information
    doc.getElementById('customerDetailsTitle').textContent = `Details - ${sub.customerName}`;
    doc.getElementById('detailCustomerName').textContent = sub.customerName || '-';
    doc.getElementById('detailDOB').textContent = details.dob || '-';
    doc.getElementById('detailNIN').textContent = details.nin || '-';
    doc.getElementById('detailPFA').textContent = details.pfa || '-';
    doc.getElementById('detailAccountNo').textContent = details.accountNo || '-';
    doc.getElementById('detailEmail').textContent = details.email || '-';
    setWhatsAppContact(doc.getElementById('detailPhone'), details.phone || '');
    doc.getElementById('detailAddress').textContent = details.address || '-';
    doc.getElementById('detailEmployer').textContent = details.employer || '-';
    doc.getElementById('detailOriginatingTP').textContent = details.originatingTP || '-';
    doc.getElementById('detailMortgageLoanApplicationFormDate').textContent = details.mortgageLoanApplicationFormDate || '-';
    doc.getElementById('detailPenNo').textContent = details.penNo || '-';
    
    // Populate RSA information
    doc.getElementById('detailRSABalance').textContent = details.rsaBalance ? formatCurrency(details.rsaBalance) : '-';
    doc.getElementById('detailRSADate').textContent = details.rsaStatementDate || '-';
    doc.getElementById('detailRSA25').textContent = details.rsa25 ? formatCurrency(details.rsa25) : '-';

    // Property / loan
    doc.getElementById('detailPropertyType').textContent = details.propertyType || '-';
    doc.getElementById('detailHouseNo').textContent = details.houseNumber || sub.houseNumber || '-';
    doc.getElementById('detailTenor').textContent = details.tenor ? `${details.tenor}` : '-';
    doc.getElementById('detailPropertyValue').textContent = details.propertyValue ? formatCurrency(details.propertyValue) : '-';
    doc.getElementById('detailFacilityFee').textContent = details.facilityFee ? formatCurrency(details.facilityFee) : '-';
    doc.getElementById('detailLoanAmount').textContent = details.loanAmount ? formatCurrency(roundUpToThousand(parseMoneyValue(details.loanAmount))) : '-';
    
    // Populate submission information
    doc.getElementById('detailUploadedBy').textContent = sub.uploadedBy || '-';
    doc.getElementById('detailUploadedAt').textContent = formatTimestamp(sub.uploadedAt);
    doc.getElementById('detailStatus').textContent = (sub.status || '-').toUpperCase();
    doc.getElementById('detailReviewedBy').textContent = sub.reviewedBy || '-';
    
    // Populate documents list
    const docsList = getEffectiveSubmissionDocuments(sub);
    
    const docListHtml = docsList.length === 0 
        ? '<p style="color: #999; text-align: center; padding: 20px;">No documents submitted</p>'
        : `<div style="display: grid; gap: 10px;">
            ${docsList.map((doc, idx) => `
                <div style="background: white; padding: 12px; border: 1px solid #e5e7eb; border-radius: 6px; display: flex; justify-content: space-between; align-items: center;">
                    <div>
                        <div style="font-weight: 600;">${DOCUMENT_TYPES[doc.documentType] || doc.documentType || 'Document'}</div>
                        <div style="font-size: 12px; color: #666;">${doc.name}</div>
                    </div>
                    <div style="display: flex; gap: 8px;">
                        <button class="action-btn" onclick="window.viewDocumentRSA('${submissionId}', ${idx})" title="View">
                            <i class="fas fa-eye"></i>
                        </button>
                        <button class="action-btn" onclick="window.downloadDocumentRSA('${submissionId}', ${idx})" title="Download">
                            <i class="fas fa-download"></i>
                        </button>
                    </div>
                </div>
            `).join('')}
        </div>`;
    
    doc.getElementById('customerDocumentsList').innerHTML = docListHtml;

    if (finalSubmitBtn) {
        const normalizedStatus = (sub.status || '').toLowerCase();
        const canSubmit = normalizedStatus === 'approved' || normalizedStatus === 'processing_to_pfa';
        finalSubmitBtn.disabled = !canSubmit;
        finalSubmitBtn.style.opacity = canSubmit ? '1' : '0.6';
        finalSubmitBtn.title = canSubmit ? 'Mark as finally submitted' : 'Only approved applications can be submitted';
    }
    
    // Show modal
    const modal = doc.getElementById('customerDetailsModal');
    if (modal) modal.classList.add('active');
};

function formatCurrency(value) {
    const num = Number(value || 0);
    try {
        return num.toLocaleString('en-NG', { style: 'currency', currency: 'NGN', maximumFractionDigits: 2 });
    } catch (e) {
        return '₦' + num.toLocaleString(undefined, { maximumFractionDigits: 2 });
    }
}

function closeCustomerDetailsModal() {
    const modal = document.getElementById('customerDetailsModal');
    if (modal) modal.classList.remove('active');
}

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

window.viewDocumentRSA = (submissionId, docIndex) => {
    const sub = allSubmissions.find(s => s.id === submissionId);
    const docs = getEffectiveSubmissionDocuments(sub);
    if (!sub || !docs[docIndex]) return;
    
    const doc = docs[docIndex];
    const viewerFileName = document.getElementById('viewerFileName');
    const documentViewer = document.getElementById('documentViewer');
    
    if (viewerFileName) viewerFileName.textContent = doc.name;
    if (documentViewer) documentViewer.src = buildViewerDocumentUrl(doc.fileUrl, sub, docIndex);
    
    const viewerModal = document.getElementById('viewerModal');
    if (viewerModal) viewerModal.classList.add('active');
};

window.downloadDocumentRSA = async (submissionId, docIndex) => {
    const sub = await getLatestSubmissionById(submissionId);
    const docs = getEffectiveSubmissionDocuments(sub);
    if (!sub || !docs[docIndex]) return;
    
    const doc = docs[docIndex];
    try {
        showLoader('Downloading document...');
        const response = await fetchWithCorsFallback(doc.fileUrl);
        const blob = await response.blob();
        await downloadBlobAsFile(blob, doc.name);
        showNotification('✅ Download started', 'success');
    } catch (error) {
        if (openDirectDocumentDownload(doc.fileUrl, doc.name)) {
            showNotification('Storage blocked secure download, so the document opened directly.', 'warning');
        } else {
            showNotification('Download failed: ' + error.message, 'error');
        }
    } finally {
        hideLoader();
    }
};

window.downloadAllRsa = async (submissionId) => {
    const sub = await getLatestSubmissionById(submissionId);
    if (!sub) return;

    const docs = getEffectiveSubmissionDocuments(sub);
    if (!docs.length) {
        showNotification('No documents available for this application', 'warning');
        return;
    }

    const safeCustomerName = (sub.customerName || 'Customer')
        .replace(/[^a-zA-Z0-9\s_-]/g, '_')
        .trim() || 'Customer';

    try {
        showLoader('Preparing document download...');
        let customerFolder = null;
        let successCount = 0;
        let directOpenedCount = 0;
        let failedCount = 0;

        if ('showDirectoryPicker' in window) {
            showNotification('Select a destination folder to save all documents', 'info');
            const rootFolder = await window.showDirectoryPicker({ mode: 'readwrite', startIn: 'downloads' });
            customerFolder = await rootFolder.getDirectoryHandle(safeCustomerName, { create: true });
        }

        for (let i = 0; i < docs.length; i++) {
            const docItem = docs[i];
            if (!docItem?.fileUrl) continue;

            showLoader(`Downloading ${i + 1} of ${docs.length}...`);
            const fileName = (docItem.name || `${safeCustomerName}_document_${i + 1}.pdf`).replace(/[\\/:*?"<>|]/g, '_');

            try {
                const response = await fetchWithCorsFallback(docItem.fileUrl);
                const blob = await response.blob();

                if (customerFolder) {
                    const fileHandle = await customerFolder.getFileHandle(fileName, { create: true });
                    const writable = await fileHandle.createWritable();
                    await writable.write(blob);
                    await writable.close();
                } else {
                    await downloadBlobAsFile(blob, fileName);
                }
                successCount++;
            } catch (docError) {
                if (openDirectDocumentDownload(docItem.fileUrl, fileName)) {
                    directOpenedCount++;
                } else {
                    failedCount++;
                }
            }
        }

        if (failedCount > 0) {
            showNotification(`Downloaded ${successCount}, opened ${directOpenedCount} directly, ${failedCount} failed`, 'warning');
        } else if (directOpenedCount > 0) {
            showNotification(`Storage blocked secure download for ${directOpenedCount} document(s), so they opened directly.`, 'warning');
        } else {
            showNotification('All documents downloaded successfully', 'success');
            alert('All documents downloaded successfully.');
        }
    } catch (error) {
        if (error?.name === 'AbortError') {
            showNotification('Download cancelled', 'info');
        } else {
            showNotification('Download failed: ' + (error?.message || 'Unknown error'), 'error');
        }
    } finally {
        hideLoader();
    }
};

window.rejectRsaSubmission = async (submissionId) => {
    if (typeof window.assertAppWritable === 'function' && !window.assertAppWritable('RSA rejection')) return;
    const sub = allSubmissions.find((s) => s.id === submissionId);
    if (!sub) return;
    currentRejectSubmissionId = submissionId;
    if (rsaRejectCustomerNameEl) rsaRejectCustomerNameEl.textContent = sub.customerName || 'This customer';
    if (rsaRejectReasonInput) {
        rsaRejectReasonInput.value = sub.latestRejectionReason || sub.comment || '';
        rsaRejectReasonInput.style.borderColor = '';
        rsaRejectReasonInput.style.boxShadow = '';
    }
    rsaRejectModal?.classList.add('active');
    setTimeout(() => rsaRejectReasonInput?.focus(), 30);
};

async function submitRsaRejection() {
    if (typeof window.assertAppWritable === 'function' && !window.assertAppWritable('RSA rejection')) return;
    const submissionId = String(currentRejectSubmissionId || '').trim();
    if (!submissionId) return;
    const sub = allSubmissions.find((s) => s.id === submissionId);
    if (!sub) {
        closeRsaRejectModal();
        return;
    }

    const trimmedReason = String(rsaRejectReasonInput?.value || '').trim();
    const systemSettings = await getSystemSettings(db, { force: true });
    const rejectionRules = systemSettings.rejectionRules || {};
    const rolePermissions = systemSettings.rolePermissions || {};
    const minRejectLength = Number(rejectionRules.minLength || 0);
    if (rolePermissions.rsaCanReject === false) {
        showNotification('RSA rejections are currently disabled by Super Admin.', 'warning');
        return;
    }
    if (rejectionRules.rsaRequired !== false && !trimmedReason) {
        if (rsaRejectReasonInput) {
            rsaRejectReasonInput.style.borderColor = '#dc2626';
            rsaRejectReasonInput.style.boxShadow = '0 0 0 4px rgba(220, 38, 38, 0.12)';
            rsaRejectReasonInput.focus();
        }
        showNotification('Rejection reason is required.', 'warning');
        return;
    }
    if (trimmedReason && minRejectLength > 0 && trimmedReason.length < minRejectLength) {
        if (rsaRejectReasonInput) rsaRejectReasonInput.focus();
        showNotification(`Rejection reason must be at least ${minRejectLength} characters.`, 'warning');
        return;
    }

    if (confirmRsaRejectBtn) {
        confirmRsaRejectBtn.disabled = true;
        confirmRsaRejectBtn.innerHTML = '<i class="fas fa-spinner fa-spin"></i> Rejecting...';
    }

    try {
        await updateDoc(doc(db, 'submissions', submissionId), {
            status: 'rejected_by_rsa',
            comment: trimmedReason,
            latestRejectionReason: trimmedReason,
            latestRejectedBy: currentUser?.email || '',
            latestRejectedAt: serverTimestamp(),
            latestRejectedStage: 'rsa',
            resubmittedAfterRejection: false,
            rsaReady: true
        });
        await addDoc(collection(db, 'audit'), {
            action: 'application_rejected_by_rsa',
            submissionId,
            customerName: sub.customerName || '',
            userEmail: sub.uploadedBy || '',
            rejectedBy: currentUser?.email || '',
            performedBy: currentUser?.email || '',
            details: trimmedReason,
            timestamp: serverTimestamp()
        });
        notifyStatusChangePush({
            currentUser,
            submissionId,
            customerName: sub.customerName || '',
            newStatus: 'rejected_by_rsa',
            statusLabel: 'Rejected by RSA',
            actionLabel: 'Application Rejected by RSA',
            message: `Application for ${sub.customerName || 'this customer'} was rejected by RSA for correction.`
        }).catch(() => {});
        currentRejectSubmissionId = '';
        showNotification('Successfully rejected by RSA.', 'success');
    } catch (error) {
        showNotification('Failed to reject application: ' + (error?.message || 'Unknown error'), 'error');
    } finally {
        if (confirmRsaRejectBtn) {
            confirmRsaRejectBtn.disabled = false;
            confirmRsaRejectBtn.innerHTML = '<i class="fas fa-times-circle"></i> Reject Application';
        }
    }
}

function closeRsaRejectModal() {
    currentRejectSubmissionId = null;
    rsaRejectModal?.classList.remove('active');
    if (rsaRejectCustomerNameEl) rsaRejectCustomerNameEl.textContent = '-';
    if (rsaRejectReasonInput) {
        rsaRejectReasonInput.value = '';
        rsaRejectReasonInput.style.borderColor = '';
        rsaRejectReasonInput.style.boxShadow = '';
    }
}

function closeViewerModal() {
    const modal = document.getElementById('viewerModal');
    if (modal) modal.classList.remove('active');
    const iframe = document.getElementById('documentViewer');
    if (iframe) iframe.src = '';
}

if (typeof window.signOutUser !== 'function') {
    window.signOutUser = () => { window.location.href = 'index.html'; };
}
window.viewMergedPDF = viewMergedPDF;

initializeRsaReportDates();

generateRsaStageReportBtn?.addEventListener('click', async () => {
    const originalHtml = generateRsaStageReportBtn.innerHTML;
    generateRsaStageReportBtn.disabled = true;
    generateRsaStageReportBtn.innerHTML = '<i class="fas fa-spinner fa-spin"></i> Generating...';
    try {
        currentRsaStageReport = await buildRsaStageReport();
        renderDashboardStageReport(currentRsaStageReport, {
            metaEl: rsaReportMeta,
            summaryBodyEl: rsaReportSummaryBody,
            detailsBodyEl: rsaReportDetailsBody
        });
        showNotification('RSA report generated successfully.', 'success');
    } catch (error) {
        showNotification(error?.message || 'Failed to generate RSA report.', 'error');
    } finally {
        generateRsaStageReportBtn.disabled = false;
        generateRsaStageReportBtn.innerHTML = originalHtml;
    }
});

exportRsaStageReportBtn?.addEventListener('click', async () => {
    if (!currentRsaStageReport) {
        showNotification('Generate a report first.', 'warning');
        return;
    }
    const originalHtml = exportRsaStageReportBtn.innerHTML;
    exportRsaStageReportBtn.disabled = true;
    exportRsaStageReportBtn.innerHTML = '<i class="fas fa-spinner fa-spin"></i> Exporting...';
    try {
        await exportDashboardStageReportExcel(currentRsaStageReport, 'CMBank RSA Dashboard');
        showNotification('RSA report Excel downloaded.', 'success');
    } catch (error) {
        showNotification(error?.message || 'Failed to export RSA report.', 'error');
    } finally {
        exportRsaStageReportBtn.disabled = false;
        exportRsaStageReportBtn.innerHTML = originalHtml;
    }
});
