// js/uploader-dashboard.js - LIGHTWEIGHT VERSION (Only Track + Calculator)
import { auth, db } from './firebase-config.js';
import {
    getCurrentUserProfile as getCurrentUserProfileShared,
    getUserFullName as getUserFullNameShared
} from './shared/user-directory.js?v=20260518a';
import { getSystemSettings } from './shared/system-settings.js?v=20260617a';
import { formatAppDateTime } from './shared/app-time.js';
import {
    getTimestampMillis as getStageTimestampMillis,
    getSubmissionCurrentStageEntryAt,
    getSubmissionReviewEntryAt,
    getSubmissionApprovalEntryAt,
    getSubmissionPaymentEntryAt,
    getSubmissionClearedEntryAt
} from './shared/submission-stage.js?v=20260609a';
import {
    collection, query, where, orderBy, onSnapshot, getDocs, getDoc, doc, limit, serverTimestamp, updateDoc
} from "https://www.gstatic.com/firebasejs/9.22.0/firebase-firestore.js";
import { signOut } from "https://www.gstatic.com/firebasejs/9.22.0/firebase-auth.js";

let currentUser = null;
let allSubmissions = [];
const userFullNames = new Map();

// ==================== PROPERTY RULES (RSA Balance Ranges) ====================
const DEFAULT_PROPERTY_RULES = [
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
let PROPERTY_RULES = [...DEFAULT_PROPERTY_RULES];

function determinePropertyByRsa(rsaAmount) {
    const n = Number(rsaAmount) || 0;
    for (const r of PROPERTY_RULES) {
        if (n >= r.min && n <= r.max) return r;
    }
    return null;
}

function formatCurrency(value) {
    const num = Number(value || 0);
    try {
        return num.toLocaleString('en-NG', { style: 'currency', currency: 'NGN', maximumFractionDigits: 2 });
    } catch (e) {
        return '₦' + num.toLocaleString(undefined, { maximumFractionDigits: 2 });
    }
}

function escapeHtml(value) {
    return String(value || '')
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&#39;');
}

function normalizeTrackSearchText(value) {
    return String(value || '').trim().toLowerCase().replace(/\s+/g, ' ');
}

function normalizePenNumber(value) {
    return String(value || '').trim().toLowerCase().replace(/[^a-z0-9]+/g, '');
}

function getSubmissionPenNo(submission = {}) {
    return normalizePenNumber(
        submission?.customerDetails?.penNoNormalized ||
        submission?.penNoNormalized ||
        submission?.customerDetails?.penNo ||
        submission?.penNo ||
        ''
    );
}

function getSubmissionCustomerName(submission = {}) {
    return String(submission?.customerName || submission?.customerDetails?.name || '').trim();
}

function getPenNumberQueryVariants(penNo = '') {
    const raw = String(penNo || '').trim();
    const normalized = normalizePenNumber(raw);
    return Array.from(new Set([
        raw,
        raw.toUpperCase(),
        raw.toLowerCase(),
        normalized,
        normalized.toUpperCase()
    ].filter(Boolean))).slice(0, 10);
}

function addTrackResult(resultMap, docSnap) {
    if (!docSnap?.exists?.()) return;
    const submission = { id: docSnap.id, ...(docSnap.data() || {}) };
    if (String(submission.status || '').toLowerCase() === 'draft') return;
    resultMap.set(docSnap.id, submission);
}

// DOM Elements
const userName = document.getElementById('userName');
const userAvatar = document.getElementById('userAvatar');
const pageTitle = document.getElementById('pageTitle');

// ==================== Track Modal Elements ====================
const globalTrackBtn = document.getElementById('globalTrackBtn');
const globalTrackModal = document.getElementById('globalTrackModal');
const closeGlobalTrackModal = document.getElementById('closeGlobalTrackModal');
const cancelGlobalTrack = document.getElementById('cancelGlobalTrack');
const searchTrackBtn = document.getElementById('searchTrackBtn');
const trackSearchInput = document.getElementById('trackSearchInput');
const trackSearchResults = document.getElementById('trackSearchResults');
const trackResultsList = document.getElementById('trackResultsList');

// ==================== Loan Calculator Elements ====================
const loanCalcBtn = document.getElementById('loanCalcBtn');
const loanCalcModal = document.getElementById('loanCalculatorModal');
const closeLoanCalcModal = document.getElementById('closeLoanCalcModal');
const closeLoanCalcBtn = document.getElementById('closeLoanCalcBtn');
const calculateLoanBtn = document.getElementById('calculateLoanBtn');
const calcRsaBalance = document.getElementById('calcRsaBalance');

// Initialize
document.addEventListener('DOMContentLoaded', () => {
    auth.onAuthStateChanged(async (user) => {
        if (user) {
            currentUser = user;
            try {
                const systemSettings = await getSystemSettings(db, { force: true });
                PROPERTY_RULES = Array.isArray(systemSettings.propertyRules) && systemSettings.propertyRules.length
                    ? [...systemSettings.propertyRules]
                    : [...DEFAULT_PROPERTY_RULES];
                const data = await getCurrentUserProfileShared(db, user);
                if (data) {
                    userName.textContent = data.fullName || user.displayName || user.email.split('@')[0];
                } else {
                    userName.textContent = user.displayName || user.email.split('@')[0];
                }
            } catch (e) {
                userName.textContent = user.displayName || user.email.split('@')[0];
            }
            userAvatar.src = user.photoURL || 'data:image/svg+xml,%3Csvg xmlns=\'http://www.w3.org/2000/svg\' width=\'40\' height=\'40\' viewBox=\'0 0 40 40\'%3E%3Ccircle cx=\'20\' cy=\'20\' r=\'20\' fill=\'%23003366\'/%3E%3Ctext x=\'20\' y=\'25\' text-anchor=\'middle\' fill=\'%23ffffff\' font-size=\'16\'%3E👤%3C/text%3E%3C/svg%3E';

            // Just load submissions for display - don't interfere with uploader
            loadSubmissionsForDisplay();
        } else {
            window.location.href = 'index.html';
        }
    });

    setupTrackListeners();
    setupLoanCalculator();
    setupTabSwitching();
    setupForceRefreshButtons();
    setupIdleLogout();
    ensureSignOutUser();
});

function ensureSignOutUser() {
    try {
        const desc = Object.getOwnPropertyDescriptor(window, 'signOutUser');
        // `js/auth.js` defines a non-writable `signOutUser`; reuse it when present.
        if (desc && desc.writable === false) return;
        window.signOutUser = async () => {
            try {
                const userId = currentUser?.uid || '';
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
    } catch (e) { /* ignore */ }
}

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

function setupTabSwitching() {
    document.querySelectorAll('.nav-item[data-tab]').forEach(item => {
        item.addEventListener('click', (e) => {
            e.preventDefault();
            const tabId = item.dataset.tab;

            // Use the full uploader tab system when available.
            if (typeof window.switchTab === 'function') {
                window.switchTab(tabId);
                return;
            }

            document.querySelectorAll('.nav-item').forEach(i => i.classList.remove('active'));
            item.classList.add('active');

            document.querySelectorAll('.tab-content').forEach(c => c.classList.remove('active'));
            const targetTab = document.getElementById(`${tabId}Tab`);
            if(targetTab) targetTab.classList.add('active');

            const titles = {
                pending: 'Pending Review',
                overview: 'Dashboard',
                approved: 'Approved Documents',
                rejected: 'Rejected Submissions',
                paid: 'Agent Commission',
                'register-agent': 'Register Agent',
                profile: 'My Profile',
                help: 'Help & SOP'
            };
            if (pageTitle) pageTitle.textContent = titles[tabId] || 'Dashboard';
        });
    });
}

function setupTrackListeners() {
    if (globalTrackBtn && globalTrackModal) {
        globalTrackBtn.addEventListener('click', () => {
            if (globalTrackModal) {
                trackSearchInput.value = '';
                trackSearchResults.style.display = 'none';
                trackResultsList.innerHTML = '';
                globalTrackModal.classList.add('active');
            }
        });
    }

    if (closeGlobalTrackModal) {
        closeGlobalTrackModal.addEventListener('click', () => {
            globalTrackModal.classList.remove('active');
        });
    }

    if (cancelGlobalTrack) {
        cancelGlobalTrack.addEventListener('click', () => {
            globalTrackModal.classList.remove('active');
        });
    }

    if (searchTrackBtn) {
        searchTrackBtn.addEventListener('click', performGlobalTrackSearch);
    }

    if (trackSearchInput) {
        trackSearchInput.addEventListener('keypress', (e) => {
            if (e.key === 'Enter') {
                performGlobalTrackSearch();
            }
        });
    }
}

function setupLoanCalculator() {
    if (loanCalcBtn && loanCalcModal) {
        loanCalcBtn.addEventListener('click', function(e) {
            e.preventDefault();
            e.stopPropagation();
            if (calcRsaBalance) calcRsaBalance.value = '';
            document.getElementById('calc25Percent').textContent = '₦0.00';
            document.getElementById('calcFacilityFee').textContent = '₦0.00';
            document.getElementById('calcPropertyType').textContent = '-';
            document.getElementById('calcPropertyValue').textContent = '₦0.00';
            document.getElementById('calcLoanAmount').textContent = '₦0.00';
            const calcResults = document.getElementById('calcResults');
            if (calcResults) calcResults.style.display = 'none';
            loanCalcModal.classList.add('active');
        });
    }

    if (closeLoanCalcModal) {
        closeLoanCalcModal.addEventListener('click', function() {
            loanCalcModal.classList.remove('active');
        });
    }

    if (closeLoanCalcBtn) {
        closeLoanCalcBtn.addEventListener('click', function() {
            const calcResults = document.getElementById('calcResults');
            if (calcResults) calcResults.style.display = 'none';
        });
    }

    if (calculateLoanBtn) {
        calculateLoanBtn.addEventListener('click', function(e) {
            e.preventDefault();
            calculateLoan();
        });
    }

    if (calcRsaBalance) {
        calcRsaBalance.addEventListener('keypress', function(e) {
            if (e.key === 'Enter') {
                e.preventDefault();
                calculateLoan();
            }
        });
        calcRsaBalance.addEventListener('input', function(e) {
            // Allow digits and a single decimal point.
            let value = String(e.target.value || '').replace(/[^0-9.]/g, '');
            const firstDot = value.indexOf('.');
            if (firstDot !== -1) {
                value = value.slice(0, firstDot + 1) + value.slice(firstDot + 1).replace(/\./g, '');
            }
            e.target.value = value;
        });
    }
}

function calculateLoan() {
    if (!calcRsaBalance) return;
    const balance = parseFloat(calcRsaBalance.value) || 0;
    if (!balance || balance < 4000000) {
        showNotification('Please enter a valid RSA balance (minimum ₦4,000,000)', 'warning');
        return;
    }
    if (balance > 299999999) {
        showNotification('RSA balance exceeds maximum allowed (₦299,999,999)', 'error');
        return;
    }
    const rule = determinePropertyByRsa(balance);
    if (!rule) {
        showNotification('RSA balance outside property range', 'error');
        return;
    }
    const twentyFivePercent = Math.max(0, Math.floor((balance * 0.25) / 1000) * 1000);
    // Both 25% contribution and loan amount stay in thousands.
    const loanAmount = Math.max(0, Math.ceil((rule.value - twentyFivePercent) / 1000) * 1000);
    document.getElementById('calc25Percent').textContent = formatCurrency(twentyFivePercent);
    document.getElementById('calcFacilityFee').textContent = formatCurrency(rule.fee);
    document.getElementById('calcPropertyType').textContent = rule.name;
    document.getElementById('calcPropertyValue').textContent = formatCurrency(rule.value);
    document.getElementById('calcLoanAmount').textContent = formatCurrency(loanAmount);
    const calcResults = document.getElementById('calcResults');
    if (calcResults) calcResults.style.display = 'block';
}

async function performGlobalTrackSearch() {
    const searchTerm = trackSearchInput.value.trim();
    if (!searchTerm) {
        showNotification('Please enter a search term', 'error');
        return;
    }
    showLoader('Searching applications...');
    try {
        const submissionsRef = collection(db, 'submissions');
        const resultMap = new Map();

        const docSnap = await getDoc(doc(db, 'submissions', searchTerm)).catch(() => null);
        addTrackResult(resultMap, docSnap);

        const normalizedPenNo = normalizePenNumber(searchTerm);
        const variants = getPenNumberQueryVariants(searchTerm);
        const penQueries = [
            query(submissionsRef, where('customerDetails.penNoNormalized', '==', normalizedPenNo)),
            query(submissionsRef, where('penNoNormalized', '==', normalizedPenNo))
        ];
        if (variants.length) {
            penQueries.push(
                query(submissionsRef, where('customerDetails.penNo', 'in', variants)),
                query(submissionsRef, where('penNo', 'in', variants))
            );
        }
        const snapshots = await Promise.all(penQueries.map((q) => getDocs(q).catch(() => null)));
        snapshots.filter(Boolean).forEach((snapshot) => {
            snapshot.forEach((docSnapItem) => addTrackResult(resultMap, docSnapItem));
        });

        const normalizedName = normalizeTrackSearchText(searchTerm);
        const allSubmissionsQuery = query(submissionsRef, limit(5000));
        const querySnapshot = await getDocs(allSubmissionsQuery);
        querySnapshot.forEach((docSnapItem) => {
            const data = docSnapItem.data();
            const customerName = normalizeTrackSearchText(getSubmissionCustomerName(data));
            const status = String(data.status || '').toLowerCase();
            if (status !== 'draft' && customerName.includes(normalizedName)) {
                resultMap.set(docSnapItem.id, { id: docSnapItem.id, ...data });
            }
        });

        const results = Array.from(resultMap.values())
            .sort((a, b) => getSubmissionCustomerName(a).localeCompare(getSubmissionCustomerName(b)));
        displayTrackResults(results);
    } catch (error) {
        showNotification('Error searching applications: ' + error.message, 'error');
    } finally {
        hideLoader();
    }
}

async function displayTrackResults(results) {
    trackResultsList.innerHTML = '';
    if (results.length === 0) {
        trackResultsList.innerHTML = '<div style="text-align: center; padding: 30px; color: #64748b;">No applications found matching your search.</div>';
        trackSearchResults.style.display = 'block';
        return;
    }
    const resultCount = document.createElement('div');
    resultCount.style.cssText = 'padding: 10px; background: #f8fafc; border-radius: 8px; margin-bottom: 15px; font-weight: 600;';
    resultCount.innerHTML = `<strong>${results.length}</strong> application(s) found`;
    trackResultsList.appendChild(resultCount);

    for (const sub of results) {
        const resultItem = document.createElement('div');
        resultItem.className = 'track-result-item';
        resultItem.style.cssText = 'padding: 15px; border: 1px solid #e2e8f0; border-radius: 8px; margin-bottom: 10px; cursor: pointer; transition: all 0.2s;';
        resultItem.onmouseover = () => { resultItem.style.backgroundColor = '#f8fafc'; };
        resultItem.onmouseout = () => { resultItem.style.backgroundColor = 'white'; };
        resultItem.onclick = () => {
            // Use the existing tracker exposed by the active uploader module.
            if (typeof window.showApplicationTrack === 'function') {
                window.showApplicationTrack(sub.id);
            } else {
                alert('Application ID: ' + sub.id);
            }
        };

        const uploadedByName = sub.uploadedBy ? await getUserFullName(sub.uploadedBy) : 'Unknown';
        const customerName = getSubmissionCustomerName(sub) || 'Unknown';
        const status = String(sub.status || 'pending').trim();
        resultItem.innerHTML = `
            <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 8px;">
                <strong style="font-size: 16px; color: #003366;">${escapeHtml(customerName)}</strong>
                <span style="padding: 4px 10px; border-radius: 20px; font-size: 11px; font-weight: 600; background: ${
                    sub.status === 'approved' ? '#d1fae5' :
                    sub.status === 'rejected' ? '#fee2e2' : '#fef3c7'
                }; color: ${
                    sub.status === 'approved' ? '#065f46' :
                    sub.status === 'rejected' ? '#991b1b' : '#92400e'
                };">${escapeHtml(status)}</span>
            </div>
            <div style="display: flex; gap: 15px; font-size: 12px; color: #64748b; margin-bottom: 8px; flex-wrap: wrap;">
                <span><i class="fas fa-id-card"></i> ID: ${escapeHtml(sub.id.substring(0, 8))}...</span>
                <span><i class="fas fa-calendar"></i> ${safeFormatDate(sub.uploadedAt)}</span>
                ${sub.uploadedBy ? `<span><i class="fas fa-user"></i> Uploader: ${escapeHtml(uploadedByName)}</span>` : ''}
            </div>
        `;
        trackResultsList.appendChild(resultItem);
    }
    trackSearchResults.style.display = 'block';
}

// ==================== GET USER FULL NAME BY EMAIL ====================
async function getUserFullName(email) {
    if (!email) return 'Unknown';
    if (userFullNames.has(email)) return userFullNames.get(email);
    try {
        const fullName = await getUserFullNameShared(db, email);
        userFullNames.set(email, fullName);
        return fullName;
    } catch (err) {
        // Silently fail
    }
    const fallbackName = email.split('@')[0];
    userFullNames.set(email, fallbackName);
    return fallbackName;
}

// ==================== LOAD SUBMISSIONS FOR DISPLAY ONLY ====================
function loadSubmissionsForDisplay() {
    if (!currentUser) return;

    const q = query(
        collection(db, 'submissions'),
        where('uploadedBy', '==', currentUser.email),
        orderBy('uploadedAt', 'desc')
    );

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

        await ensureUserFullNames(Array.from(emails));
        allSubmissions.sort((a, b) => getStageTimestampMillis(getSubmissionCurrentStageEntryAt(b)) - getStageTimestampMillis(getSubmissionCurrentStageEntryAt(a)));

        // Update dashboard counts only - don't render tables that conflict
        updateDashboardCards();
        renderRecentTable();
    }, (error) => {
        // Silently fail
    });
}

async function ensureUserFullNames(emails) {
    if (!emails || emails.length === 0) return;
    for (const email of emails) {
        if (!email || userFullNames.has(email)) continue;
        try {
            const fullName = await getUserFullNameShared(db, email);
            userFullNames.set(email, fullName);
        } catch (e) {
            userFullNames.set(email, email.split('@')[0]);
        }
    }
}

function safeFormatDate(dateValue) {
    return formatAppDateTime(dateValue, 'N/A');
}

function formatStatusLabel(status) {
    const text = String(status || '').trim();
    if (!text) return '-';
    return text
        .replace(/_/g, ' ')
        .replace(/\b\w/g, (char) => char.toUpperCase());
}

function getDisplayNameByEmail(email) {
    const raw = String(email || '').trim();
    if (!raw) return '-';
    return userFullNames.get(raw) || raw;
}

function getSubmissionAgentLabel(sub = {}) {
    return String(
        sub.agentName ||
        sub.agentFullName ||
        sub.agent?.fullName ||
        sub.agent?.name ||
        sub.customerDetails?.agentName ||
        ''
    ).trim() || '-';
}

function getTrackStatusBadgeClass(status) {
    const normalized = String(status || '').trim().toLowerCase();
    if (['rejected', 'rejected_by_reviewer', 'rejected_by_rsa'].includes(normalized)) return 'status-rejected';
    if (['pending', 'submitted', 'resubmitted', 'draft'].includes(normalized)) return 'status-pending';
    return 'status-approved';
}

function getApplicationCurrentStage(submission = {}) {
    const status = String(submission.status || '').trim().toLowerCase();
    if (status === 'draft') return { key: 'draft', label: 'Draft' };
    if (['pending', 'submitted', 'resubmitted', 'rejected', 'rejected_by_reviewer'].includes(status)) {
        return { key: 'reviewer', label: status === 'rejected' || status === 'rejected_by_reviewer' ? 'Reviewer Rejected' : 'Reviewer Stage' };
    }
    if (['approved', 'processing_to_pfa', 'rejected_by_rsa'].includes(status)) {
        return { key: 'rsa', label: status === 'rejected_by_rsa' ? 'RSA Rejected' : 'RSA Stage' };
    }
    if (['sent_to_pfa', 'rsa_submitted', 'paid'].includes(status) || submission.finalSubmitted === true || submission.rsaSubmitted === true) {
        return { key: 'payment', label: status === 'paid' ? 'Payment Confirmed' : 'Payment Stage' };
    }
    if (status === 'cleared') return { key: 'cleared', label: 'Cleared' };
    return { key: 'unknown', label: formatStatusLabel(status) };
}

function getTrackStageTimestamp(submission = {}, stageKey) {
    if (stageKey === 'upload') {
        return submission.reuploadedAt || submission.uploadedAt || submission.submittedAt || submission.createdAt || null;
    }
    if (stageKey === 'reviewer') {
        return submission.reviewedAt || getSubmissionReviewEntryAt(submission) || null;
    }
    if (stageKey === 'rsa') {
        return submission.rsaSubmittedAt ||
            submission.finalSubmittedAt ||
            submission.rsaAssignedAt ||
            getSubmissionApprovalEntryAt(submission) ||
            null;
    }
    if (stageKey === 'payment') {
        return submission.paidAt ||
            submission.paymentAssignedAt ||
            getSubmissionPaymentEntryAt(submission) ||
            null;
    }
    return null;
}

function getTrackTimelineState(submission = {}, stageKey) {
    const currentStage = getApplicationCurrentStage(submission);
    const normalizedStatus = String(submission.status || '').trim().toLowerCase();
    const order = { upload: 0, reviewer: 1, rsa: 2, payment: 3, cleared: 4 };
    const currentOrder = order[currentStage.key] ?? -1;
    const stageOrder = order[stageKey] ?? -1;

    if ((normalizedStatus === 'rejected' || normalizedStatus === 'rejected_by_reviewer') && stageKey === 'reviewer') {
        return { label: 'Attention', className: 'attention' };
    }
    if (normalizedStatus === 'rejected_by_rsa' && stageKey === 'rsa') {
        return { label: 'Attention', className: 'attention' };
    }
    if (currentStage.key === 'cleared') {
        return { label: 'Completed', className: 'completed' };
    }
    if (stageOrder < currentOrder) {
        return { label: 'Completed', className: 'completed' };
    }
    if (stageOrder === currentOrder) {
        return { label: 'Current', className: 'current' };
    }
    return { label: 'Pending', className: 'pending' };
}

function renderTrackSummaryCard(label, value) {
    return `
        <div class="track-modal-summary-card">
            <span class="label">${escapeHtml(label)}</span>
            <div class="value">${escapeHtml(value || '-')}</div>
        </div>
    `;
}

function updateDashboardCards() {
    const pending = allSubmissions.filter(s => s.status === 'pending').length;
    const approved = allSubmissions.filter((s) => {
        const status = String(s.status || '').toLowerCase();
        return status === 'processing_to_pfa' ||
            status === 'approved' ||
            status === 'sent_to_pfa' ||
            status === 'rsa_submitted' ||
            status === 'paid' ||
            status === 'cleared' ||
            s.finalSubmitted === true ||
            s.rsaSubmitted === true;
    }).length;
    const rejected = allSubmissions.filter(s => ['rejected', 'rejected_by_rsa'].includes(String(s.status || '').toLowerCase())).length;
    const paid = allSubmissions.filter(s => String(s.status || '').toLowerCase() === 'paid').length;
    document.getElementById('cardPendingCount') && (document.getElementById('cardPendingCount').textContent = pending);
    document.getElementById('cardApprovedCount') && (document.getElementById('cardApprovedCount').textContent = approved);
    document.getElementById('cardRejectedCount') && (document.getElementById('cardRejectedCount').textContent = rejected);
    const setBadge = (id, count) => {
        const badge = document.getElementById(id);
        if (badge) {
            badge.textContent = String(count);
            badge.style.display = 'inline-block';
        }
    };
    setBadge('pendingCount', pending);
    setBadge('approvedCount', approved);
    setBadge('rejectedCount', rejected);
    setBadge('paidCount', paid);
}

function renderRecentTable() {
    const rows = document.getElementById('recentTableBody');
    if (!rows) return;

    const recent = allSubmissions
        .slice()
        .sort((a, b) => getStageTimestampMillis(getSubmissionCurrentStageEntryAt(b)) - getStageTimestampMillis(getSubmissionCurrentStageEntryAt(a)))
        .slice(0, 10);

    if (recent.length === 0) {
        rows.innerHTML = '<tr><td colspan="5" class="no-data">No recent applications</td></tr>';
        return;
    }

    rows.innerHTML = recent.map(sub => {
        const uploadDate = formatAppDateTime(getSubmissionCurrentStageEntryAt(sub), 'N/A');
        const uploaderName = (sub.uploadedBy && userFullNames.get(sub.uploadedBy)) ? userFullNames.get(sub.uploadedBy) : (sub.uploadedBy || '-');
        return `
            <tr>
                <td><strong>${sub.customerName || 'Unknown'}</strong></td>
                <td>${uploadDate}</td>
                <td><span class="status-badge status-${sub.status}" style="padding:4px 12px; border-radius:20px; font-size:12px; font-weight:600; background:${
                    sub.status === 'approved' ? '#d1fae5' :
                    sub.status === 'rejected' ? '#fee2e2' : '#fef3c7'
                }; color:${
                    sub.status === 'approved' ? '#065f46' :
                    sub.status === 'rejected' ? '#991b1b' : '#92400e'
                };">${sub.status ? sub.status.charAt(0).toUpperCase() + sub.status.slice(1) : 'Pending'}</span></td>
                <td>${uploaderName}</td>
                <td>
                    <button class="action-btn view-btn-small" onclick="window.viewSubmissionDocs('${sub.id}')"><i class="fas fa-eye"></i></button>
                    <button class="action-btn track-btn" onclick="window.showApplicationTrack('${sub.id}')" style="margin-left:5px; background:#8b5cf6; color:white; border:none; padding:5px 10px; border-radius:4px;"><i class="fas fa-map-marker-alt"></i></button>
                </td>
            </tr>
        `;
    }).join('');
}

// Simple notification
function showNotification(msg, type = 'info') {
    const el = document.getElementById('notification');
    if(!el) return;
    el.textContent = msg;
    el.className = `notification ${type}`;
    el.style.display = 'block';
    if (type === 'success') el.style.background = '#10b981';
    else if (type === 'error') el.style.background = '#ef4444';
    else if (type === 'warning') el.style.background = '#f59e0b';
    else el.style.background = '#3b82f6';
    el.style.color = 'white';
    setTimeout(() => { el.style.display = 'none'; }, 3000);
}

// Loader Functions
function showLoader(msg) {
    const loader = document.getElementById('globalLoader');
    const text = document.getElementById('loaderText');
    if(loader && text) {
        text.textContent = msg || "Processing...";
        loader.style.display = 'flex';
    }
}

function hideLoader() {
    const loader = document.getElementById('globalLoader');
    if(loader) {
        loader.style.display = 'none';
    }
}

// ==================== FIXED TRACK FUNCTION ====================
// This overrides any existing track function and ensures the modal shows
window.showApplicationTrack = async function(submissionId) {

    // Find the submission in our local data
    let sub = allSubmissions.find(s => s.id === submissionId);
    if (!sub && submissionId) {
        try {
            const subSnap = await getDoc(doc(db, 'submissions', submissionId));
            if (subSnap.exists()) {
                sub = { id: subSnap.id, ...(subSnap.data() || {}) };
            }
        } catch (error) {
            console.warn('Could not fetch application for tracking', error);
        }
    }
    if (!sub) {
        showNotification('Application not found', 'error');
        return;
    }

    const currentStage = getApplicationCurrentStage(sub);
    const statusLabel = formatStatusLabel(sub.status || '-');
    const statusClass = getTrackStatusBadgeClass(sub.status);
    const uploaderName = sub.uploadedBy ? await getUserFullName(sub.uploadedBy) : '-';
    const reviewerName = sub.assignedTo ? await getUserFullName(sub.assignedTo) : 'Unassigned';
    const rsaName = sub.assignedToRSA ? await getUserFullName(sub.assignedToRSA) : 'Unassigned';
    const paymentName = sub.assignedToPayment ? await getUserFullName(sub.assignedToPayment) : 'Unassigned';
    const agentName = getSubmissionAgentLabel(sub);
    const lastStageTime = safeFormatDate(getSubmissionCurrentStageEntryAt(sub));
    const timelineItems = [
        { key: 'upload', title: 'Uploaded', time: getTrackStageTimestamp(sub, 'upload'), meta: `Uploader: ${uploaderName}` },
        { key: 'reviewer', title: 'Reviewer', time: getTrackStageTimestamp(sub, 'reviewer'), meta: `Assigned Reviewer: ${reviewerName}` },
        { key: 'rsa', title: 'RSA', time: getTrackStageTimestamp(sub, 'rsa'), meta: `Assigned RSA: ${rsaName}` },
        { key: 'payment', title: 'Payment', time: getTrackStageTimestamp(sub, 'payment'), meta: `Assigned Payment: ${paymentName}` }
    ];

    const modal = document.createElement('div');
    modal.className = 'modal active';
    modal.id = 'trackModal-' + Date.now();
    modal.style.cssText = 'display: flex !important; position: fixed; top: 0; left: 0; width: 100%; height: 100%; background: rgba(0,0,0,0.5); z-index: 9999999; align-items: center; justify-content: center;';

    modal.innerHTML = `
        <div class="modal-content large-modal" style="position: relative; z-index: 10000000;">
            <div class="modal-header">
                <h2>Application Tracking</h2>
                <button class="close-btn" onclick="this.closest('.modal').remove()" aria-label="Close modal">&times;</button>
            </div>
            <div class="modal-body">
                <div class="track-modal-hero">
                    <div>
                        <h3>${escapeHtml(sub.customerName || 'Unknown Application')}</h3>
                        <p>Application ID: ${escapeHtml(sub.id || '-')} | Last stage update: ${escapeHtml(lastStageTime)}</p>
                    </div>
                    <div class="track-modal-statuses">
                        <span class="status-badge ${escapeHtml(statusClass)}">${escapeHtml(statusLabel)}</span>
                        <span class="track-stage-pill current">${escapeHtml(currentStage.label)}</span>
                    </div>
                </div>
                <div class="track-modal-summary">
                    ${[
                        renderTrackSummaryCard('Uploaded By', uploaderName),
                        renderTrackSummaryCard('Agent', agentName),
                        renderTrackSummaryCard('Current Stage', currentStage.label),
                        renderTrackSummaryCard('Current Stage Time', lastStageTime),
                        renderTrackSummaryCard('Assigned Reviewer', reviewerName),
                        renderTrackSummaryCard('Assigned RSA', rsaName),
                        renderTrackSummaryCard('Assigned Payment', paymentName),
                        renderTrackSummaryCard('Cleared Time', safeFormatDate(getSubmissionClearedEntryAt(sub)))
                    ].join('')}
                </div>
                <div class="track-modal-timeline">
                    ${timelineItems.map((item) => {
                        const state = getTrackTimelineState(sub, item.key);
                        return `
                            <div class="track-modal-timeline-card">
                                <div class="track-modal-timeline-head">
                                    <div>
                                        <span class="label">${escapeHtml(item.title)} Time</span>
                                        <h4>${escapeHtml(item.title)}</h4>
                                    </div>
                                    <span class="track-stage-pill ${escapeHtml(state.className)}">${escapeHtml(state.label)}</span>
                                </div>
                                <div class="time">${escapeHtml(safeFormatDate(item.time))}</div>
                                <div class="meta">${escapeHtml(item.meta)}</div>
                            </div>
                        `;
                    }).join('')}
                </div>
            </div>
            <div class="modal-footer">
                <button class="cancel-btn" onclick="this.closest('.modal').remove()">Close</button>
            </div>
        </div>
    `;

    document.body.appendChild(modal);

    // Close on backdrop click
    modal.addEventListener('click', (e) => {
        if (e.target === modal) modal.remove();
    });
};

// Simple stage determination function
function getApplicationStageSimple(submission) {
    if (!submission) return 'Unknown';
    const status = String(submission.status || '').toLowerCase();
    if (status === 'cleared') return 'Cleared';
    if (status === 'paid') return 'Paid';
    if (status === 'sent_to_pfa' || status === 'rsa_submitted') return 'Sent to PFA';
    if (status === 'processing_to_pfa' || status === 'approved') return 'Processing to PFA';
    if (status === 'rejected') return 'Rejected - Fix Required';
    if (submission.assignedTo) return `With Reviewer: ${submission.assignedTo}`;
    return 'Pending Assignment';
}

// Make functions global
window.showLoader = showLoader;
window.hideLoader = hideLoader;
