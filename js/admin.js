// js/admin.js - COMPLETE UPDATED VERSION WITH FIXED DOWNLOAD ALL
import { auth, db } from './firebase-config.js';
import { ADMIN_API_BASE_URL } from './admin-api-config.js';
import { formatAppDateTime, getTrustedDateKey } from './shared/app-time.js';
import {
    collection,
    addDoc,
    query,
    where,
    onSnapshot,
    getDocs,
    getDoc,
    doc,
    setDoc,
    updateDoc,
    deleteDoc,
    orderBy,
    serverTimestamp,
    limit
} from "https://www.gstatic.com/firebasejs/9.22.0/firebase-firestore.js";
import {
    getAuth,
    createUserWithEmailAndPassword,
    signOut
} from "https://www.gstatic.com/firebasejs/9.22.0/firebase-auth.js";
import {
    getSubmissionCommissionAmount,
    resolveSubmissionCommissionRate
} from './shared/commission-config.js?v=20260507a';

// Security: suppress console output in admin dashboard.
(() => {
    const noop = () => {};
    try {
        console.log = noop;
        console.warn = noop;
        console.error = noop;
        console.info = noop;
        console.debug = noop;
    } catch (e) { /* ignore */ }
})();

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
    'benefit_application_form': 'Benefit Application Form'
};

// ==================== GLOBAL VARIABLES ====================
let currentAdmin = null;
let currentAdminProfileData = null;
let selectedUserId = null;
let selectedSubmissionId = null;
let allUsers = [];
let allAudits = [];
let allSubmissions = [];
let allPendingAgents = [];
let allApprovedAgents = [];
let allPendingUsers = [];
let allEscalations = [];
let adminNames = {};
let uploaderNames = {};
let userIdNameCache = new Map();
let userEmailNameCache = new Map();
const TRACK_PAGE_SIZE = 10;
let trackAppsPage = 1;
let currentParentTab = 'user-management';
let currentLeafTab = 'users';
let selectedLeaveUserId = '';

const TAB_GROUPS = {
    'user-management': ['users', 'pending-users', 'pending-agents', 'registered-agents'],
    'application-management': ['draft-docs', 'pending-docs', 'approved-docs', 'rejected-docs', 'escalations', 'track-apps', 'finally-submitted', 'payments', 'agent-commissions']
};

const TAB_LABELS = {
    users: 'Users',
    'pending-users': 'Pending Users',
    'pending-agents': 'Pending Agents',
    'registered-agents': 'Registered Agents',
    'draft-docs': 'Draft',
    'pending-docs': 'Pending',
    'approved-docs': 'Approved',
    'rejected-docs': 'Rejected',
    escalations: 'Escalations',
    'track-apps': 'Track Applications',
    'finally-submitted': 'Final Submission',
    payments: 'Payment',
    'agent-commissions': 'Agent Commissions'
};

function setCountBadge(id, count) {
    const badge = document.getElementById(id);
    if (!badge) return;
    badge.textContent = String(count);
    badge.style.display = 'inline-block';
}

function getAdminSubTabCount(tabId) {
    if (tabId === 'users') return allUsers.length;
    if (tabId === 'pending-users') return allPendingUsers.length;
    if (tabId === 'pending-agents') return allPendingAgents.length;
    if (tabId === 'registered-agents') return allApprovedAgents.length;
    if (tabId === 'draft-docs') return allSubmissions.filter((s) => String(s.status || '').toLowerCase() === 'draft').length;
    if (tabId === 'pending-docs') return allSubmissions.filter((s) => String(s.status || '').toLowerCase() === 'pending').length;
    if (tabId === 'approved-docs') return allSubmissions.filter((s) => {
        const status = String(s.status || '').toLowerCase();
        return status === 'processing_to_pfa' || status === 'approved';
    }).length;
    if (tabId === 'rejected-docs') return allSubmissions.filter((s) => ['rejected', 'rejected_by_rsa'].includes(String(s.status || '').toLowerCase())).length;
    if (tabId === 'escalations') return allEscalations.filter((item) => item?.escalationHandled !== true).length;
    if (tabId === 'track-apps') return allSubmissions.filter((s) => String(s.status || '').toLowerCase() !== 'draft').length;
    if (tabId === 'finally-submitted') return allSubmissions.filter((s) => s.finalSubmitted === true || s.rsaSubmitted === true).length;
    if (tabId === 'payments') return allSubmissions.filter((s) => {
        const status = String(s.status || '').toLowerCase();
        return status === 'sent_to_pfa' || status === 'rsa_submitted';
    }).length;
    if (tabId === 'agent-commissions') {
        const keys = new Set();
        allApprovedAgents.forEach((agent) => keys.add(getAgentCommissionKey(agent)));
        allSubmissions.forEach((sub) => {
            if (getSubmissionHasAgent(sub) && isSubmissionCommissionTrackable(sub)) keys.add(getSubmissionAgentCommissionKey(sub));
        });
        return keys.size;
    }
    return 0;
}

function updateAdminNavigationCounts() {
    const userManagementCount = getAdminSubTabCount('users') + getAdminSubTabCount('pending-users') + getAdminSubTabCount('pending-agents') + getAdminSubTabCount('registered-agents');
    const applicationManagementCount = allSubmissions.length;
    setCountBadge('userManagementCount', userManagementCount);
    setCountBadge('applicationManagementCount', applicationManagementCount);
}

const REVIEWER_ROLE_ALIASES = new Set(['reviewer']);
const SUPER_ADMIN_ONLY_ACTIONS = new Set([
    'admin_management_updated',
    'application_redirected',
    'round_robin_counter_reset',
    'super_admin_settings_updated'
]);

function normalizeUserRole(role) {
    const normalized = String(role || '').trim().toLowerCase();
    if (REVIEWER_ROLE_ALIASES.has(normalized)) return 'reviewer';
    return normalized || 'uploader';
}

function getRoleLabel(role) {
    const normalized = normalizeUserRole(role);
    if (normalized === 'super_admin') return 'Super Admin';
    if (normalized === 'reports_monitoring') return 'Reports Monitoring';
    if (normalized === 'rsa') return 'RSA';
    if (normalized === 'payment') return 'Payment';
    return normalized.charAt(0).toUpperCase() + normalized.slice(1);
}

async function ensureCurrentUserProfileAtUid(user, profileDocSnap) {
    if (!user?.uid || !profileDocSnap?.exists()) return profileDocSnap;
    if (profileDocSnap.id === user.uid) return profileDocSnap;

    const profileData = profileDocSnap.data() || {};
    const normalizedEmail = String(user.email || '').trim().toLowerCase();
    const profileEmail = String(profileData.email || '').trim().toLowerCase();

    if (!normalizedEmail || profileEmail !== normalizedEmail) return profileDocSnap;
    if (String(profileData.uid || '').trim() !== user.uid) return profileDocSnap;

    const uidRef = doc(db, 'users', user.uid);
    const uidSnap = await getDoc(uidRef);
    if (uidSnap.exists()) return uidSnap;

    await setDoc(uidRef, {
        ...profileData,
        uid: user.uid,
        email: normalizedEmail,
        updatedAt: serverTimestamp(),
        migratedFromDocId: profileDocSnap.id
    }, { merge: true });

    return await getDoc(uidRef);
}

async function ensureCurrentAdminWritableProfile() {
    if (!currentAdmin?.uid || !currentAdmin?.email) return null;

    const uidRef = doc(db, 'users', currentAdmin.uid);
    const uidSnap = await getDoc(uidRef);
    if (uidSnap.exists()) return uidSnap;

    const normalizedEmail = String(currentAdmin.email || '').trim().toLowerCase();
    const normalizedRole = normalizeUserRole(currentAdminProfileData?.role || 'admin');
    const normalizedStatus = String(currentAdminProfileData?.status || 'active').trim().toLowerCase();
    const allowedRoles = new Set(['uploader', 'admin', 'super_admin', 'reviewer', 'rsa', 'payment']);
    const allowedStatuses = new Set(['pending', 'active', 'deactivated']);

    await setDoc(uidRef, {
        ...(currentAdminProfileData || {}),
        uid: currentAdmin.uid,
        email: normalizedEmail,
        role: allowedRoles.has(normalizedRole) ? normalizedRole : 'admin',
        status: allowedStatuses.has(normalizedStatus) ? normalizedStatus : 'active',
        updatedAt: serverTimestamp(),
        migratedFromDocId: currentAdminProfileData?.uid && currentAdminProfileData.uid !== currentAdmin.uid
            ? currentAdminProfileData.uid
            : (currentAdminProfileData?.migratedFromDocId || '')
    }, { merge: true });

    return await getDoc(uidRef);
}

async function getReviewerUsersForRoundRobin() {
    const usersSnap = await getDocs(collection(db, 'users'));
    return usersSnap.docs
        .map((d) => {
            const data = d.data() || {};
            return {
                id: d.id,
                email: data.email,
                fullName: data.fullName || data.email || 'Unknown',
                role: normalizeUserRole(data.role),
                status: String(data.status || 'active').toLowerCase(),
                leaveStatus: String(data.leaveStatus || '').toLowerCase()
            };
        })
        .filter((u) => u.email && u.role === 'reviewer' && u.status !== 'deactivated' && u.leaveStatus !== 'on_leave')
        .sort((a, b) => a.email.localeCompare(b.email));
}

async function getPaymentUsersForRoundRobin() {
    const usersSnap = await getDocs(collection(db, 'users'));
    return usersSnap.docs
        .map((d) => {
            const data = d.data() || {};
            return {
                id: d.id,
                email: data.email,
                fullName: data.fullName || data.email || 'Unknown',
                role: normalizeUserRole(data.role),
                status: String(data.status || 'active').toLowerCase(),
                leaveStatus: String(data.leaveStatus || '').toLowerCase()
            };
        })
        .filter((u) => u.email && u.role === 'payment' && u.status !== 'deactivated' && u.leaveStatus !== 'on_leave')
        .sort((a, b) => a.email.localeCompare(b.email));
}

// ==================== DOM ELEMENTS ====================
const adminName = document.getElementById('adminName');
const adminAvatar = document.getElementById('adminAvatar');
const pageTitle = document.getElementById('pageTitle');
const pendingUserCountBadge = document.getElementById('pendingUserCount');
const pendingDocCountBadge = document.getElementById('pendingDocCount');
const rejectedDocCountBadge = document.getElementById('rejectedDocCount');
const paymentPendingCountBadge = document.getElementById('paymentPendingCount');
const usersTableBody = document.getElementById('usersTableBody');
const pendingUsersGrid = document.getElementById('pendingUsersGrid');
const pendingAgentsTableBody = document.getElementById('pendingAgentsTableBody');
const approvedAgentsTableBody = document.getElementById('approvedAgentsTableBody');
const draftDocsTableBody = document.getElementById('draftDocsTableBody');
const pendingDocsTableBody = document.getElementById('pendingDocsTableBody');
const approvedDocsTableBody = document.getElementById('approvedDocsTableBody');
const rejectedDocsTableBody = document.getElementById('rejectedDocsTableBody');
const escalationsTableBody = document.getElementById('escalationsTableBody');
const paymentsTableBody = document.getElementById('paymentsTableBody');
const agentCommissionTableBody = document.getElementById('agentCommissionTableBody');
const trackAppsTableBody = document.getElementById('trackAppsTableBody');
const trackPrevPageBtn = document.getElementById('trackPrevPageBtn');
const trackNextPageBtn = document.getElementById('trackNextPageBtn');
const trackPageInfo = document.getElementById('trackPageInfo');
const trackJumpPageInput = document.getElementById('trackJumpPageInput');
const trackJumpPageBtn = document.getElementById('trackJumpPageBtn');
const auditTableBody = document.getElementById('auditTableBody');
const notification = document.getElementById('notification');
let notificationTimer = null;
const viewerModal = document.getElementById('viewerModal');
const viewerFileName = document.getElementById('viewerFileName');
const documentViewer = document.getElementById('documentViewer');
const adminRejectionReasonModal = document.getElementById('adminRejectionReasonModal');
const closeAdminRejectionReasonModal = document.getElementById('closeAdminRejectionReasonModal');
const closeAdminRejectionReasonBtn = document.getElementById('closeAdminRejectionReasonBtn');
const adminRejectionReasonCustomerName = document.getElementById('adminRejectionReasonCustomerName');
const adminRejectionReasonHistory = document.getElementById('adminRejectionReasonHistory');
const agentCommissionModal = document.getElementById('agentCommissionModal');
const agentCommissionModalTitle = document.getElementById('agentCommissionModalTitle');
const agentCommissionModalSummary = document.getElementById('agentCommissionModalSummary');
const agentCommissionBreakdownBody = document.getElementById('agentCommissionBreakdownBody');
const agentCommissionSentTabBtn = document.getElementById('agentCommissionSentTabBtn');
const agentCommissionActiveTabBtn = document.getElementById('agentCommissionActiveTabBtn');
const agentCommissionClearedTabBtn = document.getElementById('agentCommissionClearedTabBtn');
const profileNameEl = document.getElementById('profileName');
const profileRegisteredAtEl = document.getElementById('profileRegisteredAt');
const profileEmailEl = document.getElementById('profileEmail');
const profileWhatsappEl = document.getElementById('profileWhatsapp');
const profileLocationEl = document.getElementById('profileLocation');
const profileRoleEl = document.getElementById('profileRole');
const profileStatusEl = document.getElementById('profileStatus');
let currentAgentCommissionGroup = null;
let currentAgentCommissionView = 'sent_to_pfa';
// Admin password reset removed (per request).

function renderProfileTab() {
    if (!profileNameEl && !profileEmailEl && !profileRoleEl && !profileStatusEl) return;
    const fullName = currentAdminProfileData?.fullName || currentAdmin?.displayName || currentAdmin?.email || 'N/A';
    const registeredAt = currentAdminProfileData?.createdAt ? formatDate(currentAdminProfileData.createdAt) : '-';
    const email = currentAdminProfileData?.email || currentAdmin?.email || 'N/A';
    const whatsapp = currentAdminProfileData?.whatsappNumber || currentAdminProfileData?.phone || '-';
    const location = currentAdminProfileData?.location || '-';
    const role = String(currentAdminProfileData?.role || 'admin');
    const status = String(currentAdminProfileData?.status || 'active');
    if (profileNameEl) profileNameEl.textContent = fullName;
    if (profileRegisteredAtEl) profileRegisteredAtEl.textContent = registeredAt;
    if (profileEmailEl) profileEmailEl.textContent = email;
    if (profileWhatsappEl) profileWhatsappEl.textContent = whatsapp;
    if (profileLocationEl) profileLocationEl.textContent = location;
    if (profileRoleEl) profileRoleEl.textContent = role.charAt(0).toUpperCase() + role.slice(1);
    if (profileStatusEl) profileStatusEl.textContent = status.charAt(0).toUpperCase() + status.slice(1);
}

function getRejectionHistoryEntries(submission) {
    const rawHistory = Array.isArray(submission?.rejectionHistory) ? submission.rejectionHistory : [];
    const normalizedHistory = rawHistory
        .map((entry) => {
            if (typeof entry === 'string') {
                const reason = entry.trim();
                return reason ? { reason, rejectedAt: null } : null;
            }
            const reason = String(entry?.reason || '').trim();
            if (!reason) return null;
            return {
                reason,
                rejectedAt: entry?.rejectedAt || null
            };
        })
        .filter(Boolean);

    if (normalizedHistory.length > 0) return normalizedHistory;

    const fallbackReason = String(
        submission?.latestRejectionReason ||
        submission?.previousRejectionReason ||
        submission?.comment ||
        ''
    ).trim();

    return fallbackReason ? [{
        reason: fallbackReason,
        rejectedAt: submission?.latestRejectedAt || submission?.previousRejectedAt || submission?.reviewedAt || null
    }] : [];
}

function getRejectCount(submission) {
    return getRejectionHistoryEntries(submission).length;
}

function isResubmittedSubmission(submission) {
    return String(submission?.status || '').toLowerCase() === 'pending' && (
        submission?.resubmittedAfterRejection === true ||
        Number(submission?.fixCount || 0) > 0 ||
        Boolean(submission?.reuploadedAt)
    );
}

function renderRejectedDocRow(sub) {
    const uploadDate = formatDate(sub.uploadedAt);
    const rejectionTime = formatDate(sub.latestRejectedAt || sub.rejectedAt || sub.reviewedAt || sub.uploadedAt);
    const uploaderFullName = uploaderNames[sub.uploadedBy] || sub.uploadedBy?.split('@')[0] || 'Unknown';
    const assignedToName = sub.assignedTo ? getDisplayNameByEmail(sub.assignedTo) : '-';
    const statusLabel = formatStatusLabel(sub.status || 'rejected');
    const rejectCount = getRejectCount(sub);
    const reasonCell = rejectCount > 0
        ? `<button class="action-btn reason-btn" onclick="window.openAdminRejectionReasonModal('${sub.id}')"><i class="fas fa-eye"></i> View</button>`
        : '-';

    return `
        <tr>
            <td><strong>${sub.customerName || 'Unknown'}</strong></td>
            <td>${uploadDate}</td>
            <td>${uploaderFullName}</td>
            <td>${assignedToName}</td>
            <td><span class="status-badge status-rejected">${statusLabel}</span></td>
            <td>${rejectionTime}</td>
            <td>${rejectCount || 0}</td>
            <td>${reasonCell}</td>
            <td>
                <button class="action-btn view-btn-small" onclick="window.viewSubmissionDocs('${sub.id}')">
                    <i class="fas fa-eye"></i> View All
                </button>
                <button class="action-btn download-all-btn" onclick="window.downloadAllSubmission('${sub.id}')">
                    <i class="fas fa-download"></i> Download All
                </button>
                <button class="action-btn" onclick="window.openApplicationChat('${sub.id}')">
                    <i class="fas fa-comments"></i> Chat
                </button>
            </td>
        </tr>
    `;
}

function ensureRoundRobinUnifiedTab() {
    const roundRobinTab = document.getElementById('round-robinTab');
    if (!roundRobinTab) return;
    if (document.getElementById('reviewerMonitorSection')) return;

    const reviewerMonitor = roundRobinTab.querySelector('.round-robin-monitor');
    const rsaLegacyTab = document.getElementById('rsa-round-robinTab');
    const rsaMonitor = rsaLegacyTab?.querySelector('.round-robin-monitor') || null;

    const monitorShell = document.createElement('div');
    monitorShell.className = 'round-robin-tabs-shell';
    monitorShell.innerHTML = `
        <div style="display:flex;gap:10px;flex-wrap:wrap;margin-bottom:14px;">
            <button id="rrReviewerBtn" class="action-btn active" type="button" style="padding:8px 14px;">Reviewer</button>
            <button id="rrRsaBtn" class="action-btn" type="button" style="padding:8px 14px;">RSA</button>
            <button id="rrPaymentBtn" class="action-btn" type="button" style="padding:8px 14px;">Payment</button>
        </div>
        <div id="reviewerMonitorSection"></div>
        <div id="rsaMonitorSection" style="display:none;"></div>
        <div id="paymentMonitorSection" style="display:none;">
            <div class="round-robin-monitor">
                <h2>Payment Round-Robin Distribution Monitor</h2>
                <p style="color:#666;margin-bottom:20px;">Monitor and manage queue distribution to payment officers</p>
                <div style="background:#f8f9fa;padding:20px;border-radius:8px;margin-bottom:25px;">
                    <h3>Current Distribution State</h3>
                    <div style="display:grid;grid-template-columns:repeat(auto-fit,minmax(200px,1fr));gap:15px;margin-top:15px;">
                        <div style="background:white;padding:15px;border-radius:8px;border:1px solid #e5e7eb;"><div style="font-size:12px;color:#999;margin-bottom:5px;">Last Distribution</div><div id="lastDistributionPayment" style="font-size:16px;font-weight:600;color:#003366;">Loading...</div></div>
                        <div style="background:white;padding:15px;border-radius:8px;border:1px solid #e5e7eb;"><div style="font-size:12px;color:#999;margin-bottom:5px;">Current Index</div><div id="currentDistributionIndexPayment" style="font-size:16px;font-weight:600;color:#003366;">Loading...</div></div>
                        <div style="background:white;padding:15px;border-radius:8px;border:1px solid #e5e7eb;"><div style="font-size:12px;color:#999;margin-bottom:5px;">Last Reset Date</div><div id="lastResetDatePayment" style="font-size:16px;font-weight:600;color:#003366;">Loading...</div></div>
                        <div style="background:white;padding:15px;border-radius:8px;border:1px solid #e5e7eb;"><div style="font-size:12px;color:#999;margin-bottom:5px;">Total Payment Officers</div><div id="totalPaymentUserCount" style="font-size:16px;font-weight:600;color:#003366;">Loading...</div></div>
                    </div>
                </div>
                <div style="background:#f8f9fa;padding:20px;border-radius:8px;margin-bottom:25px;">
                    <h3>Distribution Statistics (Today)</h3>
                    <div class="table-container" style="max-height:400px;overflow-y:auto;">
                        <table class="documents-table" style="margin-bottom:0;">
                            <thead><tr><th>Payment Officer Name</th><th>Email</th><th>Assigned Today</th><th>Completed</th><th>Pending</th></tr></thead>
                            <tbody id="distributionStatsPaymentBody"></tbody>
                        </table>
                    </div>
                </div>
                <div style="background:#f8f9fa;padding:20px;border-radius:8px;margin-bottom:25px;">
                    <h3>Recent Payment Assignment History (Last 20)</h3>
                    <div class="table-container" style="max-height:400px;overflow-y:auto;">
                        <table class="documents-table" style="margin-bottom:0;">
                            <thead><tr><th>Timestamp</th><th>Customer</th><th>Assigned To (Payment)</th><th>Assigned By</th></tr></thead>
                            <tbody id="assignmentHistoryPaymentBody"></tbody>
                        </table>
                    </div>
                </div>
                <div style="background:#fff3cd;padding:20px;border-radius:8px;border-left:4px solid #ffc107;">
                    <h3 style="margin-top:0;">Administrator Controls</h3>
                    <p style="color:#333;margin:10px 0;">Use these tools to reset payment distribution or test assignment order</p>
                    <div style="display:flex;gap:10px;flex-wrap:wrap;margin-top:15px;">
                        <button id="resetRoundRobinPaymentBtn" class="action-btn" style="background:#ef4444;color:white;padding:10px 20px;border:none;border-radius:6px;cursor:pointer;font-weight:600;"><i class="fas fa-redo"></i> Reset Counter (Today)</button>
                        <button id="testDistributionPaymentBtn" class="action-btn" style="background:#3b82f6;color:white;padding:10px 20px;border:none;border-radius:6px;cursor:pointer;font-weight:600;"><i class="fas fa-flask"></i> Test Payment Distribution</button>
                    </div>
                </div>
            </div>
        </div>
    `;

    roundRobinTab.innerHTML = '';
    roundRobinTab.appendChild(monitorShell);

    if (reviewerMonitor) {
        const reviewerSection = document.getElementById('reviewerMonitorSection');
        reviewerSection?.appendChild(reviewerMonitor);
    }
    if (rsaMonitor) {
        const rsaSection = document.getElementById('rsaMonitorSection');
        rsaSection?.appendChild(rsaMonitor);
    }
    if (rsaLegacyTab) {
        rsaLegacyTab.remove();
    }

    document.getElementById('rrReviewerBtn')?.addEventListener('click', () => window.switchRoundRobinMonitor('reviewer'));
    document.getElementById('rrRsaBtn')?.addEventListener('click', () => window.switchRoundRobinMonitor('rsa'));
    document.getElementById('rrPaymentBtn')?.addEventListener('click', () => window.switchRoundRobinMonitor('payment'));
    document.getElementById('resetRoundRobinPaymentBtn')?.addEventListener('click', handleResetRoundRobinPayment);
    document.getElementById('testDistributionPaymentBtn')?.addEventListener('click', handleTestDistributionPayment);
}

// ==================== DOCUMENT FETCH HELPER ====================
async function fetchWithCorsFallback(url) {
    const cleanUrl = url?.toString().trim().replace(/[\s\n\r\t]+/g, '');
    if (!cleanUrl) throw new Error('Invalid URL');

    try {
        const response = await fetch(cleanUrl, { mode: 'cors', credentials: 'omit' });
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

// ==================== DOWNLOAD HELPER FUNCTIONS ====================
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
}

async function downloadBlobAsFile(blob, fileName) {
    return saveFileWithLocationPicker(blob, fileName);
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

function showLoader(msg) {
    const loader = document.getElementById('globalLoader');
    const text = document.getElementById('loaderText');
    if (loader && text) {
        text.textContent = msg || 'Processing...';
        loader.style.display = 'flex';
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

async function saveBlobToFolderPicker(blob, defaultFileName, customerName = 'Customer') {
    if (!('showDirectoryPicker' in window)) {
        showNotification('Folder picker not supported. Falling back to save dialog...', 'info');
        await saveFileWithLocationPicker(blob, defaultFileName);
        return true;
    }
    try {
        showNotification('📁 Please select a destination folder...', 'info');
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
        showNotification(`✅ Saved to ${safeCustomerName}/${defaultFileName}`, 'success');
        return true;
    } catch (error) {
        if (error.name === 'AbortError') {
            showNotification('Save cancelled', 'info');
        } else {
            showNotification('Save failed: ' + error.message, 'error');
            await saveFileWithLocationPicker(blob, defaultFileName);
        }
        return false;
    }
}

// ==================== INITIALIZATION ====================
document.addEventListener('DOMContentLoaded', () => {
    checkAdminAuth();
    setupEventListeners();
    setupMobileSidebar();
    setupForceRefreshButtons();
    setupIdleLogout();
});

window.signOutUser = async () => {
    try { await signOut(auth); } catch (e) { }
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

// ==================== CHECK ADMIN AUTH ====================
async function checkAdminAuth() {
    auth.onAuthStateChanged(async (user) => {
        if (!user) {
            window.location.href = 'index.html';
            return;
        }

        try {
            let userData = null;
            let userProfileDoc = null;
            const userByUid = query(collection(db, 'users'), where('uid', '==', user.uid));
            const uidSnapshot = await getDocs(userByUid);

            if (!uidSnapshot.empty) {
                userProfileDoc = uidSnapshot.docs[0];
                userData = userProfileDoc.data();
            } else if (user.email) {
                const normalizedEmail = user.email.toLowerCase();
                const userByEmail = query(collection(db, 'users'), where('email', '==', normalizedEmail));
                const emailSnapshot = await getDocs(userByEmail);
                if (!emailSnapshot.empty) {
                    userProfileDoc = emailSnapshot.docs[0];
                    userData = userProfileDoc.data();
                } else {
                    const allUsersSnapshot = await getDocs(collection(db, 'users'));
                    const matchedDoc = allUsersSnapshot.docs.find((docSnap) => {
                        const data = docSnap.data();
                        return (
                            data?.uid === user.uid ||
                            (typeof data?.email === 'string' && data.email.toLowerCase() === normalizedEmail)
                        );
                    });
                    if (matchedDoc) {
                        userProfileDoc = matchedDoc;
                        userData = matchedDoc.data();
                    }
                }
            }

            if (!userData) {
                showNotification('User profile not found in database.', 'error');
                window.location.href = 'index.html';
                return;
            }

            if (userProfileDoc) {
                userProfileDoc = await ensureCurrentUserProfileAtUid(user, userProfileDoc);
                userData = userProfileDoc.data() || userData;
            }

            if (userData.role === 'super_admin') {
                window.location.href = 'super-admin-dashboard.html';
                return;
            }

            if (userData.role === 'reports_monitoring') {
                window.location.href = 'reports-monitoring-dashboard.html';
                return;
            }

            if (userData.role === 'admin') {
                currentAdmin = user;
                currentAdminProfileData = userData;
                ensureRoundRobinUnifiedTab();
                adminName.textContent = userData.fullName || user.email;
                adminAvatar.src = user.photoURL || 'data:image/svg+xml,%3Csvg xmlns=\'http://www.w3.org/2000/svg\' width=\'40\' height=\'40\' viewBox=\'0 0 40 40\'%3E%3Ccircle cx=\'20\' cy=\'20\' r=\'20\' fill=\'%23003366\'/%3E%3Ctext x=\'20\' y=\'25\' text-anchor=\'middle\' fill=\'%23ffffff\' font-size=\'16\'%3EUser%3C/text%3E%3C/svg%3E';
                renderProfileTab();

                loadUsers();
                loadPendingUsers();
                loadPendingAgents();
                loadApprovedAgents();
                loadSubmissions();
                loadEscalations();
                loadAuditLog();
                loadAdminNames();
                loadUploaderNames();
                switchTab('user-management');
            } else {
                showNotification('Access denied. Admin privileges required.', 'error');
                window.location.href = 'index.html';
            }
        } catch (error) {
            showNotification('Could not validate admin session. Please login again.', 'error');
        }
    });
}

// ==================== LOAD ADMIN NAMES ====================
async function loadAdminNames() {
    const usersQuery = query(collection(db, 'users'), where('role', '==', 'admin'));
    const snapshot = await getDocs(usersQuery);
    snapshot.forEach(doc => {
        const data = doc.data();
        const email = String(data.email || '').toLowerCase();
        if (!email) return;
        adminNames[email] = data.fullName || email.split('@')[0];
        userEmailNameCache.set(email, adminNames[email]);
    });
}

// ==================== LOAD UPLOADER NAMES ====================
async function loadUploaderNames() {
    const usersQuery = query(collection(db, 'users'));
    const snapshot = await getDocs(usersQuery);
    snapshot.forEach(doc => {
        const data = doc.data();
        const email = String(data.email || '').toLowerCase();
        if (!email) return;
        uploaderNames[email] = data.fullName || email.split('@')[0];
        userEmailNameCache.set(email, uploaderNames[email]);
        userIdNameCache.set(doc.id, uploaderNames[email]);
    });
}

// ==================== MOBILE SIDEBAR TOGGLE ====================
function setupMobileSidebar() {
    const menuToggle = document.getElementById('menuToggle');
    const sidebarClose = document.getElementById('sidebarClose');
    const sidebarOverlay = document.getElementById('sidebarOverlay');
    const sidebar = document.getElementById('sidebar');

    function openSidebar() {
        if (sidebar) sidebar.classList.add('active');
        if (sidebarOverlay) sidebarOverlay.classList.add('active');
        document.body.style.overflow = 'hidden';
        if (menuToggle) menuToggle.setAttribute('aria-expanded', 'true');
    }

    function closeSidebar() {
        if (sidebar) sidebar.classList.remove('active');
        if (sidebarOverlay) sidebarOverlay.classList.remove('active');
        document.body.style.overflow = '';
        if (menuToggle) menuToggle.setAttribute('aria-expanded', 'false');
    }

    if (menuToggle) {
        menuToggle.addEventListener('click', (e) => {
            e.stopPropagation();
            if (sidebar && sidebar.classList.contains('active')) {
                closeSidebar();
            } else {
                openSidebar();
            }
        });
    }

    if (sidebarClose) {
        sidebarClose.addEventListener('click', closeSidebar);
    }

    if (sidebarOverlay) {
        sidebarOverlay.addEventListener('click', closeSidebar);
    }

    document.addEventListener('keydown', (e) => {
        if (e.key === 'Escape' && sidebar && sidebar.classList.contains('active')) {
            closeSidebar();
        }
    });

    document.querySelectorAll('.nav-item[data-tab]').forEach(item => {
        item.addEventListener('click', () => {
            if (window.innerWidth <= 768) {
                setTimeout(closeSidebar, 150);
            }
        });
    });

    let resizeTimer;
    window.addEventListener('resize', () => {
        clearTimeout(resizeTimer);
        resizeTimer = setTimeout(() => {
            if (window.innerWidth > 768) {
                if (sidebar) sidebar.classList.remove('active');
                if (sidebarOverlay) sidebarOverlay.classList.remove('active');
                document.body.style.overflow = '';
                if (menuToggle) menuToggle.setAttribute('aria-expanded', 'false');
            }
        }, 250);
    });
}

// ==================== EVENT LISTENERS ====================
function setupEventListeners() {
    document.querySelectorAll('.nav-item[data-tab]').forEach(item => {
        item.addEventListener('click', (e) => {
            e.preventDefault();
            const tabId = item.dataset.tab;
            switchTab(tabId);
        });
    });

    document.getElementById('closeViewUserModal')?.addEventListener('click', closeViewUserModal);
    document.getElementById('closeViewAgentModal')?.addEventListener('click', closeViewAgentModal);
    document.getElementById('closeAgentDetailsBtn')?.addEventListener('click', closeViewAgentModal);
    document.getElementById('userForm')?.addEventListener('submit', saveUser);
    document.getElementById('closeUserModalBtn')?.addEventListener('click', closeUserModal);
    document.getElementById('cancelUserModalBtn')?.addEventListener('click', closeUserModal);
    document.getElementById('closeLeaveModalBtn')?.addEventListener('click', closeLeaveModal);
    document.getElementById('cancelLeaveModalBtn')?.addEventListener('click', closeLeaveModal);
    document.getElementById('confirmLeaveBtn')?.addEventListener('click', confirmActivateLeave);

    document.getElementById('closeConfirmModal')?.addEventListener('click', closeConfirmModal);
    document.getElementById('cancelConfirm')?.addEventListener('click', closeConfirmModal);

    document.getElementById('userSearch')?.addEventListener('input', filterUsers);
    document.getElementById('userRoleFilter')?.addEventListener('change', filterUsers);
    document.getElementById('userStatusFilter')?.addEventListener('change', filterUsers);

    document.getElementById('draftDocSearch')?.addEventListener('input', filterDraftDocs);
    document.getElementById('draftDocDate')?.addEventListener('change', filterDraftDocs);
    document.getElementById('pendingDocSearch')?.addEventListener('input', filterPendingDocs);
    document.getElementById('pendingDocDate')?.addEventListener('change', filterPendingDocs);
    document.getElementById('approvedDocSearch')?.addEventListener('input', filterApprovedDocs);
    document.getElementById('approvedDocDate')?.addEventListener('change', filterApprovedDocs);
    document.getElementById('rejectedDocSearch')?.addEventListener('input', filterRejectedDocs);
    document.getElementById('rejectedDocDate')?.addEventListener('change', filterRejectedDocs);
    closeAdminRejectionReasonModal?.addEventListener('click', closeAdminRejectionReasonModalFn);
    closeAdminRejectionReasonBtn?.addEventListener('click', closeAdminRejectionReasonModalFn);
    document.getElementById('closeAgentCommissionModal')?.addEventListener('click', closeAgentCommissionModalFn);
    document.getElementById('closeAgentCommissionModalFooterBtn')?.addEventListener('click', closeAgentCommissionModalFn);
    agentCommissionSentTabBtn?.addEventListener('click', () => switchAgentCommissionBreakdownTab('sent_to_pfa'));
    agentCommissionActiveTabBtn?.addEventListener('click', () => switchAgentCommissionBreakdownTab('active'));
    agentCommissionClearedTabBtn?.addEventListener('click', () => switchAgentCommissionBreakdownTab('cleared'));

    document.getElementById('auditDate')?.addEventListener('change', filterAudit);
    document.getElementById('auditAction')?.addEventListener('change', filterAudit);

    document.getElementById('trackUserSearch')?.addEventListener('input', () => { trackAppsPage = 1; renderTrackApplications(); });
    document.getElementById('trackStartDate')?.addEventListener('change', () => { trackAppsPage = 1; renderTrackApplications(); });
    document.getElementById('trackEndDate')?.addEventListener('change', () => { trackAppsPage = 1; renderTrackApplications(); });
    document.getElementById('trackStatusFilter')?.addEventListener('change', () => { trackAppsPage = 1; renderTrackApplications(); });
    document.getElementById('trackFilterBtn')?.addEventListener('click', (e) => {
        e.preventDefault();
        trackAppsPage = 1;
        renderTrackApplications();
    });
    trackPrevPageBtn?.addEventListener('click', (e) => {
        e.preventDefault();
        trackAppsPage = Math.max(1, trackAppsPage - 1);
        renderTrackApplications();
    });
    trackNextPageBtn?.addEventListener('click', (e) => {
        e.preventDefault();
        trackAppsPage = trackAppsPage + 1;
        renderTrackApplications();
    });
    trackJumpPageBtn?.addEventListener('click', (e) => {
        e.preventDefault();
        const raw = parseInt(String(trackJumpPageInput?.value || ''), 10);
        if (!raw || raw < 1) return;
        trackAppsPage = raw;
        renderTrackApplications();
    });
    trackJumpPageInput?.addEventListener('keypress', (e) => {
        if (e.key === 'Enter') {
            e.preventDefault();
            trackJumpPageBtn?.click();
        }
    });

    document.getElementById('closeViewer')?.addEventListener('click', closeViewerModal);

    window.addEventListener('click', (e) => {
        if (e.target === viewerModal) closeViewerModal();
        if (e.target === adminRejectionReasonModal) closeAdminRejectionReasonModalFn();
        if (e.target === agentCommissionModal) closeAgentCommissionModalFn();
        const userModal = document.getElementById('userModal');
        if (e.target === userModal) closeUserModal();
        const leaveModal = document.getElementById('leaveModal');
        if (e.target === leaveModal) closeLeaveModal();
        const testResultModal = document.getElementById('testResultModal');
        if (e.target === testResultModal) closeTestResultModal();
    });

    document.getElementById('resetRoundRobinBtn')?.addEventListener('click', handleResetRoundRobin);
    document.getElementById('testDistributionBtn')?.addEventListener('click', handleTestDistribution);
    document.getElementById('resetRoundRobinRSABtn')?.addEventListener('click', handleResetRoundRobinRSA);
    document.getElementById('testDistributionRSABtn')?.addEventListener('click', handleTestDistributionRSA);
    document.getElementById('resetRoundRobinPaymentBtn')?.addEventListener('click', handleResetRoundRobinPayment);
    document.getElementById('testDistributionPaymentBtn')?.addEventListener('click', handleTestDistributionPayment);
    document.getElementById('closeTestResultModal')?.addEventListener('click', closeTestResultModal);
    document.getElementById('closeTestResultBtn')?.addEventListener('click', closeTestResultModal);
}

function closeAdminRejectionReasonModalFn() {
    if (adminRejectionReasonModal) adminRejectionReasonModal.classList.remove('active');
    if (adminRejectionReasonCustomerName) adminRejectionReasonCustomerName.textContent = '-';
    if (adminRejectionReasonHistory) {
        adminRejectionReasonHistory.innerHTML = '';
        adminRejectionReasonHistory.style.display = 'none';
    }
}

function closeAgentCommissionModalFn() {
    currentAgentCommissionGroup = null;
    currentAgentCommissionView = 'sent_to_pfa';
    if (agentCommissionModal) agentCommissionModal.classList.remove('active');
    if (agentCommissionModalTitle) agentCommissionModalTitle.textContent = 'Agent Commission Breakdown';
    if (agentCommissionModalSummary) agentCommissionModalSummary.innerHTML = '';
    if (agentCommissionBreakdownBody) agentCommissionBreakdownBody.innerHTML = '';
    agentCommissionSentTabBtn?.classList.add('active');
    agentCommissionActiveTabBtn?.classList.remove('active');
    agentCommissionClearedTabBtn?.classList.remove('active');
}

// ==================== TAB SWITCHING ====================
function getParentTabForLeaf(tabId) {
    const entries = Object.entries(TAB_GROUPS);
    for (const [parent, leaves] of entries) {
        if (leaves.includes(tabId)) return parent;
    }
    return null;
}

function renderAdminSubTabs(parentTabId) {
    const host = document.getElementById('adminSubTabs');
    if (!host) return;

    const leaves = TAB_GROUPS[parentTabId] || [];
    if (!leaves.length) {
        host.innerHTML = '';
        host.style.display = 'none';
        return;
    }

    host.style.display = 'flex';
    host.classList.add('admin-subtab-strip');
    host.innerHTML = leaves.map((leaf) => {
        const active = leaf === currentLeafTab;
        const label = TAB_LABELS[leaf] || leaf;
        const count = getAdminSubTabCount(leaf);
        const badgeId = `${leaf.replace(/[^a-zA-Z0-9]+/g, '-')}Count`;
        return `
            <button
                type="button"
                class="action-btn admin-subtab-btn ${active ? 'active' : ''}"
                data-subtab="${leaf}"
            >${label}<span class="badge" id="${badgeId}">${count}</span></button>
        `;
    }).join('');

    host.querySelectorAll('[data-subtab]').forEach((btn) => {
        btn.addEventListener('click', () => switchLeafTab(btn.getAttribute('data-subtab')));
    });
}

function runTabEffects(tabId) {
    if (tabId === 'draft-docs') {
        renderDraftDocs();
    }
    if (tabId === 'escalations') {
        renderEscalations();
    }
    if (tabId === 'track-apps') {
        renderTrackApplications();
    }
    if (tabId === 'finally-submitted') {
        renderFinallySubmitted();
    }
    if (tabId === 'payments') {
        renderPaymentQueue();
    }
    if (tabId === 'agent-commissions') {
        renderAgentCommissions();
    }
    if (tabId === 'round-robin') {
        loadRoundRobinMonitor();
        loadRSARoundRobinMonitor();
        loadPaymentRoundRobinMonitor();
        window.switchRoundRobinMonitor('reviewer');
    }
}

function switchLeafTab(tabId) {
    currentLeafTab = tabId;
    document.querySelectorAll('.tab-content').forEach(content => content.classList.remove('active'));
    document.getElementById(`${tabId}Tab`)?.classList.add('active');

    const titles = {
        users: 'User Management - Users',
        'pending-users': 'User Management - Pending Users',
        'pending-agents': 'User Management - Pending Agents',
        'registered-agents': 'User Management - Registered Agents',
        'draft-docs': 'Application Management - Draft',
        'pending-docs': 'Application Management - Pending',
        'approved-docs': 'Application Management - Approved',
        'rejected-docs': 'Application Management - Rejected',
        escalations: 'Application Management - Escalations',
        'track-apps': 'Application Management - Track Applications',
        'finally-submitted': 'Application Management - Final Submission',
        payments: 'Application Management - Payment',
        'agent-commissions': 'Application Management - Agent Commissions',
        audit: 'Audit Log',
        profile: 'My Profile',
        'round-robin': 'Round Robin Monitor',
        help: 'Help & SOP'
    };
    if (pageTitle) pageTitle.textContent = titles[tabId] || 'Admin Dashboard';

    renderAdminSubTabs(currentParentTab);
    runTabEffects(tabId);
}

function switchTab(tabId) {
    const directParent = TAB_GROUPS[tabId] ? tabId : null;
    const inferredParent = getParentTabForLeaf(tabId);
    const parentTab = directParent || inferredParent;

    if (parentTab) {
        currentParentTab = parentTab;
        document.querySelectorAll('.nav-item').forEach(nav => nav.classList.remove('active'));
        document.querySelector(`[data-tab="${parentTab}"]`)?.classList.add('active');

        const defaultLeaf = TAB_GROUPS[parentTab][0];
        const leafToShow = TAB_GROUPS[parentTab].includes(currentLeafTab) ? currentLeafTab : defaultLeaf;
        switchLeafTab(inferredParent ? tabId : leafToShow);
        return;
    }

    currentParentTab = '';
    document.querySelectorAll('.nav-item').forEach(nav => nav.classList.remove('active'));
    document.querySelector(`[data-tab="${tabId}"]`)?.classList.add('active');
    document.querySelectorAll('.tab-content').forEach(content => content.classList.remove('active'));
    document.getElementById(`${tabId}Tab`)?.classList.add('active');
    const titles = {
        payments: 'Application Management - Payment',
        audit: 'Audit Log',
        profile: 'My Profile',
        'round-robin': 'Round Robin Monitor',
        help: 'Help & SOP'
    };
    if (pageTitle) pageTitle.textContent = titles[tabId] || 'Admin Dashboard';
    renderAdminSubTabs('');
    runTabEffects(tabId);
}

window.switchRoundRobinMonitor = (type = 'reviewer') => {
    const reviewerSection = document.getElementById('reviewerMonitorSection');
    const rsaSection = document.getElementById('rsaMonitorSection');
    const paymentSection = document.getElementById('paymentMonitorSection');
    const reviewerBtn = document.getElementById('rrReviewerBtn');
    const rsaBtn = document.getElementById('rrRsaBtn');
    const paymentBtn = document.getElementById('rrPaymentBtn');
    const mode = String(type || '').toLowerCase();
    const isReviewer = mode === 'reviewer';
    const isRSA = mode === 'rsa';
    const isPayment = mode === 'payment';

    if (reviewerSection) reviewerSection.style.display = isReviewer ? 'block' : 'none';
    if (rsaSection) rsaSection.style.display = isRSA ? 'block' : 'none';
    if (paymentSection) paymentSection.style.display = isPayment ? 'block' : 'none';

    [reviewerBtn, rsaBtn, paymentBtn].forEach((btn) => btn?.classList.remove('active'));
    if (isReviewer) reviewerBtn?.classList.add('active');
    if (isRSA) rsaBtn?.classList.add('active');
    if (isPayment) paymentBtn?.classList.add('active');
};

// ==================== LOAD USERS ====================
function loadUsers() {
    const q = query(collection(db, 'users'), orderBy('createdAt', 'desc'));

    onSnapshot(q, (snapshot) => {
        const users = [];
        snapshot.forEach((doc) => {
            const userData = doc.data();
            const normalizedRole = normalizeUserRole(userData.role);
            if (userData.status !== 'pending' && normalizedRole !== 'super_admin') {
                const email = String(userData.email || '').toLowerCase();
                const displayName = userData.fullName || (email ? email.split('@')[0] : 'Unknown');
                if (email) userEmailNameCache.set(email, displayName);
                userIdNameCache.set(doc.id, displayName);
                users.push({
                    id: doc.id,
                    ...userData,
                    fullName: displayName
                });
            }
        });
        renderUsersTable(users);
        allUsers = users;
        updateAdminNavigationCounts();
        renderAdminSubTabs(currentParentTab);
        if (Array.isArray(allAudits) && allAudits.length > 0) {
            renderAuditTable(allAudits);
        }
    }, (error) => {
        showNotification('Error loading users', 'error');
    });
}

// ==================== LOAD PENDING USERS ====================
function loadPendingUsers() {
    const q = query(collection(db, 'users'), where('status', '==', 'pending'));

    onSnapshot(q, (snapshot) => {
        const pendingUsers = [];
        snapshot.forEach((doc) => {
            const data = doc.data() || {};
            if (normalizeUserRole(data.role) === 'super_admin') return;
            pendingUsers.push({ id: doc.id, ...data });
        });
        allPendingUsers = pendingUsers;
        renderPendingUsersGrid(pendingUsers);
        updatePendingUserCount(pendingUsers);
    }, (error) => {
        if (error.code === 'failed-precondition') {
            loadPendingUsersFallback();
        }
    });
}

function loadPendingAgents() {
    const q = query(collection(db, 'agents'), where('status', '==', 'pending'));
    onSnapshot(q, (snapshot) => {
        allPendingAgents = snapshot.docs.map((docSnap) => ({ id: docSnap.id, ...(docSnap.data() || {}) }));
        renderPendingAgentsTable(allPendingAgents);
        updateAdminNavigationCounts();
        renderAdminSubTabs(currentParentTab);
    }, () => {
        allPendingAgents = [];
        renderPendingAgentsTable([]);
        updateAdminNavigationCounts();
        renderAdminSubTabs(currentParentTab);
    });
}

function loadApprovedAgents() {
    const q = query(collection(db, 'agents'), where('status', '==', 'approved'));
    onSnapshot(q, (snapshot) => {
        allApprovedAgents = snapshot.docs.map((docSnap) => ({ id: docSnap.id, ...(docSnap.data() || {}) }));
        renderApprovedAgentsTable(allApprovedAgents);
        renderAgentCommissions();
        updateAdminNavigationCounts();
        renderAdminSubTabs(currentParentTab);
    }, () => {
        allApprovedAgents = [];
        renderApprovedAgentsTable([]);
        renderAgentCommissions();
        updateAdminNavigationCounts();
        renderAdminSubTabs(currentParentTab);
    });
}

function renderPendingAgentsTable(items) {
    if (!pendingAgentsTableBody) return;

    if (!items.length) {
        pendingAgentsTableBody.innerHTML = '<tr><td colspan="5" class="no-data">No pending agent registrations</td></tr>';
        return;
    }

    pendingAgentsTableBody.innerHTML = items.map((agent) => {
        const registeredByRaw = String(agent.createdBy || '').trim();
        const registeredByName = registeredByRaw ? getDisplayNameByEmail(registeredByRaw) : '-';
        const registeredBy = registeredByName && registeredByName !== registeredByRaw ? registeredByName : (registeredByName || registeredByRaw || '-');
        return `
            <tr>
                <td><strong>${escapeHtml(agent.fullName || '-')}</strong></td>
                <td>${escapeHtml(registeredBy)}</td>
                <td>${formatDate(agent.createdAt)}</td>
                <td>
                    <button class="action-btn view-btn pending-agent-btn" onclick="window.viewPendingAgent('${agent.id}')" title="View Agent">
                        <i class="fas fa-eye"></i> View
                    </button>
                </td>
                <td>
                    <div class="action-buttons">
                        <button class="action-btn activate-btn pending-agent-btn" onclick="window.approveAgentRegistration('${agent.id}', this)" title="Approve">
                            <i class="fas fa-check"></i> Approve
                        </button>
                        <button class="action-btn reject-btn pending-agent-btn" onclick="window.rejectAgentRegistration('${agent.id}', this)" title="Reject">
                            <i class="fas fa-times"></i> Reject
                        </button>
                    </div>
                </td>
            </tr>
        `;
    }).join('');
}

function renderApprovedAgentsTable(items) {
    if (!approvedAgentsTableBody) return;

    if (!items.length) {
        approvedAgentsTableBody.innerHTML = '<tr><td colspan="5" class="no-data">No approved agents yet</td></tr>';
        return;
    }

    approvedAgentsTableBody.innerHTML = items.map((agent) => {
        const registeredByRaw = String(agent.createdBy || '').trim();
        const registeredByName = registeredByRaw ? getDisplayNameByEmail(registeredByRaw) : '-';
        const registeredBy = registeredByName && registeredByName !== registeredByRaw ? registeredByName : (registeredByName || registeredByRaw || '-');
        const approvedByRaw = String(agent.approvedBy || '').trim();
        const approvedByName = approvedByRaw ? getDisplayNameByEmail(approvedByRaw) : '-';
        const approvedBy = approvedByName && approvedByName !== approvedByRaw ? approvedByName : (approvedByName || approvedByRaw || '-');

        return `
            <tr>
                <td><strong>${escapeHtml(agent.fullName || '-')}</strong></td>
                <td>${escapeHtml(registeredBy)}</td>
                <td>${escapeHtml(approvedBy)}</td>
                <td>${formatDate(agent.approvedAt || agent.updatedAt || agent.createdAt)}</td>
                <td>
                    <button class="action-btn view-btn pending-agent-btn" onclick="window.viewApprovedAgent('${agent.id}')" title="View Agent">
                        <i class="fas fa-eye"></i> View
                    </button>
                </td>
            </tr>
        `;
    }).join('');
}

async function loadPendingUsersFallback() {
    const q = query(collection(db, 'users'));
    const snapshot = await getDocs(q);
    const pendingUsers = [];

    snapshot.forEach((doc) => {
        const userData = doc.data();
        if (userData.status === 'pending' && normalizeUserRole(userData.role) !== 'super_admin') {
            pendingUsers.push({ id: doc.id, ...userData });
        }
    });

    allPendingUsers = pendingUsers;
    renderPendingUsersGrid(pendingUsers);
    updatePendingUserCount(pendingUsers);
}

// ==================== RENDER PENDING USERS GRID ====================
function renderPendingUsersGrid(pendingUsers) {
    if (!pendingUsersGrid) return;

    if (pendingUsers.length === 0) {
        pendingUsersGrid.innerHTML = `
            <div class="no-pending">
                <i class="fas fa-check-circle"></i>
                <p>No pending user activations at this time</p>
            </div>
        `;
        return;
    }

    pendingUsersGrid.innerHTML = pendingUsers.map(user => {
        const fullName = user.fullName || user.email?.split('@')[0] || 'Unknown';
        const requestedRoleLabel = getRoleLabel(user.role || '');
        const joinDate = user.createdAt ?
            (user.createdAt.toDate ? user.createdAt.toDate() : new Date(user.createdAt)).toLocaleDateString() : 'N/A';

        return `
            <div class="pending-card">
                <div class="pending-card-header">
                    <h3>${fullName}</h3>
                    <span class="pending-badge">Pending Activation</span>
                </div>
                <div class="pending-card-body">
                    <p><i class="fas fa-envelope"></i> ${user.email}</p>
                    <p><i class="fas fa-building"></i> Dept: ${formatDepartment(user.department)}</p>
                    <p><i class="fas fa-phone"></i> ${user.phone || 'N/A'}</p>
                    <p><i class="fas fa-user-tag"></i> Requested Role: ${requestedRoleLabel || 'Not set'}</p>
                    <p><i class="fas fa-clock"></i> Registered: ${joinDate}</p>
                    <div class="pending-role-picker">
                        <label for="pendingRole-${user.id}">Assign Role Before Approval</label>
                        <select id="pendingRole-${user.id}" class="pending-role-select" onchange="window.togglePendingApproval('${user.id}')">
                            <option value="">Select role...</option>
                            <option value="uploader">Uploader</option>
                            <option value="reviewer">Reviewer</option>
                            <option value="payment">Payment</option>
                            <option value="rsa">RSA</option>
                            <option value="admin">Admin</option>
                        </select>
                    </div>
                </div>
                <div class="pending-card-footer">
                    <button id="approvePending-${user.id}" class="approve-btn" onclick="window.activatePendingUser('${user.id}')" disabled title="Select role first">
                        <i class="fas fa-check"></i> Activate
                    </button>
                    <button class="reject-btn" onclick="window.rejectPendingUser('${user.id}')">
                        <i class="fas fa-times"></i> Reject
                    </button>
                </div>
            </div>
        `;
    }).join('');
}

// ==================== RENDER USERS TABLE ====================
function renderUsersTable(users) {
    if (!usersTableBody) return;

    if (users.length === 0) {
        usersTableBody.innerHTML = '<tr><td colspan="8" class="no-data">No users found</td></tr>';
        return;
    }

    usersTableBody.innerHTML = users.map(user => {
        const fullName = user.fullName || user.email?.split('@')[0] || 'Unknown';
        const normalizedRole = normalizeUserRole(user.role);
        const roleLabel = getRoleLabel(user.role);
        const leaveStage = getLeaveStageForRole(user.role);
        const leaveActive = isUserOnLeave(user);
        const leaveReliever = leaveActive ? normalizeEmailValue(user.leaveRelieverEmail) : '';
        const leaveStatusHtml = leaveActive
            ? `<span class="status-badge pending">On Leave${leaveReliever ? ` -> ${escapeHtml(getDisplayNameByEmail(leaveReliever))}` : ''}</span>`
            : `<span class="status-badge ${user.status || 'active'}">${user.status || 'active'}</span>`;
        const canManageLeave = leaveStage && String(user.status || 'active').toLowerCase() === 'active';
        const leaveButtonHtml = canManageLeave
            ? (leaveActive
                ? `<button class="action-btn activate-btn" onclick="window.resumeUserFromLeave('${user.id}')" title="Resume from leave"><i class="fas fa-person-walking-arrow-right"></i></button>`
                : `<button class="action-btn leave-btn" onclick="window.openLeaveModal('${user.id}')" title="Activate leave"><i class="fas fa-person-walking-luggage"></i></button>`)
            : '';
        return `
            <tr data-user-id="${user.id}">
                <td><strong>${fullName}</strong></td>
                <td>${user.email}</td>
                <td>${renderWhatsAppContactCell(user)}</td>
                <td><span class="role-badge ${normalizedRole}">${roleLabel}</span></td>
                <td>${leaveStatusHtml}</td>
                <td>${formatDate(user.createdAt)}</td>
                <td>${formatDate(user.lastLoginAt)}</td>
                <td>
                    <div class="action-buttons">
                        <button class="action-btn view-btn" onclick="window.viewUser('${user.id}')" title="View">
                            <i class="fas fa-eye"></i>
                        </button>
                        ${user.status === 'active' ?
                            `<button class="action-btn deactivate-btn" onclick="window.deactivateUser('${user.id}')" title="Deactivate">
                                <i class="fas fa-ban"></i>
                            </button>` :
                            user.status === 'deactivated' ?
                            `<button class="action-btn activate-btn" onclick="window.activateUser('${user.id}')" title="Activate">
                                <i class="fas fa-check-circle"></i>
                            </button>` : ''
                        }
                        ${leaveButtonHtml}
                        <button class="action-btn delete-btn" onclick="window.deleteUser('${user.id}')" title="Delete">
                            <i class="fas fa-trash"></i>
                        </button>
                    </div>
                </td>
            </tr>
        `;
    }).join('');
}

// ==================== LOAD SUBMISSIONS ====================
function loadSubmissions() {
    const q = query(collection(db, 'submissions'), orderBy('uploadedAt', 'desc'));

    onSnapshot(q, (snapshot) => {
        allSubmissions = [];
        snapshot.forEach((doc) => {
            allSubmissions.push({ id: doc.id, ...doc.data() });
        });

        renderDraftDocs();
        renderPendingDocs();
        renderApprovedDocs();
        renderRejectedDocs();
        renderEscalations();
        renderTrackApplications();
        renderFinallySubmitted();
        renderPaymentQueue();
        renderAgentCommissions();
        updatePendingDocCount(allSubmissions.filter(s => s.status === 'pending'));
        updateRejectedDocCount(allSubmissions.filter(s => ['rejected', 'rejected_by_rsa'].includes(String(s.status || '').toLowerCase())));
        updateFinallySubmittedCount();
        updatePaymentPendingCount();
        updateAdminNavigationCounts();
        renderAdminSubTabs(currentParentTab);
    }, (error) => {
        // Silent fail
    });
}

function loadEscalations() {
    const q = query(collection(db, 'applicationChats'), where('escalated', '==', true));

    onSnapshot(q, (snapshot) => {
        allEscalations = snapshot.docs.map((docSnap) => ({ id: docSnap.id, ...docSnap.data() }));
        renderEscalations();
        updateAdminNavigationCounts();
        renderAdminSubTabs(currentParentTab);
    }, () => {
        allEscalations = [];
        renderEscalations();
    });
}

window.toggleEscalationHandled = async (chatId, handled) => {
    const normalizedId = String(chatId || '').trim();
    if (!normalizedId || !currentAdmin?.email) return;

    try {
        const payload = handled
            ? {
                escalationHandled: true,
                escalationHandledAt: serverTimestamp(),
                escalationHandledBy: String(currentAdmin.email || '').trim().toLowerCase()
            }
            : {
                escalationHandled: false,
                escalationHandledAt: null,
                escalationHandledBy: ''
            };

        await updateDoc(doc(db, 'applicationChats', normalizedId), payload);
        showNotification(handled ? 'Escalation marked as handled' : 'Escalation reopened', 'success');
    } catch (error) {
        showNotification(`Failed to update escalation: ${error?.message || 'Unknown error'}`, 'error');
    }
};

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

// ==================== UPDATE COUNTS ====================
function updatePendingUserCount(users) {
    const pendingCount = users.length;
    setCountBadge('pendingUserCount', pendingCount);
    setCountBadge('pending-usersCount', pendingCount);
    updateAdminNavigationCounts();
    renderAdminSubTabs(currentParentTab);
}

function updatePendingDocCount(items) {
    const pendingCount = Array.isArray(items) ? items.filter(u => u.status === 'pending').length : 0;
    setCountBadge('pendingDocCount', pendingCount);
    setCountBadge('pending-docsCount', pendingCount);
    updateAdminNavigationCounts();
}

function updateRejectedDocCount(items) {
    const rejectedCount = Array.isArray(items) ? items.filter(u => ['rejected', 'rejected_by_rsa'].includes(String(u.status || '').toLowerCase())).length : 0;
    setCountBadge('rejectedDocCount', rejectedCount);
    setCountBadge('rejected-docsCount', rejectedCount);
    updateAdminNavigationCounts();
}

// ==================== RENDER PENDING DOCUMENTS ====================
function renderDraftDocs() {
    filterDraftDocs();
}

function renderPendingDocs() {
    filterPendingDocs();
}

// ==================== RENDER APPROVED DOCUMENTS ====================
function renderApprovedDocs() {
    filterApprovedDocs();
}

// ==================== RENDER REJECTED DOCUMENTS ====================
function renderRejectedDocs() {
    filterRejectedDocs();
}

function formatEscalationStageLabel(stageKey) {
    const normalized = String(stageKey || '').trim().toLowerCase();
    if (normalized === 'rsa') return 'RSA Stage';
    if (normalized === 'payment') return 'Payment Stage';
    return 'Reviewer Stage';
}

function getEscalationSubmission(chatItem) {
    return allSubmissions.find((sub) => sub.id === chatItem.id) || null;
}

function getEscalationDisplayStatus(chatItem) {
    return chatItem?.escalationHandled === true ? 'Handled' : 'Open';
}

function renderEscalations() {
    if (!escalationsTableBody) return;

    const sortedItems = [...allEscalations].sort((a, b) => {
        const handledDiff = Number(Boolean(a?.escalationHandled)) - Number(Boolean(b?.escalationHandled));
        if (handledDiff !== 0) return handledDiff;
        const aTime = a?.escalatedAt?.toMillis ? a.escalatedAt.toMillis() : new Date(a?.escalatedAt || 0).getTime();
        const bTime = b?.escalatedAt?.toMillis ? b.escalatedAt.toMillis() : new Date(b?.escalatedAt || 0).getTime();
        return (Number.isFinite(bTime) ? bTime : 0) - (Number.isFinite(aTime) ? aTime : 0);
    });

    if (!sortedItems.length) {
        escalationsTableBody.innerHTML = '<tr><td colspan="8" class="no-data">No escalated applications</td></tr>';
        return;
    }

    escalationsTableBody.innerHTML = sortedItems.map((chatItem) => {
        const submission = getEscalationSubmission(chatItem);
        const customerName = submission?.customerName || chatItem.customerName || 'Unknown';
        const uploadedBy = getDisplayNameByEmail(submission?.uploadedBy || chatItem.uploadedBy || '');
        const escalatedBy = getDisplayNameByEmail(chatItem.escalatedBy || '');
        const stageLabel = formatEscalationStageLabel(chatItem.currentStage);
        const applicationStatus = formatStatusLabel(submission?.status || chatItem.currentStage || '-');
        const escalatedAt = formatDate(chatItem.escalatedAt || chatItem.updatedAt);
        const handledBy = chatItem.escalationHandledBy ? getDisplayNameByEmail(chatItem.escalationHandledBy) : '-';
        const handledAt = chatItem.escalationHandledAt ? formatDate(chatItem.escalationHandledAt) : '-';
        const handlingText = chatItem.escalationHandled === true
            ? `Handled by ${escapeHtml(handledBy)}<br><small>${escapeHtml(handledAt)}</small>`
            : '<span class="status-badge status-pending">Open</span>';
        const actionBtn = chatItem.escalationHandled === true
            ? `<button class="action-btn" onclick="window.toggleEscalationHandled('${chatItem.id}', false)"><i class="fas fa-rotate-left"></i> Reopen</button>`
            : `<button class="action-btn" onclick="window.toggleEscalationHandled('${chatItem.id}', true)"><i class="fas fa-check"></i> Mark Handled</button>`;

        return `
            <tr>
                <td><strong>${escapeHtml(customerName)}</strong></td>
                <td>${escapeHtml(uploadedBy)}</td>
                <td>${escapeHtml(escalatedBy)}</td>
                <td>${escapeHtml(stageLabel)}</td>
                <td><span class="status-badge ${chatItem?.escalationHandled === true ? 'status-approved' : 'status-rejected'}">${escapeHtml(getEscalationDisplayStatus(chatItem))}</span><br><small>${escapeHtml(applicationStatus)}</small></td>
                <td>${escapeHtml(escalatedAt)}</td>
                <td>${handlingText}</td>
                <td>
                    <button class="action-btn" onclick="window.openApplicationChat('${chatItem.id}')">
                        <i class="fas fa-comments"></i> Open Chat
                    </button>
                    ${actionBtn}
                </td>
            </tr>
        `;
    }).join('');
}

// ==================== RENDER FINALLY SUBMITTED APPLICATIONS ====================
function renderFinallySubmitted() {
    const finallySubmittedTableBody = document.getElementById('finallySubmittedTableBody');
    if (!finallySubmittedTableBody) return;

    const finallySubmitted = allSubmissions.filter(s => s.finalSubmitted === true || s.rsaSubmitted === true);

    if (finallySubmitted.length === 0) {
        finallySubmittedTableBody.innerHTML = '<tr><td colspan="10" class="no-data">No finally submitted applications</td></tr>';
        updateFinallySubmittedCount();
        return;
    }

    finallySubmittedTableBody.innerHTML = finallySubmitted.map(sub => {
        const uploaderEmail = sub.uploadedBy || '';
        const uploaderName = uploaderEmail ? getDisplayNameByEmail(uploaderEmail) : '-';

        const reviewerEmail = sub.reviewedBy || '';
        const reviewerName = reviewerEmail ? getDisplayNameByEmail(reviewerEmail) : '-';

        const rsaEmail = sub.finalSubmittedBy || sub.rsaSubmittedBy || sub.assignedToRSA || '';
        const rsaName = rsaEmail ? getDisplayNameByEmail(rsaEmail) : '-';
        const paymentEmail = sub.assignedToPayment || '';
        const paymentName = paymentEmail ? getDisplayNameByEmail(paymentEmail) : '-';

        const uploadedAt = formatDate(sub.uploadedAt);
        const approvedAt = formatDate(sub.reviewedAt);
        const submittedAt = formatDate(sub.finalSubmittedAt || sub.rsaSubmittedAt || sub.uploadedAt);
        const currentStatus = String(sub.status || '').toLowerCase();
        const statusLabel = currentStatus === 'paid'
            ? 'Paid'
            : currentStatus === 'cleared'
                ? 'Cleared'
                : 'Sent to PFA';

        return `
            <tr>
                <td><strong>${escapeHtml(sub.customerName || 'Unknown')}</strong></td>
                <td>${escapeHtml(uploaderName)}</td>
                <td>${escapeHtml(reviewerName)}</td>
                <td>${escapeHtml(rsaName)}</td>
                <td>${escapeHtml(paymentName)}</td>
                <td>${uploadedAt}</td>
                <td>${approvedAt}</td>
                <td>${submittedAt}</td>
                <td><span class="status-badge status-approved">${statusLabel}</span></td>
                <td>
                    <button class="action-btn view-btn-small" onclick="window.viewSubmissionDocs('${sub.id}')">
                        <i class="fas fa-eye"></i> View All
                    </button>
                    <button class="action-btn download-all-btn" onclick="window.downloadAllSubmission('${sub.id}')">
                        <i class="fas fa-download"></i> Download All
                    </button>
                    <button class="action-btn" onclick="window.openApplicationChat('${sub.id}')">
                        <i class="fas fa-comments"></i> Chat
                    </button>
                </td>
            </tr>
        `;
    }).join('');

    updateFinallySubmittedCount();
}

function renderPaymentQueue() {
    if (!paymentsTableBody) return;

    const paymentQueue = allSubmissions.filter((sub) => {
        const status = String(sub.status || '').toLowerCase();
        const isFinal = sub.finalSubmitted === true || sub.rsaSubmitted === true || status === 'sent_to_pfa' || status === 'rsa_submitted' || status === 'paid';
        return isFinal && status !== 'cleared';
    });

    if (paymentQueue.length === 0) {
        paymentsTableBody.innerHTML = '<tr><td colspan="8" class="no-data">No payment records available</td></tr>';
        return;
    }

    paymentsTableBody.innerHTML = paymentQueue.map((sub) => {
        const { pfa, twentyFive, commission2 } = getSubmissionFinancials(sub);
        const status = String(sub.status || '').toLowerCase();
        const isPaid = status === 'paid';
        const statusLabel = isPaid ? 'Paid' : 'Sent to PFA';
        const uploaderEmail = String(sub.uploadedBy || '').trim();
        const uploaderName = uploaderEmail ? getDisplayNameByEmail(uploaderEmail) : '-';
        const paymentEmail = String(sub.assignedToPayment || '').trim();
        const paymentName = paymentEmail ? getDisplayNameByEmail(paymentEmail) : '-';

        return `
            <tr>
                <td><strong>${escapeHtml(sub.customerName || 'Unknown')}</strong></td>
                <td>${escapeHtml(pfa)}</td>
                <td>${escapeHtml(uploaderName)}</td>
                <td>${escapeHtml(getSubmissionAgentLabel(sub) || '-')}</td>
                <td>${escapeHtml(paymentName)}</td>
                <td>${formatCurrency(twentyFive)}</td>
                <td>${formatCurrency(commission2)}</td>
                <td><span class="status-badge status-approved">${statusLabel}</span></td>
            </tr>
        `;
    }).join('');
}

function getSubmissionHasAgent(sub) {
    return Boolean(String(sub?.agentId || '').trim());
}

function getSubmissionAgentLabel(sub) {
    if (!String(sub?.agentId || '').trim()) return 'No Agent';
    return String(sub?.agentName || '').trim() || 'Unknown Agent';
}

function isSubmissionCommissionTrackable(sub) {
    const status = String(sub?.status || '').toLowerCase();
    return status === 'sent_to_pfa' || status === 'rsa_submitted' || status === 'paid' || status === 'cleared' || sub?.finalSubmitted === true || sub?.rsaSubmitted === true;
}

function getSubmissionAgentCommissionKey(sub) {
    const agentId = String(sub?.agentId || '').trim();
    return agentId ? `agent:${agentId}` : '';
}

function getAgentCommissionKey(agent) {
    const agentId = String(agent?.id || agent?.agentId || '').trim();
    if (agentId) return `agent:${agentId}`;
    return `name:${String(agent?.fullName || agent?.agentName || '').trim().toLowerCase()}`;
}

function buildAgentCommissionGroups() {
    const groups = new Map();

    allApprovedAgents.forEach((agent) => {
        const key = getAgentCommissionKey(agent);
        groups.set(key, {
            key,
            agentId: String(agent.id || agent.agentId || '').trim(),
            agentName: String(agent.fullName || agent.agentName || 'Unknown Agent').trim(),
            registeredByEmail: String(agent.createdBy || '').trim(),
            registeredByName: String(agent.createdBy || '').trim() ? getDisplayNameByEmail(String(agent.createdBy || '').trim()) : '-',
            bankName: String(agent.bankName || agent.accountBank || '').trim() || '-',
            accountNumber: String(agent.accountNumber || agent.agentAccountNumber || '').trim() || '-',
            sentToPfaSubmissions: [],
            activeSubmissions: [],
            clearedSubmissions: [],
            sentToPfaCommission: 0,
            activeCommission: 0,
            clearedCommission: 0,
            totalCommission: 0
        });
    });

    allSubmissions.forEach((sub) => {
        if (!getSubmissionHasAgent(sub) || !isSubmissionCommissionTrackable(sub)) return;

        const key = getSubmissionAgentCommissionKey(sub);
        if (!key) return;
        const existing = groups.get(key) || {
            key,
            agentId: String(sub.agentId || '').trim(),
            agentName: getSubmissionAgentLabel(sub),
            registeredByEmail: String(sub.uploadedBy || '').trim(),
            registeredByName: String(sub.uploadedBy || '').trim() ? getDisplayNameByEmail(String(sub.uploadedBy || '').trim()) : '-',
            bankName: String(sub.agentAccountBank || '').trim() || '-',
            accountNumber: String(sub.agentAccountNumber || '').trim() || '-',
            sentToPfaSubmissions: [],
            activeSubmissions: [],
            clearedSubmissions: [],
            sentToPfaCommission: 0,
            activeCommission: 0,
            clearedCommission: 0,
            totalCommission: 0
        };

        const normalizedStatus = String(sub.status || '').toLowerCase();
        const { commission2 } = getSubmissionFinancials(sub);

        if (normalizedStatus === 'cleared') {
            existing.clearedSubmissions.push(sub);
            existing.clearedCommission += commission2;
        } else if (normalizedStatus === 'paid') {
            existing.activeSubmissions.push(sub);
            existing.activeCommission += commission2;
        } else {
            existing.sentToPfaSubmissions.push(sub);
            existing.sentToPfaCommission += commission2;
        }

        existing.totalCommission = existing.sentToPfaCommission + existing.activeCommission + existing.clearedCommission;
        if (existing.bankName === '-' && String(sub.agentAccountBank || '').trim()) {
            existing.bankName = String(sub.agentAccountBank || '').trim();
        }
        if (existing.accountNumber === '-' && String(sub.agentAccountNumber || '').trim()) {
            existing.accountNumber = String(sub.agentAccountNumber || '').trim();
        }
        if ((!existing.registeredByName || existing.registeredByName === '-') && String(sub.uploadedBy || '').trim()) {
            existing.registeredByEmail = String(sub.uploadedBy || '').trim();
            existing.registeredByName = getDisplayNameByEmail(existing.registeredByEmail);
        }
        groups.set(key, existing);
    });

    return Array.from(groups.values()).sort((a, b) => {
        if (b.totalCommission !== a.totalCommission) return b.totalCommission - a.totalCommission;
        return a.agentName.localeCompare(b.agentName);
    });
}

function renderAgentCommissions() {
    if (!agentCommissionTableBody) return;

    const groups = buildAgentCommissionGroups();
    if (!groups.length) {
        agentCommissionTableBody.innerHTML = '<tr><td colspan="6" class="no-data">No agent commission records available</td></tr>';
        return;
    }

    agentCommissionTableBody.innerHTML = groups.map((group) => `
        <tr>
            <td><strong>${escapeHtml(group.agentName)}</strong></td>
            <td>${escapeHtml(group.registeredByName || group.registeredByEmail || '-')}</td>
            <td>${escapeHtml(group.bankName)}</td>
            <td>${escapeHtml(group.accountNumber)}</td>
            <td><strong>${formatCurrency(group.totalCommission)}</strong></td>
            <td>
                <button class="action-btn view-btn-small" onclick="window.openAgentCommissionModal('${encodeURIComponent(group.key)}')">
                    <i class="fas fa-eye"></i> View
                </button>
            </td>
        </tr>
    `).join('');
}

function renderAgentCommissionModalSummary(group) {
    if (!agentCommissionModalSummary) return;
    agentCommissionModalSummary.innerHTML = `
        <div class="agent-commission-summary-card">
            <span class="agent-commission-summary-label">Agent</span>
            <strong>${escapeHtml(group.agentName)}</strong>
        </div>
        <div class="agent-commission-summary-card">
            <span class="agent-commission-summary-label">Sent to PFA Commission</span>
            <strong>${formatCurrency(group.sentToPfaCommission)}</strong>
        </div>
        <div class="agent-commission-summary-card">
            <span class="agent-commission-summary-label">Active Commission</span>
            <strong>${formatCurrency(group.activeCommission)}</strong>
        </div>
        <div class="agent-commission-summary-card">
            <span class="agent-commission-summary-label">Cleared Commission</span>
            <strong>${formatCurrency(group.clearedCommission)}</strong>
        </div>
        <div class="agent-commission-summary-card">
            <span class="agent-commission-summary-label">Total Commission</span>
            <strong>${formatCurrency(group.totalCommission)}</strong>
        </div>
    `;
}

function renderAgentCommissionBreakdown(view = 'active') {
    if (!agentCommissionBreakdownBody || !currentAgentCommissionGroup) return;

    const list = view === 'cleared'
        ? currentAgentCommissionGroup.clearedSubmissions
        : view === 'active'
            ? currentAgentCommissionGroup.activeSubmissions
            : currentAgentCommissionGroup.sentToPfaSubmissions;
    const emptyLabel = view === 'cleared' ? 'cleared' : view === 'active' ? 'active' : 'sent to PFA';
    const rows = list
        .map((sub) => {
            const { twentyFive, commission2 } = getSubmissionFinancials(sub);
            return `
                <tr>
                    <td><strong>${escapeHtml(sub.customerName || 'Unknown')}</strong></td>
                    <td>${formatCurrency(twentyFive)}</td>
                    <td>${formatCurrency(commission2)}</td>
                    <td>${escapeHtml(formatStatusLabel(sub.status || '-'))}</td>
                </tr>
            `;
        }).join('');

    agentCommissionBreakdownBody.innerHTML = rows || `<tr><td colspan="4" class="no-data">No ${emptyLabel} commission records for this agent</td></tr>`;
}

function switchAgentCommissionBreakdownTab(view = 'sent_to_pfa') {
    currentAgentCommissionView = view === 'cleared' ? 'cleared' : view === 'active' ? 'active' : 'sent_to_pfa';
    agentCommissionSentTabBtn?.classList.toggle('active', currentAgentCommissionView === 'sent_to_pfa');
    agentCommissionActiveTabBtn?.classList.toggle('active', currentAgentCommissionView === 'active');
    agentCommissionClearedTabBtn?.classList.toggle('active', currentAgentCommissionView === 'cleared');
    renderAgentCommissionBreakdown(currentAgentCommissionView);
}

window.openAgentCommissionModal = (groupKey) => {
    const normalizedKey = decodeURIComponent(String(groupKey || '').trim());
    const group = buildAgentCommissionGroups().find((entry) => entry.key === normalizedKey);
    if (!group) {
        showNotification('Agent commission breakdown not found', 'error');
        return;
    }

    currentAgentCommissionGroup = group;
    currentAgentCommissionView = 'sent_to_pfa';
    if (agentCommissionModalTitle) {
        agentCommissionModalTitle.textContent = `${group.agentName} - Commission Breakdown`;
    }
    renderAgentCommissionModalSummary(group);
    switchAgentCommissionBreakdownTab('sent_to_pfa');
    agentCommissionModal?.classList.add('active');
};

// ==================== UPDATE FINALLY SUBMITTED COUNT ====================
function updateFinallySubmittedCount() {
    const cnt = allSubmissions.filter(s => s.finalSubmitted === true || s.rsaSubmitted === true).length;
    const badge = document.getElementById('finallySubmittedCount');
    if (badge) {
        badge.textContent = cnt;
        badge.style.display = 'inline-block';
    }
    setCountBadge('finally-submittedCount', cnt);
    updateAdminNavigationCounts();
}

function updatePaymentPendingCount() {
    const cnt = allSubmissions.filter((s) => {
        const status = String(s.status || '').toLowerCase();
        return status === 'sent_to_pfa' || status === 'rsa_submitted';
    }).length;
    setCountBadge('paymentPendingCount', cnt);
    setCountBadge('paymentsCount', cnt);
    updateAdminNavigationCounts();
}

function parseMoney(value) {
    const raw = String(value ?? '').replace(/[^0-9.\-]/g, '');
    const num = Number(raw);
    return Number.isFinite(num) ? num : 0;
}

function roundDownToNearestThousand(value) {
    const num = Number(value || 0);
    return Math.max(0, Math.floor(num / 1000) * 1000);
}

function formatCurrency(value) {
    const num = Number(value || 0);
    try {
        return num.toLocaleString('en-NG', { style: 'currency', currency: 'NGN', maximumFractionDigits: 2 });
    } catch (e) {
        return `₦${num.toLocaleString()}`;
    }
}

function getSubmissionFinancials(sub) {
    const details = sub?.customerDetails || {};
    const rsaBalance = parseMoney(details.rsaBalance || sub?.rsaBalance || 0);
    const computed25 = roundDownToNearestThousand(rsaBalance * 0.25);
    const stored25 = parseMoney(details.rsa25Percent || sub?.rsa25Percent || 0);
    const twentyFive = stored25 ? roundDownToNearestThousand(stored25) : computed25;
    const commissionRate = resolveSubmissionCommissionRate(sub);
    const commission2 = getSubmissionCommissionAmount(sub, twentyFive);
    const pfa = String(details.pfa || sub?.pfa || '').trim() || '-';
    return { pfa, twentyFive, commission2, commissionRate };
}

function formatStatusLabel(status) {
    const normalized = String(status || '').toLowerCase().trim();
    if (!normalized) return '-';
    if (normalized === 'processing_to_pfa' || normalized === 'approved') return 'Processing to PFA';
    if (normalized === 'sent_to_pfa' || normalized === 'rsa_submitted') return 'Sent to PFA';
    if (normalized === 'rejected_by_rsa') return 'Rejected by RSA';
    if (normalized === 'paid') return 'Paid';
    if (normalized === 'cleared') return 'Cleared';
    return normalized.charAt(0).toUpperCase() + normalized.slice(1);
}

// ==================== GET DISPLAY NAME BY EMAIL ====================
function getDisplayNameByEmail(email) {
    if (!email) return '-';
    const normalized = email.toLowerCase().trim();

    if (userEmailNameCache.has(normalized)) {
        return userEmailNameCache.get(normalized);
    }

    return email.split('@')[0] || email;
}

// ==================== ESCAPE HTML ====================
function escapeHtml(text) {
    if (!text) return '';
    const div = document.createElement('div');
    div.textContent = text;
    return div.innerHTML;
}

function getUserWhatsApp(user = {}) {
    const direct = String(user.whatsappNumber || user.phone || '').trim();
    if (direct) return direct;

    const code = String(user.whatsappCode || '').trim();
    const local = String(user.whatsappLocalNumber || '').trim();
    return `${code}${local}`.trim();
}

function renderWhatsAppContactCell(user = {}) {
    const raw = getUserWhatsApp(user);
    const digits = raw.replace(/\D/g, '');
    if (!digits) return '-';
    return `<a href="https://wa.me/${digits}" target="_blank" rel="noopener noreferrer">${escapeHtml(raw)}</a>`;
}

function normalizeEmailValue(email) {
    return String(email || '').trim().toLowerCase();
}

function getLeaveStageForRole(role) {
    const normalizedRole = normalizeUserRole(role);
    if (normalizedRole === 'reviewer') {
        return { key: 'reviewer', label: 'Reviewer', assignmentField: 'assignedTo', relieverRoles: ['reviewer'] };
    }
    if (normalizedRole === 'rsa') {
        return { key: 'rsa', label: 'RSA', assignmentField: 'assignedToRSA', relieverRoles: ['rsa'] };
    }
    if (normalizedRole === 'payment') {
        return { key: 'payment', label: 'Payment', assignmentField: 'assignedToPayment', relieverRoles: ['payment'] };
    }
    return null;
}

function isUserOnLeave(user = {}) {
    return String(user.leaveStatus || '').toLowerCase() === 'on_leave';
}

function getEligibleRelievers(user = {}) {
    const stage = getLeaveStageForRole(user.role);
    const userEmail = normalizeEmailValue(user.email);
    if (!stage || !userEmail) return [];
    const allowedRoles = new Set(stage.relieverRoles);
    return allUsers
        .filter((candidate) => {
            const role = normalizeUserRole(candidate.role);
            const email = normalizeEmailValue(candidate.email);
            const status = String(candidate.status || 'active').toLowerCase();
            return allowedRoles.has(role) && email && email !== userEmail && status !== 'deactivated' && !isUserOnLeave(candidate);
        })
        .sort((a, b) => normalizeEmailValue(a.email).localeCompare(normalizeEmailValue(b.email)));
}

function isSubmissionFinalizedForLeave(stageKey, sub = {}) {
    const status = String(sub.status || '').toLowerCase();
    if (stageKey === 'reviewer') return status !== 'pending';
    if (stageKey === 'rsa') return Boolean(sub.finalSubmitted || sub.rsaSubmitted) || ['sent_to_pfa', 'rsa_submitted', 'paid', 'cleared'].includes(status);
    if (stageKey === 'payment') return ['paid', 'cleared'].includes(status);
    return true;
}

function isSubmissionActiveForLeave(stageKey, sub = {}, userEmail = '') {
    if (!stageKey || !userEmail) return false;
    if (stageKey === 'reviewer') return normalizeEmailValue(sub.assignedTo) === userEmail && String(sub.status || '').toLowerCase() === 'pending';
    if (stageKey === 'rsa') {
        const status = String(sub.status || '').toLowerCase();
        return normalizeEmailValue(sub.assignedToRSA) === userEmail && ['processing_to_pfa', 'approved'].includes(status) && !sub.finalSubmitted && !sub.rsaSubmitted;
    }
    if (stageKey === 'payment') {
        const status = String(sub.status || '').toLowerCase();
        return normalizeEmailValue(sub.assignedToPayment) === userEmail && ['sent_to_pfa', 'rsa_submitted'].includes(status);
    }
    return false;
}

function getLeaveActionLabel(user = {}) {
    if (!getLeaveStageForRole(user.role)) return '';
    return isUserOnLeave(user) ? 'Resume from Leave' : 'Activate Leave';
}

// ==================== VIEW ALL DOCUMENTS FOR A SUBMISSION ====================
window.viewSubmissionDocs = async (submissionId) => {
    const sub = await getLatestSubmissionById(submissionId);
    const docs = getEffectiveSubmissionDocuments(sub);
    if (!sub || docs.length === 0) {
        showNotification('No documents available', 'error');
        return;
    }

    const firstDoc = docs[0];
    const docTypeLabel = DOCUMENT_TYPES[firstDoc.documentType] || firstDoc.documentType || 'Document';

    if (viewerModal && viewerFileName && documentViewer) {
        viewerFileName.textContent = `${sub.customerName} - ${docTypeLabel}`;
        documentViewer.src = buildViewerDocumentUrl(firstDoc.fileUrl, sub, 0);
        viewerModal.classList.add('active');
    }

    if (docs.length > 1) {
        let currentIndex = 0;

        const showDoc = (index) => {
            const doc = docs[index];
            const docTypeLabel = DOCUMENT_TYPES[doc.documentType] || doc.documentType || 'Document';
            viewerFileName.textContent = `${sub.customerName} - ${docTypeLabel} (${index + 1}/${docs.length})`;
            documentViewer.src = buildViewerDocumentUrl(doc.fileUrl, sub, index);
        };

        const addViewerNav = () => {
            const viewerHeader = document.querySelector('.viewer-header');
            if (!viewerHeader) return;

            const existingNav = viewerHeader.querySelector('.viewer-nav');
            if (existingNav) existingNav.remove();

            const nav = document.createElement('div');
            nav.className = 'viewer-nav';
            nav.style.cssText = 'display: flex; gap: 10px; align-items: center;';
            nav.innerHTML = `
                <button id="prevDoc" class="action-btn" ${currentIndex === 0 ? 'disabled' : ''}>
                    <i class="fas fa-chevron-left"></i> Prev
                </button>
                <span id="docCounter" style="font-size: 14px; color: #666;">${currentIndex + 1}/${docs.length}</span>
                <button id="nextDoc" class="action-btn" ${currentIndex === docs.length - 1 ? 'disabled' : ''}>
                    <i class="fas fa-chevron-right"></i> Next
                </button>
            `;

            viewerHeader.insertBefore(nav, document.getElementById('closeViewer'));

            document.getElementById('prevDoc').onclick = () => {
                if (currentIndex > 0) {
                    currentIndex--;
                    showDoc(currentIndex);
                    addViewerNav();
                }
            };

            document.getElementById('nextDoc').onclick = () => {
                if (currentIndex < docs.length - 1) {
                    currentIndex++;
                    showDoc(currentIndex);
                    addViewerNav();
                }
            };
        };

        addViewerNav();
    }
};

// ==================== DOWNLOAD ALL SUBMISSION WITH PROGRESS ====================
window.downloadAllSubmission = async (submissionId) => {
    const sub = await getLatestSubmissionById(submissionId);
    if (!sub) {
        showNotification('Submission not found', 'error');
        return;
    }

    const docs = getEffectiveSubmissionDocuments(sub);
    if (!docs.length) {
        showNotification('No documents available for this application', 'warning');
        return;
    }

    const safeCustomerName = (sub.customerName || 'Customer')
        .replace(/[^a-zA-Z0-9\s_-]/g, '_')
        .trim() || 'Customer';

    try {
        // Create and show progress modal
        const progressModal = document.createElement('div');
        progressModal.id = 'downloadProgressModal';
        progressModal.className = 'modal';
        progressModal.innerHTML = `
            <div class="modal-content" style="max-width: 400px;">
                <div class="modal-header">
                    <h2>Downloading Documents</h2>
                </div>
                <div class="modal-body" style="padding: 20px;">
                    <div style="margin-bottom: 20px;">
                        <div style="display: flex; justify-content: space-between; margin-bottom: 8px;">
                            <span id="downloadStatus">Preparing download...</span>
                            <span id="downloadPercentage">0%</span>
                        </div>
                        <div style="width: 100%; height: 8px; background: #e5e7eb; border-radius: 4px; overflow: hidden;">
                            <div id="downloadProgressBar" style="width: 0%; height: 100%; background: #003366; transition: width 0.3s ease;"></div>
                        </div>
                    </div>
                    <div id="currentFileInfo" style="font-size: 12px; color: #666; word-break: break-all; margin-bottom: 20px;">
                        Initializing...
                    </div>
                    <div style="display: flex; justify-content: flex-end;">
                        <button id="cancelDownloadBtn" class="cancel-btn" style="padding: 8px 16px;">Cancel</button>
                    </div>
                </div>
            </div>
        `;
        document.body.appendChild(progressModal);

        // Show the modal
        setTimeout(() => progressModal.classList.add('active'), 10);

        let cancelled = false;
        const cancelBtn = document.getElementById('cancelDownloadBtn');
        cancelBtn.addEventListener('click', () => {
            cancelled = true;
            showNotification('Download cancelled', 'info');
            progressModal.classList.remove('active');
            setTimeout(() => progressModal.remove(), 300);
        });

        // Update progress function
        const updateProgress = (current, total, status, fileName = '') => {
            const percentage = Math.round((current / total) * 100);
            document.getElementById('downloadPercentage').textContent = `${percentage}%`;
            document.getElementById('downloadProgressBar').style.width = `${percentage}%`;
            document.getElementById('downloadStatus').textContent = status || `Downloading ${current} of ${total}`;
            if (fileName) {
                document.getElementById('currentFileInfo').innerHTML = `
                    <strong>Current file:</strong> ${fileName}<br>
                    <strong>Progress:</strong> ${current}/${total} documents
                `;
            }
        };

        updateProgress(0, docs.length, 'Preparing download...');

        let customerFolder = null;
        let useFolderPicker = false;

        // Try folder picker if supported
        if ('showDirectoryPicker' in window && !cancelled) {
            try {
                updateProgress(0, docs.length, 'Waiting for folder selection...');
                const rootFolder = await window.showDirectoryPicker({
                    mode: 'readwrite',
                    startIn: 'downloads'
                });
                if (!cancelled) {
                    customerFolder = await rootFolder.getDirectoryHandle(safeCustomerName, { create: true });
                    useFolderPicker = true;
                    updateProgress(0, docs.length, 'Folder selected, starting download...');
                }
            } catch (folderError) {
                if (folderError.name === 'AbortError') {
                    showNotification('Download cancelled', 'info');
                    progressModal.classList.remove('active');
                    setTimeout(() => progressModal.remove(), 300);
                    return;
                }
            }
        }

        let successCount = 0;
        let failedCount = 0;
        let directOpenedCount = 0;

        // Download each document
        for (let i = 0; i < docs.length; i++) {
            if (cancelled) break;

            const docItem = docs[i];
            if (!docItem?.fileUrl) {
                failedCount++;
                updateProgress(i + 1, docs.length, `Skipping document ${i + 1} (no URL)...`);
                continue;
            }

            try {
                const docType = DOCUMENT_TYPES[docItem.documentType] || docItem.documentType || 'document';
                const fileExt = (docItem.name || '').split('.').pop() || 'pdf';
                const fileName = `${safeCustomerName}_${docType}.${fileExt}`.replace(/[\\/:*?"<>|]/g, '_');

                updateProgress(i, docs.length, `Downloading document ${i + 1} of ${docs.length}...`, docItem.name || fileName);

                const response = await fetchWithCorsFallback(docItem.fileUrl);

                if (!response.ok) {
                    throw new Error(`HTTP ${response.status}`);
                }

                // Get content length if available for more accurate progress
                const contentLength = response.headers.get('content-length');
                let loadedBytes = 0;
                const totalBytes = contentLength ? parseInt(contentLength) : null;

                if (totalBytes && useFolderPicker) {
                    // Stream with progress for folder picker
                    const reader = response.body.getReader();
                    const chunks = [];

                    while (true) {
                        if (cancelled) {
                            reader.cancel();
                            break;
                        }
                        const { done, value } = await reader.read();
                        if (done) break;
                        chunks.push(value);
                        loadedBytes += value.length;
                        const percent = Math.round((loadedBytes / totalBytes) * 100);
                        updateProgress(i, docs.length,
                            `Downloading ${docType}... ${percent}%`,
                            docItem.name || fileName
                        );
                    }

                    if (cancelled) break;

                    const blob = new Blob(chunks);

                    // Save to folder
                    const fileHandle = await customerFolder.getFileHandle(fileName, { create: true });
                    const writable = await fileHandle.createWritable();
                    await writable.write(blob);
                    await writable.close();
                } else {
                    // Simple download without progress for individual files
                    const blob = await response.blob();

                    if (useFolderPicker) {
                        const fileHandle = await customerFolder.getFileHandle(fileName, { create: true });
                        const writable = await fileHandle.createWritable();
                        await writable.write(blob);
                        await writable.close();
                    } else {
                        await downloadBlobAsFile(blob, fileName);
                    }
                }

                successCount++;
                updateProgress(i + 1, docs.length, `Completed ${i + 1} of ${docs.length}`, '');

            } catch (docError) {
                const docType = DOCUMENT_TYPES[docItem.documentType] || docItem.documentType || 'document';
                const fileExt = (docItem.name || '').split('.').pop() || 'pdf';
                const fileName = `${safeCustomerName}_${docType}.${fileExt}`.replace(/[\\/:*?"<>|]/g, '_');

                if (openDirectDocumentDownload(docItem.fileUrl, fileName)) {
                    directOpenedCount++;
                    updateProgress(i + 1, docs.length, `Opened document ${i + 1} directly`, docItem.name || fileName);
                } else {
                    failedCount++;
                    updateProgress(i + 1, docs.length, `Failed document ${i + 1}`, docItem.name || 'Unknown');
                    showNotification(`Failed to download: ${docItem.name}`, 'error');
                }
            }
        }

        // Close progress modal
        progressModal.classList.remove('active');
        setTimeout(() => progressModal.remove(), 300);

        if (!cancelled) {
            if (failedCount === 0) {
                if (directOpenedCount > 0) {
                    showNotification(`Downloaded ${successCount}, opened ${directOpenedCount} directly because storage blocked secure fetch`, 'warning');
                } else
                if (useFolderPicker) {
                    showNotification(`✅ All ${successCount} documents saved to folder: ${safeCustomerName}`, 'success');
                } else {
                    showNotification(`✅ All ${successCount} documents downloaded successfully`, 'success');
                }
            } else {
                showNotification(`Downloaded ${successCount}, opened ${directOpenedCount} directly, ${failedCount} failed`, 'warning');
            }
        }

    } catch (error) {
        if (error?.name === 'AbortError') {
            showNotification('Download cancelled', 'info');
        } else {
            showNotification('Download failed: ' + (error?.message || 'Unknown error'), 'error');
        }

        // Remove progress modal if it exists
        const progressModal = document.getElementById('downloadProgressModal');
        if (progressModal) {
            progressModal.classList.remove('active');
            setTimeout(() => progressModal.remove(), 300);
        }
    }
};

// ==================== PENDING USER ACTIONS ====================
window.togglePendingApproval = (userId) => {
    const selectEl = document.getElementById(`pendingRole-${userId}`);
    const approveBtn = document.getElementById(`approvePending-${userId}`);
    if (!approveBtn) return;

    const hasRole = Boolean(selectEl?.value);
    approveBtn.disabled = !hasRole;
    approveBtn.title = hasRole ? 'Activate user' : 'Select role first';
};

window.activatePendingUser = (userId) => {
    const selectEl = document.getElementById(`pendingRole-${userId}`);
    const selectedRole = (selectEl?.value || '').trim().toLowerCase();
    const roleToSet = normalizeUserRole(selectedRole);
    const validRoles = new Set(['uploader', 'reviewer', 'rsa', 'payment', 'admin']);

    if (!selectedRole) {
        showNotification('Please select a role before approving this user', 'warning');
        return;
    }
    if (!validRoles.has(roleToSet)) {
        showNotification('Invalid role selected', 'error');
        return;
    }

    const roleLabel = getRoleLabel(roleToSet);

    showConfirmModal(`Activate User as ${roleLabel}?`, 'They will be able to login immediately after approval.', async () => {
        closeConfirmModal();
        try {
            await runWithButtonSpinner(`approvePending-${userId}`, 'Approving...', async () => {
                const updateData = {
                    status: 'active',
                    role: roleToSet,
                    approvedAt: serverTimestamp(),
                    approvedBy: currentAdmin?.email
                };

                await updateDoc(doc(db, 'users', userId), updateData);

                const auditEntry = {
                    action: 'user_approved',
                    userId: userId,
                    newRole: roleToSet,
                    performedBy: currentAdmin?.email,
                    timestamp: serverTimestamp()
                };
                await addDoc(collection(db, 'audit'), auditEntry);
            });
            showNotification('User activated successfully', 'success');
        } catch (error) {
            showNotification('Failed to activate user', 'error');
        }
    });
};

window.rejectPendingUser = (userId) => {
    showConfirmModal('Reject User', 'Are you sure you want to reject this user? This will delete their account permanently.', async () => {
        try {
            const userDoc = await getDoc(doc(db, 'users', userId));
            const userData = userDoc.data();

            await deleteDoc(doc(db, 'users', userId));

            await addDoc(collection(db, 'audit'), {
                action: 'user_rejected',
                userEmail: userData?.email,
                performedBy: currentAdmin?.email,
                timestamp: serverTimestamp()
            });

            showNotification('User rejected and removed', 'success');
            closeConfirmModal();
        } catch (error) {
            showNotification('Failed to reject user', 'error');
        }
    });
};

function openAgentDetailsModal(selected) {
    if (!selected) {
        showNotification('Agent record not found', 'error');
        return;
    }

    const detailsDiv = document.getElementById('agentDetails');
    const modal = document.getElementById('viewAgentModal');
    if (!detailsDiv || !modal) return;
    const registeredByRaw = String(selected.createdBy || '').trim();
    const registeredByName = registeredByRaw ? getDisplayNameByEmail(registeredByRaw) : '-';
    const registeredBy = registeredByName && registeredByName !== registeredByRaw ? registeredByName : (registeredByName || registeredByRaw || '-');
    const statusRaw = String(selected.status || 'pending').toLowerCase().trim();
    const normalizedAgentStatus = statusRaw === 'processing_to_pfa' ? 'approved' : statusRaw;
    const statusClass = normalizedAgentStatus === 'approved'
        ? 'status-approved'
        : (normalizedAgentStatus === 'rejected' ? 'status-rejected' : 'pending');
    const statusLabel = normalizedAgentStatus === 'approved'
        ? 'Approved'
        : (normalizedAgentStatus === 'rejected' ? 'Rejected' : 'Pending');

    detailsDiv.innerHTML = `
        <div class="detail-section">
            <h4>Agent Information</h4>
            <div class="detail-row">
                <span class="detail-label">Full Name:</span>
                <span class="detail-value">${escapeHtml(selected.fullName || '-')}</span>
            </div>
            <div class="detail-row">
                <span class="detail-label">Contact Number:</span>
                <span class="detail-value">${escapeHtml(selected.contactNumber || '-')}</span>
            </div>
            <div class="detail-row">
                <span class="detail-label">Account Number:</span>
                <span class="detail-value">${escapeHtml(selected.accountNumber || '-')}</span>
            </div>
            <div class="detail-row">
                <span class="detail-label">Account Bank:</span>
                <span class="detail-value">${escapeHtml(selected.accountBank || '-')}</span>
            </div>
        </div>
        <div class="detail-section">
            <h4>Registration Information</h4>
            <div class="detail-row">
                <span class="detail-label">Registered By:</span>
                <span class="detail-value">${escapeHtml(registeredBy)}</span>
            </div>
            <div class="detail-row">
                <span class="detail-label">Submitted At:</span>
                <span class="detail-value">${formatDate(selected.createdAt)}</span>
            </div>
            <div class="detail-row">
                <span class="detail-label">Status:</span>
                <span class="detail-value"><span class="status-badge ${statusClass}">${escapeHtml(statusLabel)}</span></span>
            </div>
        </div>
    `;

    modal.classList.add('active');
}

window.viewPendingAgent = (agentId) => {
    const selected = allPendingAgents.find((a) => a.id === agentId);
    openAgentDetailsModal(selected);
};

window.viewApprovedAgent = (agentId) => {
    const selected = allApprovedAgents.find((a) => a.id === agentId);
    openAgentDetailsModal(selected);
};

function setAgentActionButtonLoading(button, text) {
    if (!button) return () => {};
    const row = button.closest('tr');
    const rowButtons = row ? Array.from(row.querySelectorAll('button')) : [button];
    const previousDisabled = rowButtons.map((btn) => ({ btn, disabled: btn.disabled }));
    const originalHtml = button.innerHTML;

    rowButtons.forEach((btn) => { btn.disabled = true; });
    button.classList.add('loading');
    button.innerHTML = `<i class="fas fa-spinner fa-spin"></i> ${escapeHtml(text || 'Processing...')}`;

    return () => {
        previousDisabled.forEach(({ btn, disabled }) => { btn.disabled = disabled; });
        button.classList.remove('loading');
        button.innerHTML = originalHtml;
    };
}

function showAgentActionSuccessModal(title, message) {
    showTestResultModal(title, [
        { label: 'Status', value: 'PASSED', ok: true },
        { label: 'Result', value: message, ok: true }
    ]);
}

window.approveAgentRegistration = async (agentId, btnEl = null) => {
    const selected = allPendingAgents.find((a) => a.id === agentId);
    if (!selected) {
        showNotification('Agent record not found', 'error');
        return;
    }
    showConfirmModal('Approve Agent Registration', `Approve ${selected.fullName || 'this agent'} for use in submissions?`, async () => {
        closeConfirmModal();
        const restoreButton = setAgentActionButtonLoading(btnEl, 'Approving...');
        try {
            await updateDoc(doc(db, 'agents', agentId), {
                status: 'approved',
                approvedAt: serverTimestamp(),
                approvedBy: currentAdmin?.email || ''
            });
            await addDoc(collection(db, 'audit'), {
                action: 'agent_registration_approved',
                agentId,
                agentName: selected.fullName || '',
                performedBy: currentAdmin?.email || '',
                timestamp: serverTimestamp()
            });
            showAgentActionSuccessModal('Agent Approval Successful', `${selected.fullName || 'Agent'} has been approved.`);
        } catch (_) {
            showNotification('Failed to approve agent', 'error');
        } finally {
            restoreButton();
        }
    });
};

window.rejectAgentRegistration = async (agentId, btnEl = null) => {
    const selected = allPendingAgents.find((a) => a.id === agentId);
    if (!selected) {
        showNotification('Agent record not found', 'error');
        return;
    }
    showConfirmModal('Reject Agent Registration', `Reject ${selected.fullName || 'this agent'}?`, async () => {
        closeConfirmModal();
        const restoreButton = setAgentActionButtonLoading(btnEl, 'Rejecting...');
        try {
            await updateDoc(doc(db, 'agents', agentId), {
                status: 'rejected',
                rejectedAt: serverTimestamp(),
                rejectedBy: currentAdmin?.email || ''
            });
            await addDoc(collection(db, 'audit'), {
                action: 'agent_registration_rejected',
                agentId,
                agentName: selected.fullName || '',
                performedBy: currentAdmin?.email || '',
                timestamp: serverTimestamp()
            });
            showAgentActionSuccessModal('Agent Rejected', `${selected.fullName || 'Agent'} has been rejected.`);
        } catch (_) {
            showNotification('Failed to reject agent', 'error');
        } finally {
            restoreButton();
        }
    });
};

// ==================== LOAD AUDIT LOG ====================
function loadAuditLog() {
    const q = query(collection(db, 'audit'), orderBy('timestamp', 'desc'), limit(500));

    onSnapshot(q, async (snapshot) => {
        const audits = [];
        for (const docSnap of snapshot.docs) {
            const data = docSnap.data();
            let adminName = data.performedBy || 'System';
            if (data.performedBy && adminNames[data.performedBy]) {
                adminName = adminNames[data.performedBy];
            }
            if (SUPER_ADMIN_ONLY_ACTIONS.has(String(data.action || '').trim())) continue;
            audits.push({ id: docSnap.id, ...data, adminName });
        }
        renderAuditTable(audits);
        allAudits = audits;
    }, (error) => {
        // Silent fail
    });
}

// ==================== RENDER AUDIT TABLE ====================
function renderAuditTable(audits) {
    if (!auditTableBody) return;

    if (audits.length === 0) {
        auditTableBody.innerHTML = '<tr><td colspan="5" class="no-data">No audit records found</td></tr>';
        return;
    }

    auditTableBody.innerHTML = audits.map(audit => `
        <tr>
            <td>${formatDate(audit.timestamp)}</td>
            <td><span class="audit-action-badge">${formatAuditAction(audit.action)}</span></td>
            <td>${formatAuditDescription(audit)}</td>
            <td>${audit.performedBy || 'System'}</td>
            <td><strong>${audit.adminName || 'N/A'}</strong></td>
        </tr>
    `).join('');
}

// ==================== DOCUMENT VIEWER ====================
window.viewDocument = (fileUrl, fileName) => {
    if (viewerModal && viewerFileName && documentViewer) {
        viewerFileName.textContent = fileName;
        const cleanUrl = fileUrl?.trim();
        documentViewer.src = cleanUrl || '';
        viewerModal.classList.add('active');
    }
};

function closeViewerModal() {
    if (viewerModal) viewerModal.classList.remove('active');
    if (documentViewer) documentViewer.src = '';

    const viewerHeader = document.querySelector('.viewer-header');
    if (viewerHeader) {
        const existingNav = viewerHeader.querySelector('.viewer-nav');
        if (existingNav) existingNav.remove();
    }
}

// ==================== USER ACTIONS ====================
window.viewUser = async (userId) => {
    selectedUserId = userId;

    try {
        const userDoc = await getDoc(doc(db, 'users', userId));
        if (userDoc.exists()) {
            const user = userDoc.data();
            if (normalizeUserRole(user.role) === 'super_admin') {
                showNotification('Access denied for Super Admin account', 'error');
                return;
            }
            const detailsDiv = document.getElementById('userDetails');
            const fullName = user.fullName || user.email?.split('@')[0] || 'Unknown';

            const docsQuery = query(collection(db, 'submissions'), where('uploadedBy', '==', user.email));
            const docsSnapshot = await getDocs(docsQuery);
            const docCount = docsSnapshot.size;

            detailsDiv.innerHTML = `
                <div class="detail-section">
                    <h4>Personal Information</h4>
                    <div class="detail-row">
                        <span class="detail-label">Full Name:</span>
                        <span class="detail-value">${fullName}</span>
                    </div>
                    <div class="detail-row">
                        <span class="detail-label">Email:</span>
                        <span class="detail-value">${user.email}</span>
                    </div>
                    <div class="detail-row">
                        <span class="detail-label">WhatsApp:</span>
                        <span class="detail-value">${renderWhatsAppContactCell(user)}</span>
                    </div>
                    <div class="detail-row">
                        <span class="detail-label">Location:</span>
                        <span class="detail-value">${user.location || 'N/A'}</span>
                    </div>
                </div>
                <div class="detail-section">
                    <h4>Account Information</h4>
                    <div class="detail-row">
                        <span class="detail-label">Role:</span>
                        <span class="detail-value"><span class="role-badge ${normalizeUserRole(user.role)}">${getRoleLabel(user.role)}</span></span>
                    </div>
                    <div class="detail-row">
                        <span class="detail-label">Status:</span>
                        <span class="detail-value"><span class="status-badge ${user.status || 'pending'}">${user.status || 'pending'}</span></span>
                    </div>
                    <div class="detail-row">
                        <span class="detail-label">Joined:</span>
                        <span class="detail-value">${formatDate(user.createdAt)}</span>
                    </div>
                    <div class="detail-row">
                        <span class="detail-label">Last Login:</span>
                        <span class="detail-value">${formatDate(user.lastLoginAt)}</span>
                    </div>
                    <div class="detail-row">
                        <span class="detail-label">Submissions:</span>
                        <span class="detail-value">${docCount} submission(s)</span>
                    </div>
                </div>
            `;

            document.getElementById('viewUserModal').classList.add('active');
        }
    } catch (error) {
        showNotification('Error loading user details', 'error');
    }
};

window.editUser = async (userId) => {
    selectedUserId = '';
    showNotification('User editing is only available in the Super Admin dashboard', 'warning');
};

async function loadUserForEdit(userId) {
    try {
        const userDoc = await getDoc(doc(db, 'users', userId));
        if (!userDoc.exists()) {
            showNotification('User not found', 'error');
            return;
        }
        const user = userDoc.data();
        if (normalizeUserRole(user.role) === 'super_admin') {
            showNotification('Access denied for Super Admin account', 'error');
            return;
        }
        document.getElementById('modalFullName').value = user.fullName || '';
        document.getElementById('modalLocation').value = user.location || '';
        document.getElementById('modalEmail').value = user.email || '';
        document.getElementById('modalEmail').setAttribute('readonly', 'true');
        document.getElementById('modalDepartment').value = user.department || '';
        document.getElementById('modalWhatsappCode').value = user.whatsappCode || '';
        document.getElementById('modalWhatsappLocalNumber').value = user.whatsappLocalNumber || '';
        document.getElementById('modalRole').value = normalizeUserRole(user.role);
        document.getElementById('modalStatus').value = user.status || 'active';
    } catch (error) {
        showNotification('Error loading user data', 'error');
    }
}

function tsToDate(ts) {
    if (!ts) return null;
    try {
        if (ts.toDate) return ts.toDate();
        const d = new Date(ts);
        return isNaN(d.getTime()) ? null : d;
    } catch (e) {
        return null;
    }
}

function matchesExactDate(ts, dateValue) {
    if (!dateValue) return true;
    const date = tsToDate(ts);
    if (!date) return false;
    const year = date.getFullYear();
    const month = String(date.getMonth() + 1).padStart(2, '0');
    const day = String(date.getDate()).padStart(2, '0');
    return `${year}-${month}-${day}` === dateValue;
}

function renderTrackApplications() {
    if (!trackAppsTableBody) return;

    const search = String(document.getElementById('trackUserSearch')?.value || '').trim().toLowerCase();
    const statusFilter = String(document.getElementById('trackStatusFilter')?.value || 'all').trim().toLowerCase();
    const startVal = document.getElementById('trackStartDate')?.value;
    const endVal = document.getElementById('trackEndDate')?.value;

    const start = startVal ? new Date(startVal) : null;
    const end = endVal ? new Date(endVal) : null;
    if (end) end.setHours(23, 59, 59, 999);

    let list = Array.isArray(allSubmissions)
        ? allSubmissions.filter((s) => String(s.status || '').toLowerCase() !== 'draft')
        : [];

    if (statusFilter && statusFilter !== 'all') {
        list = list.filter((s) => String(s.status || '').toLowerCase() === statusFilter);
    }

    if (start || end) {
        list = list.filter((s) => {
            const uploaded = tsToDate(s.uploadedAt);
            if (!uploaded) return false;
            if (start && uploaded.getTime() < start.getTime()) return false;
            if (end && uploaded.getTime() > end.getTime()) return false;
            return true;
        });
    }

    if (search) {
        list = list.filter((s) => {
            const uploaderEmail = String(s.uploadedBy || '').toLowerCase();
            const uploaderName = getDisplayNameByEmail(uploaderEmail).toLowerCase();
            const customerName = String(s.customerName || '').toLowerCase();
            return uploaderEmail.includes(search) || uploaderName.includes(search) || customerName.includes(search);
        });
    }

    const totalRecords = list.length;
    const totalPages = Math.max(1, Math.ceil(totalRecords / TRACK_PAGE_SIZE));
    trackAppsPage = Math.min(Math.max(1, trackAppsPage), totalPages);

    if (trackPageInfo) trackPageInfo.textContent = `Page ${trackAppsPage} of ${totalPages}`;
    if (trackPrevPageBtn) trackPrevPageBtn.disabled = trackAppsPage <= 1;
    if (trackNextPageBtn) trackNextPageBtn.disabled = trackAppsPage >= totalPages;
    if (trackJumpPageInput) trackJumpPageInput.max = String(totalPages);

    if (list.length === 0) {
        trackAppsTableBody.innerHTML = '<tr><td colspan="8" class="no-data">No matching applications</td></tr>';
        return;
    }

    const startIdx = (trackAppsPage - 1) * TRACK_PAGE_SIZE;
    const pageItems = list.slice(startIdx, startIdx + TRACK_PAGE_SIZE);

    trackAppsTableBody.innerHTML = pageItems.map((s) => {
        const uploaderEmail = String(s.uploadedBy || '').toLowerCase();
        const uploaderLabel = uploaderEmail ? `${getDisplayNameByEmail(uploaderEmail)}` : 'Unknown';
        const assignedReviewer = s.assignedTo ? getDisplayNameByEmail(s.assignedTo) : 'Unassigned';
        const attendedAt = s.reviewedAt ? formatDate(s.reviewedAt) : '-';
        const finalizedAt = s.rsaSubmittedAt ? formatDate(s.rsaSubmittedAt) : '-';
        const assignedRsa = s.assignedToRSA ? getDisplayNameByEmail(s.assignedToRSA) : '-';
        const status = String(s.status || '').trim() || '-';

        return `
            <tr>
                <td><strong>${s.customerName || 'Unknown'}</strong></td>
                <td>${uploaderLabel}</td>
                <td>${formatStatusLabel(status)}</td>
                <td>${formatDate(s.uploadedAt)}</td>
                <td>${assignedReviewer}</td>
                <td>${attendedAt}</td>
                <td>${assignedRsa}</td>
                <td>${finalizedAt}</td>
            </tr>
        `;
    }).join('');
}

window.markSubmissionPaid = async (submissionId) => {
    showNotification('Admin payment tab is read-only', 'warning');
};

window.clearPaidSubmissions = async () => {
    showNotification('Admin payment tab is read-only', 'warning');
};

window.openLeaveModal = (userId) => {
    const user = allUsers.find((u) => u.id === userId);
    if (!user) return showNotification('User not found', 'error');
    const stage = getLeaveStageForRole(user.role);
    if (!stage) return showNotification('Leave reassignment is only available for reviewer, RSA, and payment users.', 'warning');
    if (isUserOnLeave(user)) return showNotification('User is already on leave', 'info');

    const relievers = getEligibleRelievers(user);
    const select = document.getElementById('leaveRelieverSelect');
    const summary = document.getElementById('leaveModalSummary');
    const title = document.getElementById('leaveModalTitle');
    const confirmBtn = document.getElementById('confirmLeaveBtn');

    selectedLeaveUserId = userId;
    if (title) title.textContent = `Activate ${stage.label} Leave`;
    if (summary) {
        summary.textContent = `${user.fullName || user.email || 'This user'} will stop receiving ${stage.label.toLowerCase()} assignments until resumed.`;
    }
    if (select) {
        select.innerHTML = relievers.length
            ? '<option value="">Select reliever...</option>' + relievers.map((r) => {
                const email = normalizeEmailValue(r.email);
                const label = `${r.fullName || r.displayName || email} (${email})`;
                return `<option value="${escapeHtml(email)}">${escapeHtml(label)}</option>`;
            }).join('')
            : '<option value="">No eligible reliever found</option>';
        select.disabled = !relievers.length;
    }
    if (confirmBtn) {
        confirmBtn.disabled = !relievers.length;
        confirmBtn.innerHTML = '<i class="fas fa-person-walking-luggage"></i> Activate Leave';
    }
    document.getElementById('leaveModal')?.classList.add('active');
};

async function confirmActivateLeave() {
    const user = allUsers.find((u) => u.id === selectedLeaveUserId);
    if (!user) return showNotification('User not found', 'error');
    const stage = getLeaveStageForRole(user.role);
    const userEmail = normalizeEmailValue(user.email);
    const relieverEmail = normalizeEmailValue(document.getElementById('leaveRelieverSelect')?.value);
    if (!stage || !userEmail) return showNotification('Invalid leave user', 'error');
    if (!relieverEmail) return showNotification('Please select a reliever', 'warning');
    if (relieverEmail === userEmail) return showNotification('Reliever cannot be the same user', 'warning');

    if (!getEligibleRelievers(user).some((u) => normalizeEmailValue(u.email) === relieverEmail)) {
        return showNotification('Selected reliever is not eligible or is already on leave', 'error');
    }

    const confirmBtn = document.getElementById('confirmLeaveBtn');
    const previousHtml = confirmBtn?.innerHTML;
    if (confirmBtn) {
        confirmBtn.disabled = true;
        confirmBtn.innerHTML = '<i class="fas fa-spinner fa-spin"></i> Activating...';
    }

    try {
        const submissionsSnap = await getDocs(query(
            collection(db, 'submissions'),
            where(stage.assignmentField, '==', userEmail)
        ));
        const activeDocs = submissionsSnap.docs.filter((docSnap) => isSubmissionActiveForLeave(stage.key, docSnap.data() || {}, userEmail));

        await updateDoc(doc(db, 'users', user.id), {
            leaveStatus: 'on_leave',
            leaveStage: stage.key,
            leaveRelieverEmail: relieverEmail,
            leaveStartedAt: serverTimestamp(),
            leaveStartedBy: currentAdmin?.email || ''
        });

        await Promise.all(activeDocs.map((docSnap) => updateDoc(doc(db, 'submissions', docSnap.id), {
            [stage.assignmentField]: relieverEmail,
            leaveCoverActive: true,
            leaveCoverStage: stage.key,
            leaveCoverOriginalEmail: userEmail,
            leaveCoverRelieverEmail: relieverEmail,
            leaveCoverStartedAt: serverTimestamp(),
            leaveCoverStartedBy: currentAdmin?.email || ''
        })));

        await addDoc(collection(db, 'audit'), {
            action: 'user_leave_activated',
            userId: user.id,
            userEmail,
            relieverEmail,
            stage: stage.key,
            movedCount: activeDocs.length,
            performedBy: currentAdmin?.email || '',
            timestamp: serverTimestamp()
        });

        showNotification(`Leave activated. ${activeDocs.length} active application(s) moved to reliever.`, 'success');
        closeLeaveModal();
    } catch (error) {
        showNotification('Failed to activate leave: ' + (error?.message || 'Unknown error'), 'error');
    } finally {
        if (confirmBtn) {
            confirmBtn.disabled = false;
            confirmBtn.innerHTML = previousHtml || '<i class="fas fa-person-walking-luggage"></i> Activate Leave';
        }
    }
}

window.resumeUserFromLeave = (userId) => {
    const user = allUsers.find((u) => u.id === userId);
    if (!user) return showNotification('User not found', 'error');
    if (!isUserOnLeave(user)) return showNotification('User is not currently on leave', 'info');
    const stage = getLeaveStageForRole(user.role);
    if (!stage) return showNotification('Invalid leave stage', 'error');

    showConfirmModal('Resume User from Leave', 'Unfinished applications covered by the reliever will return to this user. Finalized applications will remain completed.', async () => {
        try {
            const userEmail = normalizeEmailValue(user.email);
            const submissionsSnap = await getDocs(query(
                collection(db, 'submissions'),
                where('leaveCoverActive', '==', true),
                where('leaveCoverStage', '==', stage.key),
                where('leaveCoverOriginalEmail', '==', userEmail)
            ));
            const coveredDocs = submissionsSnap.docs;
            const returnDocs = coveredDocs.filter((docSnap) => !isSubmissionFinalizedForLeave(stage.key, docSnap.data() || {}));
            const finalizedDocs = coveredDocs.filter((docSnap) => isSubmissionFinalizedForLeave(stage.key, docSnap.data() || {}));

            await Promise.all(returnDocs.map((docSnap) => updateDoc(doc(db, 'submissions', docSnap.id), {
                [stage.assignmentField]: userEmail,
                leaveCoverActive: false,
                leaveCoverReturnedAt: serverTimestamp(),
                leaveCoverReturnedBy: currentAdmin?.email || ''
            })));

            await Promise.all(finalizedDocs.map((docSnap) => updateDoc(doc(db, 'submissions', docSnap.id), {
                leaveCoverActive: false,
                leaveCoverFinalizedAt: serverTimestamp(),
                leaveCoverFinalizedBy: currentAdmin?.email || ''
            })));

            await updateDoc(doc(db, 'users', user.id), {
                leaveStatus: '',
                leaveEndedAt: serverTimestamp(),
                leaveEndedBy: currentAdmin?.email || '',
                leavePreviousRelieverEmail: user.leaveRelieverEmail || '',
                leaveRelieverEmail: ''
            });

            await addDoc(collection(db, 'audit'), {
                action: 'user_leave_resumed',
                userId: user.id,
                userEmail,
                stage: stage.key,
                returnedCount: returnDocs.length,
                finalizedCount: finalizedDocs.length,
                performedBy: currentAdmin?.email || '',
                timestamp: serverTimestamp()
            });

            closeConfirmModal();
            showNotification(`User resumed. ${returnDocs.length} unfinished application(s) returned; ${finalizedDocs.length} finalized application(s) left closed.`, 'success');
        } catch (error) {
            showNotification('Failed to resume user: ' + (error?.message || 'Unknown error'), 'error');
            throw error;
        }
    });
};

window.deactivateUser = (userId) => {
    showConfirmModal('Deactivate User', 'Are you sure you want to deactivate this user? They will not be able to login.', async () => {
        try {
            const userSnap = await getDoc(doc(db, 'users', userId));
            if (userSnap.exists() && normalizeUserRole(userSnap.data()?.role) === 'super_admin') {
                showNotification('Access denied for Super Admin account', 'error');
                closeConfirmModal();
                return;
            }
            await updateDoc(doc(db, 'users', userId), {
                status: 'deactivated',
                deactivatedAt: serverTimestamp(),
                deactivatedBy: currentAdmin?.email
            });

            await addDoc(collection(db, 'audit'), {
                action: 'user_deactivated',
                userId: userId,
                performedBy: currentAdmin?.email,
                timestamp: serverTimestamp()
            });

            showNotification('User deactivated', 'success');
            closeConfirmModal();
        } catch (error) {
            showNotification('Failed to deactivate user', 'error');
        }
    });
};

window.activateUser = (userId) => {
    showConfirmModal('Activate User', 'Activate this user? They will be able to login again.', async () => {
        try {
            const userSnap = await getDoc(doc(db, 'users', userId));
            if (userSnap.exists() && normalizeUserRole(userSnap.data()?.role) === 'super_admin') {
                showNotification('Access denied for Super Admin account', 'error');
                closeConfirmModal();
                return;
            }
            await updateDoc(doc(db, 'users', userId), {
                status: 'active',
                activatedAt: serverTimestamp(),
                activatedBy: currentAdmin?.email
            });

            await addDoc(collection(db, 'audit'), {
                action: 'user_activated',
                userId: userId,
                performedBy: currentAdmin?.email,
                timestamp: serverTimestamp()
            });

            showNotification('User activated', 'success');
            closeConfirmModal();
        } catch (error) {
            showNotification('Failed to activate user', 'error');
        }
    });
};

window.resetUserPassword = (userId) => {
    showConfirmModal(
        'Reset Password',
        'Reset this user password directly to 123456?',
        async () => {
            try {
                const userRef = doc(db, 'users', userId);
                const userSnap = await getDoc(userRef);
                if (!userSnap.exists()) {
                    showNotification('User not found', 'error');
                    return;
                }

                const userData = userSnap.data() || {};
                const userEmail = String(userData.email || '').trim().toLowerCase();
                if (!userEmail) {
                    showNotification('User email is missing', 'error');
                    return;
                }

                const baseUrl = String(ADMIN_API_BASE_URL || '').trim().replace(/\/+$/, '');
                if (!baseUrl || baseUrl.includes('YOUR-RENDER-URL')) {
                    showNotification('Admin API URL is not configured', 'error');
                    return;
                }

                const currentUser = getAuth().currentUser;
                const idToken = currentUser ? await currentUser.getIdToken() : '';
                if (!idToken) {
                    showNotification('Admin session token is missing. Please login again.', 'error');
                    return;
                }

                const response = await fetch(`${baseUrl}/api/admin/reset-password`, {
                    method: 'POST',
                    headers: {
                        'Content-Type': 'application/json',
                        'Authorization': `Bearer ${idToken}`
                    },
                    body: JSON.stringify({
                        userId,
                        newPassword: '123456'
                    })
                });
                const result = await response.json().catch(() => ({}));
                if (!response.ok || !result.ok) {
                    throw new Error(result.error || `Reset API failed (${response.status})`);
                }

                await updateDoc(userRef, {
                    passwordResetAt: serverTimestamp(),
                    passwordResetBy: currentAdmin?.email || '',
                    passwordResetDefault: true
                });

                await addDoc(collection(db, 'audit'), {
                    action: 'user_password_reset_direct',
                    userId,
                    userEmail,
                    resetTo: '123456',
                    performedBy: currentAdmin?.email,
                    timestamp: serverTimestamp()
                });

                showNotification(`Password reset to 123456 for ${userEmail}`, 'success');
                closeConfirmModal();
            } catch (error) {
                showNotification('Failed to reset password: ' + (error?.message || 'Unknown error'), 'error');
            }
        }
    );
};

window.deleteUser = (userId) => {
    showConfirmModal('Delete User', 'Are you sure you want to permanently delete this user?', async () => {
        try {
            const userDoc = await getDoc(doc(db, 'users', userId));
            const userData = userDoc.data();
            if (userDoc.exists() && normalizeUserRole(userData?.role) === 'super_admin') {
                showNotification('Access denied for Super Admin account', 'error');
                closeConfirmModal();
                return;
            }

            await deleteDoc(doc(db, 'users', userId));

            await addDoc(collection(db, 'audit'), {
                action: 'user_deleted',
                userEmail: userData?.email,
                performedBy: currentAdmin?.email,
                timestamp: serverTimestamp()
            });

            showNotification('User deleted', 'success');
            closeConfirmModal();
        } catch (error) {
            showNotification('Failed to delete user', 'error');
        }
    });
};

// ==================== CREATE USER MODAL ====================
function openCreateUserModal() {
    selectedUserId = null;
    const userForm = document.getElementById('userForm');
    if (userForm) userForm.reset();

    const modalTitle = document.getElementById('modalTitle');
    if (modalTitle) modalTitle.textContent = 'Create New User';

    const emailInput = document.getElementById('modalEmail');
    if (emailInput) emailInput.removeAttribute('readonly');
    document.getElementById('modalStatus').value = 'active';
    const locationInput = document.getElementById('modalLocation');
    if (locationInput) locationInput.value = '';
    const deptInput = document.getElementById('modalDepartment');
    if (deptInput) deptInput.value = '';
    const whatsappCodeInput = document.getElementById('modalWhatsappCode');
    if (whatsappCodeInput) whatsappCodeInput.value = '+234';
    const whatsappLocalInput = document.getElementById('modalWhatsappLocalNumber');
    if (whatsappLocalInput) whatsappLocalInput.value = '';
    document.getElementById('userModal').classList.add('active');
}

// ==================== SAVE USER ====================
async function saveUser(e) {
    e.preventDefault();

    const statusValue = document.getElementById('modalStatus').value || 'active';
    const saveBtn = document.getElementById('saveUserModalBtn');
    const originalSaveBtnHtml = saveBtn ? saveBtn.innerHTML : '';
    const whatsappCodeRaw = (document.getElementById('modalWhatsappCode')?.value || '').trim();
    const whatsappLocalRaw = (document.getElementById('modalWhatsappLocalNumber')?.value || '').trim();
    const whatsappLocalDigits = whatsappLocalRaw.replace(/\D/g, '');
    const normalizedWhatsappCode = whatsappCodeRaw
        ? (whatsappCodeRaw.startsWith('+') ? whatsappCodeRaw : `+${whatsappCodeRaw}`)
        : '';
    const normalizedWhatsappNumber = (normalizedWhatsappCode && whatsappLocalDigits)
        ? `${normalizedWhatsappCode}${whatsappLocalDigits}`
        : '';

    const userData = {
        fullName: document.getElementById('modalFullName').value.trim(),
        location: document.getElementById('modalLocation')?.value.trim() || '',
        email: document.getElementById('modalEmail').value.trim().toLowerCase(),
        department: document.getElementById('modalDepartment')?.value.trim() || '',
        whatsappCode: normalizedWhatsappCode,
        whatsappLocalNumber: whatsappLocalDigits,
        whatsappNumber: normalizedWhatsappNumber,
        phone: normalizedWhatsappNumber || whatsappLocalDigits || '',
        role: document.getElementById('modalRole').value,
        status: statusValue,
        updatedAt: serverTimestamp()
    };

    if (normalizedWhatsappCode && !/^\+\d{1,4}$/.test(normalizedWhatsappCode)) {
        showNotification('WhatsApp country code is invalid', 'error');
        return;
    }
    if (whatsappLocalDigits && !/^\d{10}$/.test(whatsappLocalDigits)) {
        showNotification('WhatsApp number must be exactly 10 digits', 'error');
        return;
    }

    if (normalizeUserRole(userData.role) === 'super_admin') {
        showNotification('Admin cannot create or assign Super Admin role', 'error');
        return;
    }

    if (saveBtn) {
        saveBtn.disabled = true;
        saveBtn.classList.add('loading');
        saveBtn.innerHTML = '<i class="fas fa-spinner fa-spin"></i> Saving...';
    }

    try {
        const emailDupSnap = await getDocs(query(collection(db, 'users'), where('email', '==', userData.email)));
        const emailDuplicate = emailDupSnap.docs.find((d) => d.id !== selectedUserId);
        if (emailDuplicate) {
            showSaveUserErrorToast('Cannot save user. Another user already has this email address.');
            return;
        }

        if (userData.whatsappNumber) {
            const waDupSnap = await getDocs(query(collection(db, 'users'), where('whatsappNumber', '==', userData.whatsappNumber)));
            const waDuplicate = waDupSnap.docs.find((d) => d.id !== selectedUserId);
            if (waDuplicate) {
                showSaveUserErrorToast('Cannot save user. Another user already has this WhatsApp number.');
                return;
            }

            const legacyPhoneDupSnap = await getDocs(query(collection(db, 'users'), where('phone', '==', userData.whatsappNumber)));
            const phoneDuplicate = legacyPhoneDupSnap.docs.find((d) => d.id !== selectedUserId);
            if (phoneDuplicate) {
                showSaveUserErrorToast('Cannot save user. Another user already has this WhatsApp number.');
                return;
            }
        }

        if (selectedUserId) {
            try {
                await updateDoc(doc(db, 'users', selectedUserId), userData);
            } catch (error) {
                if (error?.code !== 'permission-denied') throw error;
                await ensureCurrentAdminWritableProfile();
                await updateDoc(doc(db, 'users', selectedUserId), userData);
            }

            await addDoc(collection(db, 'audit'), {
                action: 'user_updated',
                userId: selectedUserId,
                userEmail: userData.email,
                userFullName: userData.fullName,
                performedBy: currentAdmin?.email,
                timestamp: serverTimestamp()
            });

            showNotification('User updated successfully', 'success');
        } else {
            const userCredential = await createUserWithEmailAndPassword(getAuth(), userData.email, 'CMBank@123');
            const createdUser = userCredential.user;

            await setDoc(doc(db, 'users', createdUser.uid), {
                ...userData,
                uid: createdUser.uid,
                createdAt: serverTimestamp(),
                createdBy: currentAdmin?.email
            }, { merge: true });

            await addDoc(collection(db, 'audit'), {
                action: 'user_created',
                userEmail: userData.email,
                performedBy: currentAdmin?.email,
                timestamp: serverTimestamp()
            });

            showNotification('User created successfully', 'success');
        }

        closeUserModal();
    } catch (error) {
        if (error?.code === 'already-exists') {
            showSaveUserErrorToast('Cannot save user. A duplicate record already exists.');
            return;
        }
        if (String(error?.message || '').toLowerCase().includes('email already')) {
            showSaveUserErrorToast('Cannot save user. Another user already has this email address.');
            return;
        }
        if (String(error?.message || '').toLowerCase().includes('whatsapp') || String(error?.message || '').toLowerCase().includes('phone')) {
            showSaveUserErrorToast('Cannot save user. Another user already has this WhatsApp number.');
            return;
        }
        showSaveUserErrorToast('Failed to save user: ' + (error?.message || 'Unknown error'));
    } finally {
        if (saveBtn) {
            saveBtn.disabled = false;
            saveBtn.classList.remove('loading');
            saveBtn.innerHTML = originalSaveBtnHtml;
        }
    }
}

// ==================== FILTER FUNCTIONS ====================
function filterUsers() {
    const searchTerm = document.getElementById('userSearch')?.value.toLowerCase() || '';
    const roleFilter = document.getElementById('userRoleFilter')?.value || 'all';
    const statusFilter = document.getElementById('userStatusFilter')?.value || 'all';

    if (!allUsers) return;

    const filtered = allUsers.filter(user => {
        const fullName = user.fullName || user.email?.split('@')[0] || '';
        const normalizedRole = normalizeUserRole(user.role);
        const matchesSearch = searchTerm === '' ||
            fullName.toLowerCase().includes(searchTerm) ||
            user.email?.toLowerCase().includes(searchTerm);
        const matchesRole = roleFilter === 'all' || normalizedRole === roleFilter;
        const matchesStatus = statusFilter === 'all' || user.status === statusFilter;

        return matchesSearch && matchesRole && matchesStatus;
    });

    renderUsersTable(filtered);
}

window.deleteDraftApplication = (submissionId, triggerBtn = null) => {
    const sub = allSubmissions.find((s) => s.id === submissionId && String(s.status || '').toLowerCase() === 'draft');
    if (!sub) {
        showNotification('Draft application not found', 'error');
        return;
    }

    const draftName = String(sub.customerName || 'Untitled Draft').trim();
    showConfirmModal(
        'Delete Draft Application',
        `Delete draft application "${draftName}" permanently?`,
        async () => {
            const originalHtml = triggerBtn?.innerHTML || '';
            try {
                if (triggerBtn) {
                    triggerBtn.disabled = true;
                    triggerBtn.classList.add('loading');
                    triggerBtn.innerHTML = '<i class="fas fa-spinner fa-spin"></i> Deleting...';
                }
                await deleteDoc(doc(db, 'submissions', submissionId));
                await addDoc(collection(db, 'audit'), {
                    action: 'draft_application_deleted',
                    submissionId,
                    customerName: sub.customerName || '',
                    uploadedBy: sub.uploadedBy || '',
                    performedBy: currentAdmin?.email || '',
                    timestamp: serverTimestamp()
                });
                showNotification('Draft application deleted', 'success');
                closeConfirmModal();
            } catch (error) {
                if (triggerBtn) {
                    triggerBtn.disabled = false;
                    triggerBtn.classList.remove('loading');
                    triggerBtn.innerHTML = originalHtml || '<i class="fas fa-trash"></i> Delete';
                }
                showNotification('Failed to delete draft application', 'error');
            }
        }
    );
};

function filterDraftDocs() {
    const searchTerm = document.getElementById('draftDocSearch')?.value.toLowerCase() || '';
    const dateFilter = document.getElementById('draftDocDate')?.value || '';
    const drafts = allSubmissions.filter((s) => String(s.status || '').toLowerCase() === 'draft');

    const filtered = drafts.filter((sub) => {
        const matchesSearch = searchTerm === '' ||
            String(sub.customerName || '').toLowerCase().includes(searchTerm) ||
            String(sub.uploadedBy || '').toLowerCase().includes(searchTerm) ||
            getSubmissionAgentLabel(sub).toLowerCase().includes(searchTerm);
        const relevantDate = sub.draftSavedAt || sub.uploadedAt || sub.updatedAt;
        const matchesDate = matchesExactDate(relevantDate, dateFilter);
        return matchesSearch && matchesDate;
    });

    if (!draftDocsTableBody) return;

    if (filtered.length === 0) {
        draftDocsTableBody.innerHTML = '<tr><td colspan="7" class="no-data">No draft applications</td></tr>';
        return;
    }

    draftDocsTableBody.innerHTML = filtered.map((sub) => {
        const uploaderEmail = String(sub.uploadedBy || '').trim().toLowerCase();
        const uploaderLabel = uploaderEmail ? (uploaderNames[uploaderEmail] || uploaderEmail.split('@')[0] || uploaderEmail) : 'Unknown';
        const lastSaved = sub.draftSavedAt || sub.uploadedAt || sub.updatedAt;
        const docCount = Array.isArray(sub.documents) ? sub.documents.length : 0;
        return `
            <tr>
                <td><strong>${escapeHtml(sub.customerName || 'Untitled Draft')}</strong></td>
                <td>${escapeHtml(uploaderLabel)}</td>
                <td>${escapeHtml(getSubmissionAgentLabel(sub))}</td>
                <td>${docCount}</td>
                <td>${formatDate(lastSaved)}</td>
                <td><span class="status-badge status-pending">Draft</span></td>
                <td>
                    <button class="action-btn view-btn-small" onclick="window.viewSubmissionDocs('${sub.id}')">
                        <i class="fas fa-eye"></i> View All
                    </button>
                    <button class="action-btn" onclick="window.deleteDraftApplication('${sub.id}', this)" style="background:#b91c1c;color:#fff;border:none;">
                        <i class="fas fa-trash"></i> Delete
                    </button>
                </td>
            </tr>
        `;
    }).join('');
}

function filterPendingDocs() {
    const searchTerm = document.getElementById('pendingDocSearch')?.value.toLowerCase() || '';
    const dateFilter = document.getElementById('pendingDocDate')?.value || '';
    const pending = allSubmissions.filter(s => s.status === 'pending');

    const filtered = pending.filter(sub => {
        const matchesSearch = searchTerm === '' ||
            sub.customerName?.toLowerCase().includes(searchTerm) ||
            sub.uploadedBy?.toLowerCase().includes(searchTerm) ||
            sub.assignedTo?.toLowerCase().includes(searchTerm);
        const relevantDate = isResubmittedSubmission(sub) ? (sub.reuploadedAt || sub.uploadedAt) : sub.uploadedAt;
        const matchesDate = matchesExactDate(relevantDate, dateFilter);
        return matchesSearch && matchesDate;
    });

    if (!pendingDocsTableBody) return;

    if (filtered.length === 0) {
        pendingDocsTableBody.innerHTML = '<tr><td colspan="7" class="no-data">No pending documents</td></tr>';
        return;
    }

    pendingDocsTableBody.innerHTML = filtered.map(sub => {
        const isResubmitted = isResubmittedSubmission(sub);
        const preferredDate = isResubmitted ? (sub.reuploadedAt || sub.uploadedAt) : sub.uploadedAt;
        const uploadDate = preferredDate ?
            (preferredDate.toDate ? preferredDate.toDate() : new Date(preferredDate)).toLocaleString() : 'N/A';
        const uploaderFullName = uploaderNames[sub.uploadedBy] || sub.uploadedBy?.split('@')[0] || 'Unknown';
        const assignedEmail = (sub.assignedTo || '').toString().trim().toLowerCase();
        const assignedLabel = assignedEmail
            ? (uploaderNames[assignedEmail] || assignedEmail.split('@')[0] || assignedEmail)
            : 'Unassigned';
        const reasonCell = getRejectCount(sub) > 0
            ? `<button class="action-btn reason-btn" onclick="window.openAdminRejectionReasonModal('${sub.id}')"><i class="fas fa-eye"></i> View</button>`
            : '-';
        const statusCell = `
            <span class="status-badge status-pending">Pending</span>
            ${isResubmitted ? '<span class="status-badge status-resubmitted">Resubmitted</span>' : ''}
        `;

        return `
            <tr>
                <td><strong>${sub.customerName || 'Unknown'}</strong></td>
                <td>${uploadDate}</td>
                <td>${uploaderFullName}</td>
                <td>${assignedLabel}</td>
                <td>${statusCell}</td>
                <td>${reasonCell}</td>
                <td>
                    <button class="action-btn view-btn-small" onclick="window.viewSubmissionDocs('${sub.id}')">
                        <i class="fas fa-eye"></i> View All
                    </button>
                    <button class="action-btn download-all-btn" onclick="window.downloadAllSubmission('${sub.id}')">
                        <i class="fas fa-download"></i> Download All
                    </button>
                    <button class="action-btn" onclick="window.openApplicationChat('${sub.id}')">
                        <i class="fas fa-comments"></i> Chat
                    </button>
                </td>
            </tr>
        `;
    }).join('');
}

function filterApprovedDocs() {
    const searchTerm = document.getElementById('approvedDocSearch')?.value.toLowerCase() || '';
    const dateFilter = document.getElementById('approvedDocDate')?.value || '';
    const approved = allSubmissions.filter(s => {
        const status = String(s.status || '').toLowerCase();
        return status === 'processing_to_pfa' || status === 'approved';
    });

    const filtered = approved.filter(sub => {
        const matchesSearch = searchTerm === '' ||
            sub.customerName?.toLowerCase().includes(searchTerm) ||
            sub.uploadedBy?.toLowerCase().includes(searchTerm);
        const matchesDate = matchesExactDate(sub.reviewedAt || sub.uploadedAt, dateFilter);
        return matchesSearch && matchesDate;
    });

    if (!approvedDocsTableBody) return;

    if (filtered.length === 0) {
        approvedDocsTableBody.innerHTML = '<tr><td colspan="6" class="no-data">No approved documents</td></tr>';
        return;
    }

    approvedDocsTableBody.innerHTML = filtered.map(sub => {
        const uploadDate = sub.uploadedAt ?
            (sub.uploadedAt.toDate ? sub.uploadedAt.toDate() : new Date(sub.uploadedAt)).toLocaleString() : 'N/A';
        const approvedDate = sub.reviewedAt ?
            (sub.reviewedAt.toDate ? sub.reviewedAt.toDate() : new Date(sub.reviewedAt)).toLocaleString() : 'N/A';
        const uploaderFullName = uploaderNames[sub.uploadedBy] || sub.uploadedBy?.split('@')[0] || 'Unknown';
        const approverFullName = adminNames[sub.reviewedBy] || sub.reviewedBy?.split('@')[0] || 'N/A';

        return `
            <tr>
                <td><strong>${sub.customerName || 'Unknown'}</strong></td>
                <td>${uploadDate}</td>
                <td>${uploaderFullName}</td>
                <td>${approverFullName}</td>
                <td>${approvedDate}</td>
                <td>
                    <button class="action-btn view-btn-small" onclick="window.viewSubmissionDocs('${sub.id}')">
                        <i class="fas fa-eye"></i> View All
                    </button>
                    <button class="action-btn download-all-btn" onclick="window.downloadAllSubmission('${sub.id}')">
                        <i class="fas fa-download"></i> Download All
                    </button>
                    <button class="action-btn" onclick="window.openApplicationChat('${sub.id}')">
                        <i class="fas fa-comments"></i> Chat
                    </button>
                </td>
            </tr>
        `;
    }).join('');
}

function filterRejectedDocs() {
    const searchTerm = document.getElementById('rejectedDocSearch')?.value.toLowerCase() || '';
    const dateFilter = document.getElementById('rejectedDocDate')?.value || '';
    const rejected = allSubmissions.filter(s => ['rejected', 'rejected_by_rsa'].includes(String(s.status || '').toLowerCase()));

    const filtered = rejected.filter(sub => {
        const matchesSearch = searchTerm === '' ||
            sub.customerName?.toLowerCase().includes(searchTerm) ||
            sub.uploadedBy?.toLowerCase().includes(searchTerm);
        const matchesDate = matchesExactDate(sub.latestRejectedAt || sub.rejectedAt || sub.reviewedAt || sub.uploadedAt, dateFilter);
        return matchesSearch && matchesDate;
    });

    if (!rejectedDocsTableBody) return;

    if (filtered.length === 0) {
        rejectedDocsTableBody.innerHTML = '<tr><td colspan="9" class="no-data">No rejected documents</td></tr>';
        return;
    }

    rejectedDocsTableBody.innerHTML = filtered.map(renderRejectedDocRow).join('');
}

function filterAudit() {
    const dateFilter = document.getElementById('auditDate')?.value;
    const actionFilter = document.getElementById('auditAction')?.value || 'all';

    if (!allAudits) return;

    const filtered = allAudits.filter(audit => {
        let matchesDate = true;
        if (dateFilter && audit.timestamp) {
            const auditDate = new Date(audit.timestamp.seconds * 1000).toISOString().split('T')[0];
            matchesDate = auditDate === dateFilter;
        }
        const matchesAction = actionFilter === 'all' || audit.action === actionFilter;

        return matchesDate && matchesAction;
    });

    renderAuditTable(filtered);
}

window.openAdminRejectionReasonModal = (submissionId) => {
    const sub = allSubmissions.find((s) => s.id === submissionId);
    if (!sub || !adminRejectionReasonModal) return;

    const entries = getRejectionHistoryEntries(sub);
    if (adminRejectionReasonCustomerName) {
        adminRejectionReasonCustomerName.textContent = sub.customerName || 'Unknown';
    }
    if (adminRejectionReasonHistory) {
        if (entries.length) {
            adminRejectionReasonHistory.innerHTML = `
                <ol class="rejection-history-list">
                    ${entries.map((entry, index) => {
                        const timeText = entry.rejectedAt ? formatDate(entry.rejectedAt) : 'Time not available';
                        return `<li><strong>Rejection ${index + 1}:</strong> ${escapeHtml(entry.reason)}<span class="rejection-history-time">${escapeHtml(timeText)}</span></li>`;
                    }).join('')}
                </ol>
            `;
        } else {
            adminRejectionReasonHistory.textContent = 'No rejection reason available.';
        }
        adminRejectionReasonHistory.style.display = 'block';
    }

    adminRejectionReasonModal.classList.add('active');
};

// ==================== MODAL CONTROLS ====================
function closeUserModal() {
    document.getElementById('userModal').classList.remove('active');
}

function closeLeaveModal() {
    selectedLeaveUserId = '';
    document.getElementById('leaveModal')?.classList.remove('active');
}

function closeViewUserModal() {
    document.getElementById('viewUserModal').classList.remove('active');
}

function closeViewAgentModal() {
    document.getElementById('viewAgentModal')?.classList.remove('active');
}

function closeTestResultModal() {
    const modal = document.getElementById('testResultModal');
    if (modal) modal.classList.remove('active');
}

function normalizeTestStatus(value) {
    const raw = String(value ?? '').replace(/\u00a0/g, ' ').trim();
    const upper = raw.toUpperCase();
    if (upper.includes('PASS')) return 'PASSED';
    if (upper.includes('FAIL')) return 'FAILED';

    const cleaned = upper
        .replace(/[^A-Z0-9 _-]/g, ' ')
        .replace(/\s+/g, ' ')
        .trim();

    return cleaned || 'UNKNOWN';
}

function showTestResultModal(title, resultRows) {
    const modal = document.getElementById('testResultModal');
    const titleEl = document.getElementById('testResultTitle');
    const bodyEl = document.getElementById('testResultBody');
    if (!modal || !titleEl || !bodyEl) return;

    const safeRows = (Array.isArray(resultRows) ? resultRows : []).map((row) => {
        const label = String(row?.label ?? 'Detail');
        const isStatusRow = /status/i.test(label);
        let value = String(row?.value ?? 'N/A');
        let ok = Boolean(row?.ok);
        let statusClass = '';

        if (isStatusRow) {
            value = normalizeTestStatus(value);
            if (value === 'PASSED') {
                ok = true;
                statusClass = 'status-passed';
            } else if (value === 'FAILED') {
                ok = false;
                statusClass = 'status-failed';
            }
        }

        return { label, value, ok, statusClass };
    });

    titleEl.textContent = String(title || 'Distribution Test Result');
    bodyEl.innerHTML = safeRows.map((row) => `
        <div class="test-result-item ${row.statusClass}">
            <span class="test-result-label">${escapeHtml(row.label)}</span>
            <span class="test-result-value ${row.ok ? 'ok' : ''} ${row.statusClass === 'status-failed' ? 'bad' : ''}">${escapeHtml(row.value)}</span>
        </div>
    `).join('');

    modal.classList.add('active');
}

async function runWithButtonSpinner(buttonId, loadingText, action) {
    const btn = document.getElementById(buttonId);
    const originalHtml = btn ? btn.innerHTML : '';
    if (btn) {
        btn.disabled = true;
        btn.classList.add('loading');
        btn.innerHTML = `<i class="fas fa-spinner fa-spin"></i> ${loadingText}`;
    }
    try {
        return await action();
    } finally {
        if (btn) {
            btn.disabled = false;
            btn.classList.remove('loading');
            btn.innerHTML = originalHtml;
        }
    }
}

function showConfirmModal(title, message, onConfirm) {
    const confirmTitle = document.getElementById('confirmTitle');
    const confirmMessage = document.getElementById('confirmMessage');
    const confirmModal = document.getElementById('confirmModal');

    if (confirmTitle) confirmTitle.textContent = title;
    if (confirmMessage) confirmMessage.textContent = message;
    if (confirmModal) confirmModal.classList.add('active');

    const confirmBtn = document.getElementById('confirmAction');
    if (confirmBtn) {
        const newConfirmBtn = confirmBtn.cloneNode(true);
        newConfirmBtn.disabled = false;
        newConfirmBtn.classList.remove('loading');
        newConfirmBtn.innerHTML = 'Confirm';
        confirmBtn.parentNode.replaceChild(newConfirmBtn, confirmBtn);
        newConfirmBtn.addEventListener('click', async () => {
            const originalHtml = newConfirmBtn.innerHTML;
            newConfirmBtn.disabled = true;
            newConfirmBtn.classList.add('loading');
            newConfirmBtn.innerHTML = '<i class="fas fa-spinner fa-spin"></i> Processing...';
            try {
                await onConfirm();
            } catch (error) {
                newConfirmBtn.disabled = false;
                newConfirmBtn.classList.remove('loading');
                newConfirmBtn.innerHTML = originalHtml;
                throw error;
            }
        });
    }
}

function closeConfirmModal() {
    document.getElementById('confirmModal').classList.remove('active');
    const confirmBtn = document.getElementById('confirmAction');
    if (confirmBtn) {
        confirmBtn.disabled = false;
        confirmBtn.classList.remove('loading');
        confirmBtn.innerHTML = 'Confirm';
    }
}

// ==================== NOTIFICATION SYSTEM ====================
function showNotification(message, type = 'info') {
    if (!notification) return;

    notification.textContent = message;
    notification.className = `notification ${type}`;
    notification.style.display = 'block';
    notification.offsetHeight;
    notification.classList.add('show');

    if (notificationTimer) {
        clearTimeout(notificationTimer);
    }

    notificationTimer = setTimeout(() => {
        notification.classList.remove('show');
        window.setTimeout(() => {
            notification.style.display = 'none';
        }, 400);
        notificationTimer = null;
    }, 3500);
}

function showToast(message, type = 'info') {
    showNotification(message, type);
}

function showSaveUserErrorToast(message) {
    showToast(message, 'error');
}

// ==================== ROUND-ROBIN MONITORING ====================

async function loadRoundRobinMonitor() {
    try {
        const RR_COUNTER_DOC = doc(db, 'counters', 'roundRobin');
        const counterSnap = await getDoc(RR_COUNTER_DOC);

        let lastIndex = -1;
        let lastDate = '';

        if (counterSnap.exists()) {
            const data = counterSnap.data();
            lastIndex = data.lastIndex ?? -1;
            lastDate = data.lastDate || 'Never';
        }

        const viewers = await getReviewerUsersForRoundRobin();

        const nextIndex = (lastIndex + 1) % (viewers.length || 1);
        const nextViewerEmail = viewers[nextIndex]?.email || 'N/A';
        const nextViewerName = viewers[nextIndex]?.fullName || 'N/A';

        const lastViewerEmail = lastIndex >= 0 && lastIndex < viewers.length ? viewers[lastIndex]?.email : 'N/A';
        const lastViewerName = lastIndex >= 0 && lastIndex < viewers.length ? viewers[lastIndex]?.fullName : 'N/A';

        document.getElementById('lastDistributionViewer').textContent = `${lastViewerName} (${lastViewerEmail})`;
        document.getElementById('currentDistributionIndex').textContent = `${nextIndex} of ${viewers.length}`;
        document.getElementById('lastResetDate').textContent = lastDate;
        document.getElementById('totalViewerCount').textContent = viewers.length;

        await loadDistributionStats(viewers);
        await loadAssignmentHistory();

    } catch (error) {
        showNotification('Error loading monitor: ' + error.message, 'error');
    }
}

async function loadDistributionStats(viewers) {
    const today = await getTrustedDateKey();
    const statsBody = document.getElementById('distributionStatsBody');
    if (!statsBody) return;

    const stats = [];

    for (const viewer of viewers) {
        const assignedQuery = query(
            collection(db, 'submissions'),
            where('assignedTo', '==', viewer.email),
            where('uploadedAt', '>=', new Date(`${today}T00:00:00+01:00`))
        );
        const assignedSnap = await getDocs(assignedQuery);
        const assignedCount = assignedSnap.size;

        const completedDocs = assignedSnap.docs.filter(d => ['approved', 'rejected', 'rejected_by_rsa'].includes(d.data().status));
        const completedCount = completedDocs.length;

        const pendingCount = assignedCount - completedCount;

        stats.push({
            email: viewer.email,
            fullName: viewer.fullName,
            assigned: assignedCount,
            completed: completedCount,
            pending: pendingCount
        });
    }

    stats.sort((a, b) => b.assigned - a.assigned);

    statsBody.innerHTML = stats.map(s => `
        <tr>
            <td>${s.fullName}</td>
            <td>${s.email}</td>
            <td style="text-align: center; font-weight: 600;">${s.assigned}</td>
            <td style="text-align: center; color: #10b981;">${s.completed}</td>
            <td style="text-align: center; color: #f59e0b;">${s.pending}</td>
        </tr>
    `).join('');
}

async function loadAssignmentHistory() {
    const historyBody = document.getElementById('assignmentHistoryBody');
    if (!historyBody) return;

    try {
        const assignmentQuery = query(
            collection(db, 'roundRobinAssignments'),
            orderBy('assignedAt', 'desc'),
            limit(20)
        );

        const assignmentSnap = await getDocs(assignmentQuery);

        if (assignmentSnap.empty) {
            historyBody.innerHTML = '<tr><td colspan="4" style="text-align: center; padding: 20px; color: #999;">No assignment history yet</td></tr>';
            return;
        }

        const rows = assignmentSnap.docs.map(doc => {
            const data = doc.data();
            const timestamp = data.assignedAt?.toDate?.() || new Date(data.assignedAt);
            const timeStr = timestamp.toLocaleString();

            return `
                <tr>
                    <td>${timeStr}</td>
                    <td>${data.customerName || 'N/A'}</td>
                    <td>${data.assignedTo || 'N/A'}</td>
                    <td>${data.assignedBy || 'System'}</td>
                </tr>
            `;
        }).join('');

        historyBody.innerHTML = rows;
    } catch (error) {
        historyBody.innerHTML = '<tr><td colspan="4" style="text-align: center; padding: 20px; color: #999;">Assignment history not available</td></tr>';
    }
}

async function handleResetRoundRobin() {
    showConfirmModal(
        'Reset Round-Robin Counter?',
        'This will reset the distribution counter to start from the first reviewer. Are you sure?',
        async () => {
            closeConfirmModal();
            await runWithButtonSpinner('resetRoundRobinBtn', 'Resetting...', async () => {
                try {
                    const RR_COUNTER_DOC = doc(db, 'counters', 'roundRobin');
                    await setDoc(RR_COUNTER_DOC, {
                        lastIndex: -1,
                        lastDate: ''
                    }, { merge: true });

                    await loadRoundRobinMonitor();

                    showTestResultModal('Round-Robin Reset Result', [
                        { label: 'Counter', value: 'Round-Robin' },
                        { label: 'Reset By', value: currentAdmin?.email || 'Admin' },
                        { label: 'Reset Time', value: new Date().toLocaleString() },
                        { label: 'Action Status', value: 'PASSED', ok: true }
                    ]);
                } catch (error) {
                    showTestResultModal('Round-Robin Reset Result', [
                        { label: 'Counter', value: 'Round-Robin' },
                        { label: 'Error', value: error.message || 'Unknown error' },
                        { label: 'Action Status', value: 'FAILED', ok: false }
                    ]);
                }
            });
        }
    );
}

async function handleTestDistribution() {
    await runWithButtonSpinner('testDistributionBtn', 'Testing...', async () => {
        try {
            showNotification('Testing round-robin distribution...', 'info');

            const viewers = (await getReviewerUsersForRoundRobin()).map((u) => u.email);

            if (viewers.length === 0) {
                throw new Error('No reviewers found in the system');
            }

            const counterRef = doc(db, 'counters', 'roundRobin');
            const counterSnap = await getDoc(counterRef);
            let lastIndex = -1;
            let lastDate = '';

            if (counterSnap.exists()) {
                const data = counterSnap.data();
                lastIndex = data.lastIndex ?? -1;
                lastDate = data.lastDate || '';
            }

            const today = await getTrustedDateKey();
            if (lastDate !== today) lastIndex = -1;

            const nextIndex = (lastIndex + 1) % viewers.length;
            const nextViewer = viewers[nextIndex];

            showTestResultModal('Round-Robin Test Result', [
                { label: 'Total Reviewers', value: String(viewers.length) },
                { label: 'Last Index', value: String(lastIndex) },
                { label: 'Next Index', value: String(nextIndex) },
                { label: 'Next Reviewer', value: nextViewer },
                { label: 'Current Date', value: today },
                { label: 'Last Reset Date', value: lastDate || 'Never' },
                { label: 'Test Status', value: 'PASSED', ok: true }
            ]);

            showNotification(`Test successful. Next reviewer: ${nextViewer}`, 'success');
        } catch (error) {
            showNotification('Test failed: ' + error.message, 'error');
        }
    });
}

// ==================== RSA ROUND-ROBIN MONITORING ====================

async function loadRSARoundRobinMonitor() {
    try {
        const RSA_COUNTER_DOC = doc(db, 'counters', 'roundRobinRSA');
        const counterSnap = await getDoc(RSA_COUNTER_DOC);

        let lastIndex = -1;
        let lastDate = '';

        if (counterSnap.exists()) {
            const data = counterSnap.data();
            lastIndex = data.lastIndex ?? -1;
            lastDate = data.lastDate || 'Never';
        }

        const rsaQuery = query(collection(db, 'users'), where('role', '==', 'rsa'));
        const rsaSnap = await getDocs(rsaQuery);
        const rsaUsers = rsaSnap.docs.map(d => ({
            id: d.id,
            email: d.data().email,
            fullName: d.data().fullName || d.data().email,
            status: String(d.data().status || 'active').toLowerCase(),
            leaveStatus: String(d.data().leaveStatus || '').toLowerCase(),
            skipRsaRoundRobin: d.data().skipRsaRoundRobin === true
        }))
            .filter((user) => user.status !== 'deactivated' && user.leaveStatus !== 'on_leave' && user.skipRsaRoundRobin !== true)
            .sort((a, b) => a.email.localeCompare(b.email));

        const nextIndex = (lastIndex + 1) % (rsaUsers.length || 1);
        const nextRSAEmail = rsaUsers[nextIndex]?.email || 'N/A';
        const nextRSAName = rsaUsers[nextIndex]?.fullName || 'N/A';

        const lastRSAEmail = lastIndex >= 0 && lastIndex < rsaUsers.length ? rsaUsers[lastIndex]?.email : 'N/A';
        const lastRSAName = lastIndex >= 0 && lastIndex < rsaUsers.length ? rsaUsers[lastIndex]?.fullName : 'N/A';

        document.getElementById('lastDistributionRSA').textContent = `${lastRSAName} (${lastRSAEmail})`;
        document.getElementById('currentDistributionIndexRSA').textContent = `${nextIndex} of ${rsaUsers.length}`;
        document.getElementById('lastResetDateRSA').textContent = lastDate;
        document.getElementById('totalRSAUserCount').textContent = rsaUsers.length;

        await loadRSADistributionStats(rsaUsers);
        await loadRSAAssignmentHistory();

    } catch (error) {
        showNotification('Error loading RSA monitor: ' + error.message, 'error');
    }
}

async function loadRSADistributionStats(rsaUsers) {
    const today = await getTrustedDateKey();
    const statsBody = document.getElementById('distributionStatsRSABody');
    if (!statsBody) return;

    const stats = [];

    for (const rsa of rsaUsers) {
        const assignedQuery = query(
            collection(db, 'submissions'),
            where('assignedToRSA', '==', rsa.email),
            where('uploadedAt', '>=', new Date(`${today}T00:00:00+01:00`))
        );
        const assignedSnap = await getDocs(assignedQuery);
        const assignedCount = assignedSnap.size;

        const completedDocs = assignedSnap.docs.filter(d => d.data().status !== 'pending');
        const completedCount = completedDocs.length;

        const pendingCount = assignedCount - completedCount;

        stats.push({
            email: rsa.email,
            fullName: rsa.fullName,
            assigned: assignedCount,
            completed: completedCount,
            pending: pendingCount
        });
    }

    stats.sort((a, b) => b.assigned - a.assigned);

    statsBody.innerHTML = stats.map(s => `
        <tr>
            <td>${s.fullName}</td>
            <td>${s.email}</td>
            <td style="text-align: center; font-weight: 600;">${s.assigned}</td>
            <td style="text-align: center; color: #10b981;">${s.completed}</td>
            <td style="text-align: center; color: #f59e0b;">${s.pending}</td>
        </tr>
    `).join('');
}

async function loadRSAAssignmentHistory() {
    const historyBody = document.getElementById('assignmentHistoryRSABody');
    if (!historyBody) return;

    try {
        const assignmentQuery = query(
            collection(db, 'roundRobinAssignmentsRSA'),
            orderBy('assignedAt', 'desc'),
            limit(20)
        );

        const assignmentSnap = await getDocs(assignmentQuery);

        if (assignmentSnap.empty) {
            historyBody.innerHTML = '<tr><td colspan="4" style="text-align: center; padding: 20px; color: #999;">No RSA assignment history yet</td></tr>';
            return;
        }

        const rows = assignmentSnap.docs.map(doc => {
            const data = doc.data();
            const timestamp = data.assignedAt?.toDate?.() || new Date(data.assignedAt);
            const timeStr = timestamp.toLocaleString();

            return `
                <tr>
                    <td>${timeStr}</td>
                    <td>${data.customerName || 'N/A'}</td>
                    <td>${data.assignedToRSA || 'N/A'}</td>
                    <td>${data.assignedBy || 'System'}</td>
                </tr>
            `;
        }).join('');

        historyBody.innerHTML = rows;
    } catch (error) {
        historyBody.innerHTML = '<tr><td colspan="4" style="text-align: center; padding: 20px; color: #999;">RSA assignment history not available</td></tr>';
    }
}

async function handleResetRoundRobinRSA() {
    showConfirmModal(
        'Reset RSA Round-Robin Counter?',
        'This will reset the RSA distribution counter to start from the first RSA user. Are you sure?',
        async () => {
            closeConfirmModal();
            await runWithButtonSpinner('resetRoundRobinRSABtn', 'Resetting...', async () => {
                try {
                    const RSA_COUNTER_DOC = doc(db, 'counters', 'roundRobinRSA');
                    await setDoc(RSA_COUNTER_DOC, {
                        lastIndex: -1,
                        lastDate: ''
                    }, { merge: true });

                    await loadRSARoundRobinMonitor();

                    showTestResultModal('RSA Round-Robin Reset Result', [
                        { label: 'Counter', value: 'RSA Round-Robin' },
                        { label: 'Reset By', value: currentAdmin?.email || 'Admin' },
                        { label: 'Reset Time', value: new Date().toLocaleString() },
                        { label: 'Action Status', value: 'PASSED', ok: true }
                    ]);
                } catch (error) {
                    showTestResultModal('RSA Round-Robin Reset Result', [
                        { label: 'Counter', value: 'RSA Round-Robin' },
                        { label: 'Error', value: error.message || 'Unknown error' },
                        { label: 'Action Status', value: 'FAILED', ok: false }
                    ]);
                }
            });
        }
    );
}

async function handleTestDistributionRSA() {
    await runWithButtonSpinner('testDistributionRSABtn', 'Testing...', async () => {
        try {
            showNotification('Testing RSA round-robin distribution...', 'info');

            const rsaQuery = query(collection(db, 'users'), where('role', '==', 'rsa'));
            const rsaSnap = await getDocs(rsaQuery);
            const rsaUsers = rsaSnap.docs
                .map(d => d.data() || {})
                .filter((user) => String(user.status || 'active').toLowerCase() !== 'deactivated' && String(user.leaveStatus || '').toLowerCase() !== 'on_leave' && user?.skipRsaRoundRobin !== true)
                .map((user) => user.email)
                .filter(Boolean)
                .sort();

            if (rsaUsers.length === 0) {
                throw new Error('No RSA users found in the system');
            }

            const counterRef = doc(db, 'counters', 'roundRobinRSA');
            const counterSnap = await getDoc(counterRef);
            let lastIndex = -1;
            let lastDate = '';

            if (counterSnap.exists()) {
                const data = counterSnap.data();
                lastIndex = data.lastIndex ?? -1;
                lastDate = data.lastDate || '';
            }

            const nextIndex = (lastIndex + 1) % rsaUsers.length;
            const nextRSA = rsaUsers[nextIndex];

            showTestResultModal('RSA Round-Robin Test Result', [
                { label: 'Total RSA Users', value: String(rsaUsers.length) },
                { label: 'Last Index', value: String(lastIndex) },
                { label: 'Next Index', value: String(nextIndex) },
                { label: 'Next RSA User', value: nextRSA },
                { label: 'Current Date', value: today },
                { label: 'Last Counter Update', value: lastDate || 'Never' },
                { label: 'Test Status', value: 'PASSED', ok: true }
            ]);

            showNotification(`Test successful. Next RSA user: ${nextRSA}`, 'success');
        } catch (error) {
            showNotification('Test failed: ' + error.message, 'error');
        }
    });
}

// ==================== PAYMENT ROUND-ROBIN MONITORING ====================

async function loadPaymentRoundRobinMonitor() {
    try {
        const PAYMENT_COUNTER_DOC = doc(db, 'counters', 'roundRobinPayment');
        const counterSnap = await getDoc(PAYMENT_COUNTER_DOC);

        let lastIndex = -1;
        let lastDate = '';

        if (counterSnap.exists()) {
            const data = counterSnap.data();
            lastIndex = data.lastIndex ?? -1;
            lastDate = data.lastDate || 'Never';
        }

        const paymentUsers = await getPaymentUsersForRoundRobin();
        const nextIndex = (lastIndex + 1) % (paymentUsers.length || 1);
        const nextPaymentEmail = paymentUsers[nextIndex]?.email || 'N/A';
        const nextPaymentName = paymentUsers[nextIndex]?.fullName || 'N/A';

        const lastPaymentEmail = lastIndex >= 0 && lastIndex < paymentUsers.length ? paymentUsers[lastIndex]?.email : 'N/A';
        const lastPaymentName = lastIndex >= 0 && lastIndex < paymentUsers.length ? paymentUsers[lastIndex]?.fullName : 'N/A';

        document.getElementById('lastDistributionPayment').textContent = `${lastPaymentName} (${lastPaymentEmail})`;
        document.getElementById('currentDistributionIndexPayment').textContent = `${nextIndex} of ${paymentUsers.length}`;
        document.getElementById('lastResetDatePayment').textContent = lastDate;
        document.getElementById('totalPaymentUserCount').textContent = paymentUsers.length;

        await loadPaymentDistributionStats(paymentUsers);
        await loadPaymentAssignmentHistory();
    } catch (error) {
        showNotification('Error loading Payment monitor: ' + error.message, 'error');
    }
}

async function loadPaymentDistributionStats(paymentUsers) {
    const today = await getTrustedDateKey();
    const statsBody = document.getElementById('distributionStatsPaymentBody');
    if (!statsBody) return;

    const stats = [];

    for (const paymentUser of paymentUsers) {
        const assignedQuery = query(
            collection(db, 'submissions'),
            where('assignedToPayment', '==', paymentUser.email),
            where('uploadedAt', '>=', new Date(`${today}T00:00:00+01:00`))
        );
        const assignedSnap = await getDocs(assignedQuery);
        const assignedCount = assignedSnap.size;

        const completedDocs = assignedSnap.docs.filter(d => ['paid', 'cleared'].includes(String(d.data().status || '').toLowerCase()));
        const completedCount = completedDocs.length;
        const pendingCount = Math.max(0, assignedCount - completedCount);

        stats.push({
            email: paymentUser.email,
            fullName: paymentUser.fullName,
            assigned: assignedCount,
            completed: completedCount,
            pending: pendingCount
        });
    }

    stats.sort((a, b) => b.assigned - a.assigned);

    statsBody.innerHTML = stats.map(s => `
        <tr>
            <td>${s.fullName}</td>
            <td>${s.email}</td>
            <td style="text-align: center; font-weight: 600;">${s.assigned}</td>
            <td style="text-align: center; color: #10b981;">${s.completed}</td>
            <td style="text-align: center; color: #f59e0b;">${s.pending}</td>
        </tr>
    `).join('');
}

async function loadPaymentAssignmentHistory() {
    const historyBody = document.getElementById('assignmentHistoryPaymentBody');
    if (!historyBody) return;

    try {
        const assignmentQuery = query(
            collection(db, 'roundRobinAssignmentsPayment'),
            orderBy('assignedAt', 'desc'),
            limit(20)
        );

        const assignmentSnap = await getDocs(assignmentQuery);

        if (assignmentSnap.empty) {
            historyBody.innerHTML = '<tr><td colspan="4" style="text-align: center; padding: 20px; color: #999;">No Payment assignment history yet</td></tr>';
            return;
        }

        const rows = assignmentSnap.docs.map(doc => {
            const data = doc.data();
            const timestamp = data.assignedAt?.toDate?.() || new Date(data.assignedAt);
            const timeStr = timestamp.toLocaleString();

            return `
                <tr>
                    <td>${timeStr}</td>
                    <td>${data.customerName || 'N/A'}</td>
                    <td>${data.assignedToPayment || 'N/A'}</td>
                    <td>${data.assignedBy || 'System'}</td>
                </tr>
            `;
        }).join('');

        historyBody.innerHTML = rows;
    } catch (error) {
        historyBody.innerHTML = '<tr><td colspan="4" style="text-align: center; padding: 20px; color: #999;">Payment assignment history not available</td></tr>';
    }
}

async function handleResetRoundRobinPayment() {
    showConfirmModal(
        'Reset Payment Round-Robin Counter?',
        'This will reset the Payment distribution counter to start from the first Payment user. Are you sure?',
        async () => {
            closeConfirmModal();
            await runWithButtonSpinner('resetRoundRobinPaymentBtn', 'Resetting...', async () => {
                try {
                    const PAYMENT_COUNTER_DOC = doc(db, 'counters', 'roundRobinPayment');
                    await setDoc(PAYMENT_COUNTER_DOC, {
                        lastIndex: -1,
                        lastDate: ''
                    }, { merge: true });

                    await loadPaymentRoundRobinMonitor();

                    showTestResultModal('Payment Round-Robin Reset Result', [
                        { label: 'Counter', value: 'Payment Round-Robin' },
                        { label: 'Reset By', value: currentAdmin?.email || 'Admin' },
                        { label: 'Reset Time', value: new Date().toLocaleString() },
                        { label: 'Action Status', value: 'PASSED', ok: true }
                    ]);
                } catch (error) {
                    showTestResultModal('Payment Round-Robin Reset Result', [
                        { label: 'Counter', value: 'Payment Round-Robin' },
                        { label: 'Error', value: error.message || 'Unknown error' },
                        { label: 'Action Status', value: 'FAILED', ok: false }
                    ]);
                }
            });
        }
    );
}

async function handleTestDistributionPayment() {
    await runWithButtonSpinner('testDistributionPaymentBtn', 'Testing...', async () => {
        try {
            showNotification('Testing Payment round-robin distribution...', 'info');

            const paymentUsers = (await getPaymentUsersForRoundRobin()).map((u) => u.email);

            if (paymentUsers.length === 0) {
                throw new Error('No Payment users found in the system');
            }

            const counterRef = doc(db, 'counters', 'roundRobinPayment');
            const counterSnap = await getDoc(counterRef);
            let lastIndex = -1;
            let lastDate = '';

            if (counterSnap.exists()) {
                const data = counterSnap.data();
                lastIndex = data.lastIndex ?? -1;
                lastDate = data.lastDate || '';
            }

            const nextIndex = (lastIndex + 1) % paymentUsers.length;
            const nextPaymentUser = paymentUsers[nextIndex];

            showTestResultModal('Payment Round-Robin Test Result', [
                { label: 'Total Payment Officers', value: String(paymentUsers.length) },
                { label: 'Last Index', value: String(lastIndex) },
                { label: 'Next Index', value: String(nextIndex) },
                { label: 'Next Payment Officer', value: nextPaymentUser },
                { label: 'Current Date', value: today },
                { label: 'Last Counter Update', value: lastDate || 'Never' },
                { label: 'Test Status', value: 'PASSED', ok: true }
            ]);

            showNotification(`Test successful. Next Payment user: ${nextPaymentUser}`, 'success');
        } catch (error) {
            showNotification('Test failed: ' + error.message, 'error');
        }
    });
}

function formatDepartment(dept) {
    if (!dept) return 'N/A';
    const depts = {
        'hr': 'Human Resources',
        'finance': 'Finance',
        'operations': 'Operations',
        'it': 'IT',
        'legal': 'Legal'
    };
    return depts[dept] || dept;
}

function formatAuditAction(action) {
    return action.split('_').map(word =>
        word.charAt(0).toUpperCase() + word.slice(1)
    ).join(' ');
}

function formatAuditDescription(audit) {
    const resolveName = (emailOrId) => {
        const key = String(emailOrId || '').trim();
        if (!key) return '';
        if (key.includes('@')) {
            const normalized = key.toLowerCase();
            return userEmailNameCache.get(normalized) || adminNames[normalized] || uploaderNames[normalized] || normalized;
        }
        return userIdNameCache.get(key) || key;
    };

    switch(audit.action) {
        case 'user_created': return `New user created: ${resolveName(audit.userEmail || audit.userId)}`;
        case 'user_approved': return `User approved: ${resolveName(audit.userEmail || audit.userId)}`;
        case 'user_rejected': return `User rejected: ${resolveName(audit.userEmail || audit.userId)}`;
        case 'user_updated': return `User details updated: ${audit.userFullName || resolveName(audit.userEmail || audit.userId)}`;
        case 'user_deactivated': return `User account deactivated`;
        case 'user_activated': return `User account activated`;
        case 'user_leave_activated': return `${resolveName(audit.userEmail || audit.userId)} placed on leave; ${audit.movedCount || 0} application(s) moved to ${resolveName(audit.relieverEmail)}`;
        case 'user_leave_resumed': return `${resolveName(audit.userEmail || audit.userId)} resumed; ${audit.returnedCount || 0} application(s) returned and ${audit.finalizedCount || 0} finalized`;
        case 'user_deleted': return `User permanently deleted: ${resolveName(audit.userEmail || audit.userId)}`;
        case 'document_uploaded': return `Document uploaded`;
        case 'document_approved': return `Document approved`;
        case 'document_rejected': return `Document rejected`;
        default: return audit.action || 'Action performed';
    }
}

function formatDate(timestamp) {
    return formatAppDateTime(timestamp, 'N/A');
}

// ==================== SIGN OUT ====================
window.signOutUser = () => {
    window.location.href = 'index.html';
};

// ==================== MAKE FUNCTIONS GLOBAL ====================
if (typeof signOutUser === 'function' && typeof window.signOutUser !== 'function') {
    window.signOutUser = signOutUser;
}
if (typeof viewDocument === 'function') window.viewDocument = viewDocument;
if (typeof viewUser === 'function') window.viewUser = viewUser;
if (typeof editUser === 'function') window.editUser = editUser;
if (typeof viewSubmissionDocs === 'function') window.viewSubmissionDocs = viewSubmissionDocs;
if (typeof downloadAllSubmission === 'function') window.downloadAllSubmission = downloadAllSubmission;
if (typeof activatePendingUser === 'function') window.activatePendingUser = activatePendingUser;
if (typeof rejectPendingUser === 'function') window.rejectPendingUser = rejectPendingUser;
if (typeof deactivateUser === 'function') window.deactivateUser = deactivateUser;
if (typeof activateUser === 'function') window.activateUser = activateUser;
if (typeof deleteUser === 'function') window.deleteUser = deleteUser;
window.closeUserModal = closeUserModal;
window.closeViewUserModal = closeViewUserModal;
window.closeTestResultModal = closeTestResultModal;
