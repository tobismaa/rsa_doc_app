// js/admin.js - COMPLETE UPDATED VERSION WITH FIXED DOWNLOAD ALL
import { auth, db } from './firebase-config.js';
import { ADMIN_API_BASE_URL } from './admin-api-config.js';
import { notifyStatusChangePush } from './status-push.js';
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
let adminNames = {};
let uploaderNames = {};
let userIdNameCache = new Map();
let userEmailNameCache = new Map();
const TRACK_PAGE_SIZE = 10;
let trackAppsPage = 1;
let currentParentTab = 'user-management';
let currentLeafTab = 'users';

const TAB_GROUPS = {
    'user-management': ['users', 'pending-users', 'pending-agents', 'registered-agents'],
    'application-management': ['pending-docs', 'approved-docs', 'rejected-docs', 'track-apps', 'finally-submitted', 'payments']
};

const TAB_LABELS = {
    users: 'Users',
    'pending-users': 'Pending Users',
    'pending-agents': 'Pending Agents',
    'registered-agents': 'Registered Agents',
    'pending-docs': 'Pending',
    'approved-docs': 'Approved',
    'rejected-docs': 'Rejected',
    'track-apps': 'Track Applications',
    'finally-submitted': 'Final Submission',
    payments: 'Payment'
};

const REVIEWER_ROLE_ALIASES = new Set(['reviewer', 'viewer']);
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
    if (normalized === 'rsa') return 'RSA';
    if (normalized === 'payment') return 'Payment';
    return normalized.charAt(0).toUpperCase() + normalized.slice(1);
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
                role: normalizeUserRole(data.role)
            };
        })
        .filter((u) => u.email && u.role === 'reviewer')
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
                role: normalizeUserRole(data.role)
            };
        })
        .filter((u) => u.email && u.role === 'payment')
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
const pendingDocsTableBody = document.getElementById('pendingDocsTableBody');
const approvedDocsTableBody = document.getElementById('approvedDocsTableBody');
const rejectedDocsTableBody = document.getElementById('rejectedDocsTableBody');
const paymentsTableBody = document.getElementById('paymentsTableBody');
const trackAppsTableBody = document.getElementById('trackAppsTableBody');
const trackPrevPageBtn = document.getElementById('trackPrevPageBtn');
const trackNextPageBtn = document.getElementById('trackNextPageBtn');
const trackPageInfo = document.getElementById('trackPageInfo');
const trackJumpPageInput = document.getElementById('trackJumpPageInput');
const trackJumpPageBtn = document.getElementById('trackJumpPageBtn');
const auditTableBody = document.getElementById('auditTableBody');
const notification = document.getElementById('notification');
const viewerModal = document.getElementById('viewerModal');
const viewerFileName = document.getElementById('viewerFileName');
const documentViewer = document.getElementById('documentViewer');
const profileNameEl = document.getElementById('profileName');
const profileRegisteredAtEl = document.getElementById('profileRegisteredAt');
const profileEmailEl = document.getElementById('profileEmail');
const profileWhatsappEl = document.getElementById('profileWhatsapp');
const profileLocationEl = document.getElementById('profileLocation');
const profileRoleEl = document.getElementById('profileRole');
const profileStatusEl = document.getElementById('profileStatus');
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
                        <div style="background:white;padding:15px;border-radius:8px;border:1px solid #e5e7eb;"><div style="font-size:12px;color:#999;margin-bottom:5px;">Total Payment Users</div><div id="totalPaymentUserCount" style="font-size:16px;font-weight:600;color:#003366;">Loading...</div></div>
                    </div>
                </div>
                <div style="background:#f8f9fa;padding:20px;border-radius:8px;margin-bottom:25px;">
                    <h3>Distribution Statistics (Today)</h3>
                    <div class="table-container" style="max-height:400px;overflow-y:auto;">
                        <table class="documents-table" style="margin-bottom:0;">
                            <thead><tr><th>Payment User Name</th><th>Email</th><th>Assigned Today</th><th>Completed</th><th>Pending</th></tr></thead>
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

    const response = await fetch(cleanUrl, { mode: 'cors', credentials: 'omit' });
    if (!response.ok) {
        throw new Error(`Document fetch failed: ${response.status}`);
    }
    return response;
}

// ==================== DOWNLOAD HELPER FUNCTIONS ====================
function downloadBlobAsFile(blob, fileName) {
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = fileName;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(url);
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
        downloadBlobAsFile(blob, defaultFileName);
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
            console.error('Save error (folder picker):', error);
            showNotification('Save failed: ' + error.message, 'error');
            downloadBlobAsFile(blob, defaultFileName);
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
            const userByUid = query(collection(db, 'users'), where('uid', '==', user.uid));
            const uidSnapshot = await getDocs(userByUid);

            if (!uidSnapshot.empty) {
                userData = uidSnapshot.docs[0].data();
            } else if (user.email) {
                const normalizedEmail = user.email.toLowerCase();
                const userByEmail = query(collection(db, 'users'), where('email', '==', normalizedEmail));
                const emailSnapshot = await getDocs(userByEmail);
                if (!emailSnapshot.empty) {
                    userData = emailSnapshot.docs[0].data();
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
                        userData = matchedDoc.data();
                    }
                }
            }

            if (!userData) {
                showNotification('User profile not found in database.', 'error');
                window.location.href = 'index.html';
                return;
            }

            if (userData.role === 'super_admin') {
                window.location.href = 'super-admin-dashboard.html';
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
    document.getElementById('editFromViewBtn')?.addEventListener('click', () => {
        if (selectedUserId) {
            window.editUser(selectedUserId);
        } else {
            showNotification('No user selected', 'error');
        }
    });
    document.getElementById('userForm')?.addEventListener('submit', saveUser);
    document.getElementById('closeUserModalBtn')?.addEventListener('click', closeUserModal);
    document.getElementById('cancelUserModalBtn')?.addEventListener('click', closeUserModal);

    document.getElementById('closeConfirmModal')?.addEventListener('click', closeConfirmModal);
    document.getElementById('cancelConfirm')?.addEventListener('click', closeConfirmModal);

    document.getElementById('userSearch')?.addEventListener('input', filterUsers);
    document.getElementById('userRoleFilter')?.addEventListener('change', filterUsers);
    document.getElementById('userStatusFilter')?.addEventListener('change', filterUsers);

    document.getElementById('pendingDocSearch')?.addEventListener('input', filterPendingDocs);
    document.getElementById('approvedDocSearch')?.addEventListener('input', filterApprovedDocs);
    document.getElementById('rejectedDocSearch')?.addEventListener('input', filterRejectedDocs);

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
        const userModal = document.getElementById('userModal');
        if (e.target === userModal) closeUserModal();
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
        return `
            <button
                type="button"
                class="action-btn admin-subtab-btn ${active ? 'active' : ''}"
                data-subtab="${leaf}"
            >${label}</button>
        `;
    }).join('');

    host.querySelectorAll('[data-subtab]').forEach((btn) => {
        btn.addEventListener('click', () => switchLeafTab(btn.getAttribute('data-subtab')));
    });
}

function runTabEffects(tabId) {
    if (tabId === 'track-apps') {
        renderTrackApplications();
    }
    if (tabId === 'finally-submitted') {
        renderFinallySubmitted();
    }
    if (tabId === 'payments') {
        renderPaymentQueue();
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
        'pending-docs': 'Application Management - Pending',
        'approved-docs': 'Application Management - Approved',
        'rejected-docs': 'Application Management - Rejected',
        'track-apps': 'Application Management - Track Applications',
        'finally-submitted': 'Application Management - Final Submission',
        payments: 'Application Management - Payment',
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
    }, () => {
        renderPendingAgentsTable([]);
    });
}

function loadApprovedAgents() {
    const q = query(collection(db, 'agents'), where('status', '==', 'approved'));
    onSnapshot(q, (snapshot) => {
        allApprovedAgents = snapshot.docs.map((docSnap) => ({ id: docSnap.id, ...(docSnap.data() || {}) }));
        renderApprovedAgentsTable(allApprovedAgents);
    }, () => {
        renderApprovedAgentsTable([]);
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
        usersTableBody.innerHTML = '<tr><td colspan="7" class="no-data">No users found</td></tr>';
        return;
    }

    usersTableBody.innerHTML = users.map(user => {
        const fullName = user.fullName || user.email?.split('@')[0] || 'Unknown';
        const normalizedRole = normalizeUserRole(user.role);
        const roleLabel = getRoleLabel(user.role);
        return `
            <tr data-user-id="${user.id}">
                <td><strong>${fullName}</strong></td>
                <td>${user.email}</td>
                <td>${renderWhatsAppContactCell(user)}</td>
                <td><span class="role-badge ${normalizedRole}">${roleLabel}</span></td>
                <td><span class="status-badge ${user.status || 'active'}">${user.status || 'active'}</span></td>
                <td>${formatDate(user.createdAt)}</td>
                <td>
                    <div class="action-buttons">
                        <button class="action-btn view-btn" onclick="window.viewUser('${user.id}')" title="View">
                            <i class="fas fa-eye"></i>
                        </button>
                        <button class="action-btn edit-btn" onclick="window.editUser('${user.id}')" title="Edit">
                            <i class="fas fa-edit"></i>
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

        renderPendingDocs();
        renderApprovedDocs();
        renderRejectedDocs();
        renderTrackApplications();
        renderFinallySubmitted();
        renderPaymentQueue();
        updatePendingDocCount(allSubmissions.filter(s => s.status === 'pending'));
        updateRejectedDocCount(allSubmissions.filter(s => s.status === 'rejected'));
        updateFinallySubmittedCount();
        updatePaymentPendingCount();
    }, (error) => {
        // Silent fail
    });
}

// ==================== UPDATE COUNTS ====================
function updatePendingUserCount(users) {
    const pendingCount = users.length;
    if (pendingUserCountBadge) {
        pendingUserCountBadge.textContent = pendingCount;
        pendingUserCountBadge.style.display = pendingCount > 0 ? 'inline' : 'none';
    }
}

function updatePendingDocCount(items) {
    const pendingCount = Array.isArray(items) ? items.filter(u => u.status === 'pending').length : 0;
    if (pendingDocCountBadge) {
        pendingDocCountBadge.textContent = pendingCount;
        pendingDocCountBadge.style.display = pendingCount > 0 ? 'inline' : 'none';
    }
}

function updateRejectedDocCount(items) {
    const rejectedCount = Array.isArray(items) ? items.filter(u => u.status === 'rejected').length : 0;
    if (rejectedDocCountBadge) {
        rejectedDocCountBadge.textContent = rejectedCount;
        rejectedDocCountBadge.style.display = rejectedCount > 0 ? 'inline' : 'none';
    }
}

// ==================== RENDER PENDING DOCUMENTS ====================
function renderPendingDocs() {
    if (!pendingDocsTableBody) return;

    const pending = allSubmissions.filter(s => s.status === 'pending');

    if (pending.length === 0) {
        pendingDocsTableBody.innerHTML = '<tr><td colspan="7" class="no-data">No pending documents</td></tr>';
        return;
    }

    pendingDocsTableBody.innerHTML = pending.map(sub => {
        const uploadDate = formatDate(sub.uploadedAt);
        const uploaderFullName = uploaderNames[sub.uploadedBy] || sub.uploadedBy?.split('@')[0] || 'Unknown';
        const assignedEmail = (sub.assignedTo || '').toString().trim().toLowerCase();
        const assignedLabel = assignedEmail
            ? (uploaderNames[assignedEmail] || assignedEmail.split('@')[0] || assignedEmail)
            : 'Unassigned';
        const rejectionReason = sub.comment || '-';

        return `
            <tr>
                <td><strong>${sub.customerName || 'Unknown'}</strong></td>
                <td>${uploadDate}</td>
                <td>${uploaderFullName}</td>
                <td>${assignedLabel}</td>
                <td><span class="status-badge status-pending">Pending</span></td>
                <td>${rejectionReason}</td>
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

// ==================== RENDER APPROVED DOCUMENTS ====================
function renderApprovedDocs() {
    if (!approvedDocsTableBody) return;

    const approved = allSubmissions.filter(s => {
        const status = String(s.status || '').toLowerCase();
        return status === 'processing_to_pfa' || status === 'approved';
    });

    if (approved.length === 0) {
        approvedDocsTableBody.innerHTML = '<tr><td colspan="6" class="no-data">No approved documents</td></tr>';
        return;
    }

    approvedDocsTableBody.innerHTML = approved.map(sub => {
        const uploadDate = formatDate(sub.uploadedAt);
        const approvedDate = formatDate(sub.reviewedAt);
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

// ==================== RENDER REJECTED DOCUMENTS ====================
function renderRejectedDocs() {
    if (!rejectedDocsTableBody) return;

    const rejected = allSubmissions.filter(s => s.status === 'rejected');

    if (rejected.length === 0) {
        rejectedDocsTableBody.innerHTML = '<tr><td colspan="7" class="no-data">No rejected documents</td></tr>';
        return;
    }

    rejectedDocsTableBody.innerHTML = rejected.map(sub => {
        const uploadDate = formatDate(sub.uploadedAt);
        const rejectedDate = formatDate(sub.reviewedAt);
        const uploaderFullName = uploaderNames[sub.uploadedBy] || sub.uploadedBy?.split('@')[0] || 'Unknown';
        const rejecterFullName = adminNames[sub.reviewedBy] || sub.reviewedBy?.split('@')[0] || 'N/A';
        const rejectionReason = sub.comment || '-';

        return `
            <tr>
                <td><strong>${sub.customerName || 'Unknown'}</strong></td>
                <td>${uploadDate}</td>
                <td>${uploaderFullName}</td>
                <td>${rejecterFullName}</td>
                <td>${rejectedDate}</td>
                <td><span class="rejection-comment">${rejectionReason}</span></td>
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

// ==================== RENDER FINALLY SUBMITTED APPLICATIONS ====================
function renderFinallySubmitted() {
    const finallySubmittedTableBody = document.getElementById('finallySubmittedTableBody');
    if (!finallySubmittedTableBody) return;

    const finallySubmitted = allSubmissions.filter(s => s.finalSubmitted === true || s.rsaSubmitted === true);

    if (finallySubmitted.length === 0) {
        finallySubmittedTableBody.innerHTML = '<tr><td colspan="9" class="no-data">No finally submitted applications</td></tr>';
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
        paymentsTableBody.innerHTML = '<tr><td colspan="7" class="no-data">No payment records available</td></tr>';
        return;
    }

    paymentsTableBody.innerHTML = paymentQueue.map((sub) => {
        const { pfa, twentyFive, commission2 } = getSubmissionFinancials(sub);
        const status = String(sub.status || '').toLowerCase();
        const isPaid = status === 'paid';
        const statusLabel = isPaid ? 'Paid' : 'Sent to PFA';
        const actionHtml = isPaid
            ? '<button class="action-btn" style="opacity:.65;cursor:not-allowed;" disabled><i class="fas fa-check"></i> Paid</button>'
            : `<button class="action-btn" style="background:#16a34a;color:#fff;border:none;" onclick="window.markSubmissionPaid('${sub.id}')"><i class="fas fa-check-circle"></i> Paid</button>`;

        return `
            <tr>
                <td><strong>${escapeHtml(sub.customerName || 'Unknown')}</strong></td>
                <td>${escapeHtml(pfa)}</td>
                <td>${escapeHtml(sub.agentName || '-')}</td>
                <td>${formatCurrency(twentyFive)}</td>
                <td>${formatCurrency(commission2)}</td>
                <td><span class="status-badge status-approved">${statusLabel}</span></td>
                <td>${actionHtml} <button class="action-btn" onclick="window.openApplicationChat('${sub.id}')"><i class="fas fa-comments"></i> Chat</button></td>
            </tr>
        `;
    }).join('');
}

// ==================== UPDATE FINALLY SUBMITTED COUNT ====================
function updateFinallySubmittedCount() {
    const cnt = allSubmissions.filter(s => s.finalSubmitted === true || s.rsaSubmitted === true).length;
    const badge = document.getElementById('finallySubmittedCount');
    if (badge) {
        badge.textContent = cnt;
        badge.style.display = cnt > 0 ? 'inline' : 'none';
    }
}

function updatePaymentPendingCount() {
    const cnt = allSubmissions.filter((s) => {
        const status = String(s.status || '').toLowerCase();
        return status === 'sent_to_pfa' || status === 'rsa_submitted';
    }).length;
    if (paymentPendingCountBadge) {
        paymentPendingCountBadge.textContent = cnt;
        paymentPendingCountBadge.style.display = cnt > 0 ? 'inline' : 'none';
    }
}

function parseMoney(value) {
    const raw = String(value ?? '').replace(/[^0-9.\-]/g, '');
    const num = Number(raw);
    return Number.isFinite(num) ? num : 0;
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
    const computed25 = Math.ceil((rsaBalance * 0.25) / 100) * 100;
    const twentyFive = parseMoney(details.rsa25Percent || sub?.rsa25Percent || computed25);
    const commission2 = twentyFive * 0.02;
    const pfa = String(details.pfa || sub?.pfa || '').trim() || '-';
    return { pfa, twentyFive, commission2 };
}

function formatStatusLabel(status) {
    const normalized = String(status || '').toLowerCase().trim();
    if (!normalized) return '-';
    if (normalized === 'processing_to_pfa' || normalized === 'approved') return 'Processing to PFA';
    if (normalized === 'sent_to_pfa' || normalized === 'rsa_submitted') return 'Sent to PFA';
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

// ==================== VIEW ALL DOCUMENTS FOR A SUBMISSION ====================
window.viewSubmissionDocs = (submissionId) => {
    const sub = allSubmissions.find(s => s.id === submissionId);
    if (!sub || !sub.documents || sub.documents.length === 0) {
        showNotification('No documents available', 'error');
        return;
    }

    const firstDoc = sub.documents[0];
    const docTypeLabel = DOCUMENT_TYPES[firstDoc.documentType] || firstDoc.documentType || 'Document';

    if (viewerModal && viewerFileName && documentViewer) {
        viewerFileName.textContent = `${sub.customerName} - ${docTypeLabel}`;
        const cleanUrl = firstDoc.fileUrl?.trim();
        documentViewer.src = cleanUrl || '';
        viewerModal.classList.add('active');
    }

    if (sub.documents.length > 1) {
        let currentIndex = 0;

        const showDoc = (index) => {
            const doc = sub.documents[index];
            const docTypeLabel = DOCUMENT_TYPES[doc.documentType] || doc.documentType || 'Document';
            viewerFileName.textContent = `${sub.customerName} - ${docTypeLabel} (${index + 1}/${sub.documents.length})`;
            const cleanUrl = doc.fileUrl?.trim();
            documentViewer.src = cleanUrl || '';
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
                <span id="docCounter" style="font-size: 14px; color: #666;">${currentIndex + 1}/${sub.documents.length}</span>
                <button id="nextDoc" class="action-btn" ${currentIndex === sub.documents.length - 1 ? 'disabled' : ''}>
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
                if (currentIndex < sub.documents.length - 1) {
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
    const sub = allSubmissions.find(s => s.id === submissionId);
    if (!sub) {
        showNotification('Submission not found', 'error');
        return;
    }

    const docs = sub.documents || [];
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
                console.log('Folder picker not supported or cancelled, falling back to individual downloads');
            }
        }

        let successCount = 0;
        let failedCount = 0;

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
                        downloadBlobAsFile(blob, fileName);
                    }
                }

                successCount++;
                updateProgress(i + 1, docs.length, `Completed ${i + 1} of ${docs.length}`, '');

            } catch (docError) {
                console.error(`Error downloading document ${i + 1}:`, docError);
                failedCount++;
                updateProgress(i + 1, docs.length, `Failed document ${i + 1}`, docItem.name || 'Unknown');
                showNotification(`Failed to download: ${docItem.name}`, 'error');
            }
        }

        // Close progress modal
        progressModal.classList.remove('active');
        setTimeout(() => progressModal.remove(), 300);

        if (!cancelled) {
            if (failedCount === 0) {
                if (useFolderPicker) {
                    showNotification(`✅ All ${successCount} documents saved to folder: ${safeCustomerName}`, 'success');
                } else {
                    showNotification(`✅ All ${successCount} documents downloaded successfully`, 'success');
                }
            } else {
                showNotification(`⚠️ Downloaded ${successCount} documents, ${failedCount} failed`, 'warning');
            }
        }

    } catch (error) {
        if (error?.name === 'AbortError') {
            showNotification('Download cancelled', 'info');
        } else {
            console.error('Download all error:', error);
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
    const statusClass = (statusRaw === 'processing_to_pfa' || statusRaw === 'approved') ? 'status-approved' : (statusRaw === 'rejected' ? 'status-rejected' : 'pending');
    const statusLabel = formatStatusLabel(statusRaw || 'pending');

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
                    <div class="detail-row">
                        <span class="detail-label">Department:</span>
                        <span class="detail-value">${user.department || 'N/A'}</span>
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
    selectedUserId = userId;
    await loadUserForEdit(userId);
    const modalTitle = document.getElementById('modalTitle');
    if (modalTitle) modalTitle.textContent = 'Edit User';
    closeViewUserModal();
    document.getElementById('userModal')?.classList.add('active');
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

function renderTrackApplications() {
    if (!trackAppsTableBody) return;

    const search = String(document.getElementById('trackUserSearch')?.value || '').trim().toLowerCase();
    const statusFilter = String(document.getElementById('trackStatusFilter')?.value || 'all').trim().toLowerCase();
    const startVal = document.getElementById('trackStartDate')?.value;
    const endVal = document.getElementById('trackEndDate')?.value;

    const start = startVal ? new Date(startVal) : null;
    const end = endVal ? new Date(endVal) : null;
    if (end) end.setHours(23, 59, 59, 999);

    let list = Array.isArray(allSubmissions) ? [...allSubmissions] : [];

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
    const sub = allSubmissions.find((s) => s.id === submissionId);
    if (!sub) {
        showNotification('Submission not found', 'error');
        return;
    }

    const confirmed = confirm(`Mark ${sub.customerName || 'this customer'} as PAID?`);
    if (!confirmed) return;

    try {
        await updateDoc(doc(db, 'submissions', submissionId), {
            status: 'paid',
            paidAt: serverTimestamp(),
            paidBy: currentAdmin?.email || ''
        });

        await addDoc(collection(db, 'audit'), {
            action: 'submission_paid',
            submissionId,
            customerName: sub.customerName || '',
            performedBy: currentAdmin?.email || '',
            timestamp: serverTimestamp()
        });
        notifyStatusChangePush({
            currentUser: auth.currentUser,
            submissionId,
            customerName: sub.customerName || '',
            newStatus: 'paid',
            statusLabel: 'Paid',
            actionLabel: 'Application Marked Paid',
            message: `Application for ${sub.customerName || 'this customer'} was marked as paid.`
        }).catch(() => {});

        showNotification('Marked as paid successfully', 'success');
    } catch (error) {
        showNotification('Failed to mark paid: ' + (error?.message || 'Unknown error'), 'error');
    }
};

window.clearPaidSubmissions = async () => {
    const paidItems = allSubmissions.filter((s) => String(s.status || '').toLowerCase() === 'paid');
    if (paidItems.length === 0) {
        showNotification('No paid records to clear', 'info');
        return;
    }

    const confirmed = confirm(`Clear ${paidItems.length} paid record(s)?\nThis will change their status to Cleared.`);
    if (!confirmed) return;

    try {
        await Promise.all(paidItems.map((sub) => updateDoc(doc(db, 'submissions', sub.id), {
            status: 'cleared',
            clearedAt: serverTimestamp(),
            clearedBy: currentAdmin?.email || ''
        })));

        await addDoc(collection(db, 'audit'), {
            action: 'paid_records_cleared',
            count: paidItems.length,
            performedBy: currentAdmin?.email || '',
            timestamp: serverTimestamp()
        });

        showNotification(`Cleared ${paidItems.length} paid record(s)`, 'success');
    } catch (error) {
        showNotification('Failed to clear paid records: ' + (error?.message || 'Unknown error'), 'error');
    }
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
            showNotification('Another user already has this email address', 'error');
            return;
        }

        if (userData.whatsappNumber) {
            const waDupSnap = await getDocs(query(collection(db, 'users'), where('whatsappNumber', '==', userData.whatsappNumber)));
            const waDuplicate = waDupSnap.docs.find((d) => d.id !== selectedUserId);
            if (waDuplicate) {
                showNotification('Another user already has this WhatsApp number', 'error');
                return;
            }

            const legacyPhoneDupSnap = await getDocs(query(collection(db, 'users'), where('phone', '==', userData.whatsappNumber)));
            const phoneDuplicate = legacyPhoneDupSnap.docs.find((d) => d.id !== selectedUserId);
            if (phoneDuplicate) {
                showNotification('Another user already has this WhatsApp number', 'error');
                return;
            }
        }

        if (selectedUserId) {
            await updateDoc(doc(db, 'users', selectedUserId), userData);

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

            await addDoc(collection(db, 'users'), {
                ...userData,
                uid: userCredential.user.uid,
                createdAt: serverTimestamp(),
                createdBy: currentAdmin?.email
            });

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
        showNotification('Failed to save user: ' + error.message, 'error');
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

function filterPendingDocs() {
    const searchTerm = document.getElementById('pendingDocSearch')?.value.toLowerCase() || '';
    const pending = allSubmissions.filter(s => s.status === 'pending');

    const filtered = pending.filter(sub => {
        const matchesSearch = searchTerm === '' ||
            sub.customerName?.toLowerCase().includes(searchTerm) ||
            sub.uploadedBy?.toLowerCase().includes(searchTerm) ||
            sub.assignedTo?.toLowerCase().includes(searchTerm);
        return matchesSearch;
    });

    if (!pendingDocsTableBody) return;

    if (filtered.length === 0) {
        pendingDocsTableBody.innerHTML = '<tr><td colspan="7" class="no-data">No pending documents</td></tr>';
        return;
    }

    pendingDocsTableBody.innerHTML = filtered.map(sub => {
        const uploadDate = sub.uploadedAt ?
            (sub.uploadedAt.toDate ? sub.uploadedAt.toDate() : new Date(sub.uploadedAt)).toLocaleString() : 'N/A';
        const uploaderFullName = uploaderNames[sub.uploadedBy] || sub.uploadedBy?.split('@')[0] || 'Unknown';
        const assignedEmail = (sub.assignedTo || '').toString().trim().toLowerCase();
        const assignedLabel = assignedEmail
            ? (uploaderNames[assignedEmail] || assignedEmail.split('@')[0] || assignedEmail)
            : 'Unassigned';
        const rejectionReason = sub.comment || '-';

        return `
            <tr>
                <td><strong>${sub.customerName || 'Unknown'}</strong></td>
                <td>${uploadDate}</td>
                <td>${uploaderFullName}</td>
                <td>${assignedLabel}</td>
                <td><span class="status-badge status-pending">Pending</span></td>
                <td>${rejectionReason}</td>
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
    const approved = allSubmissions.filter(s => {
        const status = String(s.status || '').toLowerCase();
        return status === 'processing_to_pfa' || status === 'approved';
    });

    const filtered = approved.filter(sub => {
        const matchesSearch = searchTerm === '' ||
            sub.customerName?.toLowerCase().includes(searchTerm) ||
            sub.uploadedBy?.toLowerCase().includes(searchTerm);
        return matchesSearch;
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
    const rejected = allSubmissions.filter(s => s.status === 'rejected');

    const filtered = rejected.filter(sub => {
        const matchesSearch = searchTerm === '' ||
            sub.customerName?.toLowerCase().includes(searchTerm) ||
            sub.uploadedBy?.toLowerCase().includes(searchTerm);
        return matchesSearch;
    });

    if (!rejectedDocsTableBody) return;

    if (filtered.length === 0) {
        rejectedDocsTableBody.innerHTML = '<tr><td colspan="7" class="no-data">No rejected documents</td></tr>';
        return;
    }

    rejectedDocsTableBody.innerHTML = filtered.map(sub => {
        const uploadDate = sub.uploadedAt ?
            (sub.uploadedAt.toDate ? sub.uploadedAt.toDate() : new Date(sub.uploadedAt)).toLocaleString() : 'N/A';
        const rejectedDate = sub.reviewedAt ?
            (sub.reviewedAt.toDate ? sub.reviewedAt.toDate() : new Date(sub.reviewedAt)).toLocaleString() : 'N/A';
        const uploaderFullName = uploaderNames[sub.uploadedBy] || sub.uploadedBy?.split('@')[0] || 'Unknown';
        const rejecterFullName = adminNames[sub.reviewedBy] || sub.reviewedBy?.split('@')[0] || 'N/A';
        const rejectionReason = sub.comment || '-';

        return `
            <tr>
                <td><strong>${sub.customerName || 'Unknown'}</strong></td>
                <td>${uploadDate}</td>
                <td>${uploaderFullName}</td>
                <td>${rejecterFullName}</td>
                <td>${rejectedDate}</td>
                <td><span class="rejection-comment">${rejectionReason}</span></td>
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

// ==================== MODAL CONTROLS ====================
function closeUserModal() {
    document.getElementById('userModal').classList.remove('active');
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
        confirmBtn.parentNode.replaceChild(newConfirmBtn, confirmBtn);
        newConfirmBtn.addEventListener('click', async () => {
            await onConfirm();
        });
    }
}

function closeConfirmModal() {
    document.getElementById('confirmModal').classList.remove('active');
}

// ==================== NOTIFICATION SYSTEM ====================
function showNotification(message, type = 'info') {
    if (!notification) return;

    notification.textContent = message;
    notification.className = `notification ${type}`;
    notification.style.display = 'block';

    setTimeout(() => {
        notification.style.display = 'none';
    }, 3000);
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
    const today = new Date().toISOString().slice(0, 10);
    const statsBody = document.getElementById('distributionStatsBody');
    if (!statsBody) return;

    const stats = [];

    for (const viewer of viewers) {
        const assignedQuery = query(
            collection(db, 'submissions'),
            where('assignedTo', '==', viewer.email),
            where('uploadedAt', '>=', new Date(today + 'T00:00:00Z'))
        );
        const assignedSnap = await getDocs(assignedQuery);
        const assignedCount = assignedSnap.size;

        const completedDocs = assignedSnap.docs.filter(d => ['approved', 'rejected'].includes(d.data().status));
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

            const today = new Date().toISOString().slice(0, 10);
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
            fullName: d.data().fullName || d.data().email
        })).sort((a, b) => a.email.localeCompare(b.email));

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
    const today = new Date().toISOString().slice(0, 10);
    const statsBody = document.getElementById('distributionStatsRSABody');
    if (!statsBody) return;

    const stats = [];

    for (const rsa of rsaUsers) {
        const assignedQuery = query(
            collection(db, 'submissions'),
            where('assignedToRSA', '==', rsa.email),
            where('uploadedAt', '>=', new Date(today + 'T00:00:00Z'))
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
            const rsaUsers = rsaSnap.docs.map(d => d.data().email).filter(Boolean).sort();

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

            const today = new Date().toISOString().slice(0, 10);
            if (lastDate !== today) lastIndex = -1;

            const nextIndex = (lastIndex + 1) % rsaUsers.length;
            const nextRSA = rsaUsers[nextIndex];

            showTestResultModal('RSA Round-Robin Test Result', [
                { label: 'Total RSA Users', value: String(rsaUsers.length) },
                { label: 'Last Index', value: String(lastIndex) },
                { label: 'Next Index', value: String(nextIndex) },
                { label: 'Next RSA User', value: nextRSA },
                { label: 'Current Date', value: today },
                { label: 'Last Reset Date', value: lastDate || 'Never' },
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
    const today = new Date().toISOString().slice(0, 10);
    const statsBody = document.getElementById('distributionStatsPaymentBody');
    if (!statsBody) return;

    const stats = [];

    for (const paymentUser of paymentUsers) {
        const assignedQuery = query(
            collection(db, 'submissions'),
            where('assignedToPayment', '==', paymentUser.email),
            where('uploadedAt', '>=', new Date(today + 'T00:00:00Z'))
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

            const today = new Date().toISOString().slice(0, 10);
            if (lastDate !== today) lastIndex = -1;

            const nextIndex = (lastIndex + 1) % paymentUsers.length;
            const nextPaymentUser = paymentUsers[nextIndex];

            showTestResultModal('Payment Round-Robin Test Result', [
                { label: 'Total Payment Users', value: String(paymentUsers.length) },
                { label: 'Last Index', value: String(lastIndex) },
                { label: 'Next Index', value: String(nextIndex) },
                { label: 'Next Payment User', value: nextPaymentUser },
                { label: 'Current Date', value: today },
                { label: 'Last Reset Date', value: lastDate || 'Never' },
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
        case 'user_deleted': return `User permanently deleted: ${resolveName(audit.userEmail || audit.userId)}`;
        case 'document_uploaded': return `Document uploaded`;
        case 'document_approved': return `Document approved`;
        case 'document_rejected': return `Document rejected`;
        default: return audit.action || 'Action performed';
    }
}

function formatDate(timestamp) {
    if (!timestamp) return 'N/A';
    try {
        const date = timestamp.toDate ? timestamp.toDate() : new Date(timestamp);
        return date.toLocaleString('en-NG', {
            day: 'numeric',
            month: 'short',
            year: 'numeric',
            hour: '2-digit',
            minute: '2-digit'
        });
    } catch (error) {
        return 'Invalid date';
    }
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
