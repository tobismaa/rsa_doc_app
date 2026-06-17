// js/admin.js - COMPLETE UPDATED VERSION WITH FIXED DOWNLOAD ALL
import { auth, db } from './firebase-config.js';
import { ADMIN_API_BASE_URL } from './admin-api-config.js';
import { notifyUserPushEvent } from './push-alerts.js';
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
import { getSystemSettings } from './shared/system-settings.js?v=20260615a';
import {
    getTimestampMillis as getStageTimestampMillis,
    getSubmissionCurrentStageEntryAt,
    getSubmissionDraftEntryAt,
    getSubmissionReviewEntryAt,
    getSubmissionApprovalEntryAt,
    getSubmissionRejectionEntryAt,
    getSubmissionFinalSubmissionEntryAt,
    getSubmissionPaymentEntryAt,
    getSubmissionPaidEntryAt,
    getSubmissionClearedEntryAt
} from './shared/submission-stage.js?v=20260609a';

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
let trackReportParsedNames = [];
let trackReportPreviewRows = [];
let trackReportUnmatchedEntries = [];
let currentDocumentGenerationSubmissionId = '';
let generatedDocumentPreviewItems = [];
let pdfFontAssetCache = null;
let adminSystemSettings = {};

const GENERATED_DOCUMENT_TYPES = [
    { id: 'offer_letter', label: 'Offer Letter', description: 'Mortgage facility offer and terms.' },
    { id: 'allocation_letter', label: 'Allocation Letter', description: 'Property allocation and estate terms.' },
    { id: 'availability_letter', label: 'Availability Letter', description: 'Confirms availability of allocated property.' },
    { id: 'indemnity_letter', label: 'Indemnity Letter', description: 'RSA equity indemnity in favour of the PFA.' },
    { id: 'readiness_letter', label: 'Readiness Letter', description: 'Confirms readiness to receive disbursement.' },
    { id: 'title_letter', label: 'Title Letter', description: 'Confirms title authenticity and search status.' },
    { id: 'verification_letter', label: 'Verification Letter', description: 'Verifies property offer and sale validity.' }
];

const GENERATED_DOCUMENT_MASTER_FILES = {
    offer_letter: 'offer-letter-master.pdf',
    allocation_letter: 'allocation-letter-master.pdf',
    availability_letter: 'availability-letter-master.pdf',
    indemnity_letter: 'indemnity-letter-master.pdf',
    readiness_letter: 'readiness-letter-master.pdf',
    title_letter: 'title-letter-master.pdf',
    verification_letter: 'verification-letter-master.pdf'
};

const GENERATED_DOCUMENT_PAGE_FORMATS_MM = {
    offer_letter: {
        0: [210, 260],
        1: [210, 275],
        2: [210, 275],
        3: [210, 275]
    },
    allocation_letter: {
        0: [215, 240],
        1: [215, 260]
    },
    indemnity_letter: {
        0: [215, 260],
        1: [220, 290]
    },
    verification_letter: {
        0: [210, 240]
    },
    availability_letter: {
        0: [215, 255]
    },
    readiness_letter: {
        0: [210, 260]
    },
    title_letter: {
        0: [215, 260]
    }
};

const PDF_TEMPLATE_CONFIGS = {
    offer_letter: {
        fileName: 'offer-letter-template.pdf',
        wipeZones: [
            { page: 0, rect: [201.26, 336.3, 401.98, 349.2] },
            { page: 0, rect: [327.0, 472.2, 535.75, 502.1] },
            { page: 0, rect: [201.26, 523.2, 535.75, 536.3] },
            { page: 1, rect: [313.0, 91.1, 512.8, 104.0] },
            { page: 1, rect: [414.5, 456.8, 536.0, 486.4] },
            { page: 1, rect: [414.8, 714.8, 536.0, 727.4] },
            { page: 2, rect: [328.0, 125.1, 531.2, 138.0] }
        ],
        fields: [
            { page: 0, key: 'currentDate', rect: [59.53, 134.19, 158.63, 145.83], fontWeight: 'bold' },
            { page: 0, key: 'customerName', rect: [59.53, 154.03, 174.94, 165.67], fontWeight: 'bold' },
            { page: 0, key: 'customerAddress', rect: [59.53, 173.87, 372.49, 185.51], fontWeight: 'regular' },
            { page: 0, key: 'customerName', rect: [201.26, 285.84, 304.01, 297.48], fontWeight: 'regular' },
            { page: 0, key: 'customerAddress', rect: [201.26, 302.85, 480.38, 314.49], fontWeight: 'regular' },
            { page: 0, key: 'offerPropertyValueLine', rect: [201.26, 336.86, 401.98, 348.5], fontWeight: 'regular' },
            { page: 0, key: 'loanAmountText', rect: [201.26, 404.89, 266.88, 416.53], fontWeight: 'regular' },
            { page: 0, key: 'offerPurposeFragment', rect: [327.0, 472.93, 535.75, 501.57], fontWeight: 'regular', allowWrap: true },
            { page: 0, key: 'offerEquityLine', rect: [201.26, 523.95, 535.75, 535.59], fontWeight: 'regular' },
            { page: 1, key: 'customerName', rect: [59.53, 148.36, 174.94, 160.0], fontWeight: 'bold' },
            { page: 1, key: 'offerSecurityFragment', rect: [313.3, 91.67, 512.45, 103.31], fontWeight: 'bold' },
            { page: 1, key: 'offerReevaluateFragment', rect: [415.0, 457.33, 535.75, 485.98], fontWeight: 'bold', allowWrap: true },
            { page: 1, key: 'offerClause6Fragment', rect: [415.11, 715.29, 535.75, 726.93], fontWeight: 'bold' },
            { page: 1, key: 'customerName', rect: [311.39, 763.48, 426.8, 775.12], fontWeight: 'bold' },
            { page: 2, key: 'offerClause8Fragment', rect: [328.48, 125.68, 531.05, 137.32], fontWeight: 'bold' },
            { page: 2, key: 'customerName', rect: [81.79, 173.87, 197.2, 185.51], fontWeight: 'bold' },
            { page: 2, key: 'customerName', rect: [283.91, 321.27, 399.32, 332.91], fontWeight: 'bold' }
        ]
    },
    allocation_letter: {
        fileName: 'allocation-letter-template.pdf',
        wipeZones: [
            { page: 0, rect: [150.0, 282.0, 459.0, 295.0] },
            { page: 1, rect: [334.6, 214.0, 529.2, 226.8] }
        ],
        fields: [
            { page: 0, key: 'currentDate', rect: [59.53, 140.84, 151.67, 152.48], fontWeight: 'bold' },
            { page: 0, key: 'customerName', rect: [59.53, 160.68, 174.94, 172.32], fontWeight: 'bold' },
            { page: 0, key: 'customerAddress', rect: [59.53, 180.52, 372.49, 192.16], fontWeight: 'regular' },
            { page: 0, key: 'allocationHouseLine', rect: [150.77, 282.57, 458.16, 294.21], fontWeight: 'regular' },
            { page: 0, key: 'currentDate', rect: [491.47, 296.74, 510.33, 308.38], fontWeight: 'bold' },
            { page: 0, key: 'currentDateShort', rect: [59.53, 310.92, 129.33, 322.56], fontWeight: 'bold' },
            { page: 1, key: 'allocationValueLine', rect: [335.07, 214.54, 528.92, 226.18], fontWeight: 'regular' }
        ]
    },
    availability_letter: {
        fileName: 'availability-letter-template.pdf',
        wipeZones: [
            { page: 0, rect: [59.0, 366.4, 488.4, 379.4] },
            { page: 0, rect: [59.0, 380.6, 399.6, 393.5] },
            { page: 0, rect: [59.0, 394.8, 527.5, 407.8] }
        ],
        fields: [
            { page: 0, key: 'currentDate', rect: [59.53, 148.84, 158.63, 160.48], fontWeight: 'bold' },
            { page: 0, key: 'pfaName', rect: [59.53, 188.52, 241.92, 200.16], fontWeight: 'bold' },
            { page: 0, key: 'availabilityLine1', rect: [59.53, 367.11, 487.93, 378.75], fontWeight: 'regular' },
            { page: 0, key: 'availabilityLine2', rect: [59.53, 381.28, 398.96, 392.92], fontWeight: 'regular' },
            { page: 0, key: 'availabilityLine3', rect: [59.53, 395.45, 526.97, 407.09], fontWeight: 'regular' }
        ]
    },
    indemnity_letter: {
        fileName: 'indemnity-letter-template.pdf',
        wipeZones: [
            { page: 0, rect: [59.0, 598.3, 191.5, 611.5] },
            { page: 1, rect: [151.8, 555.8, 274.0, 568.9] },
            { page: 1, rect: [147.0, 575.5, 259.2, 588.8] }
        ],
        fields: [
            { page: 0, key: 'pfaName', rect: [59.53, 171.43, 241.92, 183.07] },
            { page: 0, key: 'pfaName', rect: [98.88, 292.16, 281.27, 303.8] },
            { page: 0, key: 'pfaName', rect: [71.35, 610.41, 253.74, 622.05] },
            { page: 0, key: 'customerName', rect: [75.43, 599.07, 190.84, 610.71] },
            { page: 1, key: 'customerName', rect: [157.35, 556.55, 272.76, 568.19] },
            { page: 1, key: 'rsaPin', rect: [152.28, 576.39, 258.33, 588.03] },
            { page: 1, key: 'pfaName', rect: [290.69, 249.1, 473.08, 260.74] },
            { page: 1, key: 'pfaName', rect: [361.1, 639.52, 535.92, 651.16] }
        ]
    },
    readiness_letter: {
        fileName: 'readiness-letter-template.pdf',
        wipeZones: [
            { page: 0, rect: [59.0, 385.5, 212.0, 398.4] },
            { page: 0, rect: [59.0, 405.3, 229.6, 418.1] },
            { page: 0, rect: [59.0, 425.2, 520.5, 438.1] },
            { page: 0, rect: [404.0, 498.7, 521.4, 523.4] }
        ],
        fields: [
            { page: 0, key: 'currentDate', rect: [59.52, 134.07, 157.98, 145.46], fontWeight: 'bold' },
            { page: 0, key: 'pfaName', rect: [59.52, 173.67, 238.7, 185.06], fontWeight: 'bold' },
            { page: 0, key: 'readinessNameLine', rect: [59.52, 386.22, 211.31, 397.61], fontWeight: 'bold' },
            { page: 0, key: 'readinessAccountLine', rect: [59.52, 406.04, 228.9, 417.43], fontWeight: 'bold' },
            { page: 0, key: 'readinessEquityLine', rect: [59.52, 425.96, 519.9, 437.35], fontWeight: 'bold' },
            { page: 0, key: 'readinessLoanLine', rect: [405.0, 499.52, 520.8, 522.55], fontWeight: 'bold', allowWrap: true }
        ]
    },
    title_letter: {
        fileName: 'title-letter-template.pdf',
        wipeZones: [
            { page: 0, rect: [127.5, 340.4, 380.2, 353.5] },
            { page: 0, rect: [442.2, 394.2, 525.3, 421.5] }
        ],
        fields: [
            { page: 0, key: 'currentDate', rect: [59.53, 134.19, 158.63, 145.83], fontWeight: 'bold' },
            { page: 0, key: 'pfaName', rect: [59.53, 173.87, 241.92, 185.51], fontWeight: 'bold' },
            { page: 0, key: 'houseType', rect: [418.49, 312.77, 517.51, 324.41] },
            { page: 0, key: 'houseType', rect: [59.53, 326.94, 189.5, 338.58] },
            { page: 0, key: 'titleHeaderLine', rect: [128.15, 341.11, 379.39, 352.75], fontWeight: 'bold' },
            { page: 0, key: 'titleBodyLine', rect: [442.9, 394.97, 524.68, 420.79], fontWeight: 'bold', allowWrap: true }
        ]
    },
    verification_letter: {
        fileName: 'verification-letter-template.pdf',
        wipeZones: [
            { page: 0, rect: [59.0, 349.3, 515.2, 362.5] },
            { page: 0, rect: [59.0, 363.5, 499.5, 376.6] },
            { page: 0, rect: [59.0, 377.7, 521.4, 390.8] }
        ],
        fields: [
            { page: 0, key: 'currentDate', rect: [59.53, 191.36, 158.63, 203.0], fontWeight: 'bold' },
            { page: 0, key: 'pfaName', rect: [59.53, 231.04, 241.92, 242.68], fontWeight: 'bold' },
            { page: 0, key: 'verificationLine1', rect: [59.53, 350.1, 514.53, 361.74], fontWeight: 'regular' },
            { page: 0, key: 'verificationLine2', rect: [59.53, 364.27, 498.82, 375.91], fontWeight: 'regular' },
            { page: 0, key: 'verificationLine3', rect: [59.53, 378.44, 520.61, 390.08], fontWeight: 'regular' }
        ]
    }
};

const TAB_GROUPS = {
    'user-management': ['users', 'pending-users', 'pending-agents', 'registered-agents'],
    'application-management': ['draft-docs', 'pending-docs', 'approved-docs', 'rejected-docs', 'escalations', 'track-apps', 'generate-documents', 'finally-submitted', 'payments', 'agent-commissions']
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
    'generate-documents': 'Generate Document',
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
    if (tabId === 'generate-documents') return allSubmissions.filter((s) => {
        const status = String(s.status || '').toLowerCase();
        return status === 'processing_to_pfa' || status === 'approved';
    }).length;
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
const generateDocumentsTableBody = document.getElementById('generateDocumentsTableBody');
const openTrackReportInputModalBtn = document.getElementById('openTrackReportInputModalBtn');
const downloadTrackTemplateBtn = document.getElementById('downloadTrackTemplateBtn');
const trackReportFileInput = document.getElementById('trackReportFileInput');
const trackReportNamesInput = document.getElementById('trackReportNamesInput');
const generateTrackReportBtn = document.getElementById('generateTrackReportBtn');
const clearTrackReportBtn = document.getElementById('clearTrackReportBtn');
const trackReportInlineStatus = document.getElementById('trackReportInlineStatus');
const trackReportInputModal = document.getElementById('trackReportInputModal');
const trackReportPreviewModal = document.getElementById('trackReportPreviewModal');
const trackReportPreviewSummary = document.getElementById('trackReportPreviewSummary');
const trackReportPreviewAlerts = document.getElementById('trackReportPreviewAlerts');
const trackReportPreviewTableBody = document.getElementById('trackReportPreviewTableBody');
const downloadTrackReportBtn = document.getElementById('downloadTrackReportBtn');
const documentGenerationModal = document.getElementById('documentGenerationModal');
const documentGenerationCustomerName = document.getElementById('documentGenerationCustomerName');
const documentGenerationMeta = document.getElementById('documentGenerationMeta');
const documentGenerationChecklist = document.getElementById('documentGenerationChecklist');
const generatedDocumentsPreviewModal = document.getElementById('generatedDocumentsPreviewModal');
const generatedDocumentsPreviewMeta = document.getElementById('generatedDocumentsPreviewMeta');
const generatedDocumentsPreviewList = document.getElementById('generatedDocumentsPreviewList');
const saveAllGeneratedDocumentsBtn = document.getElementById('saveAllGeneratedDocumentsBtn');
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
const trackApplicationModal = document.getElementById('trackApplicationModal');
const trackApplicationCustomerName = document.getElementById('trackApplicationCustomerName');
const trackApplicationMeta = document.getElementById('trackApplicationMeta');
const trackApplicationStatusBadges = document.getElementById('trackApplicationStatusBadges');
const trackApplicationSummary = document.getElementById('trackApplicationSummary');
const trackApplicationTimeline = document.getElementById('trackApplicationTimeline');
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
    try {
        const userId = currentAdminProfileData?.id || currentAdmin?.uid || '';
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

async function loadAdminSystemSettings() {
    try {
        adminSystemSettings = await getSystemSettings(db, { force: true });
    } catch (_) {
        adminSystemSettings = {};
    }
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
                loadAdminSystemSettings();
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
    document.getElementById('userOnlineFilter')?.addEventListener('change', filterUsers);
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
    document.getElementById('closeTrackApplicationModal')?.addEventListener('click', closeTrackApplicationModalFn);
    document.getElementById('closeTrackApplicationModalFooterBtn')?.addEventListener('click', closeTrackApplicationModalFn);
    document.getElementById('closeDocumentGenerationModal')?.addEventListener('click', closeDocumentGenerationModalFn);
    document.getElementById('closeDocumentGenerationFooterBtn')?.addEventListener('click', closeDocumentGenerationModalFn);
    document.getElementById('selectAllDocumentsBtn')?.addEventListener('click', toggleAllDocumentSelections);
    document.getElementById('generateSelectedDocumentsBtn')?.addEventListener('click', generateSelectedDocumentsForPreview);
    document.getElementById('closeGeneratedDocumentsPreviewModal')?.addEventListener('click', closeGeneratedDocumentsPreviewModalFn);
    document.getElementById('closeGeneratedDocumentsPreviewFooterBtn')?.addEventListener('click', closeGeneratedDocumentsPreviewModalFn);
    saveAllGeneratedDocumentsBtn?.addEventListener('click', saveAllGeneratedDocumentsToFolder);
    openTrackReportInputModalBtn?.addEventListener('click', openTrackReportInputModalFn);
    document.getElementById('closeTrackReportInputModal')?.addEventListener('click', closeTrackReportInputModalFn);
    document.getElementById('closeTrackReportInputFooterBtn')?.addEventListener('click', closeTrackReportInputModalFn);
    document.getElementById('closeTrackReportPreviewModal')?.addEventListener('click', closeTrackReportPreviewModalFn);
    document.getElementById('closeTrackReportPreviewFooterBtn')?.addEventListener('click', closeTrackReportPreviewModalFn);
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
    downloadTrackTemplateBtn?.addEventListener('click', downloadTrackReportTemplate);
    trackReportFileInput?.addEventListener('change', handleTrackReportFileSelected);
    generateTrackReportBtn?.addEventListener('click', generateTrackReportPreview);
    clearTrackReportBtn?.addEventListener('click', clearTrackReportInputs);
    downloadTrackReportBtn?.addEventListener('click', downloadTrackReportWorkbook);

    document.getElementById('closeViewer')?.addEventListener('click', closeViewerModal);

    window.addEventListener('click', (e) => {
        if (e.target === viewerModal) closeViewerModal();
        if (e.target === adminRejectionReasonModal) closeAdminRejectionReasonModalFn();
        if (e.target === agentCommissionModal) closeAgentCommissionModalFn();
        if (e.target === trackApplicationModal) closeTrackApplicationModalFn();
        if (e.target === documentGenerationModal) closeDocumentGenerationModalFn();
        if (e.target === generatedDocumentsPreviewModal) closeGeneratedDocumentsPreviewModalFn();
        if (e.target === trackReportInputModal) closeTrackReportInputModalFn();
        if (e.target === trackReportPreviewModal) closeTrackReportPreviewModalFn();
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

function closeTrackApplicationModalFn() {
    if (trackApplicationModal) trackApplicationModal.classList.remove('active');
    if (trackApplicationCustomerName) trackApplicationCustomerName.textContent = '-';
    if (trackApplicationMeta) trackApplicationMeta.textContent = '-';
    if (trackApplicationStatusBadges) trackApplicationStatusBadges.innerHTML = '';
    if (trackApplicationSummary) trackApplicationSummary.innerHTML = '';
    if (trackApplicationTimeline) trackApplicationTimeline.innerHTML = '';
}

function closeDocumentGenerationModalFn() {
    currentDocumentGenerationSubmissionId = '';
    if (documentGenerationModal) documentGenerationModal.classList.remove('active');
    if (documentGenerationCustomerName) documentGenerationCustomerName.textContent = '-';
    if (documentGenerationMeta) documentGenerationMeta.textContent = '-';
    if (documentGenerationChecklist) documentGenerationChecklist.innerHTML = '';
}

function closeGeneratedDocumentsPreviewModalFn() {
    if (generatedDocumentsPreviewModal) generatedDocumentsPreviewModal.classList.remove('active');
    if (generatedDocumentsPreviewList) generatedDocumentsPreviewList.innerHTML = '';
    if (generatedDocumentsPreviewMeta) generatedDocumentsPreviewMeta.textContent = '';
    resetGeneratedDocumentPreviewItems();
}

function openTrackReportInputModalFn() {
    if (trackReportInputModal) trackReportInputModal.classList.add('active');
}

function closeTrackReportInputModalFn() {
    if (trackReportInputModal) trackReportInputModal.classList.remove('active');
}

function closeTrackReportPreviewModalFn() {
    if (trackReportPreviewModal) trackReportPreviewModal.classList.remove('active');
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
    if (tabId === 'generate-documents') {
        renderGenerateDocumentsTable();
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
        'generate-documents': 'Application Management - Generate Document',
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
function formatUserLastLoginCell(user = {}) {
    if (user.isOnline === true) {
        return '<span class="status-badge approved">Online</span>';
    }
    return formatDate(user.lastLogoutAt || user.lastSeenAt || user.lastLoginAt);
}

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
                <td>${formatUserLastLoginCell(user)}</td>
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

    const finallySubmitted = allSubmissions
        .filter(s => s.finalSubmitted === true || s.rsaSubmitted === true)
        .slice()
        .sort((a, b) => getStageTimestampMillis(getSubmissionFinalSubmissionEntryAt(b)) - getStageTimestampMillis(getSubmissionFinalSubmissionEntryAt(a)));

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
    }).slice().sort((a, b) => getStageTimestampMillis(getSubmissionCurrentStageEntryAt(b)) - getStageTimestampMillis(getSubmissionCurrentStageEntryAt(a)));

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
            await notifyUserPushEvent({
                currentUser: auth.currentUser,
                recipientUserId: String(selected.createdByUid || '').trim(),
                recipientEmail: String(selected.createdBy || '').trim(),
                eventType: 'agent_registration_approved',
                title: 'Agent Registration Approved',
                body: `${selected.fullName || 'Your agent'} has been approved and is now available for submissions.`,
                clickUrl: '/dashboard.html',
                meta: {
                    agentId,
                    agentName: selected.fullName || '',
                    approvedBy: currentAdmin?.email || '',
                    createdBy: selected.createdBy || '',
                    createdByUid: selected.createdByUid || ''
                }
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
                        <span class="detail-value">${user.isOnline === true ? 'Online' : formatDate(user.lastLogoutAt || user.lastSeenAt || user.lastLoginAt)}</span>
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
        return submission.rsaSubmittedAt
            || submission.finalSubmittedAt
            || submission.rsaAssignedAt
            || getSubmissionApprovalEntryAt(submission)
            || null;
    }
    if (stageKey === 'payment') {
        return submission.paidAt
            || submission.paymentAssignedAt
            || getSubmissionPaymentEntryAt(submission)
            || null;
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

function matchesTrackStatusFilter(submission = {}, filterValue = 'all') {
    const filter = String(filterValue || 'all').trim().toLowerCase();
    const status = String(submission.status || '').trim().toLowerCase();
    if (filter === 'all') return true;
    if (filter === 'pending') return ['pending', 'submitted', 'resubmitted'].includes(status);
    if (filter === 'approved') return ['approved', 'processing_to_pfa', 'sent_to_pfa', 'rsa_submitted', 'paid', 'cleared'].includes(status) || submission.finalSubmitted === true || submission.rsaSubmitted === true;
    if (filter === 'rejected') return ['rejected', 'rejected_by_reviewer', 'rejected_by_rsa'].includes(status);
    return status === filter;
}

function toTitleCaseWords(value) {
    const text = String(value ?? '').trim().toLowerCase();
    if (!text) return '';
    return text.replace(/\b([a-z])([a-z']*)/g, (_, first, rest) => `${first.toUpperCase()}${rest}`);
}

function formatLetterDate(date = new Date()) {
    const d = date instanceof Date ? date : new Date(date);
    const day = d.getDate();
    const month = d.toLocaleString('en-US', { month: 'long' });
    const year = d.getFullYear();
    const suffix = day % 10 === 1 && day !== 11 ? 'st'
        : day % 10 === 2 && day !== 12 ? 'nd'
            : day % 10 === 3 && day !== 13 ? 'rd'
                : 'th';
    return `${day}${suffix} ${month}, ${year}`;
}

function amountToWords(value) {
    const num = Math.round(Number(value || 0));
    if (!Number.isFinite(num) || num <= 0) return 'Zero naira only';
    const ones = ['', 'one', 'two', 'three', 'four', 'five', 'six', 'seven', 'eight', 'nine', 'ten', 'eleven', 'twelve', 'thirteen', 'fourteen', 'fifteen', 'sixteen', 'seventeen', 'eighteen', 'nineteen'];
    const tens = ['', '', 'twenty', 'thirty', 'forty', 'fifty', 'sixty', 'seventy', 'eighty', 'ninety'];
    const chunkToWords = (n) => {
        const hundred = Math.floor(n / 100);
        const rest = n % 100;
        const parts = [];
        if (hundred) parts.push(`${ones[hundred]} hundred`);
        if (rest) {
            if (hundred) parts.push('and');
            if (rest < 20) parts.push(ones[rest]);
            else {
                const ten = Math.floor(rest / 10);
                const one = rest % 10;
                parts.push(one ? `${tens[ten]}-${ones[one]}` : tens[ten]);
            }
        }
        return parts.join(' ');
    };
    const scales = [
        { value: 1_000_000_000, label: 'billion' },
        { value: 1_000_000, label: 'million' },
        { value: 1_000, label: 'thousand' }
    ];
    let remaining = num;
    const parts = [];
    scales.forEach((scale) => {
        if (remaining >= scale.value) {
            const chunk = Math.floor(remaining / scale.value);
            parts.push(`${chunkToWords(chunk)} ${scale.label}`);
            remaining %= scale.value;
        }
    });
    if (remaining) {
        if (parts.length && remaining < 100) parts.push('and');
        parts.push(chunkToWords(remaining));
    }
    const sentence = parts.join(' ').replace(/\s+/g, ' ').trim();
    return `${sentence.charAt(0).toUpperCase()}${sentence.slice(1)} naira only`;
}

function getSubmissionDetailValue(submission = {}, keys = [], fallback = '') {
    const details = submission?.customerDetails || {};
    for (const key of keys) {
        const value = details?.[key] ?? submission?.[key];
        if (value !== undefined && value !== null && String(value).trim() !== '') return String(value).trim();
    }
    return fallback;
}

function getPfaAddress(pfaName = '') {
    const normalized = String(pfaName || '').trim().toLowerCase();
    const configured = getConfiguredPfaAddress(pfaName);
    if (configured) return configured;
    if (normalized.includes('pal pension')) return 'Plot 289, Adeogun Street, Victoria Island, Lagos.';
    return 'Pension Fund Administrator Address.';
}

function getConfiguredPfaAddress(pfaName = '') {
    const sourceName = String(pfaName || '').trim();
    const addresses = adminSystemSettings?.pfaAddresses || {};
    if (!sourceName || !addresses || typeof addresses !== 'object') return '';
    if (addresses[sourceName]) {
        const entry = normalizeConfiguredPfaAddressEntry(addresses[sourceName]);
        return formatAddressLineParts([entry.address, entry.landmark, entry.state], { trailingPeriod: false });
    }
    const normalized = sourceName.toLowerCase();
    const match = Object.entries(addresses).find(([name]) => String(name || '').trim().toLowerCase() === normalized);
    if (!match) return '';
    const entry = normalizeConfiguredPfaAddressEntry(match[1]);
    return formatAddressLineParts([entry.address, entry.landmark, entry.state], { trailingPeriod: false });
}

function getConfiguredPfaAddressEntry(pfaName = '') {
    const sourceName = String(pfaName || '').trim();
    const addresses = adminSystemSettings?.pfaAddresses || {};
    if (!sourceName || !addresses || typeof addresses !== 'object') return null;
    const direct = addresses[sourceName];
    if (direct) return normalizeConfiguredPfaAddressEntry(direct);
    const normalized = sourceName.toLowerCase();
    const match = Object.entries(addresses).find(([name]) => String(name || '').trim().toLowerCase() === normalized);
    return match ? normalizeConfiguredPfaAddressEntry(match[1]) : null;
}

function normalizeConfiguredPfaAddressEntry(value = {}) {
    if (value && typeof value === 'object' && !Array.isArray(value)) {
        return {
            address: String(value.address || value.addressLine || '').trim(),
            landmark: String(value.landmark || '').trim(),
            state: String(value.state || '').trim()
        };
    }
    const bits = String(value || '').split(',').map((item) => item.trim()).filter(Boolean);
    return {
        address: bits[0] || String(value || '').trim(),
        landmark: bits[1] || '',
        state: bits.slice(2).join(', ')
    };
}

function getPfaAddressParts(pfaName = '', fallbackAddress = '') {
    const normalized = String(pfaName || '').trim().toLowerCase();
    const configured = getConfiguredPfaAddressEntry(pfaName);
    if (configured && (configured.address || configured.landmark || configured.state)) {
        return {
            addressLine: configured.address,
            landmark: configured.landmark,
            state: configured.state
        };
    }
    if (normalized.includes('pal pension')) {
        return {
            addressLine: 'Plot 289, Adeogun Street',
            landmark: 'Victoria Island',
            state: 'Lagos'
        };
    }

    const source = String(fallbackAddress || getPfaAddress(pfaName) || '').trim();
    const bits = source.split(',').map((item) => item.trim()).filter(Boolean);
    return {
        addressLine: bits[0] || source || 'Pension Fund Administrator Address',
        landmark: bits[1] || '',
        state: bits[2] || ''
    };
}

function buildPropertyDescription(documentData = {}) {
    const houseNumber = documentData.houseNumber || 'the allocated house';
    const houseType = documentData.houseType || 'Residential Property';
    const estateName = documentData.estateName || 'the Estate';
    const estateAddress = documentData.estateAddress || '';
    return `House ${houseNumber}, a ${houseType} at ${estateName}${estateAddress ? `, ${estateAddress}` : ''}`.replace(/\s+/g, ' ').trim();
}

function formatAddressLineParts(parts = [], { trailingPeriod = false, separator = ', ' } = {}) {
    const cleaned = parts
        .map((part) => String(part || '').trim().replace(/^[,.\s]+|[,.\s]+$/g, ''))
        .filter(Boolean);
    const text = cleaned.join(separator).replace(/\s+,/g, ',').replace(/,\s*,/g, ',').trim();
    if (!text) return '';
    return trailingPeriod ? `${text}.` : text;
}

function formatMultilinePfaAddress(r = {}, { includePfaName = false } = {}) {
    const lines = [];
    if (includePfaName && r.pfa) lines.push(String(r.pfa || '').trim().replace(/^[,.\s]+|[,.\s]+$/g, ''));
    const address = formatAddressLineParts([r.pfaAddress], { trailingPeriod: false });
    const location = formatAddressLineParts([r.landmark, r.state], { trailingPeriod: true });
    if (address) lines.push(address);
    if (location) lines.push(location);
    return lines.filter(Boolean).join('\n');
}

function addDays(dateValue, days = 0) {
    const base = dateValue instanceof Date ? new Date(dateValue.getTime()) : new Date(dateValue);
    if (Number.isNaN(base.getTime())) return new Date();
    base.setDate(base.getDate() + Number(days || 0));
    return base;
}

function getSubmissionDate(submission = {}) {
    const candidates = [
        submission.submittedAt,
        submission.uploadedAt,
        submission.createdAt,
        submission.customerDetails?.submittedAt,
        submission.customerDetails?.uploadedAt,
        submission.customerDetails?.createdAt
    ];
    for (const candidate of candidates) {
        const millis = getStageTimestampMillis(candidate);
        if (Number.isFinite(millis) && millis > 0) return new Date(millis);
        const parsed = new Date(candidate);
        if (!Number.isNaN(parsed.getTime())) return parsed;
    }
    return new Date();
}

function extractBirthYear(value = '') {
    const raw = String(value || '').trim();
    if (!raw) return 0;
    const directYear = raw.match(/\b(19|20)\d{2}\b/);
    if (directYear) return Number(directYear[0]);
    const parsed = new Date(raw);
    if (!Number.isNaN(parsed.getTime())) return parsed.getFullYear();
    return 0;
}

function getCustomerBirthYear(submission = {}) {
    const details = submission.customerDetails || {};
    const keys = ['birthYear', 'yearOfBirth', 'dateOfBirth', 'dob', 'birthDate'];
    for (const key of keys) {
        const year = extractBirthYear(details?.[key] ?? submission?.[key]);
        if (year) return year;
    }
    return 0;
}

function calculateRepaymentAmount(submission = {}, loanAmount = 0) {
    const birthYear = getCustomerBirthYear(submission);
    const currentYear = new Date().getFullYear();
    const age = birthYear ? currentYear - birthYear : 0;
    const remainingYears = birthYear ? Math.max(1, 60 - age) : 0;
    if (remainingYears > 0) return Math.round(Number(loanAmount || 0) / remainingYears);
    return parseMoney(getSubmissionDetailValue(submission, ['monthlyRepayment'], '208310'));
}

function buildLetterDocumentData(submission = {}) {
    const details = submission.customerDetails || {};
    const propertyValue = parseMoney(getSubmissionDetailValue(submission, ['propertyValue'], '0'));
    const loanAmount = parseMoney(getSubmissionDetailValue(submission, ['loanAmount'], '0'));
    const rsaBalance = parseMoney(getSubmissionDetailValue(submission, ['rsaBalance'], '0'));
    const stored25 = parseMoney(getSubmissionDetailValue(submission, ['rsa25Percent', 'rsa25'], '0'));
    const equityContribution = stored25 || roundDownToNearestThousand(rsaBalance * 0.25);
    const pfaName = getSubmissionDetailValue(submission, ['pfa', 'pfaName'], 'Pension Fund Administrator');
    const houseType = getSubmissionDetailValue(submission, ['propertyType', 'houseType'], 'Residential Property');
    const houseNumber = getSubmissionDetailValue(submission, ['houseNumber'], 'N/A');
    const address = getSubmissionDetailValue(submission, ['address'], '');
    const estateName = getSubmissionDetailValue(submission, ['estateName'], 'Pacesetter Gardens Estate');
    const estateAddress = getSubmissionDetailValue(submission, ['estateAddress'], 'Adegbayi Area, Off Ibadan-Ife Expressway, Ajoda, Ibadan, Oyo State');
    const repaymentScheduleAmount = calculateRepaymentAmount(submission, loanAmount);
    const tenorYears = getSubmissionDetailValue(submission, ['tenor'], '5');
    const interestRate = getSubmissionDetailValue(submission, ['interestRate'], '6% per annum subject to review in line with changes in money market rate.');
    const facilityFee = parseMoney(getSubmissionDetailValue(submission, ['facilityFee'], '4000'));
    const bankAccountNumber = getSubmissionDetailValue(submission, ['accountNo', 'bankAccountNumber'], '5980207331');
    const currentDate = formatLetterDate(new Date());
    const postApprovalDate = formatLetterDate(addDays(getSubmissionDate(submission), -3));
    const customerName = toTitleCaseWords(submission.customerName || getSubmissionDetailValue(submission, ['name'], 'Customer')).toUpperCase();
    const pfaAddressParts = getPfaAddressParts(
        pfaName,
        getSubmissionDetailValue(submission, ['pfaAddress', 'pfa_address'], '')
    );
    const propertyDescription = buildPropertyDescription({
        houseNumber,
        houseType,
        estateName,
        estateAddress
    });
    return {
        currentDate,
        postApprovalDate,
        customerName,
        customerAddress: address || 'Customer Address',
        propertyValue,
        propertyValueText: formatCurrency(propertyValue),
        propertyValueWords: amountToWords(propertyValue),
        loanAmount,
        loanAmountText: formatCurrency(loanAmount),
        loanAmountWords: amountToWords(loanAmount),
        equityContribution,
        equityContributionText: formatCurrency(equityContribution),
        equityContributionWords: amountToWords(equityContribution),
        houseNumber,
        houseType: toTitleCaseWords(houseType),
        estateName: toTitleCaseWords(estateName),
        estateAddress,
        propertyDescription,
        pfaName: String(pfaName || 'Pension Fund Administrator').trim(),
        pfaBaseName: String(pfaName || 'Pension Fund Administrator').trim().replace(/\s+limited$/i, '').trim(),
        pfaAddress: getPfaAddress(pfaName),
        pfaAddressLine: pfaAddressParts.addressLine,
        pfaLandmark: getSubmissionDetailValue(submission, ['landmark'], pfaAddressParts.landmark),
        pfaState: getSubmissionDetailValue(submission, ['state'], pfaAddressParts.state),
        rsaPin: getSubmissionDetailValue(submission, ['penNo', 'rsaPin', 'pin'], 'PEN000000000000'),
        bankName: 'Cooperative Mortgage Bank Limited',
        bankAccountNumber,
        propertyValueAmount: propertyValue,
        loanAmountAmount: loanAmount,
        equityContributionAmount: equityContribution,
        monthlyRepaymentAmount: repaymentScheduleAmount,
        facilityFeeAmount: facilityFee,
        managementFeeAmount: loanAmount * 0.01,
        tenorYears,
        repaymentScheduleText: `${formatCurrency(repaymentScheduleAmount)} monthly principal and interest repayment`,
        facilityFeeText: `${formatCurrency(facilityFee)} monthly payment`,
        sourceOfRepayment: getSubmissionDetailValue(submission, ['sourceOfRepayment'], 'Monthly Salary'),
        tenorText: `${tenorYears} Years`,
        interestRate,
        managementFeeText: `1% of the facility amount (${formatCurrency(loanAmount * 0.01)}) upfront payment upon booking`,
        commencementDateText: 'This facility shall commence upon drawdown or on the date of disbursement notwithstanding the date on the offer letter or date of execution.',
        availabilityText: 'Upon satisfactory compliance with all conditions precedent to drawdown but not later than 14 days from date of offer letter.',
        readinessBankLabel: 'Cooperative Mortgage Bank Ltd.',
        titleDescription: `Deed of Sublease over property situated at ${toTitleCaseWords(estateName)}, ${estateAddress}.`,
        uploadedByName: getDisplayNameByEmail(submission.uploadedBy || ''),
        rsaOfficerName: submission.assignedToRSA ? getDisplayNameByEmail(submission.assignedToRSA) : 'RSA Officer',
        paymentOfficerName: submission.assignedToPayment ? getDisplayNameByEmail(submission.assignedToPayment) : 'Payment Officer',
        stampLabel: 'COOPERATIVE MORTGAGE BANK LTD',
        submissionId: submission.id || ''
    };
}

function buildShortLetterDate(dateValue = new Date()) {
    const date = dateValue instanceof Date ? dateValue : new Date(dateValue);
    if (Number.isNaN(date.getTime())) return '';
    return date.toLocaleDateString('en-GB', {
        day: 'numeric',
        month: 'long',
        year: 'numeric'
    });
}

function toPdfSafeText(value = '') {
    return String(value || '')
        .replace(/₦/g, '#')
        .replace(/[“”]/g, '"')
        .replace(/[‘’]/g, "'")
        .replace(/–/g, '-');
}

function formatAmountForLetter(amount) {
    const numeric = parseMoney(amount);
    return `#${numeric.toLocaleString('en-NG', { minimumFractionDigits: 0, maximumFractionDigits: 0 })} (${amountToWords(numeric)})`;
}

function formatAmountWithTwoDecimals(amount) {
    const numeric = Number(amount || 0);
    return `#${numeric.toLocaleString('en-NG', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`;
}

async function ensurePdfFontAssets() {
    if (pdfFontAssetCache) return pdfFontAssetCache;
    if (!window.PDFLib?.PDFDocument) {
        throw new Error('PDF engine is unavailable. Please refresh and try again.');
    }
    if (!window.fontkit) {
        throw new Error('PDF font engine is unavailable. Please refresh and try again.');
    }

    const fontFiles = {
        arial: 'assets/fonts/arial.ttf',
        arialBold: 'assets/fonts/arialbd.ttf',
        dejavu: 'assets/fonts/DejaVuSans.ttf',
        dejavuBold: 'assets/fonts/DejaVuSans-Bold.ttf',
        trebuchet: 'assets/fonts/trebuc.ttf',
        trebuchetBold: 'assets/fonts/trebucbd.ttf'
    };

    const entries = await Promise.all(Object.entries(fontFiles).map(async ([key, path]) => {
        const response = await fetch(path, { cache: 'force-cache' });
        if (!response.ok) throw new Error(`Unable to load font asset: ${path}`);
        return [key, await response.arrayBuffer()];
    }));

    pdfFontAssetCache = Object.fromEntries(entries);
    return pdfFontAssetCache;
}

function getPdfDocumentFontProfile(documentType, field = {}) {
    const weight = String(field.fontWeight || '').toLowerCase() === 'regular' ? 'regular' : 'bold';
    return weight === 'regular'
        ? { family: 'dejavu', assetKey: 'dejavu' }
        : { family: 'dejavuBold', assetKey: 'dejavuBold' };
}

function getPdfFontProfileForTemplateLine(line = {}) {
    const fontName = String(line.font || '').toLowerCase();
    if (fontName.includes('trebuchet')) {
        return fontName.includes('bold')
            ? { family: 'trebuchetBold', assetKey: 'trebuchetBold' }
            : { family: 'trebuchet', assetKey: 'trebuchet' };
    }
    if (fontName.includes('bold') || Number(line.flags || 0) >= 16) {
        return { family: 'dejavuBold', assetKey: 'dejavuBold' };
    }
    return { family: 'dejavu', assetKey: 'dejavu' };
}

function buildDocumentContent(documentType, data = {}) {
    const r = {
        name: data.customerName,
        address: data.customerAddress,
        pfa: data.pfaName,
        pfaBase: data.pfaBaseName || data.pfaName,
        pfaAddress: data.pfaAddressLine,
        landmark: data.pfaLandmark,
        state: data.pfaState,
        houseNumber: data.houseNumber,
        houseType: data.houseType,
        propertyValue: data.propertyValueAmount,
        loanAmount: data.loanAmountAmount,
        equityContribution: data.equityContributionAmount,
        repayment: data.monthlyRepaymentAmount,
        facilityFee: data.facilityFeeAmount,
        managementFee: data.managementFeeAmount,
        tenor: data.tenorYears,
        accountNumber: data.bankAccountNumber,
        pen: data.rsaPin,
        ovDate: data.currentDate,
        aDate: data.currentDate,
        date: data.currentDate
    };

    const offerBoldWords = [
        r.address,
        r.ovDate,
        r.pfa,
        r.pfaAddress,
        r.state,
        r.landmark,
        r.name,
        formatAmountWithTwoDecimals(r.managementFee),
        r.houseNumber,
        r.houseType,
        formatAmountForLetter(r.equityContribution),
        formatAmountForLetter(r.loanAmount),
        formatAmountForLetter(r.propertyValue),
        formatAmountForLetter(r.repayment),
        formatAmountForLetter(r.facilityFee),
        'PROVISIONAL OFFER OF MORTGAGE FACILITY',
        'COOPERATIVE MORTGAGE BANK LIMITED',
        'CONDITIONS PRECEDENT TO DRAWDOWN:',
        '2.5%',
        '15th day of the succeeding month.',
        "invoking the Bank's right of sale on the mortgaged property",
        'Title over Property:',
        'Cost of Registration of Title:',
        '3 (three) consecutive months.',
        'Note: All the conditions must be complied with or else this offer shall be withdrawn.',
        'OTHER CONDITIONS',
        'DEFECT LIABILITY PERIOD',
        'EVENTS OF DEFAULT',
        'POLICY CLAUSES',
        'Cooperative Mortgage Bank Limited',
        'Security/Comfort:'
    ].filter(Boolean);

    const generators = {
        offer_letter: () => ({
            pages: [
                {
                    paragraphs: [
                        `${r.ovDate}`,
                        `${r.name}`,
                        `${r.address}`,
                        'Dear Sir/Ma,',
                        'PROVISIONAL OFFER OF MORTGAGE FACILITY',
                        'We are pleased to inform you that Cooperative Mortgage Bank Limited ("the Bank") has approved a Mortgage Facility in your favor under the following terms and conditions:'
                    ],
                    table: [
                        ['Lender:', 'COOPERATIVE MORTGAGE BANK LIMITED'],
                        ['Borrower:', r.name],
                        ['Address:', r.address],
                        ['Facility Type:', 'Mortgage Facility'],
                        ['Property Value:', formatAmountForLetter(r.propertyValue)],
                        ['Commencement Date:', 'This facility shall commence upon drawdown or on the date of disbursement notwithstanding the date on the offer letter or date of execution'],
                        ['Loan Amount:', formatAmountForLetter(r.loanAmount)],
                        ['Availability:', 'Upon satisfactory compliance with all conditions precedent to drawdown but not later than 14 days from date of offer letter.'],
                        ['Purpose:', `To part finance the purchase House ${r.houseNumber} of a ${r.houseType} at Pacesetter Gardens Estate, Adegbayi Area, Off Ibadan Ife Expressway Ajoda, Ibadan Oyo State`],
                        ['Equity Contribution:', formatAmountForLetter(r.equityContribution)]
                    ],
                    offerTermsTable: 'primary',
                    boldWords: offerBoldWords
                },
                {
                    tablePosition: 'before',
                    table: [
                        ['Repayment Schedule:', `${formatAmountForLetter(r.repayment)} Monthly principal and interest repayment`],
                        ['Facility Fee:', `${formatAmountForLetter(r.facilityFee)} Monthly payment`],
                        ['Source of Repayment:', 'Monthly Salary'],
                        ['Tenor:', `${r.tenor} Years`],
                        ['Interest Rate:', '6% per annum subject to review in line with changes in Money market rate.'],
                        ['Management Fee:', '1% of the facility amount (Upfront payment upon booking)'],
                        ['Prepayment:', 'Voluntary prepayment is allowed during the term of the facility without penal charge.'],
                        ['Security/Comfort:', `Deed of Sublease on House ${r.houseNumber} of a ${r.houseType} at Pacesetter Gardens Estate, Adegbayi Area, Off Ibadan-Ife Expressway Ajoda, Ibadan Oyo State. Comprehensive Fire and Other Perils Insurance Policy with Cooperative Mortgage Bank Ltd noted as the first loss payee. Mortgage Protection Policy on behalf of ${r.name}.`],
                        ['Title over Property:', 'Deed of Sublease over property situated at Pacesetter Gardens Estate, Adegbayi Area, off Ibadan-Ife Expressway Ajoda Ibadan Oyo State.'],
                        ['Cost of Registration of Title:', 'The cost and expenses on preparation and perfection of legal mortgage shall be borne by the Borrower. This is inclusive of other incidental expenses necessary for perfection.']
                    ],
                    offerTermsTable: 'continuation',
                    paragraphs: [
                        'CONDITIONS PRECEDENT TO DRAWDOWN:',
                        '1. Duly executed offer letter accepting the terms and conditions of the facility unconditionally.',
                        '2. Upfront payment of fees:',
                        `a. Management fees (1% of loan amount) ${formatAmountWithTwoDecimals(r.managementFee)}.`,
                        'b. Preparation of Deed of Sublease #25,000.00.',
                        'c. Preparation of Deed of Legal Mortgage #25,000.00.',
                        '3. All documents necessary to perfect Legal Mortgage which includes:',
                        "   - Executed copy of Deed of Legal Mortgage between the Bank and the Borrower.",
                        "   - Original Title documents in Bank's custody.",
                        `4. Letter of authority granting CMBank or any of its appointed Professional Valuer unrestricted access to periodically re-evaluate the property financed which is House ${r.houseNumber}, a ${r.houseType} at Pacesetter Gardens Estate, Adegbayi Area, Off Ibadan-Ife Expressway Ajoda, Ibadan, Oyo State.`,
                        '5. Acceptance of the following conditions as it relates to the facility:',
                        'a. In the event of default beyond the due date as stated above, the payments so in arrears shall henceforth bear interest at a daily default rate of 1% computed from the date same became payable to the 15th day of the succeeding month.',
                        'b. A monthly late payment charge of 2.5% shall also accrue on any unpaid amount after the 15th day of the succeeding month.',
                        "c. Such default shall warrant invoking the Bank's right of sale on the mortgaged property, including but not limited to the occurrence of any of the following events:",
                        'd. Non-repayment of both principal and interest on due date for 3 (three) consecutive months.',
                        'e. Default in payment of yearly insurance premium and other fees as stipulated in the offer.',
                        'f. Part payment of monthly repayment due for any current month.',
                        `6. That the security for the facility is the Deed of Sublease on House ${r.houseNumber} of the ${r.houseType} at Pacesetter Gardens Estate, Adegbayi Area, Off Ibadan-Ife Expressway, Ajoda, Ibadan, Oyo State.`,
                        `7. That CMBank shall arrange at the expense of ${r.name} a fire insurance cover for the property and a mortgage protection policy for the borrower.`
                    ],
                    boldWords: offerBoldWords
                },
                {
                    paragraphs: [
                        'The premium payable on the fire insurance and mortgage protection policy is on per annual basis.',
                        `8. The Bank shall sell the collateral/property (House ${r.houseNumber}, a ${r.houseType} at Pacesetter Gardens Estate, Adegbayi Area, Off Ibadan-Ife Expressway, Ajoda, Ibadan) if the Borrower is in default of any of the terms and conditions therein stated.`,
                        `9. ${r.name}, or any person(s) designated by the borrower shall be responsible for all cost incurred in the recovery process in case of default.`,
                        '10. This offer may be withdrawn:',
                        'a. If it is not accepted within two weeks of the receipt of the offer letter.',
                        'b. If any irregularity is discovered anytime.',
                        'Note: All the conditions must be complied with or else this offer shall be withdrawn.',
                        'OTHER CONDITIONS',
                        `1. The Bank shall, without any recourse to ${r.name}, debit the following charges to his account annually throughout the tenor of the facility except where the evidence of payment is provided:`,
                        `a. Facility Maintenance Fee: ${formatAmountWithTwoDecimals(r.facilityFee)} monthly.`,
                        'b. Mortgage Protection premium per annum will be communicated to the customer in future.',
                        'c. Ground rent per annum will be communicated to the customer in future.',
                        '2. No waiver of interest or accumulated interest on this facility will be entertained.',
                        "3. The Bank reserves the right to debit the Borrower's Account with perfection costs without recourse to the Borrower.",
                        '4. Notwithstanding any repayment condition stated herein, this facility shall become repayable upon the occurrence of any of the following events:',
                        'a. If there should, in the opinion of the Bank, be a material adverse change in the financial condition of the Borrower.',
                        'b. The Bank reserves the right to cancel and/or reduce the facility in line with its ability to accommodate it within its legal lending limits and/or policy or portfolio constraints.',
                        '',
                        'DEFECT LIABILITY PERIOD',
                        "1.0. The following provisions set out the developer's responsibilities in the event of damage or destruction to permanent fixtures within the allocated housing unit.",
                        '1.1. Where there is a defect as to the installation and/or functionality of electrical fittings/lights, internal doors, sanitary wares (inclusive of the water closet, wash-hand basin, floor, drain, kitchen sink etc.), painting and wardrobes, the developer shall be liable for repairs no later than three months after the handover date.',
                        '1.2. Where there is a defect as to the installation and/or functionality of the septic tank, other external works (inclusive of inspection chambers, compound landscaping interlocks/concreting), the ceiling, roof and external doors, the developer shall be liable for repairs no later than six months after the handover date.'
                    ],
                    boldWords: offerBoldWords
                },
                {
                    paragraphs: [
                        "1.3. Where such damage or destruction results in a need for the replacement of permanent fixtures, the developer shall not replace such fixture more than twice during the defect liability period.",
                        '1.4. Where such damage or destruction occurs towards the end of the defect liability period, the Developer shall be notified through Cooperative Mortgage Bank no later than 5 working days before the lapse of the defect liability period.',
                        '1.5. In the event that there is noncompliance with clause 1.4 above, the developer shall no longer be required to effect any repair on the allocated housing unit, and any claim of whatever nature shall be deemed to have been waived and shall become absolutely barred.',
                        '',
                        'EVENTS OF DEFAULT',
                        "1. In the event of default to meet the Borrower's obligations to the Bank in respect of the Facility granted, the Bank shall be at liberty to take immediate physical possession (where possible) of the property with a view to selling same towards liquidation of the Borrower's indebtedness to the Bank. Such recourse to the property shall however be without prejudice to other rights which the Bank may have against the Borrower.",
                        "2. Without prejudice to CMBank's right to demand repayment of outstanding amounts of this facility at any time, the occurrence of any of the following events shall cause all outstanding amounts under the facility to become immediately payable if:",
                        '2.1. The Borrower commits any breach or default under the terms of this facility or of any other credit facilities granted to the Borrower by CMBank or any other creditors.',
                        "2.2. In the opinion of CMBank, there are significant material adverse changes in the Borrower's income or financial conditions.",
                        '2.3. The Bank is compelled by any Central Bank of Nigeria rules and regulations or directive to call in the loan.',
                        '',
                        'POLICY CLAUSES',
                        "1. It is the Bank's policy to review facilities from time to time in the light of changing market conditions.",
                        '2. The Bank may decide to refinance the loan and all right of the Bank under this loan may be transferred to the underwriter with notice to the Borrower. The Borrower shall honor its obligations under this loan as it would with the Bank.',
                        '3. Any dispute, question or difference arising in connection with this agreement shall be referred to arbitration under the Arbitration and Conciliation Act Cap 19, Laws of the Federation of Nigeria 1990. The arbitration shall be conducted by a single arbitrator to be appointed by the Chairman of the Nigerian Branch of the Chartered Institute of Arbitrator.'
                    ],
                    boldWords: offerBoldWords
                }
            ]
        }),
        allocation_letter: () => ({
            pages: [
                {
                    paragraphs: [
                        `${r.aDate}`,
                        `${r.name}`,
                        `${r.address}`,
                        'Dear Sir/Ma,',
                        `LETTER OF ALLOCATION FOR A ${r.houseType} AT PACESETTER GARDEN ESTATE ADEGBAYI AREA, OFF IBADAN-IFE EXPRESSWAY IBADAN`,
                        `With reference to your application for a house at PACESETTER GARDEN ESTATE, Adegbayi Area, off Ibadan-Ife Expressway Ibadan. We are pleased to inform you that you have been formally allocated HOUSE ${r.houseNumber}, a unit of ${r.houseType} in the Estate which is allocated to you for residential purpose on a leasehold basis with effect from ${r.aDate}.`,
                        '1. This allocation is subject to the following terms and condition:',
                        'a. Not to use any or whole of the premises allocated to you for any other purpose(s) except residential and the occupying of the same either as an office, shop, light industry or any other use not in accordance with the user clause of this allocation, would automatically attract forfeiture of the allocation;',
                        'b. Not to construct servant quarters, perimeter fence or any other structure around or within the premises without prior approval of the Company being sought and obtained in writing, the violation of which shall attract demolition, penalty/or both;',
                        'c. Not to alter or cause to be altered the external design or structure of the house/flat. Internal alteration may however, be effected only after prior approval of the Company has been sought and obtained in writing;',
                        'd. To comply with the rules and regulations that the Company may make from time to time, as they affect ownership, possession, occupation and use of housing unit(s);',
                        'e. Not to mortgage, sublet, assign, transfer or part with possession of the housing unit or any part thereof without the consent of the Company being sought and obtained in writing, such consent shall not be unreasonably withheld;',
                        'f. To enter into a Legal mortgage agreement with Cooperative Mortgage Bank Limited and pay such fees for Title Document as may be prescribed;',
                        'g. Payment of ground rent as it will be communicated to you at a later date;',
                        'h. To pay levies, rates or other service charges as may from time to time be levied in respect of the property for the maintenance of the estate by the Company in charge of such maintenance;',
                        'i. Maintain the property hereby allocated as well as its environment in good sanitary and tenantable condition to the satisfaction of the Company;'
                    ],
                    boldWords: [
                        r.aDate, r.address, r.name, r.houseNumber, r.houseType, formatAmountForLetter(r.propertyValue),
                        'LETTER OF ALLOCATION FOR A', 'AT PACESETTER GARDEN ESTATE ADEGBAYI AREA, OFF IBADAN-IFE EXPRESSWAY IBADAN', 'HOUSE'
                    ]
                },
                {
                    paragraphs: [
                        'j. Refrain from erecting, exhibiting or permitting the erection or exhibition of any bills, or notice boards on the premises;',
                        'k. Identify with the Resident Association established within the estate charged with the maintenance and management of the estate as it affects commercial, cultural facilities and/or recreation;',
                        'l. Not to rear animals or pets in the apartment and its premises without the written consent of the Company. Such consent however not to be unreasonably withheld.',
                        `2. This letter represents an offer by Coop Property Development Company Limited in respect of the sale of the said dwelling house to you at a price of ${formatAmountForLetter(r.propertyValue)}.`,
                        '',
                        'Please accept our congratulations.',
                        'Yours faithfully,',
                        'For: Coop Property Development Company Limited'
                    ],
                    boldWords: [
                        r.aDate, r.address, r.name, r.houseNumber, r.houseType, formatAmountForLetter(r.propertyValue),
                        'LETTER OF ALLOCATION FOR A', 'AT PACESETTER GARDEN ESTATE ADEGBAYI AREA, OFF IBADAN-IFE EXPRESSWAY IBADAN', 'HOUSE'
                    ]
                }
            ]
        }),
        indemnity_letter: () => ({
            pages: [
                {
                    paragraphs: [
                        formatMultilinePfaAddress(r, { includePfaName: true }),
                        '',
                        'LETTER OF INDEMNITY IN RESPECT OF RETIREMENT SAVINGS ACCOUNT EQUITY CONTRIBUTION',
                        `THIS INDEMNITY is issued by Cooperative Mortgage Bank Ltd. of 11B University Crescent, UI-Secretariat road, Old Bodija, Ibadan (hereinafter called "Cooperative Mortgage Bank" which expression shall where the context so admits include its successors-in-title and assigns), to ${r.pfaBase} Limited having its Head Office at ${formatAddressLineParts([r.pfaAddress, r.landmark, r.state], { trailingPeriod: true })} (hereinafter called "${r.pfaBase} (PFA)" which expression shall where the context so admits include its successors-in-title and assigns).`,
                        'WHEREAS:',
                        `1. ${r.pfaBase} is a Pension Fund Administrator licensed by the National Pension Commission ("the Commission") to manage individual Retirement Savings Accounts (RSA) in accordance with the provisions of the Pension Reform Act, 2014 ("the Act").`,
                        '2. Pursuant to the Act, the Commission has issued Guidelines on Accessing RSA Balance towards Payment of Equity Contributions for Residential Mortgage ("Guidelines"), which permit RSA Account Holders to access a portion of their RSA balance as equity contribution towards acquiring a residential property from a licensed financial institution.',
                        '3. Cooperative Mortgage Bank is duly licensed by the Central Bank of Nigeria (CBN) and meets all the eligibility criteria for the provision of mortgage lending services to RSA holders under the Guidelines issued by the Commission.',
                        `4. Cooperative Mortgage Bank has received applications from several RSA holders who maintain RSAs with ${r.pfaBase} and, in line with the Guidelines, the indemnifier has agreed to finance the purchase of a residential property on behalf of the RSA holders, subject to the execution of relevant mortgage contracts.`,
                        `5. ${r.name}, the RSA holder, has requested and authorized that a portion of his/her RSA balance should be utilized as equity contribution to facilitate his/her mortgage and has applied to ${r.pfaBase} to release to the indemnifier a portion not more than 25% of his/her RSA balance as equity contribution towards the acquisition of a residential mortgage.`,
                        `6. Cooperative Mortgage Bank has conducted all requisite due diligence to confirm the authenticity of the RSA holder's intended acquisition of a residential property and undertakes to ${r.pfaBase} that the portion of the RSA holder's fund which shall be released to the Cooperative Mortgage Bank on the RSA holder's request shall be applied solely as equity contribution towards the acquisition of a residential mortgage, under the mortgage.`
                    ],
                    boldWords: [
                        r.pen, r.pfaBase, r.pfaAddress, r.state, r.landmark, r.name,
                        'THIS INDEMNITY', 'Cooperative Mortgage Bank Ltd', 'Cooperative Mortgage Bank',
                        'WHEREAS:', 'NOW THEREFORE', 'We, Cooperative Mortgage Bank Ltd',
                        'Furthermore, Cooperative Mortgage Bank', 'PIN', 'RSA HOLDER',
                        'LETTER OF INDEMNITY IN RESPECT OF RETIREMENT SAVINGS ACCOUNT EQUITY CONTRIBUTION', '(PFA)'
                    ]
                },
                {
                    paragraphs: [
                        'Agreement between the Indemnifier and the RSA holder shall not, under any condition, be released to the RSA holder or any third party for any other purpose.',
                        `7. ${r.pfaBase} agrees to release the relevant portion of the RSA holder's fund to the Cooperative Mortgage Bank and has requested that the Indemnifier provides this indemnity in favour of ${r.pfaBase} in the manner hereinafter stated.`,
                        `NOW THEREFORE, in consideration of ${r.pfaBase} effecting the transfer of the portion of the RSA holder's balance to the Indemnifier as equity contribution by the RSA holder whose RSA PIN herein appears, towards acquisition of residential property in accordance with the Guidelines, We, Cooperative Mortgage Bank Ltd hereby irrevocably and unconditionally covenant that we shall at all times hereafter, indemnify ${r.pfaBase} and keep ${r.pfaBase} fully indemnified against all claims, demands, liabilities, actions, damages, penalties and legal proceedings (including any cost of litigation) which may be incurred by ${r.pfaBase} in the event that the fund so released by ${r.pfaBase} is not utilized as equity contribution towards the purchase of a residential property on behalf of the RSA holder.`,
                        `Furthermore, Cooperative Mortgage Bank undertakes to pay ${r.pfaBase} on demand without cavil or contention, all payments, liabilities, damages and expenses (including but not limited to legal fees) incurred by ${r.pfaBase} from acceding to the RSA holder's request to release funds from his/her RSA, in the event that the fund so released by ${r.pfaBase} is not utilized as equity contribution towards the purchase of a residential property on behalf of the RSA holder.`,
                        `This Indemnity shall be a continuing security and shall be enforceable from the date of disbursement of the portion of the RSA holder's balance to Cooperative Mortgage Bank and shall inure to the benefit of ${r.pfaBase} up until the utilization of the funds so released, as equity contribution towards the acquisition of the residential property in favour of the RSA holder by ${r.pfaBase}.`,
                        `This Indemnity shall not be enforceable against Cooperative Mortgage Bank for any claims, demands, liabilities, losses, actions, damages, penalties, and legal proceedings (including any cost of litigation), arising from any dispute or breach howsoever, in the contractual relationship between ${r.pfaBase} and the RSA holder.`,
                        `The indemnification provided under this Indemnity by Cooperative Mortgage Bank shall be completely discharged upon the utilization of the portion of the RSA holder's balance towards the purchase of a residential property in favour of the RSA holder by Cooperative Mortgage Bank.`,
                        `RSA HOLDER: ${r.name}`,
                        `PIN: ${r.pen}`,
                        'This indemnity shall be governed by and construed in accordance with the extant laws of the Federal Republic of Nigeria.',
                        'THE COMMON SEAL of the within named INDEMNIFIER Cooperative Mortgage Bank was hereunto affixed in the presence of:'
                    ],
                    boldWords: [
                        r.pen, r.pfaBase, r.pfaAddress, r.state, r.landmark, r.name,
                        'THIS INDEMNITY', 'Cooperative Mortgage Bank Ltd', 'Cooperative Mortgage Bank',
                        'WHEREAS:', 'NOW THEREFORE', 'We, Cooperative Mortgage Bank Ltd',
                        'Furthermore, Cooperative Mortgage Bank', 'PIN', 'RSA HOLDER',
                        'LETTER OF INDEMNITY IN RESPECT OF RETIREMENT SAVINGS ACCOUNT EQUITY CONTRIBUTION', '(PFA)'
                    ]
                }
            ]
        }),
        verification_letter: () => ({
            pages: [{
                paragraphs: [
                    `${r.ovDate}`,
                    '',
                    formatMultilinePfaAddress(r, { includePfaName: true }),
                    '',
                    'VERIFICATION OF PROPERTY',
                    `This is to confirm the authenticity of the property offer allocated to ${r.name}, which is House ${r.houseNumber}, a ${r.houseType} situated at Pacesetter Gardens Estate, Adegbayi Area, off Ibadan-Ife Express way, Ibadan valued at ${formatAmountForLetter(r.propertyValue)} and available for sale. The allocation letter and offer of property is hereby confirmed as genuine and valid.`,
                    '',
                    'Thank you.',
                    'Yours Faithfully,'
                ],
                boldWords: ['VERIFICATION OF PROPERTY', r.ovDate, r.pfa, r.name, formatAmountForLetter(r.propertyValue)]
            }]
        }),
        availability_letter: () => ({
            pages: [{
                paragraphs: [
                    `${r.date}`,
                    '',
                    formatMultilinePfaAddress(r, { includePfaName: true }),
                    '',
                    'Dear Sir/Ma,',
                    '',
                    'CONFIRMATION OF AVAILABILITY OF PROPERTY',
                    '',
                    `This is to confirm the authenticity of the property allocated to ${r.name}, which is a ${r.houseType}, House number ${r.houseNumber} situated at Pacesetter Gardens Estate, Adegbayi Area, off Ibadan- Ife Express way, Ibadan, valued at ${formatAmountForLetter(r.propertyValue)}.`,
                    '',
                    'Thank you.',
                    'Yours Faithfully,'
                ],
                boldWords: [
                    r.pfa, r.date, r.houseType, r.houseNumber, 'CONFIRMATION OF AVAILABILITY OF PROPERTY',
                    r.name, formatAmountForLetter(r.propertyValue), r.pfaAddress, r.landmark, r.state
                ]
            }]
        }),
        readiness_letter: () => ({
            pages: [{
                paragraphs: [
                    `${r.date}`,
                    '',
                    formatMultilinePfaAddress(r, { includePfaName: true }),
                    '',
                    'Dear Sir/Ma,',
                    'READINESS TO RECEIVE DISBURSEMENT AS EQUITY FOR MORTGAGE',
                    '',
                    'We hereby assert the availability of the aforementioned property and our readiness to receive disbursement on behalf of the above-named customer into the account details below:',
                    '',
                    `Name: ${r.name}\nAccount Number: ${r.accountNumber}\nEquity Amount: ${formatAmountForLetter(r.equityContribution)}\nBank: Cooperative Mortgage Bank Ltd.`,
                    '',
                    `We also affirm our readiness to grant a Mortgage Loan of ${formatAmountForLetter(r.loanAmount)} only to ${r.name} for the aforementioned property.`,
                    '',
                    'Thank you.'
                ],
                boldWords: [
                    r.date, r.name, r.pfa, r.accountNumber, formatAmountForLetter(r.equityContribution), 'Cooperative Mortgage Bank Ltd.',
                    r.pfaAddress, r.state, r.landmark, formatAmountForLetter(r.loanAmount),
                    'READINESS TO RECEIVE DISBURSEMENT AS EQUITY FOR MORTGAGE',
                    'Name:', 'Account Number:', 'Equity Amount:', 'Bank:'
                ]
            }]
        }),
        title_letter: () => ({
            pages: [{
                paragraphs: [
                    `${r.date}`,
                    '',
                    formatMultilinePfaAddress(r, { includePfaName: true }),
                    '',
                    'Dear Sir/Ma.',
                    '',
                    `CONFIRMATION OF TITLE DOCUMENT OF PROPERTY KNOWN AS ${String(r.houseType || '').toUpperCase()} LOCATED AT PACESETTER GARDENS ESTATE, AJODA IBADAN IN FAVOUR OF ${r.name} PIN:${r.pen}`,
                    '',
                    `Following the application and subsequent approval of the National Pension Commission (PenCom) for the release of 25% RSA balance as equity for the mortgage of ${r.name}. We hereby confirm that a search was conducted to confirm the authenticity of the property document and it was confirmed that the property is free from all encumbrances.`,
                    'Please find enclosed the search report on the property.',
                    '',
                    'Thank you.',
                    'Yours Faithfully,'
                ],
                boldWords: [
                    String(r.houseType || '').toUpperCase(), r.pfa, r.name, r.date, r.pen, r.pfaAddress, r.state, r.landmark,
                    'CONFIRMATION OF TITLE DOCUMENT OF PROPERTY KNOWN AS',
                    'LOCATED AT PACESETTER GARDENS ESTATE, AJODA IBADAN IN FAVOUR OF',
                    'PIN:'
                ]
            }]
        })
    };

    return generators[documentType]?.() || { pages: [] };
}

function ensureDocxZipLibrary() {
    if (window.PizZip) return true;
    showNotification('Word document engine is unavailable. Please refresh and try again.', 'error');
    return false;
}

function escapeXml(value = '') {
    return String(value ?? '')
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&apos;');
}

function escapeAttribute(value = '') {
    return escapeXml(value).replace(/\n/g, ' ');
}

function normalizeWordText(value = '') {
    return String(value ?? '')
        .replace(/[“”]/g, '"')
        .replace(/[‘’]/g, "'")
        .replace(/–/g, '-')
        .replace(/\u00a0/g, ' ');
}

function splitTextIntoWordRuns(text = '', boldWords = []) {
    const source = normalizeWordText(text);
    if (!source) return [];
    const phrases = normalizeBoldPhraseList(boldWords)
        .filter((phrase) => phrase.length >= 2)
        .map((phrase) => normalizeWordText(phrase));
    if (!phrases.length) return [{ text: source, bold: false }];

    const runs = [];
    let cursor = 0;
    const lowerSource = source.toLowerCase();
    while (cursor < source.length) {
        let bestMatch = null;
        phrases.forEach((phrase) => {
            const index = lowerSource.indexOf(phrase.toLowerCase(), cursor);
            if (index === -1) return;
            if (!bestMatch || index < bestMatch.index || (index === bestMatch.index && phrase.length > bestMatch.phrase.length)) {
                bestMatch = { index, phrase };
            }
        });

        if (!bestMatch) {
            runs.push({ text: source.slice(cursor), bold: false });
            break;
        }
        if (bestMatch.index > cursor) {
            runs.push({ text: source.slice(cursor, bestMatch.index), bold: false });
        }
        runs.push({
            text: source.slice(bestMatch.index, bestMatch.index + bestMatch.phrase.length),
            bold: true
        });
        cursor = bestMatch.index + bestMatch.phrase.length;
    }
    return runs.filter((run) => run.text);
}

function isDocumentHeading(text = '') {
    const value = String(text || '').trim();
    if (!value) return false;
    const normalized = value.replace(/\s+/g, ' ').replace(/[.,;]+$/g, '').toUpperCase();
    const exactHeadings = new Set([
        'PROVISIONAL OFFER OF MORTGAGE FACILITY',
        'LETTER OF INDEMNITY IN RESPECT OF RETIREMENT SAVINGS ACCOUNT EQUITY CONTRIBUTION',
        'VERIFICATION OF PROPERTY'
    ]);
    if (exactHeadings.has(normalized)) return true;
    if (normalized.startsWith('LETTER OF ALLOCATION FOR ')) return true;
    if (normalized.startsWith('CONFIRMATION OF TITLE DOCUMENT OF PROPERTY KNOWN AS ')) return true;
    return false;
}

function getDocxPageTopSpacerTwips(documentType, pageIndex) {
    if (pageIndex === 0) {
        const firstPageSpacers = {
            offer_letter: 420,
            allocation_letter: 760,
            availability_letter: 1100,
            indemnity_letter: 1100,
            readiness_letter: 900,
            title_letter: 760,
            verification_letter: 1000
        };
        if (Object.prototype.hasOwnProperty.call(firstPageSpacers, documentType)) {
            return firstPageSpacers[documentType];
        }
    }
    if (documentType === 'indemnity_letter' && pageIndex === 1) return 0;
    return 180;
}

function getDocxPreviewPaddingPx(documentType, pageIndex) {
    const spacer = getDocxPageTopSpacerTwips(documentType, pageIndex);
    if (spacer === 0) return 104;
    return 132 + Math.round(spacer / 12);
}

function createWordSpacerParagraphXml(twips = 0) {
    const value = Math.max(0, Number(twips || 0));
    if (!value) return '';
    return `<w:p><w:pPr><w:spacing w:before="0" w:after="${value}"/></w:pPr></w:p>`;
}

function createWordRunXml(run = {}, options = {}) {
    const text = normalizeWordText(run.text || '');
    if (!text) return '';
    const bold = run.bold || options.bold;
    const underline = options.underline;
    const size = Number(options.size || 20);
    return `
      <w:r>
        <w:rPr>
          <w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>
          <w:sz w:val="${size}"/><w:szCs w:val="${size}"/>
          ${bold ? '<w:b/><w:bCs/>' : ''}
          ${underline ? '<w:u w:val="single"/>' : ''}
        </w:rPr>
        <w:t xml:space="preserve">${escapeXml(text)}</w:t>
      </w:r>`;
}

function createWordParagraphXml(text = '', boldWords = [], options = {}) {
    const rawText = normalizeWordText(text);
    const lines = rawText.split(/\n/);
    const heading = options.heading ?? isDocumentHeading(rawText);
    const center = options.center ?? heading;
    const bold = options.bold ?? heading;
    const underline = options.underline ?? heading;
    const spacingAfter = Number(options.spacingAfter ?? (heading ? 180 : 140));
    const size = Number(options.size || 20);
    const lineHeight = Number(options.lineHeight || 276);
    const alignment = center ? '<w:jc w:val="center"/>' : '';

    if (!rawText.trim()) {
        return '<w:p><w:pPr><w:spacing w:after="120"/></w:pPr></w:p>';
    }

    const runXml = lines.map((line, index) => {
        const runs = splitTextIntoWordRuns(line, boldWords);
        const xml = runs.length
            ? runs.map((run) => createWordRunXml(run, { bold, underline, size })).join('')
            : createWordRunXml({ text: line, bold }, { bold, underline, size });
        return `${index > 0 ? '<w:r><w:br/></w:r>' : ''}${xml}`;
    }).join('');

    return `
    <w:p>
      <w:pPr>
        <w:spacing w:before="0" w:after="${spacingAfter}" w:line="${lineHeight}" w:lineRule="auto"/>
        ${alignment}
      </w:pPr>
      ${runXml}
    </w:p>`;
}

function createWordTableXml(rows = [], boldWords = [], options = {}) {
    if (!Array.isArray(rows) || !rows.length) return '';
    const compact = Boolean(options.compact);
    const offerTermsMode = options.offerTerms || '';
    const offerTerms = Boolean(offerTermsMode);
    const offerPrimaryTerms = offerTermsMode === 'primary';
    const hasFourColumns = rows.some((row) => Array.isArray(row) && row.length > 2);
    const columnWidths = hasFourColumns
        ? (compact ? [1500, 3100, 1500, 3100] : [1700, 2900, 1700, 2900])
        : (offerTerms ? [2600, 6600] : [compact ? 2500 : 3000, compact ? 6700 : 6200]);
    const fontSize = offerTerms ? 18 : (compact ? 14 : 18);
    const lineHeight = offerPrimaryTerms ? 285 : (offerTerms ? 235 : (compact ? 160 : 218));
    const spacingAfter = (compact || offerTerms) ? 0 : 8;
    const cellTopBottom = offerPrimaryTerms ? 54 : (offerTerms ? 34 : (compact ? 8 : 35));
    const cell = (content, width, bold = false) => `
      <w:tc>
        <w:tcPr><w:tcW w:w="${width}" w:type="dxa"/><w:vAlign w:val="top"/></w:tcPr>
        ${createWordParagraphXml(content, boldWords, { bold, spacingAfter, size: fontSize, lineHeight, center: false, underline: false })}
      </w:tc>`;
    const rowXml = rows.map((row) => {
        const safeRow = Array.isArray(row) ? row : [];
        const cells = hasFourColumns
            ? [
                cell(safeRow[0] || '', columnWidths[0], true),
                cell(safeRow[1] || '', columnWidths[1], false),
                cell(safeRow[2] || '', columnWidths[2], true),
                cell(safeRow[3] || '', columnWidths[3], false)
            ].join('')
            : [
                cell(safeRow[0] || '', columnWidths[0], true),
                cell(safeRow[1] || '', columnWidths[1], false)
            ].join('');
        return `<w:tr><w:trPr><w:trHeight w:val="0" w:hRule="auto"/></w:trPr>${cells}</w:tr>`;
    }).join('');
    const gridXml = columnWidths.map((width) => `<w:gridCol w:w="${width}"/>`).join('');
    return `
    <w:tbl>
      <w:tblPr>
        <w:tblW w:w="9200" w:type="dxa"/>
        <w:tblBorders>
          <w:top w:val="single" w:sz="4" w:space="0" w:color="D9E2EC"/>
          <w:left w:val="single" w:sz="4" w:space="0" w:color="D9E2EC"/>
          <w:bottom w:val="single" w:sz="4" w:space="0" w:color="D9E2EC"/>
          <w:right w:val="single" w:sz="4" w:space="0" w:color="D9E2EC"/>
          <w:insideH w:val="single" w:sz="4" w:space="0" w:color="E5E7EB"/>
          <w:insideV w:val="single" w:sz="4" w:space="0" w:color="E5E7EB"/>
        </w:tblBorders>
        <w:tblCellMar>
          <w:top w:w="${cellTopBottom}" w:type="dxa"/><w:left w:w="70" w:type="dxa"/>
          <w:bottom w:w="${cellTopBottom}" w:type="dxa"/><w:right w:w="70" w:type="dxa"/>
        </w:tblCellMar>
      </w:tblPr>
      <w:tblGrid>${gridXml}</w:tblGrid>
      ${rowXml}
    </w:tbl>`;
}

function createWordPageBreakXml() {
    return '<w:p><w:r><w:br w:type="page"/></w:r></w:p>';
}

const DOCX_TEMPLATE_PAGE_IMAGES = {
    offer_letter: { stem: 'offer-letter-master', pages: 4 },
    allocation_letter: { stem: 'allocation-letter-master', pages: 2 },
    availability_letter: { stem: 'availability-letter-master', pages: 1 },
    indemnity_letter: { stem: 'indemnity-letter-master', pages: 2 },
    readiness_letter: { stem: 'readiness-letter-master', pages: 1 },
    title_letter: { stem: 'title-letter-master', pages: 1 },
    verification_letter: { stem: 'verification-letter-master', pages: 1 }
};

async function loadDocxTemplatePageAssets(documentType, pageCount) {
    const config = DOCX_TEMPLATE_PAGE_IMAGES[documentType];
    if (!config) return [];
    const assets = [];
    for (let index = 0; index < pageCount; index += 1) {
        const imageIndex = Math.min(index + 1, config.pages);
        const fileName = `${config.stem}-page-${imageIndex}.png`;
        const url = `assets/document-templates/docx-page-images/${fileName}`;
        const response = await fetch(url, { cache: 'force-cache' });
        if (!response.ok) {
            assets.push(null);
            continue;
        }
        assets.push({
            rId: `rIdTemplatePage${index + 1}`,
            fileName,
            zipName: `template-page-${index + 1}.png`,
            url,
            bytes: await response.arrayBuffer()
        });
    }
    return assets;
}

function createWordBackgroundImageXml(asset = null, pageIndex = 0) {
    if (!asset?.rId) return '';
    const widthEmu = 7560310;
    const heightEmu = 10692130;
    return `
    <w:p>
      <w:pPr><w:spacing w:before="0" w:after="0"/></w:pPr>
      <w:r>
        <w:drawing>
          <wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="0" behindDoc="1" locked="0" layoutInCell="1" allowOverlap="1">
            <wp:simplePos x="0" y="0"/>
            <wp:positionH relativeFrom="page"><wp:posOffset>0</wp:posOffset></wp:positionH>
            <wp:positionV relativeFrom="page"><wp:posOffset>0</wp:posOffset></wp:positionV>
            <wp:extent cx="${widthEmu}" cy="${heightEmu}"/>
            <wp:effectExtent l="0" t="0" r="0" b="0"/>
            <wp:wrapNone/>
            <wp:docPr id="${9000 + pageIndex}" name="Template Page ${pageIndex + 1}"/>
            <wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr>
            <a:graphic>
              <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
                <pic:pic>
                  <pic:nvPicPr><pic:cNvPr id="${9000 + pageIndex}" name="${escapeAttribute(asset.fileName)}"/><pic:cNvPicPr/></pic:nvPicPr>
                  <pic:blipFill><a:blip r:embed="${asset.rId}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>
                  <pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${widthEmu}" cy="${heightEmu}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>
                </pic:pic>
              </a:graphicData>
            </a:graphic>
          </wp:anchor>
        </w:drawing>
      </w:r>
    </w:p>`;
}

function createDocxDocumentXml(documentType, contentModel = {}, pageAssets = []) {
    const body = [];
    const pages = Array.isArray(contentModel.pages) ? contentModel.pages : [];
    pages.forEach((page, pageIndex) => {
        const boldWords = normalizeBoldPhraseList(page.boldWords || []);
        body.push(createWordBackgroundImageXml(pageAssets[pageIndex], pageIndex));
        body.push(createWordSpacerParagraphXml(getDocxPageTopSpacerTwips(documentType, pageIndex)));
        if (page.tablePosition === 'before' && Array.isArray(page.table) && page.table.length) {
            body.push(createWordTableXml(page.table, boldWords, { compact: page.compactTable, offerTerms: page.offerTermsTable }));
        }
        (page.paragraphs || []).forEach((paragraph) => {
            body.push(createWordParagraphXml(paragraph, boldWords));
        });
        if (page.tablePosition !== 'before' && Array.isArray(page.table) && page.table.length) {
            body.push(createWordTableXml(page.table, boldWords, { compact: page.compactTable, offerTerms: page.offerTermsTable }));
        }
        if (pageIndex < pages.length - 1) body.push(createWordPageBreakXml());
    });

    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
  <w:body>
    ${body.join('')}
    <w:sectPr>
      <w:pgSz w:w="11906" w:h="16838"/>
      <w:pgMar w:top="1600" w:right="1134" w:bottom="1134" w:left="1134" w:header="0" w:footer="0" w:gutter="0"/>
      <w:cols w:space="720"/>
      <w:docGrid w:linePitch="360"/>
    </w:sectPr>
  </w:body>
</w:document>`;
}

function createDocxStylesXml() {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal">
    <w:name w:val="Normal"/>
    <w:qFormat/>
    <w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr>
  </w:style>
</w:styles>`;
}

async function generateDocxDocumentFromContent(documentType, data = {}) {
    if (!ensureDocxZipLibrary()) throw new Error('Word document engine is unavailable.');
    const contentModel = buildDocumentContent(documentType, data);
    const contentPages = Array.isArray(contentModel.pages) ? contentModel.pages : [];
    const pageAssets = await loadDocxTemplatePageAssets(documentType, contentPages.length);
    const zip = new window.PizZip();
    zip.file('[Content_Types].xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="png" ContentType="image/png"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>`);
    zip.folder('_rels').file('.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>`);
    zip.folder('word').file('document.xml', createDocxDocumentXml(documentType, contentModel, pageAssets));
    zip.folder('word').file('styles.xml', createDocxStylesXml());
    const imageRelationships = pageAssets
        .filter(Boolean)
        .map((asset) => `  <Relationship Id="${asset.rId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/${asset.zipName}"/>`)
        .join('\n');
    zip.folder('word').folder('_rels').file('document.xml.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
${imageRelationships}
</Relationships>`);
    pageAssets.filter(Boolean).forEach((asset) => {
        zip.folder('word').folder('media').file(asset.zipName, asset.bytes, { binary: true });
    });
    zip.folder('docProps').file('app.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><Application>CMBank RSA Admin Dashboard</Application></Properties>`);
    zip.folder('docProps').file('core.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:title>${escapeXml(GENERATED_DOCUMENT_TYPES.find((item) => item.id === documentType)?.label || documentType)}</dc:title><dc:creator>CMBank RSA Admin Dashboard</dc:creator><cp:lastModifiedBy>CMBank RSA Admin Dashboard</cp:lastModifiedBy></cp:coreProperties>`);

    const blob = zip.generate({
        type: 'blob',
        mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        compression: 'DEFLATE'
    });
    return {
        blob,
        previewHtml: renderDocxPreviewHtml(contentModel, pageAssets, documentType),
        contentModel
    };
}

function ensurePreviewPdfLibraries() {
    if (!window.html2canvas || !window.jspdf?.jsPDF) {
        throw new Error('PDF preview export engine is unavailable. Please refresh and try again.');
    }
}

async function createPdfBlobFromPreviewHtml(previewHtml = '') {
    ensurePreviewPdfLibraries();
    const host = document.createElement('div');
    host.className = 'generated-pdf-capture-host word-generated-doc-render';
    host.style.position = 'fixed';
    host.style.left = '-10000px';
    host.style.top = '0';
    host.style.width = '850px';
    host.style.background = '#eef4fb';
    host.style.padding = '26px';
    host.style.zIndex = '0';
    host.style.pointerEvents = 'none';
    host.innerHTML = previewHtml;
    document.body.appendChild(host);

    try {
        await new Promise((resolve) => requestAnimationFrame(() => resolve()));
        await withTimeout(waitForPreviewBackgroundImages(host), 10000, 'PDF template images took too long to load.');
        const pages = Array.from(host.querySelectorAll('.word-preview-page'));
        if (!pages.length) throw new Error('Generated document preview is empty.');

        const pdf = new window.jspdf.jsPDF({ orientation: 'portrait', unit: 'pt', format: 'a4' });
        const pageWidth = pdf.internal.pageSize.getWidth();
        const pageHeight = pdf.internal.pageSize.getHeight();

        for (let index = 0; index < pages.length; index += 1) {
            const page = pages[index];
            const canvas = await withTimeout(window.html2canvas(page, {
                scale: 2.5,
                backgroundColor: '#ffffff',
                useCORS: true,
                allowTaint: true,
                logging: false
            }), 45000, `PDF page ${index + 1} took too long to render.`);
            const imgData = canvas.toDataURL('image/png');
            if (index > 0) pdf.addPage();
            pdf.addImage(imgData, 'PNG', 0, 0, pageWidth, pageHeight, undefined, 'SLOW');
        }

        return pdf.output('blob');
    } finally {
        host.remove();
    }
}

function withTimeout(promise, timeoutMs = 30000, message = 'Operation timed out.') {
    let timer = null;
    return Promise.race([
        promise,
        new Promise((_, reject) => {
            timer = setTimeout(() => reject(new Error(message)), timeoutMs);
        })
    ]).finally(() => {
        if (timer) clearTimeout(timer);
    });
}

async function waitForPreviewBackgroundImages(host) {
    const urls = Array.from(host.querySelectorAll('.word-preview-page'))
        .map((page) => {
            const value = page.style.backgroundImage || '';
            const match = value.match(/url\(["']?(.+?)["']?\)/);
            return match ? match[1] : '';
        })
        .filter(Boolean);
    if (!urls.length) return;
    await Promise.all(urls.map((url) => new Promise((resolve) => {
        const img = new Image();
        img.onload = resolve;
        img.onerror = resolve;
        img.src = url;
    })));
}

async function generatePdfDocumentFromPreview(documentType, data = {}) {
    const generated = await generateDocxDocumentFromContent(documentType, data);
    const blob = await createPdfBlobFromPreviewHtml(generated.previewHtml);
    return {
        blob,
        previewUrl: URL.createObjectURL(blob),
        previewHtml: '',
        contentModel: generated.contentModel
    };
}

function renderDocxPreviewHtml(contentModel = {}, pageAssets = [], documentType = '') {
    const pages = Array.isArray(contentModel.pages) ? contentModel.pages : [];
    return pages.map((page, pageIndex) => {
        const boldWords = normalizeBoldPhraseList(page.boldWords || []);
        const buildParagraphHtml = () => (page.paragraphs || []).map((paragraph) => {
            const text = String(paragraph || '');
            if (!text.trim()) return '<p class="word-preview-spacer"></p>';
            const className = isDocumentHeading(text) ? ' class="word-preview-heading"' : '';
            const html = text.split(/\n/).map((line) => {
                const runs = splitTextIntoWordRuns(line, boldWords);
                return runs.map((run) => run.bold
                    ? `<strong>${escapeHtml(run.text)}</strong>`
                    : escapeHtml(run.text)).join('');
            }).join('<br>');
            return `<p${className}>${html}</p>`;
        }).join('');
        let tableHtml = '';
        if (Array.isArray(page.table) && page.table.length) {
            const hasFourColumns = page.table.some((row) => Array.isArray(row) && row.length > 2);
            const tableClasses = [
                page.compactTable ? 'compact-word-table' : '',
                page.offerTermsTable ? 'offer-terms-word-table' : '',
                page.offerTermsTable === 'primary' ? 'offer-terms-primary-table' : '',
                page.offerTermsTable === 'continuation' ? 'offer-terms-continuation-table' : '',
                hasFourColumns ? 'four-column-word-table' : ''
            ].filter(Boolean).join(' ');
            const tableRows = page.table.map((row) => {
                const safeRow = Array.isArray(row) ? row : [];
                return hasFourColumns
                    ? `<tr><th>${escapeHtml(safeRow[0] || '')}</th><td>${escapeHtml(safeRow[1] || '')}</td><th>${escapeHtml(safeRow[2] || '')}</th><td>${escapeHtml(safeRow[3] || '')}</td></tr>`
                    : `<tr><th>${escapeHtml(safeRow[0] || '')}</th><td>${escapeHtml(safeRow[1] || '')}</td></tr>`;
            }).join('');
            tableHtml = `<table class="${tableClasses}">${tableRows}</table>`;
        }
        const paragraphHtml = buildParagraphHtml();
        const contentHtml = page.tablePosition === 'before'
            ? `${tableHtml}${paragraphHtml}`
            : `${paragraphHtml}${tableHtml}`;
        const styleParts = [`padding-top: ${getDocxPreviewPaddingPx(documentType, pageIndex)}px`];
        if (pageAssets[pageIndex]?.url) {
            styleParts.push(`background-image: url('${escapeAttribute(pageAssets[pageIndex].url)}')`);
        }
        const styleAttr = ` style="${styleParts.join('; ')}"`;
        const spacer = getDocxPageTopSpacerTwips(documentType, pageIndex);
        const spacingClass = spacer >= 500 ? ' page-spaced-down' : (spacer === 0 ? ' page-spaced-up' : '');
        return `<section class="word-preview-page${spacingClass}"${styleAttr}><div class="word-preview-page-number">Page ${pageIndex + 1}</div>${contentHtml}</section>`;
    }).join('');
}

function resolveGeneratedDocumentFieldValue(fieldKey, data = {}) {
    const upperHouseType = String(data.houseType || '').toUpperCase();
    const upperCustomerName = String(data.customerName || '').toUpperCase();
    const safeEstateName = String(data.estateName || '').trim();
    const safeEstateAddress = String(data.estateAddress || '').trim();
    const valueMap = {
        currentDate: data.currentDate,
        currentDateShort: buildShortLetterDate(new Date()),
        customerName: data.customerName,
        customerAddress: data.customerAddress,
        propertyValueText: data.propertyValueText,
        loanAmountText: data.loanAmountText,
        equityContributionText: data.equityContributionText,
        houseNumber: data.houseNumber,
        houseType: data.houseType,
        pfaName: data.pfaName,
        rsaPin: data.rsaPin,
        bankAccountNumber: data.bankAccountNumber,
        offerPropertyValueLine: `${data.propertyValueText} (${data.propertyValueWords})`,
        offerPurposeFragment: `House ${data.houseNumber} of a ${upperHouseType}`,
        offerEquityLine: `${data.equityContributionText} (${data.equityContributionWords})`,
        offerSecurityFragment: `House ${data.houseNumber} of a ${upperHouseType}`,
        offerReevaluateFragment: `House ${data.houseNumber}, a ${upperHouseType}`,
        offerClause6Fragment: `House ${data.houseNumber} of the ${upperHouseType}`,
        offerClause8Fragment: `House ${data.houseNumber}, a ${upperHouseType}`,
        allocationHouseLine: `HOUSE ${data.houseNumber}, a unit of ${upperHouseType} in the Estate`,
        allocationValueLine: `${data.propertyValueText} (${data.propertyValueWords})`,
        availabilityLine1: `This is to confirm the authenticity of the property allocated to ${upperCustomerName}, which is a`,
        availabilityLine2: `${upperHouseType}, House number ${data.houseNumber} situated at ${safeEstateName}`,
        availabilityLine3: `${safeEstateName}, ${safeEstateAddress}, valued at ${data.propertyValueText}`,
        readinessNameLine: `Name: ${upperCustomerName}`,
        readinessAccountLine: `Account Number: ${data.bankAccountNumber}`,
        readinessEquityLine: `Equity Amount: ${data.equityContributionText} (${data.equityContributionWords})`,
        readinessLoanLine: upperCustomerName,
        titleHeaderLine: `${upperCustomerName} PIN:${data.rsaPin}`,
        titleBodyLine: upperCustomerName,
        verificationLine1: `This is to confirm the authenticity of the property offer allocated to ${upperCustomerName},`,
        verificationLine2: `which is House ${data.houseNumber}, a ${upperHouseType} situated at ${safeEstateName}`,
        verificationLine3: `${safeEstateName}, ${safeEstateAddress} valued at ${data.propertyValueText}`
    };
    return toPdfSafeText(valueMap[fieldKey] ?? '').trim();
}

function wrapPdfText(text, font, fontSize, maxWidth) {
    const content = String(text || '').trim();
    if (!content) return [''];
    const words = content.split(/\s+/);
    const lines = [];
    let currentLine = '';

    words.forEach((word) => {
        const candidate = currentLine ? `${currentLine} ${word}` : word;
        if (font.widthOfTextAtSize(candidate, fontSize) <= maxWidth || !currentLine) {
            currentLine = candidate;
        } else {
            lines.push(currentLine);
            currentLine = word;
        }
    });

    if (currentLine) lines.push(currentLine);
    return lines;
}

function pointsFromMillimeters(mmValue) {
    return Number(mmValue || 0) * 2.8346456693;
}

function normalizeBoldPhraseList(items = []) {
    return Array.from(new Set(
        items
            .map((item) => String(item || '').trim())
            .filter(Boolean)
    )).sort((a, b) => b.length - a.length);
}

function tokenizeRichText(text = '', boldPhrases = []) {
    const tokens = [];
    const source = String(text || '');
    let cursor = 0;
    while (cursor < source.length) {
        let match = null;
        for (const phrase of boldPhrases) {
            if (phrase && source.slice(cursor, cursor + phrase.length) === phrase) {
                match = phrase;
                break;
            }
        }
        if (match) {
            tokens.push({ text: match, bold: true });
            cursor += match.length;
            continue;
        }
        let next = cursor + 1;
        while (next < source.length) {
            const upcoming = boldPhrases.find((phrase) => phrase && source.slice(next, next + phrase.length) === phrase);
            if (upcoming) break;
            next += 1;
        }
        tokens.push({ text: source.slice(cursor, next), bold: false });
        cursor = next;
    }
    return tokens;
}

function splitRichTokensIntoWords(tokens = []) {
    const pieces = [];
    tokens.forEach((token) => {
        const parts = String(token.text || '').match(/\S+|\s+/g) || [];
        parts.forEach((part) => {
            pieces.push({ text: part, bold: token.bold });
        });
    });
    return pieces;
}

function measureRichWordWidth(word, fonts, fontSize) {
    const font = word.bold ? fonts.bold : fonts.regular;
    return font.widthOfTextAtSize(word.text, fontSize);
}

function wrapRichTextLine(text = '', fonts, fontSize, maxWidth, boldPhrases = []) {
    const tokens = splitRichTokensIntoWords(tokenizeRichText(text, boldPhrases));
    const lines = [];
    let currentLine = [];
    let lineWidth = 0;

    tokens.forEach((word) => {
        const wordWidth = measureRichWordWidth(word, fonts, fontSize);
        const nextWidth = lineWidth + wordWidth;
        const isLeadingWhitespace = currentLine.length === 0 && /^\s+$/.test(word.text);
        if (!isLeadingWhitespace && currentLine.length > 0 && nextWidth > maxWidth) {
            lines.push(currentLine);
            currentLine = /^\s+$/.test(word.text) ? [] : [word];
            lineWidth = /^\s+$/.test(word.text) ? 0 : wordWidth;
        } else {
            if (!isLeadingWhitespace) {
                currentLine.push(word);
                lineWidth = nextWidth;
            }
        }
    });

    if (currentLine.length) lines.push(currentLine);
    if (!lines.length) lines.push([]);
    return lines;
}

function drawRichTextLine(page, lineTokens, x, y, fonts, fontSize) {
    let cursorX = x;
    lineTokens.forEach((token) => {
        const font = token.bold ? fonts.bold : fonts.regular;
        page.drawText(token.text, {
            x: cursorX,
            y,
            size: fontSize,
            font,
            color: window.PDFLib.rgb(0, 0, 0)
        });
        cursorX += font.widthOfTextAtSize(token.text, fontSize);
    });
}

function renderParagraphBlock(page, paragraph = '', boldWords = [], layout, fonts) {
    const sourceLines = String(paragraph || '').split('\n');
    const contentWidth = layout.contentWidth;
    sourceLines.forEach((sourceLine) => {
        if (!sourceLine) {
            layout.cursorY -= layout.paragraphSpacing;
            return;
        }
        const wrappedLines = wrapRichTextLine(sourceLine, fonts, layout.fontSize, contentWidth, boldWords);
        wrappedLines.forEach((lineTokens) => {
            drawRichTextLine(page, lineTokens, layout.marginLeft, layout.cursorY, fonts, layout.fontSize);
            layout.cursorY -= layout.lineHeight;
        });
        layout.cursorY -= Math.max(0, layout.paragraphSpacing - layout.lineHeight);
    });
}

function wrapPlainText(text = '', font, fontSize, maxWidth) {
    return wrapPdfText(String(text || ''), font, fontSize, maxWidth);
}

function renderTableBlock(page, rows = [], layout, fonts) {
    const labelWidth = pointsFromMillimeters(50);
    const valueWidth = layout.contentWidth - labelWidth;
    rows.forEach(([label, value]) => {
        const valueLines = wrapPlainText(String(value || ''), fonts.regular, layout.tableFontSize, valueWidth - 8);
        const rowHeight = Math.max(layout.tableLineHeight, valueLines.length * layout.tableLineHeight);
        const rowBottom = layout.cursorY - rowHeight;
        page.drawRectangle({
            x: layout.marginLeft,
            y: rowBottom,
            width: labelWidth,
            height: rowHeight,
            borderWidth: 1,
            borderColor: window.PDFLib.rgb(0, 0, 0)
        });
        page.drawRectangle({
            x: layout.marginLeft + labelWidth,
            y: rowBottom,
            width: valueWidth,
            height: rowHeight,
            borderWidth: 1,
            borderColor: window.PDFLib.rgb(0, 0, 0)
        });

        page.drawText(String(label || ''), {
            x: layout.marginLeft + 4,
            y: layout.cursorY - layout.tableFontSize - 3,
            size: layout.tableFontSize,
            font: fonts.bold,
            color: window.PDFLib.rgb(0, 0, 0)
        });

        valueLines.forEach((line, index) => {
            page.drawText(line, {
                x: layout.marginLeft + labelWidth + 4,
                y: layout.cursorY - layout.tableFontSize - 3 - (index * layout.tableLineHeight),
                size: layout.tableFontSize,
                font: fonts.regular,
                color: window.PDFLib.rgb(0, 0, 0)
            });
        });

        layout.cursorY = rowBottom - pointsFromMillimeters(1.4);
    });
}

function getOverlayPageFormatPoints(documentType, pageIndex, page) {
    const format = GENERATED_DOCUMENT_PAGE_FORMATS_MM?.[documentType]?.[pageIndex];
    if (!format) {
        return { width: page.getWidth(), height: page.getHeight(), offsetX: 0, offsetY: 0 };
    }
    const width = pointsFromMillimeters(format[0]);
    const height = pointsFromMillimeters(format[1]);
    return {
        width,
        height,
        offsetX: 0,
        offsetY: Math.max(0, page.getHeight() - height)
    };
}

function pdfRectFromTopLeft(page, rect = []) {
    const [x1, y1, x2, y2] = rect.map((value) => Number(value || 0));
    const padding = 1.8;
    return {
        x: Math.max(0, x1 - padding),
        y: Math.max(0, page.getHeight() - y2 - padding),
        width: Math.max(1, x2 - x1 + (padding * 2)),
        height: Math.max(1, y2 - y1 + (padding * 2)),
        textX: x1,
        textY: page.getHeight() - y2 + 2,
        maxWidth: Math.max(1, x2 - x1)
    };
}

function wipePdfRect(page, rect = []) {
    const box = pdfRectFromTopLeft(page, rect);
    page.drawRectangle({
        x: box.x,
        y: box.y,
        width: box.width,
        height: box.height,
        color: window.PDFLib.rgb(1, 1, 1),
        borderWidth: 0
    });
}

function drawPdfField(page, field = {}, value = '', embeddedFonts = {}) {
    const content = toPdfSafeText(value).trim();
    if (!content) return;
    const box = pdfRectFromTopLeft(page, field.rect);
    const fontSize = Number(field.fontSize || 10);
    const fieldForFont = field.key === 'customerName'
        ? { ...field, fontWeight: 'bold' }
        : field;
    const profile = getPdfDocumentFontProfile('', fieldForFont);
    const font = embeddedFonts[profile.family] || embeddedFonts.dejavu || embeddedFonts.dejavuBold;
    const color = window.PDFLib.rgb(0, 0, 0);
    const lines = field.allowWrap
        ? wrapPdfText(content, font, fontSize, box.maxWidth)
        : [content];

    lines.slice(0, 3).forEach((line, index) => {
        page.drawText(line, {
            x: box.textX,
            y: box.textY - (index * (fontSize + 2)),
            size: fontSize,
            font,
            color,
            maxWidth: box.maxWidth
        });
    });
}

async function generatePdfDocumentFromTemplate(documentType, data = {}) {
    const config = PDF_TEMPLATE_CONFIGS[documentType];
    if (!config) throw new Error(`Unsupported document template: ${documentType}`);
    if (!window.PDFLib?.PDFDocument) throw new Error('PDF engine is unavailable. Please refresh and try again.');

    const fontAssets = await ensurePdfFontAssets();
    const templateResponse = await fetch(`assets/document-templates/${config.fileName}`, { cache: 'no-store' });
    if (!templateResponse.ok) {
        throw new Error(`Template not found for ${documentType}.`);
    }

    const templateBytes = await templateResponse.arrayBuffer();
    const templateDoc = await window.PDFLib.PDFDocument.load(templateBytes);
    templateDoc.registerFontkit(window.fontkit);
    const templatePages = templateDoc.getPages();
    const embeddedFonts = {
        dejavu: await templateDoc.embedFont(fontAssets.dejavu),
        dejavuBold: await templateDoc.embedFont(fontAssets.dejavuBold),
        trebuchet: fontAssets.trebuchet ? await templateDoc.embedFont(fontAssets.trebuchet) : null,
        trebuchetBold: fontAssets.trebuchetBold ? await templateDoc.embedFont(fontAssets.trebuchetBold) : null
    };

    (config.wipeZones || []).forEach((zone) => {
        const page = templatePages[zone.page];
        if (page) wipePdfRect(page, zone.rect);
    });

    (config.fields || []).forEach((field) => {
        const page = templatePages[field.page];
        if (!page) return;
        wipePdfRect(page, field.rect);
        const value = resolveGeneratedDocumentFieldValue(field.key, data);
        drawPdfField(page, field, value, embeddedFonts);
    });

    const pdfBytes = await templateDoc.save();
    return new Blob([pdfBytes], { type: 'application/pdf' });
}

function resetGeneratedDocumentPreviewItems() {
    generatedDocumentPreviewItems.forEach((item) => {
        if (item?.previewUrl) URL.revokeObjectURL(item.previewUrl);
    });
    generatedDocumentPreviewItems = [];
}

function renderDocumentGenerationChecklist(submission = {}) {
    if (!documentGenerationChecklist) return;
    documentGenerationChecklist.innerHTML = GENERATED_DOCUMENT_TYPES.map((item) => `
        <label class="document-check-card">
            <input type="checkbox" class="document-check-input" value="${item.id}" checked>
            <span class="document-check-copy">
                <strong>${escapeHtml(item.label)}</strong>
                <span>${escapeHtml(item.description)}</span>
            </span>
        </label>
    `).join('');
    if (documentGenerationCustomerName) documentGenerationCustomerName.textContent = submission.customerName || 'Customer';
    if (documentGenerationMeta) {
        const details = submission.customerDetails || {};
        documentGenerationMeta.textContent = `Application ID: ${submission.id || '-'} | PFA: ${details.pfa || submission.pfa || '-'} | House No: ${details.houseNumber || submission.houseNumber || '-'}`;
    }
}

function renderGenerateDocumentsTable() {
    if (!generateDocumentsTableBody) return;
    const items = allSubmissions
        .filter((sub) => {
            const status = String(sub.status || '').toLowerCase();
            return status === 'processing_to_pfa' || status === 'approved';
        })
        .slice()
        .sort((a, b) => getStageTimestampMillis(getSubmissionApprovalEntryAt(b)) - getStageTimestampMillis(getSubmissionApprovalEntryAt(a)));

    if (!items.length) {
        generateDocumentsTableBody.innerHTML = '<tr><td colspan="8" class="no-data">No Processing to PFA applications available for document generation.</td></tr>';
        return;
    }

    generateDocumentsTableBody.innerHTML = items.map((sub) => {
        const details = sub.customerDetails || {};
        return `
            <tr>
                <td><strong>${escapeHtml(sub.customerName || 'Unknown')}</strong></td>
                <td>${escapeHtml(getDisplayNameByEmail(sub.uploadedBy || ''))}</td>
                <td>${escapeHtml(details.pfa || sub.pfa || '-')}</td>
                <td>${escapeHtml(details.houseNumber || sub.houseNumber || '-')}</td>
                <td>${escapeHtml(details.propertyType || details.houseType || '-')}</td>
                <td>${escapeHtml(sub.assignedToRSA ? getDisplayNameByEmail(sub.assignedToRSA) : '-')}</td>
                <td><span class="status-badge status-approved">${escapeHtml(formatStatusLabel(sub.status || '-'))}</span></td>
                <td>
                    <button class="action-btn view-btn-small" onclick="window.openDocumentGenerationModal('${sub.id}', this)">
                        <i class="fas fa-file-circle-plus"></i> Generate
                    </button>
                </td>
            </tr>
        `;
    }).join('');
}

function toggleAllDocumentSelections() {
    const checks = Array.from(documentGenerationChecklist?.querySelectorAll('.document-check-input') || []);
    if (!checks.length) return;
    const shouldCheckAll = checks.some((check) => !check.checked);
    checks.forEach((check) => { check.checked = shouldCheckAll; });
}

window.openDocumentGenerationModal = async (submissionId, triggerBtn) => {
    const submission = allSubmissions.find((item) => item.id === submissionId);
    if (!submission || !documentGenerationModal) return;
    currentDocumentGenerationSubmissionId = submissionId;
    const originalHtml = triggerBtn?.innerHTML || '';
    if (triggerBtn) {
        triggerBtn.disabled = true;
        triggerBtn.classList.add('loading');
        triggerBtn.innerHTML = '<i class="fas fa-spinner fa-spin"></i> Opening...';
    }
    try {
        renderDocumentGenerationChecklist(submission);
        documentGenerationModal.classList.add('active');
    } finally {
        if (triggerBtn) {
            triggerBtn.disabled = false;
            triggerBtn.classList.remove('loading');
            triggerBtn.innerHTML = originalHtml || '<i class="fas fa-file-circle-plus"></i> Generate';
        }
    }
};

function getSelectedDocumentTypes() {
    return Array.from(documentGenerationChecklist?.querySelectorAll('.document-check-input:checked') || []).map((input) => input.value);
}

async function generateSelectedDocumentsForPreview() {
    const submission = allSubmissions.find((item) => item.id === currentDocumentGenerationSubmissionId);
    if (!submission) {
        showNotification('Application not found for document generation.', 'error');
        return;
    }
    const selectedTypes = getSelectedDocumentTypes();
    if (!selectedTypes.length) {
        showNotification('Select at least one document to generate.', 'warning');
        return;
    }

    const generateBtn = document.getElementById('generateSelectedDocumentsBtn');
    const originalBtnHtml = generateBtn?.innerHTML || '';
    if (generateBtn) {
        generateBtn.disabled = true;
        generateBtn.innerHTML = '<i class="fas fa-spinner fa-spin"></i> Generating...';
    }

    try {
        if (!adminSystemSettings?.pfaOptions) {
            await loadAdminSystemSettings();
        }
        const data = buildLetterDocumentData(submission);
        resetGeneratedDocumentPreviewItems();
        const previewItems = [];

        for (const type of selectedTypes) {
            const config = GENERATED_DOCUMENT_TYPES.find((item) => item.id === type);
            const isPostApprovalDocument = ['availability_letter', 'readiness_letter', 'title_letter'].includes(type);
            const documentData = {
                ...data,
                currentDate: isPostApprovalDocument ? data.postApprovalDate : data.currentDate
            };
            const generated = await generateDocxDocumentFromContent(type, documentData);
            previewItems.push({
                type,
                label: config?.label || type,
                blob: null,
                previewUrl: '',
                previewHtml: generated.previewHtml,
                fileExtension: 'pdf',
                mimeType: 'application/pdf',
                customerName: data.customerName
            });
        }

        generatedDocumentPreviewItems = previewItems;
        renderGeneratedDocumentsPreview(data.customerName);
        closeDocumentGenerationModalFn();
        if (generatedDocumentsPreviewModal) generatedDocumentsPreviewModal.classList.add('active');
    } catch (error) {
        showNotification(`Document generation failed: ${error.message || 'Unknown error'}`, 'error');
    } finally {
        if (generateBtn) {
            generateBtn.disabled = false;
            generateBtn.innerHTML = originalBtnHtml || '<i class="fas fa-file-circle-plus"></i> Generate Selected';
        }
    }
}

function renderGeneratedDocumentsPreview(customerName = 'Customer') {
    if (!generatedDocumentsPreviewList || !generatedDocumentsPreviewMeta) return;
    generatedDocumentsPreviewMeta.textContent = `${customerName} | ${generatedDocumentPreviewItems.length} document(s) generated`;
    generatedDocumentsPreviewList.innerHTML = generatedDocumentPreviewItems.map((item, index) => `
        <div class="generated-doc-card">
            <div class="generated-doc-card-head">
                <h3>${escapeHtml(item.label)}</h3>
                <button class="action-btn" onclick="window.saveGeneratedDocumentPdf(${index}, this)">
                    <i class="fas fa-file-pdf"></i> Save as PDF
                </button>
            </div>
            <div class="generated-doc-render ${item.previewHtml ? 'word-generated-doc-render' : ''}" data-generated-doc-index="${index}">
                ${item.previewUrl
                    ? `<iframe class="generated-doc-frame" src="${escapeAttribute(item.previewUrl)}" title="${escapeAttribute(item.label)} preview"></iframe>`
                    : (item.previewHtml || '<div class="no-data">Preview unavailable.</div>')}
            </div>
        </div>
    `).join('');
}

function sanitizeFileNamePart(value = '') {
    return String(value || '').replace(/[\\/:*?"<>|]/g, '_').replace(/\s+/g, ' ').trim();
}

async function saveGeneratedDocumentAtIndex(index, directoryHandle = null) {
    const item = generatedDocumentPreviewItems[index];
    if (!item?.blob && item?.previewHtml) {
        item.blob = await createPdfBlobFromPreviewHtml(item.previewHtml);
        item.mimeType = 'application/pdf';
        item.fileExtension = 'pdf';
        generatedDocumentPreviewItems[index] = item;
    }
    if (!item?.blob) throw new Error('Generated document preview is unavailable.');
    const customerName = sanitizeFileNamePart(item.customerName || 'Customer') || 'Customer';
    const extension = sanitizeFileNamePart(item.fileExtension || 'docx') || 'docx';
    const fileName = `${customerName} - ${sanitizeFileNamePart(item.label)}.${extension}`;

    if (directoryHandle && 'getDirectoryHandle' in directoryHandle) {
        const customerFolder = await directoryHandle.getDirectoryHandle(customerName, { create: true });
        const fileHandle = await customerFolder.getFileHandle(fileName, { create: true });
        const writable = await fileHandle.createWritable();
        await writable.write(item.blob);
        await writable.close();
        return true;
    }

    return saveFileWithLocationPicker(item.blob, fileName);
}

async function ensureDirectoryWritePermission(directoryHandle) {
    if (!directoryHandle || typeof directoryHandle.queryPermission !== 'function') return true;
    const descriptor = { mode: 'readwrite' };
    const current = await directoryHandle.queryPermission(descriptor);
    if (current === 'granted') return true;
    if (typeof directoryHandle.requestPermission !== 'function') return false;
    const requested = await directoryHandle.requestPermission(descriptor);
    return requested === 'granted';
}

window.saveGeneratedDocumentPdf = async (index, triggerBtn = null) => {
    const originalHtml = triggerBtn?.innerHTML || '';
    try {
        if (triggerBtn) {
            triggerBtn.disabled = true;
            triggerBtn.innerHTML = '<i class="fas fa-spinner fa-spin"></i> Preparing PDF...';
        }
        await saveGeneratedDocumentAtIndex(index);
    } catch (error) {
        showNotification(`Failed to save document: ${error.message || 'Unknown error'}`, 'error');
    } finally {
        if (triggerBtn) {
            triggerBtn.disabled = false;
            triggerBtn.innerHTML = originalHtml || '<i class="fas fa-file-pdf"></i> Save as PDF';
        }
    }
};

async function saveAllGeneratedDocumentsToFolder() {
    if (!generatedDocumentPreviewItems.length) {
        showNotification('No generated documents available to save.', 'warning');
        return;
    }
    const originalHtml = saveAllGeneratedDocumentsBtn?.innerHTML || '';
    if (saveAllGeneratedDocumentsBtn) {
        saveAllGeneratedDocumentsBtn.disabled = true;
        saveAllGeneratedDocumentsBtn.innerHTML = '<i class="fas fa-spinner fa-spin"></i> Preparing PDFs...';
    }
    if (!('showDirectoryPicker' in window)) {
        try {
            for (let index = 0; index < generatedDocumentPreviewItems.length; index += 1) {
                await window.saveGeneratedDocumentPdf(index);
            }
        } finally {
            if (saveAllGeneratedDocumentsBtn) {
                saveAllGeneratedDocumentsBtn.disabled = false;
                saveAllGeneratedDocumentsBtn.innerHTML = originalHtml || '<i class="fas fa-folder-plus"></i> Save All as PDF';
            }
        }
        return;
    }

    try {
        const rootFolder = await window.showDirectoryPicker({ mode: 'readwrite', startIn: 'downloads' });
        const canWrite = await ensureDirectoryWritePermission(rootFolder);
        if (!canWrite) throw new Error('Write permission was not granted for the selected folder.');
        let successCount = 0;
        const failures = [];
        for (let index = 0; index < generatedDocumentPreviewItems.length; index += 1) {
            showNotification(`Preparing PDF ${index + 1} of ${generatedDocumentPreviewItems.length}...`, 'info');
            try {
                await saveGeneratedDocumentAtIndex(index, rootFolder);
                successCount += 1;
                showNotification(`Saved ${index + 1} of ${generatedDocumentPreviewItems.length}`, 'success');
            } catch (error) {
                failures.push(`${generatedDocumentPreviewItems[index]?.label || `Document ${index + 1}`}: ${error.message || 'Failed'}`);
            }
        }
        if (successCount) {
            showNotification(`${successCount} document(s) exported to folder successfully.`, failures.length ? 'warning' : 'success');
        }
        if (failures.length) {
            showNotification(`Some documents failed: ${failures.slice(0, 2).join(' | ')}`, 'error');
        }
    } catch (error) {
        if (error?.name === 'AbortError') {
            showNotification('Save cancelled', 'info');
            return;
        }
        showNotification(`Failed to save generated documents: ${error.message || 'Unknown error'}`, 'error');
    } finally {
        if (saveAllGeneratedDocumentsBtn) {
            saveAllGeneratedDocumentsBtn.disabled = false;
            saveAllGeneratedDocumentsBtn.innerHTML = originalHtml || '<i class="fas fa-folder-plus"></i> Save All as PDF';
        }
    }
}

function setTrackReportInlineStatus(message, type = '') {
    if (!trackReportInlineStatus) return;
    trackReportInlineStatus.textContent = String(message || '');
    trackReportInlineStatus.className = `track-report-inline-status ${type}`.trim();
}

function ensureXlsxLibrary() {
    if (window.XLSX) return true;
    showNotification('Excel library is unavailable. Please refresh and try again.', 'error');
    return false;
}

function ensureExcelJsLibrary() {
    if (window.ExcelJS) return true;
    showNotification('Styled Excel export is unavailable right now. Please refresh and try again.', 'error');
    return false;
}

function normalizeCustomerName(name) {
    return String(name || '')
        .toLowerCase()
        .replace(/[^\p{L}\p{N}\s]/gu, ' ')
        .replace(/\s+/g, ' ')
        .trim();
}

function getCustomerNameMatchKeys(name) {
    const normalized = normalizeCustomerName(name);
    const compact = normalized.replace(/\s+/g, '');
    const tokens = normalized.split(' ').filter(Boolean);
    const sortedTokens = tokens.slice().sort();
    return { normalized, compact, tokens, sortedCompact: sortedTokens.join('') };
}

function scoreCustomerNameMatch(inputName, candidateName) {
    const input = getCustomerNameMatchKeys(inputName);
    const candidate = getCustomerNameMatchKeys(candidateName);
    if (!input.normalized || !candidate.normalized) return 0;
    if (input.normalized === candidate.normalized) return 100;
    if (input.compact && input.compact === candidate.compact) return 95;
    if (input.sortedCompact && input.sortedCompact === candidate.sortedCompact) return 93;
    if (candidate.normalized.includes(input.normalized) || input.normalized.includes(candidate.normalized)) return 88;

    if (!input.tokens.length || !candidate.tokens.length) return 0;
    const shared = input.tokens.filter((token) => candidate.tokens.includes(token));
    if (!shared.length) return 0;

    const coverage = shared.length / Math.max(input.tokens.length, candidate.tokens.length);
    if (shared.length >= Math.min(input.tokens.length, candidate.tokens.length) && coverage >= 0.66) {
        return 82 + Math.round(coverage * 10);
    }
    if (shared.length === input.tokens.length || shared.length === candidate.tokens.length) {
        return 75 + Math.round(coverage * 10);
    }
    return Math.round(coverage * 70);
}

function splitCustomerNames(rawText) {
    return String(rawText || '')
        .split(/\r?\n|,|;|\t/)
        .map((item) => item.trim())
        .filter(Boolean);
}

async function handleTrackReportFileSelected(event) {
    const file = event?.target?.files?.[0];
    if (!file) {
        trackReportParsedNames = [];
        return;
    }
    if (!ensureXlsxLibrary()) return;

    try {
        const buffer = await file.arrayBuffer();
        const workbook = window.XLSX.read(buffer, { type: 'array' });
        const firstSheetName = workbook.SheetNames[0];
        const sheet = workbook.Sheets[firstSheetName];
        const rows = window.XLSX.utils.sheet_to_json(sheet, { header: 1, blankrows: false });
        const names = rows
            .flat()
            .map((value) => String(value || '').trim())
            .filter(Boolean)
            .filter((value) => normalizeCustomerName(value) !== 'customer name');

        trackReportParsedNames = Array.from(new Set(names));
        setTrackReportInlineStatus(`${trackReportParsedNames.length} customer name(s) loaded from ${file.name}.`, 'success');
    } catch (error) {
        trackReportParsedNames = [];
        setTrackReportInlineStatus('Could not read the uploaded file. Please use the provided template.', 'error');
    }
}

function clearTrackReportInputs() {
    trackReportParsedNames = [];
    trackReportPreviewRows = [];
    trackReportUnmatchedEntries = [];
    if (trackReportFileInput) trackReportFileInput.value = '';
    if (trackReportNamesInput) trackReportNamesInput.value = '';
    setTrackReportInlineStatus('');
}

function getPaymentHandlingStatus(submission = {}) {
    const status = String(submission.status || '').trim().toLowerCase();
    const paymentOfficer = submission.assignedToPayment ? getDisplayNameByEmail(submission.assignedToPayment) : '';
    if (status === 'cleared') return paymentOfficer ? `Cleared by ${paymentOfficer}` : 'Cleared';
    if (status === 'paid') return paymentOfficer ? `Paid and handled by ${paymentOfficer}` : 'Paid';
    if (['sent_to_pfa', 'rsa_submitted'].includes(status) || submission.finalSubmitted === true || submission.rsaSubmitted === true) {
        return paymentOfficer ? `Assigned to ${paymentOfficer}` : 'Awaiting payment assignment';
    }
    return 'Not yet in payment stage';
}

function buildTrackReportRow(submission = {}) {
    const currentStage = getApplicationCurrentStage(submission);
    const reviewApprovedAt = submission.reviewedAt || null;
    const rsaQueueAt = submission.rsaAssignedAt || getSubmissionApprovalEntryAt(submission) || null;
    const rsaApprovedAt = submission.rsaSubmittedAt || submission.finalSubmittedAt || null;
    const paymentQueueAt = submission.paymentAssignedAt || getSubmissionPaymentEntryAt(submission) || null;

    return {
        'Customer Name': submission.customerName || 'Unknown',
        'Upload Timestamp': formatDate(getTrackStageTimestamp(submission, 'upload')),
        'Assigned Reviewer': submission.assignedTo ? getDisplayNameByEmail(submission.assignedTo) : 'Unassigned',
        'Reviewer Approval Timestamp': formatDate(reviewApprovedAt),
        'RSA Queue Timestamp': formatDate(rsaQueueAt),
        'Assigned RSA Officer': submission.assignedToRSA ? getDisplayNameByEmail(submission.assignedToRSA) : 'Unassigned',
        'RSA Approval Timestamp': formatDate(rsaApprovedAt),
        'Payment Queue Timestamp': formatDate(paymentQueueAt),
        'Payment Handling Status': getPaymentHandlingStatus(submission),
        'Current Processing Stage': currentStage.label
    };
}

function sortSubmissionsByLatest(submissions = []) {
    return submissions.slice().sort((a, b) => getStageTimestampMillis(getSubmissionCurrentStageEntryAt(b)) - getStageTimestampMillis(getSubmissionCurrentStageEntryAt(a)));
}

function getBestCustomerNameSuggestion(inputName, submissions = []) {
    const ranked = submissions
        .map((submission) => ({
            customerName: submission.customerName || '',
            score: scoreCustomerNameMatch(inputName, submission.customerName || '')
        }))
        .filter((item) => item.customerName)
        .sort((a, b) => b.score - a.score);
    return ranked[0] || null;
}

function findBestSubmissionByCustomerName(inputName, submissions = []) {
    const ranked = submissions
        .map((submission) => ({
            submission,
            score: scoreCustomerNameMatch(inputName, submission.customerName || '')
        }))
        .filter((item) => item.score >= 68)
        .sort((a, b) => {
            if (b.score !== a.score) return b.score - a.score;
            return getStageTimestampMillis(getSubmissionCurrentStageEntryAt(b.submission)) - getStageTimestampMillis(getSubmissionCurrentStageEntryAt(a.submission));
        });

    return ranked[0]?.submission || null;
}

function renderTrackReportPreview() {
    if (!trackReportPreviewSummary || !trackReportPreviewAlerts || !trackReportPreviewTableBody) return;

    trackReportPreviewSummary.innerHTML = `
        <div class="track-report-preview-card">
            <span class="label">Names Submitted</span>
            <div class="value">${trackReportPreviewRows.length + trackReportUnmatchedEntries.length}</div>
        </div>
        <div class="track-report-preview-card">
            <span class="label">Matched Records</span>
            <div class="value">${trackReportPreviewRows.length}</div>
        </div>
        <div class="track-report-preview-card">
            <span class="label">Unmatched Names</span>
            <div class="value">${trackReportUnmatchedEntries.length}</div>
        </div>
    `;

    const alerts = [];
    if (trackReportUnmatchedEntries.length) {
        alerts.push(`
            <div class="track-report-preview-alert warning">
                ${trackReportUnmatchedEntries.map((entry) => {
                    const suggestion = entry?.suggestion?.customerName
                        ? ` (closest match: ${entry.suggestion.customerName})`
                        : '';
                    return escapeHtml(`${entry.inputName}${suggestion}`);
                }).join('<br>')}
            </div>
        `);
    }
    if (trackReportPreviewRows.length) {
        alerts.push(`
            <div class="track-report-preview-alert info">
                Review the preview below, then download the Excel report when you are satisfied.
            </div>
        `);
    }
    trackReportPreviewAlerts.innerHTML = alerts.join('');

    if (!trackReportPreviewRows.length) {
        trackReportPreviewTableBody.innerHTML = '<tr><td colspan="10" class="no-data">No matching applications found for the submitted customer names.</td></tr>';
        return;
    }

    trackReportPreviewTableBody.innerHTML = trackReportPreviewRows.map((row) => `
        <tr>
            <td><strong>${escapeHtml(row['Customer Name'])}</strong></td>
            <td>${escapeHtml(row['Upload Timestamp'])}</td>
            <td>${escapeHtml(row['Assigned Reviewer'])}</td>
            <td>${escapeHtml(row['Reviewer Approval Timestamp'])}</td>
            <td>${escapeHtml(row['RSA Queue Timestamp'])}</td>
            <td>${escapeHtml(row['Assigned RSA Officer'])}</td>
            <td>${escapeHtml(row['RSA Approval Timestamp'])}</td>
            <td>${escapeHtml(row['Payment Queue Timestamp'])}</td>
            <td>${escapeHtml(row['Payment Handling Status'])}</td>
            <td>${escapeHtml(row['Current Processing Stage'])}</td>
        </tr>
    `).join('');
}

async function generateTrackReportPreview() {
    const typedNames = splitCustomerNames(trackReportNamesInput?.value || '');
    const sourceNames = Array.from(new Set([...trackReportParsedNames, ...typedNames]));

    if (!sourceNames.length) {
        setTrackReportInlineStatus('Please upload a customer list or paste at least one customer name.', 'error');
        return;
    }

    const sortableSubmissions = sortSubmissionsByLatest(allSubmissions);
    const matchedRows = [];
    const unmatched = [];

    sourceNames.forEach((name) => {
        const match = findBestSubmissionByCustomerName(name, sortableSubmissions);
        if (!match) {
            unmatched.push({
                inputName: name,
                suggestion: getBestCustomerNameSuggestion(name, sortableSubmissions)
            });
            return;
        }
        matchedRows.push(buildTrackReportRow(match));
    });

    trackReportPreviewRows = matchedRows;
    trackReportUnmatchedEntries = unmatched;
    renderTrackReportPreview();
    closeTrackReportInputModalFn();
    if (trackReportPreviewModal) trackReportPreviewModal.classList.add('active');
    setTrackReportInlineStatus(`${matchedRows.length} record(s) prepared for preview.`, matchedRows.length ? 'success' : 'error');
}

function getTrackingReportHeaders() {
    return [
        'Customer Name',
        'Upload Timestamp',
        'Assigned Reviewer',
        'Reviewer Approval Timestamp',
        'RSA Queue Timestamp',
        'Assigned RSA Officer',
        'RSA Approval Timestamp',
        'Payment Queue Timestamp',
        'Payment Handling Status',
        'Current Processing Stage'
    ];
}

function applyTrackingHeaderStyle(row) {
    row.eachCell((cell) => {
        cell.font = { name: 'Calibri', size: 11, bold: true, color: { argb: 'FFFFFFFF' } };
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF0F3B67' } };
        cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
        cell.border = {
            top: { style: 'thin', color: { argb: 'FFD1D5DB' } },
            left: { style: 'thin', color: { argb: 'FFD1D5DB' } },
            bottom: { style: 'thin', color: { argb: 'FFD1D5DB' } },
            right: { style: 'thin', color: { argb: 'FFD1D5DB' } }
        };
    });
}

function applyTrackingBodyStyle(row, fillColor = 'FFFFFFFF') {
    row.eachCell((cell) => {
        cell.font = { name: 'Calibri', size: 11, color: { argb: 'FF0F172A' } };
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: fillColor } };
        cell.alignment = { vertical: 'middle', wrapText: true };
        cell.border = {
            top: { style: 'thin', color: { argb: 'FFE5E7EB' } },
            left: { style: 'thin', color: { argb: 'FFE5E7EB' } },
            bottom: { style: 'thin', color: { argb: 'FFE5E7EB' } },
            right: { style: 'thin', color: { argb: 'FFE5E7EB' } }
        };
    });
}

function applyTrackingSummaryCellStyle(cell, { fill = 'FFF8FAFC', fontColor = 'FF0F172A', bold = true, align = 'left' } = {}) {
    cell.font = { name: 'Calibri', size: 11, bold, color: { argb: fontColor } };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: fill } };
    cell.alignment = { vertical: 'middle', horizontal: align, wrapText: true };
    cell.border = {
        top: { style: 'thin', color: { argb: 'FFE5E7EB' } },
        left: { style: 'thin', color: { argb: 'FFE5E7EB' } },
        bottom: { style: 'thin', color: { argb: 'FFE5E7EB' } },
        right: { style: 'thin', color: { argb: 'FFE5E7EB' } }
    };
}

function styleTrackingStageCell(cell, value) {
    const normalized = String(value || '').toLowerCase();
    cell.font = { name: 'Calibri', size: 11, bold: true, color: { argb: 'FF0F172A' } };
    cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
    if (normalized.includes('rejected')) {
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFEE2E2' } };
        cell.font = { name: 'Calibri', size: 11, bold: true, color: { argb: 'FF991B1B' } };
    } else if (normalized.includes('payment') || normalized.includes('paid') || normalized.includes('cleared')) {
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFDCFCE7' } };
        cell.font = { name: 'Calibri', size: 11, bold: true, color: { argb: 'FF166534' } };
    } else if (normalized.includes('rsa')) {
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE0F2FE' } };
        cell.font = { name: 'Calibri', size: 11, bold: true, color: { argb: 'FF0C4A6E' } };
    } else if (normalized.includes('reviewer') || normalized.includes('pending')) {
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFEF3C7' } };
        cell.font = { name: 'Calibri', size: 11, bold: true, color: { argb: 'FF92400E' } };
    }
}

async function downloadWorkbookFromRows(fileName, rows) {
    if (!ensureExcelJsLibrary()) return;
    const headers = getTrackingReportHeaders();
    const workbook = new window.ExcelJS.Workbook();
    workbook.creator = 'CMBank RSA Admin Dashboard';
    workbook.company = 'CMBank';
    workbook.created = new Date();
    workbook.modified = new Date();

    const worksheet = workbook.addWorksheet('Tracking Report', {
        views: [{ state: 'frozen', ySplit: 5 }]
    });

    worksheet.columns = [
        { header: headers[0], key: headers[0], width: 30 },
        { header: headers[1], key: headers[1], width: 22 },
        { header: headers[2], key: headers[2], width: 24 },
        { header: headers[3], key: headers[3], width: 24 },
        { header: headers[4], key: headers[4], width: 22 },
        { header: headers[5], key: headers[5], width: 24 },
        { header: headers[6], key: headers[6], width: 22 },
        { header: headers[7], key: headers[7], width: 22 },
        { header: headers[8], key: headers[8], width: 30 },
        { header: headers[9], key: headers[9], width: 24 }
    ];

    const titleRow = worksheet.addRow(['Customer Application Tracking Report']);
    worksheet.mergeCells(`A${titleRow.number}:J${titleRow.number}`);
    titleRow.height = 28;
    titleRow.getCell(1).font = { name: 'Calibri', size: 16, bold: true, color: { argb: 'FFFFFFFF' } };
    titleRow.getCell(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF0F3B67' } };
    titleRow.getCell(1).alignment = { horizontal: 'center', vertical: 'middle' };

    const metaRow = worksheet.addRow([`Generated At: ${new Date().toLocaleString()}`]);
    worksheet.mergeCells(`A${metaRow.number}:J${metaRow.number}`);
    metaRow.height = 20;
    metaRow.getCell(1).font = { name: 'Calibri', size: 11, italic: true, color: { argb: 'FF334155' } };
    metaRow.getCell(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE2E8F0' } };
    metaRow.getCell(1).alignment = { horizontal: 'left', vertical: 'middle' };

    worksheet.addRow([]);

    const summaryLabelRow = worksheet.addRow(['Report Summary', '', 'Matched Records', rows.length, 'Unmatched Names', trackReportUnmatchedEntries.length]);
    worksheet.mergeCells(`A${summaryLabelRow.number}:B${summaryLabelRow.number}`);
    summaryLabelRow.height = 22;
    applyTrackingSummaryCellStyle(summaryLabelRow.getCell(1), { fill: 'FFE8F5E9', fontColor: 'FF166534', align: 'left' });
    applyTrackingSummaryCellStyle(summaryLabelRow.getCell(2), { fill: 'FFE8F5E9', fontColor: 'FF166534', align: 'left' });
    applyTrackingSummaryCellStyle(summaryLabelRow.getCell(3), { fill: 'FFE0F2FE', fontColor: 'FF0C4A6E', align: 'center' });
    applyTrackingSummaryCellStyle(summaryLabelRow.getCell(4), { fill: 'FFE0F2FE', fontColor: 'FF0C4A6E', align: 'center' });
    applyTrackingSummaryCellStyle(summaryLabelRow.getCell(5), { fill: 'FFFFF7ED', fontColor: 'FF9A3412', align: 'center' });
    applyTrackingSummaryCellStyle(summaryLabelRow.getCell(6), { fill: 'FFFFF7ED', fontColor: 'FF9A3412', align: 'center' });

    const sectionRow = worksheet.addRow(['Tracking Details']);
    worksheet.mergeCells(`A${sectionRow.number}:J${sectionRow.number}`);
    sectionRow.height = 20;
    sectionRow.getCell(1).font = { name: 'Calibri', size: 12, bold: true, color: { argb: 'FF166534' } };
    sectionRow.getCell(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE8F5E9' } };
    sectionRow.getCell(1).alignment = { horizontal: 'left', vertical: 'middle' };
    sectionRow.getCell(1).border = {
        top: { style: 'thin', color: { argb: 'FFD1D5DB' } },
        left: { style: 'thin', color: { argb: 'FFD1D5DB' } },
        bottom: { style: 'thin', color: { argb: 'FFD1D5DB' } },
        right: { style: 'thin', color: { argb: 'FFD1D5DB' } }
    };

    const headerRow = worksheet.addRow(headers);
    headerRow.height = 22;
    applyTrackingHeaderStyle(headerRow);

    rows.forEach((row, index) => {
        const dataRow = worksheet.addRow(headers.map((header) => row[header] || 'N/A'));
        dataRow.height = 20;
        applyTrackingBodyStyle(dataRow, index % 2 === 0 ? 'FFFFFFFF' : 'FFF8FAFC');
        styleTrackingStageCell(dataRow.getCell(10), row['Current Processing Stage']);
        styleTrackingStageCell(dataRow.getCell(9), row['Payment Handling Status']);
    });

    worksheet.autoFilter = {
        from: { row: headerRow.number, column: 1 },
        to: { row: headerRow.number, column: headers.length }
    };

    if (trackReportUnmatchedEntries.length) {
        const unmatchedSheet = workbook.addWorksheet('Unmatched Names');
        unmatchedSheet.columns = [
            { header: 'Submitted Name', key: 'inputName', width: 34 },
            { header: 'Closest Match Suggestion', key: 'suggestion', width: 34 },
            { header: 'Match Score', key: 'score', width: 16 }
        ];
        const unmatchedTitleRow = unmatchedSheet.addRow(['Unmatched Tracking Report Names']);
        unmatchedSheet.mergeCells(`A${unmatchedTitleRow.number}:C${unmatchedTitleRow.number}`);
        unmatchedTitleRow.getCell(1).font = { name: 'Calibri', size: 14, bold: true, color: { argb: 'FFFFFFFF' } };
        unmatchedTitleRow.getCell(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF7C2D12' } };
        unmatchedTitleRow.getCell(1).alignment = { horizontal: 'center', vertical: 'middle' };
        unmatchedTitleRow.height = 24;
        const unmatchedHeaderRow = unmatchedSheet.addRow(['Submitted Name', 'Closest Match Suggestion', 'Match Score']);
        applyTrackingHeaderStyle(unmatchedHeaderRow);
        trackReportUnmatchedEntries.forEach((entry, index) => {
            const dataRow = unmatchedSheet.addRow([
                entry.inputName || '',
                entry?.suggestion?.customerName || 'No suggestion',
                Number.isFinite(entry?.suggestion?.score) ? entry.suggestion.score : ''
            ]);
            applyTrackingBodyStyle(dataRow, index % 2 === 0 ? 'FFFFFFFF' : 'FFFFF7ED');
        });
    }

    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], {
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = fileName;
    link.click();
    URL.revokeObjectURL(url);
}

function downloadTrackReportTemplate() {
    if (!ensureXlsxLibrary()) return;
    const worksheetData = [
        ['Customer Tracking Upload Template'],
        ['Enter one customer name per row in Column A'],
        [],
        ['Customer Name'],
        ['Sample Customer Name 1'],
        ['Sample Customer Name 2']
    ];
    const worksheet = window.XLSX.utils.aoa_to_sheet(worksheetData);
    worksheet['!merges'] = [
        { s: { r: 0, c: 0 }, e: { r: 0, c: 2 } },
        { s: { r: 1, c: 0 }, e: { r: 1, c: 2 } }
    ];
    worksheet['!cols'] = [{ wch: 40 }];
    ['A1', 'A2', 'A4'].forEach((ref) => {
        if (!worksheet[ref]) return;
        worksheet[ref].s = ref === 'A4'
            ? {
                font: { bold: true, color: { rgb: 'FFFFFF' } },
                fill: { fgColor: { rgb: '1D4ED8' } },
                alignment: { horizontal: 'center' }
            }
            : {
                font: { bold: ref === 'A1', sz: ref === 'A1' ? 14 : 11, color: { rgb: ref === 'A1' ? 'FFFFFF' : '334155' } },
                fill: { fgColor: { rgb: ref === 'A1' ? '0F3B67' : 'E2E8F0' } },
                alignment: { horizontal: ref === 'A1' ? 'center' : 'left' }
            };
    });
    const workbook = window.XLSX.utils.book_new();
    window.XLSX.utils.book_append_sheet(workbook, worksheet, 'Template');
    window.XLSX.writeFile(workbook, 'customer-tracking-template.xlsx');
}

function downloadTrackReportWorkbook() {
    if (!trackReportPreviewRows.length) {
        showNotification('There is no report data to download yet.', 'warning');
        return;
    }
    downloadWorkbookFromRows('customer-tracking-report.xlsx', trackReportPreviewRows);
}

window.openTrackApplicationModal = (submissionId) => {
    const submission = allSubmissions.find((item) => item.id === submissionId);
    if (!submission || !trackApplicationModal) return;

    const currentStage = getApplicationCurrentStage(submission);
    const statusLabel = formatStatusLabel(submission.status || '-');
    const statusClass = getTrackStatusBadgeClass(submission.status);
    const uploaderName = getDisplayNameByEmail(submission.uploadedBy || '');
    const reviewerName = submission.assignedTo ? getDisplayNameByEmail(submission.assignedTo) : 'Unassigned';
    const rsaName = submission.assignedToRSA ? getDisplayNameByEmail(submission.assignedToRSA) : 'Unassigned';
    const paymentName = submission.assignedToPayment ? getDisplayNameByEmail(submission.assignedToPayment) : 'Unassigned';
    const agentName = getSubmissionAgentLabel(submission);
    const lastStageTime = formatDate(getSubmissionCurrentStageEntryAt(submission));
    const timelineItems = [
        { key: 'upload', title: 'Uploaded', time: getTrackStageTimestamp(submission, 'upload'), meta: `Uploader: ${uploaderName}` },
        { key: 'reviewer', title: 'Reviewer', time: getTrackStageTimestamp(submission, 'reviewer'), meta: `Assigned Reviewer: ${reviewerName}` },
        { key: 'rsa', title: 'RSA', time: getTrackStageTimestamp(submission, 'rsa'), meta: `Assigned RSA: ${rsaName}` },
        { key: 'payment', title: 'Payment', time: getTrackStageTimestamp(submission, 'payment'), meta: `Assigned Payment: ${paymentName}` }
    ];

    if (trackApplicationCustomerName) trackApplicationCustomerName.textContent = submission.customerName || 'Unknown Application';
    if (trackApplicationMeta) {
        trackApplicationMeta.textContent = `Application ID: ${submission.id || '-'} | Last stage update: ${lastStageTime}`;
    }
    if (trackApplicationStatusBadges) {
        trackApplicationStatusBadges.innerHTML = `
            <span class="status-badge ${statusClass}">${escapeHtml(statusLabel)}</span>
            <span class="track-stage-pill current">${escapeHtml(currentStage.label)}</span>
        `;
    }
    if (trackApplicationSummary) {
        trackApplicationSummary.innerHTML = [
            renderTrackSummaryCard('Uploaded By', uploaderName),
            renderTrackSummaryCard('Agent', agentName),
            renderTrackSummaryCard('Current Stage', currentStage.label),
            renderTrackSummaryCard('Current Stage Time', lastStageTime),
            renderTrackSummaryCard('Assigned Reviewer', reviewerName),
            renderTrackSummaryCard('Assigned RSA', rsaName),
            renderTrackSummaryCard('Assigned Payment', paymentName),
            renderTrackSummaryCard('Cleared Time', formatDate(getSubmissionClearedEntryAt(submission)))
        ].join('');
    }
    if (trackApplicationTimeline) {
        trackApplicationTimeline.innerHTML = timelineItems.map((item) => {
            const state = getTrackTimelineState(submission, item.key);
            return `
                <div class="track-modal-timeline-card">
                    <div class="track-modal-timeline-head">
                        <div>
                            <span class="label">${escapeHtml(item.title)} Time</span>
                            <h4>${escapeHtml(item.title)}</h4>
                        </div>
                        <span class="track-stage-pill ${state.className}">${escapeHtml(state.label)}</span>
                    </div>
                    <div class="time">${escapeHtml(formatDate(item.time))}</div>
                    <div class="meta">${escapeHtml(item.meta)}</div>
                </div>
            `;
        }).join('');
    }

    trackApplicationModal.classList.add('active');
};

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
        list = list.filter((s) => matchesTrackStatusFilter(s, statusFilter));
    }

    if (start || end) {
        list = list.filter((s) => {
            const entryMs = getStageTimestampMillis(getSubmissionCurrentStageEntryAt(s));
            if (!entryMs) return false;
            if (start && entryMs < start.getTime()) return false;
            if (end && entryMs > end.getTime()) return false;
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
    list = list.slice().sort((a, b) => getStageTimestampMillis(getSubmissionCurrentStageEntryAt(b)) - getStageTimestampMillis(getSubmissionCurrentStageEntryAt(a)));

    const totalRecords = list.length;
    const totalPages = Math.max(1, Math.ceil(totalRecords / TRACK_PAGE_SIZE));
    trackAppsPage = Math.min(Math.max(1, trackAppsPage), totalPages);

    if (trackPageInfo) trackPageInfo.textContent = `Page ${trackAppsPage} of ${totalPages}`;
    if (trackPrevPageBtn) trackPrevPageBtn.disabled = trackAppsPage <= 1;
    if (trackNextPageBtn) trackNextPageBtn.disabled = trackAppsPage >= totalPages;
    if (trackJumpPageInput) trackJumpPageInput.max = String(totalPages);

    if (list.length === 0) {
        trackAppsTableBody.innerHTML = '<tr><td colspan="7" class="no-data">No matching applications</td></tr>';
        return;
    }

    const startIdx = (trackAppsPage - 1) * TRACK_PAGE_SIZE;
    const pageItems = list.slice(startIdx, startIdx + TRACK_PAGE_SIZE);

    trackAppsTableBody.innerHTML = pageItems.map((s) => {
        const uploaderEmail = String(s.uploadedBy || '').toLowerCase();
        const uploaderLabel = uploaderEmail ? `${getDisplayNameByEmail(uploaderEmail)}` : 'Unknown';
        const currentStage = getApplicationCurrentStage(s);
        const lastStageTime = formatDate(getSubmissionCurrentStageEntryAt(s));
        const status = String(s.status || '').trim() || '-';
        const agentLabel = getSubmissionAgentLabel(s);

        return `
            <tr>
                <td><strong>${escapeHtml(s.customerName || 'Unknown')}</strong></td>
                <td>${escapeHtml(uploaderLabel)}</td>
                <td>${escapeHtml(agentLabel)}</td>
                <td><span class="status-badge ${getTrackStatusBadgeClass(status)}">${escapeHtml(formatStatusLabel(status))}</span></td>
                <td>${escapeHtml(currentStage.label)}</td>
                <td>${escapeHtml(lastStageTime)}</td>
                <td>
                    <button class="action-btn view-btn-small" onclick="window.openTrackApplicationModal('${s.id}')">
                        <i class="fas fa-route"></i> Track
                    </button>
                </td>
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
    const onlineFilter = document.getElementById('userOnlineFilter')?.value || 'all';

    if (!allUsers) return;

    const filtered = allUsers.filter(user => {
        const fullName = user.fullName || user.email?.split('@')[0] || '';
        const normalizedRole = normalizeUserRole(user.role);
        const presence = user.isOnline === true ? 'online' : 'offline';
        const matchesSearch = searchTerm === '' ||
            fullName.toLowerCase().includes(searchTerm) ||
            user.email?.toLowerCase().includes(searchTerm);
        const matchesRole = roleFilter === 'all' || normalizedRole === roleFilter;
        const matchesStatus = statusFilter === 'all' || user.status === statusFilter;
        const matchesOnline = onlineFilter === 'all' || presence === onlineFilter;

        return matchesSearch && matchesRole && matchesStatus && matchesOnline;
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
        const relevantDate = getSubmissionDraftEntryAt(sub);
        const matchesDate = matchesExactDate(relevantDate, dateFilter);
        return matchesSearch && matchesDate;
    }).sort((a, b) => getStageTimestampMillis(getSubmissionDraftEntryAt(b)) - getStageTimestampMillis(getSubmissionDraftEntryAt(a)));

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
        const relevantDate = getSubmissionReviewEntryAt(sub);
        const matchesDate = matchesExactDate(relevantDate, dateFilter);
        return matchesSearch && matchesDate;
    }).sort((a, b) => getStageTimestampMillis(getSubmissionReviewEntryAt(b)) - getStageTimestampMillis(getSubmissionReviewEntryAt(a)));

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
        const matchesDate = matchesExactDate(getSubmissionApprovalEntryAt(sub), dateFilter);
        return matchesSearch && matchesDate;
    }).sort((a, b) => getStageTimestampMillis(getSubmissionApprovalEntryAt(b)) - getStageTimestampMillis(getSubmissionApprovalEntryAt(a)));

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
        const matchesDate = matchesExactDate(getSubmissionRejectionEntryAt(sub), dateFilter);
        return matchesSearch && matchesDate;
    }).sort((a, b) => getStageTimestampMillis(getSubmissionRejectionEntryAt(b)) - getStageTimestampMillis(getSubmissionRejectionEntryAt(a)));

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
if (typeof window.signOutUser !== 'function') {
    window.signOutUser = () => {
        window.location.href = 'index.html';
    };
}

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
