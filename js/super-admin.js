import { auth, db } from './firebase-config.js';
import { formatAppDateTime } from './shared/app-time.js';
import { signOut } from "https://www.gstatic.com/firebasejs/9.22.0/firebase-auth.js";
import {
    collection,
    addDoc,
    doc,
    getDoc,
    getDocs,
    onSnapshot,
    orderBy,
    limit,
    query,
    serverTimestamp,
    setDoc,
    updateDoc,
    where
} from "https://www.gstatic.com/firebasejs/9.22.0/firebase-firestore.js";
import {
    clearCommissionSettingsCache,
    formatCommissionRateLabel,
    getDefaultCommissionSettings,
    resolveSubmissionCommissionRate
} from './shared/commission-config.js?v=20260507a';
import { clearSystemSettingsCache, getDefaultSystemSettings, getSystemSettings, normalizeAgentBankOptions } from './shared/system-settings.js?v=20260508a';
import { getCurrentUserProfile as getCurrentUserProfileShared } from './shared/user-directory.js?v=20260518a';

let currentUser = null;
let currentUserData = null;
let allUsers = [];
let allSubmissions = [];
let allAudits = [];
let allRoutingRules = [];
let allAgents = [];
let currentTab = 'global';
let currentAgentSubTab = 'application-agents';
let currentRoutingSubTab = 'normal';
let selectedSuperUserId = '';
let selectedSuperAgentId = '';
let currentSettingsSubTab = 'system';
let currentAgentBankOptions = [];
let currentPfaOptions = [];
let currentDocumentRequirements = [];
let currentDocumentRequirementRoles = {};
let currentWorkflowLabels = {};
let currentPropertyRules = [];
let currentRolePermissions = {};
let currentRoutingPolicies = {};
let currentNotificationTemplates = {};
let currentHouseNumberRules = [];
let settingsCardsInitialized = false;
let activeSettingsModalPanelId = '';
let currentAccessUserSearch = '';
let activeSettingsDropdownTab = '';

const ROLE_PERMISSION_FIELDS = [
    { key: 'uploaderCanUpload', label: 'Uploader Can Upload' },
    { key: 'reviewerCanApprove', label: 'Reviewer Can Approve' },
    { key: 'reviewerCanReject', label: 'Reviewer Can Reject' },
    { key: 'rsaCanApprove', label: 'RSA Can Approve' },
    { key: 'rsaCanReject', label: 'RSA Can Reject' },
    { key: 'superAdminCanEditAgentRecords', label: 'Super Admin Can Edit Agent Records' },
    { key: 'superAdminCanClearCache', label: 'Super Admin Can Clear Cache' }
];

const DOCUMENT_REQUIREMENT_ROLE_FIELDS = [
    { key: 'uploader_level_1', label: 'Uploader Level 1' },
    { key: 'uploader_level_2', label: 'Uploader Level 2' },
    { key: 'reviewer', label: 'Reviewer' },
    { key: 'rsa_level_1', label: 'RSA Level 1' },
    { key: 'rsa_level_2', label: 'RSA Level 2' },
    { key: 'admin', label: 'Admin' },
    { key: 'super_admin', label: 'Super Admin' },
    { key: 'payment', label: 'Payment' }
];

const NOTIFICATION_TEMPLATE_FIELDS = [
    { key: 'submissionReceived', label: 'Submission Received' },
    { key: 'reviewerApproved', label: 'Reviewer Approved' },
    { key: 'reviewerRejected', label: 'Reviewer Rejected' },
    { key: 'rsaRejected', label: 'RSA Rejected' },
    { key: 'sentToPfa', label: 'Sent To PFA' },
    { key: 'paid', label: 'Paid' }
];

const SETTINGS_SUBTAB_META = {
    system: {
        title: 'General Settings',
        description: 'Open a section from General to manage system availability, announcements, and upload limits in a focused modal workspace.'
    },
    workflow: {
        title: 'Workflow Settings',
        description: 'Open a workflow section to manage round robin, routing, review rules, and workflow wording.'
    },
    commission: {
        title: 'Data Rules Settings',
        description: 'Open a data rules section to manage PFAs, banks, document boxes, pricing, import columns, and house number setup.'
    },
    notifications: {
        title: 'Notifications Settings',
        description: 'Open a notifications section to control alert channels and edit notification copy without crowding the page.'
    },
    access: {
        title: 'Access Settings',
        description: 'Open an access section to manage user levels, RSA round robin inclusion, and workflow permissions.'
    },
    security: {
        title: 'Security Settings',
        description: 'Open a security section to manage session timing, audit retention, cache refresh, and forced sign-outs.'
    }
};

const SETTINGS_DROPDOWN_TABS = [
    { id: 'system', label: 'General', icon: 'fa-sliders' },
    { id: 'workflow', label: 'Workflow', icon: 'fa-gears' },
    { id: 'commission', label: 'Data Rules', icon: 'fa-database' },
    { id: 'notifications', label: 'Notifications', icon: 'fa-bell' },
    { id: 'access', label: 'Access', icon: 'fa-key' },
    { id: 'security', label: 'Security', icon: 'fa-lock' }
];

const SETTINGS_CARD_ICON_MAP = {
    'General Controls': 'fa-sliders',
    'Upload Limits': 'fa-cloud-arrow-up',
    'Bank Management': 'fa-building-columns',
    'PFA List': 'fa-list-check',
    'Round Robin Controls': 'fa-rotate',
    'Review Rules': 'fa-list-check',
    'Document Requirements Manager': 'fa-folder-tree',
    'Bulk Import Required Columns': 'fa-file-arrow-up',
    'Routing Policy Settings': 'fa-route',
    'Status / Workflow Labels': 'fa-tag',
    'Commission and Pricing': 'fa-percent',
    'Property Type Amount Range': 'fa-chart-column',
    'House Number Rules': 'fa-house-circle-check',
    'Messages and Alerts': 'fa-bell',
    'Notification Templates': 'fa-envelope-open-text',
    'Access Levels': 'fa-user-shield',
    'RSA Round Robin Control': 'fa-user-check',
    'Role Permissions': 'fa-key',
    'Security and Audit': 'fa-lock',
    'Cache and Session Control': 'fa-shield-halved'
};

const superAdminName = document.getElementById('superAdminName');
const pageTitle = document.getElementById('pageTitle');
const notification = document.getElementById('notification');
const settingsSectionModal = document.getElementById('settingsSectionModal');
const settingsSectionModalTitle = document.getElementById('settingsSectionModalTitle');
const settingsSectionModalHeroTitle = document.getElementById('settingsSectionModalHeroTitle');
const settingsSectionModalDescription = document.getElementById('settingsSectionModalDescription');
const settingsSectionModalIcon = document.getElementById('settingsSectionModalIcon');
const settingsSectionModalContentHost = document.getElementById('settingsSectionModalContentHost');
const settingsModalStorage = document.getElementById('settingsModalStorage');
const settingsDropdownNav = document.getElementById('settingsDropdownNav');

function setCountBadge(id, count) {
    const badge = document.getElementById(id);
    if (!badge) return;
    badge.textContent = String(count);
    badge.style.display = 'inline-block';
}

function updateNavigationCounts() {
    const securityUsers = allUsers.filter((u) => {
        const role = String(u.role || '').toLowerCase();
        const status = String(u.status || 'active').toLowerCase();
        return ['super_admin', 'admin', 'reviewer', 'rsa', 'payment'].includes(role) && status !== 'active';
    }).length;
    setCountBadge('globalNavCount', allSubmissions.length);
    setCountBadge('adminsNavCount', allUsers.length);
    setCountBadge('routingRulesNavCount', allRoutingRules.length);
    setCountBadge('auditNavCount', allAudits.length);
    setCountBadge('securityNavCount', securityUsers);
}

function showNotification(message, type = 'info') {
    if (!notification) return;
    const iconMap = {
        success: 'fa-circle-check',
        error: 'fa-circle-xmark',
        warning: 'fa-triangle-exclamation',
        info: 'fa-circle-info'
    };
    const titleMap = {
        success: 'Success',
        error: 'Action Failed',
        warning: 'Attention',
        info: 'Notice'
    };
    const safeType = ['success', 'error', 'warning', 'info'].includes(type) ? type : 'info';
    notification.innerHTML = `
        <div class="notification-icon"><i class="fas ${iconMap[safeType]}"></i></div>
        <div class="notification-copy">
            <strong>${titleMap[safeType]}</strong>
            <span>${escapeHtml(message)}</span>
        </div>
    `;
    notification.className = `notification ${safeType}`;
    notification.style.display = 'block';
    requestAnimationFrame(() => notification.classList.add('show'));
    clearTimeout(showNotification._timer);
    showNotification._timer = setTimeout(() => {
        notification.classList.remove('show');
        setTimeout(() => { notification.style.display = 'none'; }, 220);
    }, 3200);
}

function normalizeAccessUserSearch(value = '') {
    return String(value || '').trim().toLowerCase();
}

function updateAccessSearchInputs() {
    const normalized = currentAccessUserSearch;
    const accessInput = document.getElementById('accessUserSearchInput');
    const rsaInput = document.getElementById('rsaRoundRobinSearchInput');
    if (accessInput && accessInput.value !== normalized) accessInput.value = normalized;
    if (rsaInput && rsaInput.value !== normalized) rsaInput.value = normalized;
}

function getAccessSearchMatch(user = {}) {
    const query = currentAccessUserSearch;
    if (!query) return true;
    const haystacks = [
        user.fullName,
        user.email,
        user.role,
        `level ${getUserRoleLevel(user)}`
    ].map((item) => String(item || '').toLowerCase());
    return haystacks.some((item) => item.includes(query));
}

function escapeHtml(text) {
    if (!text) return '';
    const div = document.createElement('div');
    div.textContent = String(text);
    return div.innerHTML;
}

function normalizeEmail(email) {
    return String(email || '').trim().toLowerCase();
}

function prettyJson(value) {
    try {
        return JSON.stringify(value, null, 2);
    } catch (_) {
        return '';
    }
}

function parseJsonTextarea(id, fallback) {
    const raw = String(document.getElementById(id)?.value || '').trim();
    if (!raw) return fallback;
    try {
        return JSON.parse(raw);
    } catch (_) {
        throw new Error(`Invalid JSON in ${id}`);
    }
}

function parseLinesTextarea(id) {
    return String(document.getElementById(id)?.value || '')
        .split(/\r?\n/)
        .map((line) => line.trim())
        .filter(Boolean);
}

function slugifyDocumentRequirementId(value) {
    return String(value || '')
        .trim()
        .toLowerCase()
        .replace(/[^a-z0-9]+/g, '_')
        .replace(/^_+|_+$/g, '')
        .slice(0, 80);
}

function formatSettingLabel(value) {
    return String(value || '')
        .replaceAll('_', ' ')
        .replace(/\b\w/g, (char) => char.toUpperCase());
}

function normalizeHouseNumberRulesForUi(value) {
    const source = value && typeof value === 'object' && !Array.isArray(value) ? value : {};
    return Object.entries(source).map(([propertyType, rule]) => ({
        propertyType,
        mode: String(rule?.mode || 'alpha_suffix').trim() || 'alpha_suffix',
        prefix: String(rule?.prefix || '').trim(),
        startNumber: Number(rule?.startNumber || 0),
        startLetter: String(rule?.startLetter || '').trim(),
        startPrefix: String(rule?.startPrefix || '').trim()
    }));
}

function getRoleLabel(role) {
    const normalized = String(role || 'uploader').toLowerCase();
    if (normalized === 'super_admin') return 'Super Admin';
    if (normalized === 'admin') return 'Admin';
    if (normalized === 'reviewer' || normalized === 'viewer') return 'Reviewer';
    if (normalized === 'rsa') return 'RSA';
    if (normalized === 'payment') return 'Payment';
    return 'Uploader';
}

function getUserWhatsApp(user = {}) {
    return String(user.whatsappNumber || user.whatsapp || user.phone || user.phoneNumber || '').trim();
}

function splitWhatsApp(user = {}) {
    const code = String(user.whatsappCode || '').trim();
    const local = String(user.whatsappLocalNumber || '').replace(/\D/g, '');
    if (code || local) return { code, local };

    const whatsapp = getUserWhatsApp(user);
    if (!whatsapp) return { code: '+234', local: '' };
    const digits = whatsapp.replace(/\D/g, '');
    return {
        code: whatsapp.startsWith('+') && digits.length > 10 ? `+${digits.slice(0, digits.length - 10)}` : '+234',
        local: digits.slice(-10)
    };
}

async function ensureCurrentUserProfileAtUid(user, profileDocSnap) {
    if (!user?.uid || !profileDocSnap?.exists()) return profileDocSnap;
    if (profileDocSnap.id === user.uid) return profileDocSnap;

    const profileData = profileDocSnap.data() || {};
    const normalizedEmail = normalizeEmail(user.email);
    const profileEmail = normalizeEmail(profileData.email);

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

async function ensureCurrentSuperAdminWritableProfile() {
    if (!currentUser?.uid || !currentUser?.email) return null;

    const uidRef = doc(db, 'users', currentUser.uid);
    const uidSnap = await getDoc(uidRef);
    if (uidSnap.exists()) return uidSnap;

    const normalizedEmail = normalizeEmail(currentUser.email);
    const normalizedRole = String(currentUserData?.role || 'super_admin').trim().toLowerCase();
    const normalizedStatus = String(currentUserData?.status || 'active').trim().toLowerCase();
    const allowedRoles = new Set(['uploader', 'admin', 'super_admin', 'reviewer', 'rsa', 'payment']);
    const allowedStatuses = new Set(['pending', 'active', 'deactivated']);

    await setDoc(uidRef, {
        ...(currentUserData || {}),
        uid: currentUser.uid,
        email: normalizedEmail,
        role: allowedRoles.has(normalizedRole) ? normalizedRole : 'super_admin',
        status: allowedStatuses.has(normalizedStatus) ? normalizedStatus : 'active',
        updatedAt: serverTimestamp()
    }, { merge: true });

    return await getDoc(uidRef);
}

function formatDate(ts) {
    return formatAppDateTime(ts, '-');
}

function roleHome(role) {
    const r = String(role || '').toLowerCase();
    if (r === 'super_admin') return 'super-admin-dashboard.html';
    if (r === 'admin') return 'admin-dashboard.html';
    if (r === 'reviewer') return 'reviewer-dashboard.html';
    if (r === 'rsa') return 'rsa-dashboard.html';
    if (r === 'payment') return 'payment-dashboard.html';
    return 'dashboard.html';
}

function switchTab(tabId) {
    currentTab = tabId;
    if (tabId !== 'settings') closeSettingsSectionModal();
    document.querySelectorAll('.nav-item').forEach((n) => n.classList.remove('active'));
    document.querySelector(`[data-tab="${tabId}"]`)?.classList.add('active');
    document.querySelectorAll('.tab-content').forEach((t) => t.classList.remove('active'));
    document.getElementById(`${tabId}Tab`)?.classList.add('active');

    const titles = {
        global: 'Global View',
        admins: 'User Management',
        'routing-rules': 'Uploader Routing',
        agents: 'Agent',
        audit: 'Audit',
        security: 'Security',
        'round-robin': 'Round Robin',
        settings: 'Settings',
        help: 'Help & SOP'
    };
    if (pageTitle) pageTitle.textContent = titles[tabId] || 'Super Admin';
    renderCurrentTab();
}

function renderAgentSubTabState() {
    const isApplication = currentAgentSubTab === 'application-agents';
    const isRecords = currentAgentSubTab === 'agent-records';
    const isReroute = currentAgentSubTab === 'agent-reroute';

    const setButtonState = (id, active) => {
        const button = document.getElementById(id);
        if (!button) return;
        button.style.background = active ? '#003366' : '';
        button.style.color = active ? '#fff' : '';
        button.style.border = active ? 'none' : '';
    };

    const setSectionState = (id, active) => {
        const section = document.getElementById(id);
        if (section) section.style.display = active ? '' : 'none';
    };

    setButtonState('agentApplicationSubTabBtn', isApplication);
    setButtonState('agentRecordsSubTabBtn', isRecords);
    setButtonState('agentRerouteSubTabBtn', isReroute);
    setSectionState('agentApplicationSection', isApplication);
    setSectionState('agentRecordsSection', isRecords);
    setSectionState('agentRerouteSection', isReroute);
}

window.switchAgentSubTab = (tabId) => {
    const allowedTabs = ['application-agents', 'agent-records', 'agent-reroute'];
    currentAgentSubTab = allowedTabs.includes(String(tabId || '').trim()) ? String(tabId).trim() : 'application-agents';
    renderAgentSubTabState();
    renderCurrentTab();
};

function renderGlobalView() {
    const usersCount = allUsers.length;
    const submissionsCount = allSubmissions.length;
    const paidCount = allSubmissions.filter((s) => String(s.status || '').toLowerCase() === 'paid').length;
    const pendingCount = allSubmissions.filter((s) => String(s.status || '').toLowerCase() === 'pending').length;

    const elUsers = document.getElementById('globalUsersCount');
    const elSubs = document.getElementById('globalSubmissionsCount');
    const elPaid = document.getElementById('globalPaidCount');
    const elPending = document.getElementById('globalPendingCount');
    if (elUsers) elUsers.textContent = String(usersCount);
    if (elSubs) elSubs.textContent = String(submissionsCount);
    if (elPaid) elPaid.textContent = String(paidCount);
    if (elPending) elPending.textContent = String(pendingCount);

    const statuses = ['pending', 'processing_to_pfa', 'approved', 'rejected', 'sent_to_pfa', 'paid', 'cleared'];
    const pipelineBody = document.getElementById('pipelineTableBody');
    if (!pipelineBody) return;
    pipelineBody.innerHTML = statuses.map((status) => {
        const count = allSubmissions.filter((s) => String(s.status || '').toLowerCase() === status).length;
        const label = status.replaceAll('_', ' ').replace(/\b\w/g, (c) => c.toUpperCase());
        return `<tr><td>${label}</td><td><strong>${count}</strong></td></tr>`;
    }).join('');
}

function renderAdminManagement() {
    const body = document.getElementById('adminManagersTableBody');
    if (!body) return;

    const search = String(document.getElementById('superUserSearch')?.value || '').trim().toLowerCase();
    const roleFilter = String(document.getElementById('superUserRoleFilter')?.value || '').trim().toLowerCase();
    const statusFilter = String(document.getElementById('superUserStatusFilter')?.value || '').trim().toLowerCase();

    const filteredUsers = allUsers
        .filter((u) => {
            const role = String(u.role || 'uploader').toLowerCase();
            const status = String(u.status || 'active').toLowerCase();
            const whatsapp = getUserWhatsApp(u);
            const searchable = [
                u.fullName || '',
                u.displayName || '',
                u.email || '',
                whatsapp,
                getRoleLabel(role),
                status
            ].join(' ').toLowerCase();
            return (!search || searchable.includes(search))
                && (!roleFilter || role === roleFilter)
                && (!statusFilter || status === statusFilter);
        })
        .sort((a, b) => String(a.fullName || a.email || '').localeCompare(String(b.fullName || b.email || '')));

    if (!filteredUsers.length) {
        body.innerHTML = '<tr><td colspan="7" class="no-data">No users found</td></tr>';
        return;
    }

    body.innerHTML = filteredUsers.map((u) => {
        const storedRole = String(u.role || 'uploader').toLowerCase();
        const role = storedRole === 'viewer' ? 'reviewer' : storedRole;
        const status = String(u.status || 'active').toLowerCase();
        const whatsapp = getUserWhatsApp(u);
        return `
            <tr data-search="${escapeHtml([u.fullName || '', u.email || '', whatsapp, role, status].join(' '))}">
                <td><strong>${escapeHtml(u.fullName || u.email || 'Unknown')}</strong></td>
                <td>${escapeHtml(u.email || '-')}</td>
                <td>${escapeHtml(whatsapp || '-')}</td>
                <td>
                    <select id="sa-role-${u.id}" style="padding:8px;border:1px solid #e2e8f0;border-radius:8px;">
                        <option value="uploader" ${role === 'uploader' ? 'selected' : ''}>Uploader</option>
                        <option value="reviewer" ${role === 'reviewer' ? 'selected' : ''}>Reviewer</option>
                        <option value="rsa" ${role === 'rsa' ? 'selected' : ''}>RSA</option>
                        <option value="payment" ${role === 'payment' ? 'selected' : ''}>Payment</option>
                        <option value="admin" ${role === 'admin' ? 'selected' : ''}>Admin</option>
                        <option value="super_admin" ${role === 'super_admin' ? 'selected' : ''}>Super Admin</option>
                    </select>
                </td>
                <td>
                    <select id="sa-status-${u.id}" style="padding:8px;border:1px solid #e2e8f0;border-radius:8px;">
                        <option value="active" ${status === 'active' ? 'selected' : ''}>Active</option>
                        <option value="pending" ${status === 'pending' ? 'selected' : ''}>Pending</option>
                        <option value="deactivated" ${status === 'deactivated' ? 'selected' : ''}>Deactivated</option>
                    </select>
                </td>
                <td>${formatDate(u.createdAt)}</td>
                <td>
                    <button class="action-btn" style="background:#003366;color:#fff;border:none;" onclick="window.saveAdminManager('${u.id}')">
                        <i class="fas fa-save"></i> Save
                    </button>
                    <button class="action-btn edit-btn" onclick="window.editSuperUser('${u.id}')" title="Edit user">
                        <i class="fas fa-edit"></i> Edit
                    </button>
                </td>
            </tr>
        `;
    }).join('');
}

function getUsersByRole(role) {
    const r = String(role || '').trim().toLowerCase();
    return allUsers.filter((u) => {
        const userRole = String(u.role || '').trim().toLowerCase();
        const status = String(u.status || 'active').trim().toLowerCase();
        const leaveStatus = String(u.leaveStatus || '').trim().toLowerCase();
        return userRole === r && status !== 'deactivated' && leaveStatus !== 'on_leave';
    });
}

function getRoutingUploaderUsers() {
    const uploadCapableRoles = new Set(['uploader', 'reviewer', 'rsa']);
    return allUsers
        .filter((u) => {
            const role = String(u.role || '').trim().toLowerCase();
            const status = String(u.status || 'active').trim().toLowerCase();
            const leaveStatus = String(u.leaveStatus || '').trim().toLowerCase();
            return uploadCapableRoles.has(role) && status !== 'deactivated' && leaveStatus !== 'on_leave' && normalizeEmail(u.email);
        })
        .sort((a, b) => normalizeEmail(a.email).localeCompare(normalizeEmail(b.email)));
}

function getActiveUsersByRoles(roles = []) {
    const set = new Set(roles.map((r) => String(r || '').trim().toLowerCase()));
    return allUsers.filter((u) => {
        const role = String(u.role || '').trim().toLowerCase();
        const status = String(u.status || 'active').trim().toLowerCase();
        const leaveStatus = String(u.leaveStatus || '').trim().toLowerCase();
        return set.has(role) && status !== 'deactivated' && leaveStatus !== 'on_leave';
    });
}

function findRoutingRuleForUploader(uploaderEmail) {
    const normalized = normalizeEmail(uploaderEmail);
    return allRoutingRules.find((r) => normalizeEmail(r.uploaderEmail) === normalized) || null;
}

function getRoutingRuleMode(rule = {}) {
    return String(rule?.routeMode || 'normal').toLowerCase() === 'skip_reviewer' ? 'skip_reviewer' : 'normal';
}

function optionsForUsers(users, currentValue) {
    const current = String(currentValue || '').toLowerCase();
    const opts = ['<option value="">Unassigned</option>'];
    for (const user of users) {
        const email = String(user.email || '').toLowerCase();
        const label = `${user.fullName || email} (${email})`;
        opts.push(`<option value="${escapeHtml(email)}" ${email === current ? 'selected' : ''}>${escapeHtml(label)}</option>`);
    }
    return opts.join('');
}

function getUserSearchText(email) {
    const normalized = normalizeEmail(email);
    if (!normalized) return '';
    const user = allUsers.find((u) => normalizeEmail(u.email) === normalized);
    return [normalized, user?.fullName || '', user?.displayName || '', user?.role || ''].join(' ');
}

function getRoutingRulesSearchTerm() {
    return String(document.getElementById('routingRulesSearch')?.value || '').trim().toLowerCase();
}

function renderRedirectTable() {
    const body = document.getElementById('redirectTableBody');
    if (!body) return;
    const reviewerUsers = getUsersByRole('reviewer');
    const rsaUsers = getUsersByRole('rsa');
    const paymentUsers = getUsersByRole('payment');

    const rows = allSubmissions.slice(0, 200);
    if (!rows.length) {
        body.innerHTML = '<tr><td colspan="7" class="no-data">No submissions found</td></tr>';
        return;
    }

    body.innerHTML = rows.map((s) => `
        <tr>
            <td><strong>${escapeHtml(s.customerName || 'Unknown')}</strong></td>
            <td>${escapeHtml(String(s.status || '-'))}</td>
            <td><select id="rd-reviewer-${s.id}" style="padding:8px;border:1px solid #e2e8f0;border-radius:8px;width:100%;">${optionsForUsers(reviewerUsers, s.assignedTo)}</select></td>
            <td><select id="rd-rsa-${s.id}" style="padding:8px;border:1px solid #e2e8f0;border-radius:8px;width:100%;">${optionsForUsers(rsaUsers, s.assignedToRSA)}</select></td>
            <td><select id="rd-payment-${s.id}" style="padding:8px;border:1px solid #e2e8f0;border-radius:8px;width:100%;">${optionsForUsers(paymentUsers, s.assignedToPayment)}</select></td>
            <td><input id="rd-reason-${s.id}" placeholder="Reason (required)" style="padding:8px;border:1px solid #e2e8f0;border-radius:8px;width:100%;"></td>
            <td><button class="action-btn" style="background:#003366;color:#fff;border:none;" onclick="window.redirectSubmission('${s.id}')"><i class="fas fa-random"></i> Apply</button></td>
        </tr>
    `).join('');
}

function renderRoutingRules() {
    renderRoutingRulesSubTabState();
    if (currentRoutingSubTab === 'skip-reviewer') return renderSkipReviewerRoutingRules();

    const body = document.getElementById('routingRulesTableBody');
    if (!body) return;

    const uploaderUsers = getRoutingUploaderUsers();
    const reviewerUsers = getActiveUsersByRoles(['reviewer']);
    const rsaUsers = getUsersByRole('rsa');
    const paymentUsers = getUsersByRole('payment');
    const searchTerm = getRoutingRulesSearchTerm();

    if (!uploaderUsers.length) {
        body.innerHTML = '<tr><td colspan="7" class="no-data">No upload-capable users found</td></tr>';
        return;
    }

    const filteredUsers = uploaderUsers.filter((uploader) => {
        const uploaderEmail = normalizeEmail(uploader.email);
        const uploaderRole = String(uploader.role || 'uploader').toLowerCase();
        const rule = findRoutingRuleForUploader(uploaderEmail) || {};
        const searchableText = [
            uploader.fullName || '',
            uploader.displayName || '',
            uploaderEmail,
            uploaderRole,
            getUserSearchText(rule.reviewerEmail),
            getUserSearchText(rule.rsaEmail),
            getUserSearchText(rule.paymentEmail),
            formatDate(rule.updatedAt || rule.createdAt)
        ].join(' ').toLowerCase();
        return !searchTerm || searchableText.includes(searchTerm);
    });

    if (!filteredUsers.length) {
        body.innerHTML = '<tr><td colspan="7" class="no-data">No matching routing rules found</td></tr>';
        return;
    }

    body.innerHTML = filteredUsers.map((uploader) => {
        const uploaderEmail = normalizeEmail(uploader.email);
        const uploaderRole = String(uploader.role || 'uploader').toLowerCase();
        const rule = findRoutingRuleForUploader(uploaderEmail) || {};
        const selectedReviewer = normalizeEmail(rule.reviewerEmail);
        const selectedRsa = normalizeEmail(rule.rsaEmail);
        const selectedPayment = normalizeEmail(rule.paymentEmail);
        const updatedAt = formatDate(rule.updatedAt || rule.createdAt);
        const searchableText = [
            uploader.fullName || '',
            uploader.displayName || '',
            uploaderEmail,
            uploaderRole,
            getUserSearchText(selectedReviewer),
            getUserSearchText(selectedRsa),
            getUserSearchText(selectedPayment),
            updatedAt
        ].join(' ');

        return `
            <tr data-search="${escapeHtml(searchableText)}">
                <td>
                    <strong>${escapeHtml(uploader.fullName || uploaderEmail || 'Unknown')}</strong>
                    <div style="font-size:12px;color:#64748b;margin-top:4px;">Primary role: ${escapeHtml(uploaderRole)}</div>
                </td>
                <td>${escapeHtml(uploaderEmail || '-')}</td>
                <td>
                    <select id="route-reviewer-${uploader.id}" style="padding:8px;border:1px solid #e2e8f0;border-radius:8px;width:100%;">
                        ${optionsForUsers(reviewerUsers, selectedReviewer)}
                    </select>
                </td>
                <td>
                    <select id="route-rsa-${uploader.id}" style="padding:8px;border:1px solid #e2e8f0;border-radius:8px;width:100%;">
                        ${optionsForUsers(rsaUsers, selectedRsa)}
                    </select>
                </td>
                <td>
                    <select id="route-payment-${uploader.id}" style="padding:8px;border:1px solid #e2e8f0;border-radius:8px;width:100%;">
                        ${optionsForUsers(paymentUsers, selectedPayment)}
                    </select>
                </td>
                <td>${escapeHtml(updatedAt)}</td>
                <td>
                    <button id="route-save-${uploader.id}" class="action-btn" style="background:#003366;color:#fff;border:none;" onclick="window.saveUploaderRoutingRule('${uploader.id}')">
                        <i class="fas fa-save"></i> Save
                    </button>
                </td>
            </tr>
        `;
    }).join('');
}

function renderSkipReviewerRoutingRules() {
    const body = document.getElementById('routingRulesSkipReviewerTableBody');
    if (!body) return;

    const uploaderUsers = getRoutingUploaderUsers();
    const rsaUsers = getUsersByRole('rsa');
    const paymentUsers = getUsersByRole('payment');
    const searchTerm = getRoutingRulesSearchTerm();

    if (!uploaderUsers.length) {
        body.innerHTML = '<tr><td colspan="6" class="no-data">No upload-capable users found</td></tr>';
        return;
    }

    const filteredUsers = uploaderUsers.filter((uploader) => {
        const uploaderEmail = normalizeEmail(uploader.email);
        const rule = findRoutingRuleForUploader(uploaderEmail) || {};
        const searchableText = [
            uploader.fullName || '',
            uploader.displayName || '',
            uploaderEmail,
            getUserSearchText(rule.rsaEmail),
            getUserSearchText(rule.paymentEmail),
            getRoutingRuleMode(rule),
            formatDate(rule.updatedAt || rule.createdAt)
        ].join(' ').toLowerCase();
        return !searchTerm || searchableText.includes(searchTerm);
    });

    if (!filteredUsers.length) {
        body.innerHTML = '<tr><td colspan="6" class="no-data">No matching skip-reviewer routes found</td></tr>';
        return;
    }

    body.innerHTML = filteredUsers.map((uploader) => {
        const uploaderEmail = normalizeEmail(uploader.email);
        const rule = findRoutingRuleForUploader(uploaderEmail) || {};
        const selectedRsa = normalizeEmail(rule.rsaEmail);
        const selectedPayment = normalizeEmail(rule.paymentEmail);
        const updatedAt = formatDate(rule.updatedAt || rule.createdAt);
        const searchableText = [
            uploader.fullName || '',
            uploader.displayName || '',
            uploaderEmail,
            getUserSearchText(selectedRsa),
            getUserSearchText(selectedPayment),
            getRoutingRuleMode(rule),
            updatedAt
        ].join(' ');

        return `
            <tr data-search="${escapeHtml(searchableText)}">
                <td>
                    <strong>${escapeHtml(uploader.fullName || uploaderEmail || 'Unknown')}</strong>
                    <div style="font-size:12px;color:#64748b;margin-top:4px;">Mode: ${getRoutingRuleMode(rule) === 'skip_reviewer' ? 'Skip Reviewer' : 'Normal'}</div>
                </td>
                <td>${escapeHtml(uploaderEmail || '-')}</td>
                <td>
                    <select id="route-skip-rsa-${uploader.id}" style="padding:8px;border:1px solid #e2e8f0;border-radius:8px;width:100%;">
                        ${optionsForUsers(rsaUsers, selectedRsa)}
                    </select>
                </td>
                <td>
                    <select id="route-skip-payment-${uploader.id}" style="padding:8px;border:1px solid #e2e8f0;border-radius:8px;width:100%;">
                        ${optionsForUsers(paymentUsers, selectedPayment)}
                    </select>
                </td>
                <td>${escapeHtml(updatedAt)}</td>
                <td>
                    <button id="route-skip-save-${uploader.id}" class="action-btn" style="background:#003366;color:#fff;border:none;" onclick="window.saveSkipReviewerRoutingRule('${uploader.id}')">
                        <i class="fas fa-save"></i> Save
                    </button>
                </td>
            </tr>
        `;
    }).join('');
}

function renderRoutingRulesSubTabState() {
    const normalSection = document.getElementById('routingNormalSection');
    const skipSection = document.getElementById('routingSkipReviewerSection');
    const normalBtn = document.getElementById('routingNormalTabBtn');
    const skipBtn = document.getElementById('routingSkipReviewerTabBtn');
    const isSkip = currentRoutingSubTab === 'skip-reviewer';

    if (normalSection) normalSection.style.display = isSkip ? 'none' : '';
    if (skipSection) skipSection.style.display = isSkip ? '' : 'none';

    if (normalBtn) {
        normalBtn.style.background = isSkip ? '' : '#003366';
        normalBtn.style.color = isSkip ? '' : '#fff';
        normalBtn.style.border = isSkip ? '' : 'none';
    }
    if (skipBtn) {
        skipBtn.style.background = isSkip ? '#003366' : '';
        skipBtn.style.color = isSkip ? '#fff' : '';
        skipBtn.style.border = isSkip ? 'none' : '';
    }
}

async function renderAudit() {
    const body = document.getElementById('superAuditTableBody');
    if (!body) return;
    const systemSettings = await getSystemSettings(db, { force: true });
    const retentionDays = Number(systemSettings.auditControls?.retentionDays || 30);
    const cutoffMs = Date.now() - (Math.max(1, retentionDays) * 24 * 60 * 60 * 1000);
    const visibleAudits = allAudits.filter((entry) => {
        const raw = entry.timestamp || entry.createdAt || null;
        try {
            const ms = typeof raw?.toMillis === 'function' ? raw.toMillis() : (raw?.toDate ? raw.toDate().getTime() : new Date(raw).getTime());
            return Number.isFinite(ms) ? ms >= cutoffMs : true;
        } catch (_) {
            return true;
        }
    });
    if (!visibleAudits.length) {
        body.innerHTML = '<tr><td colspan="4" class="no-data">No audit records</td></tr>';
        return;
    }
    body.innerHTML = visibleAudits.slice(0, 200).map((a) => `
        <tr>
            <td>${formatDate(a.timestamp || a.createdAt)}</td>
            <td>${escapeHtml(a.action || '-')}</td>
            <td>${escapeHtml(a.customerName || a.submissionId || a.details || '-')}</td>
            <td>${escapeHtml(a.performedBy || a.userEmail || '-')}</td>
        </tr>
    `).join('');
}

function renderSecurity() {
    const privileged = allUsers.filter((u) => ['super_admin', 'admin', 'reviewer', 'rsa', 'payment'].includes(String(u.role || '').toLowerCase()));
    const activePrivileged = privileged.filter((u) => String(u.status || 'active').toLowerCase() === 'active').length;
    const deactivated = allUsers.filter((u) => String(u.status || '').toLowerCase() === 'deactivated').length;
    const c1 = document.getElementById('securityPrivilegedCount');
    const c2 = document.getElementById('securityDeactivatedCount');
    if (c1) c1.textContent = String(activePrivileged);
    if (c2) c2.textContent = String(deactivated);

    const body = document.getElementById('securityUsersBody');
    if (!body) return;
    body.innerHTML = privileged.map((u) => `
        <tr>
            <td>${escapeHtml(u.fullName || '-')}</td>
            <td>${escapeHtml(u.email || '-')}</td>
            <td>${escapeHtml(String(u.role || '-'))}</td>
            <td>${escapeHtml(String(u.status || 'active'))}</td>
        </tr>
    `).join('');
}

async function renderRoundRobin() {
    const body = document.getElementById('rrCountersBody');
    if (!body) return;
    const queues = [
        { label: 'Reviewer', docId: 'roundRobin' },
        { label: 'RSA', docId: 'roundRobinRSA' },
        { label: 'Payment', docId: 'roundRobinPayment' }
    ];
    const rows = [];
    for (const q of queues) {
        const snap = await getDoc(doc(db, 'counters', q.docId));
        const data = snap.exists() ? (snap.data() || {}) : {};
        rows.push(`
            <tr>
                <td>${q.label}</td>
                <td>${q.docId}</td>
                <td>${data.lastIndex ?? -1}</td>
                <td>${data.lastDate || 'Never'}</td>
                <td><button class="action-btn" style="background:#ef4444;color:#fff;border:none;" onclick="window.resetRoundRobinCounter('${q.docId}')"><i class="fas fa-redo"></i> Reset</button></td>
            </tr>
        `);
    }
    body.innerHTML = rows.join('');
}

async function loadSettings() {
    const snap = await getDoc(doc(db, 'settings', 'system'));
    const data = snap.exists() ? (snap.data() || {}) : {};
    const defaultCommissionSettings = getDefaultCommissionSettings();
    const defaultSystemSettings = getDefaultSystemSettings();
    const systemSettings = await getSystemSettings(db, { force: true });
    const maintenance = document.getElementById('settingMaintenanceMode');
    const maintenanceMessage = document.getElementById('settingMaintenanceMessage');
    const commissionRate = document.getElementById('settingCommissionRate');
    const commissionEffectiveFrom = document.getElementById('settingCommissionEffectiveFrom');
    const maxImageUploadMb = document.getElementById('settingMaxImageUploadMb');
    const maxPdfUploadMb = document.getElementById('settingMaxPdfUploadMb');
    const reviewerRoundRobinEnabled = document.getElementById('settingReviewerRoundRobinEnabled');
    const rsaRoundRobinEnabled = document.getElementById('settingRsaRoundRobinEnabled');
    const paymentRoundRobinEnabled = document.getElementById('settingPaymentRoundRobinEnabled');
    const agentEditSyncEnabled = document.getElementById('settingAgentEditSyncEnabled');
    const notificationsEmailEnabled = document.getElementById('settingNotificationsEmailEnabled');
    const notificationsPushEnabled = document.getElementById('settingNotificationsPushEnabled');
    const announcementEnabled = document.getElementById('settingAnnouncementEnabled');
    const announcementTone = document.getElementById('settingAnnouncementTone');
    const announcementMessage = document.getElementById('settingAnnouncementMessage');
    const fallbackAssignmentMode = document.getElementById('settingFallbackAssignmentMode');
    const rejectionMinLength = document.getElementById('settingRejectionMinLength');
    const reviewerRejectRequired = document.getElementById('settingReviewerRejectRequired');
    const rsaRejectRequired = document.getElementById('settingRsaRejectRequired');
    const agentRegistrationApprovalRequired = document.getElementById('settingAgentRegistrationApprovalRequired');
    const bulkImportRequiredColumns = document.getElementById('settingBulkImportRequiredColumns');
    const sessionTimeoutMinutes = document.getElementById('settingSessionTimeoutMinutes');
    const auditRetentionDays = document.getElementById('settingAuditRetentionDays');
    const defaultRouteMode = document.getElementById('settingDefaultRouteMode');
    if (maintenance) maintenance.value = data.maintenanceMode ? 'true' : 'false';
    if (maintenanceMessage) maintenanceMessage.value = String(data.maintenanceMessage || defaultSystemSettings.maintenanceMessage);
    if (commissionRate) commissionRate.value = String(Number(data.commissionRate ?? defaultCommissionSettings.rate) * 100);
    if (commissionEffectiveFrom) {
        const rawDate = String(data.commissionRateEffectiveFrom || defaultCommissionSettings.effectiveFromIso).trim();
        commissionEffectiveFrom.value = rawDate ? rawDate.slice(0, 10) : '2026-05-07';
    }
    if (maxImageUploadMb) maxImageUploadMb.value = String(Number(data.maxImageUploadMb ?? defaultSystemSettings.maxImageUploadMb));
    if (maxPdfUploadMb) maxPdfUploadMb.value = String(Number(data.maxPdfUploadMb ?? defaultSystemSettings.maxPdfUploadMb));
    if (reviewerRoundRobinEnabled) reviewerRoundRobinEnabled.value = String((data.reviewerRoundRobinEnabled ?? defaultSystemSettings.reviewerRoundRobinEnabled) ? 'true' : 'false');
    if (rsaRoundRobinEnabled) rsaRoundRobinEnabled.value = String((data.rsaRoundRobinEnabled ?? defaultSystemSettings.rsaRoundRobinEnabled) ? 'true' : 'false');
    if (paymentRoundRobinEnabled) paymentRoundRobinEnabled.value = String((data.paymentRoundRobinEnabled ?? defaultSystemSettings.paymentRoundRobinEnabled) ? 'true' : 'false');
    if (agentEditSyncEnabled) agentEditSyncEnabled.value = String((data.agentEditSyncEnabled ?? defaultSystemSettings.agentEditSyncEnabled) ? 'true' : 'false');
    if (notificationsEmailEnabled) notificationsEmailEnabled.value = String((data.notificationsEmailEnabled ?? defaultSystemSettings.notificationsEmailEnabled) ? 'true' : 'false');
    if (notificationsPushEnabled) notificationsPushEnabled.value = String((data.notificationsPushEnabled ?? defaultSystemSettings.notificationsPushEnabled) ? 'true' : 'false');
    if (announcementEnabled) announcementEnabled.value = String(systemSettings.dashboardAnnouncement.enabled ? 'true' : 'false');
    if (announcementTone) announcementTone.value = String(systemSettings.dashboardAnnouncement.tone || 'info');
    if (announcementMessage) announcementMessage.value = String(systemSettings.dashboardAnnouncement.message || '');
    currentPfaOptions = Array.isArray(systemSettings.pfaOptions) ? [...systemSettings.pfaOptions] : [];
    renderPfaManagement();
    currentDocumentRequirements = Array.isArray(systemSettings.documentRequirements) ? systemSettings.documentRequirements.map((doc) => ({ ...doc })) : [];
    renderDocumentRequirementsManager();
    currentDocumentRequirementRoles = { ...(systemSettings.documentRequirementRoles || defaultSystemSettings.documentRequirementRoles || {}) };
    renderDocumentRequirementRoleManager();
    currentRoutingPolicies = { ...(systemSettings.routingPolicies || {}) };
    if (fallbackAssignmentMode) fallbackAssignmentMode.value = String(systemSettings.routingPolicies?.fallbackAssignmentMode || defaultSystemSettings.routingPolicies.fallbackAssignmentMode || 'round_robin');
    currentWorkflowLabels = { ...(systemSettings.workflowLabels || {}) };
    renderWorkflowLabelsManager();
    if (rejectionMinLength) rejectionMinLength.value = String(systemSettings.rejectionRules.minLength ?? 10);
    if (reviewerRejectRequired) reviewerRejectRequired.value = String(systemSettings.rejectionRules.reviewerRequired ? 'true' : 'false');
    if (rsaRejectRequired) rsaRejectRequired.value = String(systemSettings.rejectionRules.rsaRequired ? 'true' : 'false');
    if (agentRegistrationApprovalRequired) agentRegistrationApprovalRequired.value = String(systemSettings.agentRegistrationRules.approvalRequired ? 'true' : 'false');
    if (bulkImportRequiredColumns) bulkImportRequiredColumns.value = (systemSettings.bulkImportRules.requiredColumns || []).join('\n');
    currentNotificationTemplates = { ...(systemSettings.notificationTemplates || {}) };
    renderNotificationTemplatesManager();
    if (sessionTimeoutMinutes) sessionTimeoutMinutes.value = String(systemSettings.securityControls.sessionTimeoutMinutes ?? 60);
    if (auditRetentionDays) auditRetentionDays.value = String(systemSettings.auditControls.retentionDays ?? 30);
    currentRolePermissions = { ...(systemSettings.rolePermissions || {}) };
    renderRolePermissionsManager();
    renderAccessLevelManager();
    currentPropertyRules = Array.isArray(systemSettings.propertyRules) ? systemSettings.propertyRules.map((rule) => ({ ...rule })) : [];
    renderPropertyRulesManager();
    currentHouseNumberRules = normalizeHouseNumberRulesForUi(systemSettings.houseNumberRules);
    renderHouseNumberRulesManager();
    if (defaultRouteMode) defaultRouteMode.value = String(systemSettings.routingPolicies.defaultRouteMode || 'normal');
    currentAgentBankOptions = Array.isArray(systemSettings.agentBankOptions) ? [...systemSettings.agentBankOptions] : [...defaultSystemSettings.agentBankOptions];
    renderAgentBankManagement();
    if (!settingsCardsInitialized) initializeSettingsCardLayout();
    renderSettingsDropdownNav();
    renderSettingsSubTabState();
}

function renderCurrentTab() {
    if (currentTab === 'global') return renderGlobalView();
    if (currentTab === 'admins') return renderAdminManagement();
    if (currentTab === 'routing-rules') return renderRoutingRules();
    if (currentTab === 'agents') {
        renderAgentSubTabState();
        if (currentAgentSubTab === 'agent-records') return renderAgentRecords();
        if (currentAgentSubTab === 'agent-reroute') return renderApplicationAgentRerouteModule();
        return renderApplicationAgentModule();
    }
    if (currentTab === 'audit') return renderAudit();
    if (currentTab === 'security') return renderSecurity();
    if (currentTab === 'round-robin') return renderRoundRobin();
    if (currentTab === 'settings') return loadSettings();
}

window.saveUserAccessLevel = async (userId) => {
    const user = allUsers.find((item) => item.id === userId);
    if (!user) return showNotification('User not found', 'error');

    const select = document.getElementById(`access-level-${userId}`);
    const rsaRoundRobinSelect = document.getElementById(`access-rsa-rr-${userId}`) || document.getElementById(`rsa-rr-${userId}`);
    const saveBtn = document.getElementById(`save-access-level-${userId}`);
    const originalHtml = saveBtn?.innerHTML || '';
    const roleLevel = Number(select?.value || 1) === 2 ? 2 : 1;
    const role = String(user?.role || '').trim().toLowerCase();
    const skipRsaRoundRobin = role === 'rsa'
        ? String(rsaRoundRobinSelect?.value || 'include') === 'skip'
        : Boolean(user?.skipRsaRoundRobin);

    try {
        if (saveBtn) {
            saveBtn.disabled = true;
            saveBtn.classList.add('loading');
            saveBtn.innerHTML = '<i class="fas fa-spinner fa-spin"></i> Saving...';
        }

        await updateDoc(doc(db, 'users', userId), {
            roleLevel,
            skipRsaRoundRobin,
            updatedAt: serverTimestamp(),
            updatedBy: currentUser?.email || ''
        });

        await addDoc(collection(db, 'audit'), {
            action: 'user_role_level_updated',
            userId,
            userEmail: user.email || '',
            userRole: user.role || '',
            newRoleLevel: roleLevel,
            skipRsaRoundRobin,
            performedBy: currentUser?.email || '',
            timestamp: serverTimestamp()
        });

        const targetIndex = allUsers.findIndex((item) => item.id === userId);
        if (targetIndex >= 0) {
            allUsers[targetIndex] = { ...allUsers[targetIndex], roleLevel, skipRsaRoundRobin };
        }
        renderAccessLevelManager();
        showNotification('Access level updated successfully.', 'success');
    } catch (error) {
        showNotification('Failed to update access level', 'error');
    } finally {
        if (saveBtn) {
            saveBtn.disabled = false;
            saveBtn.classList.remove('loading');
            saveBtn.innerHTML = originalHtml;
        }
    }
};

window.saveRsaRoundRobinSetting = async (userId) => {
    const user = allUsers.find((item) => item.id === userId);
    if (!user) return showNotification('RSA user not found', 'error');

    const role = String(user?.role || '').trim().toLowerCase();
    if (role !== 'rsa') return showNotification('Only RSA users can use this setting', 'error');

    const rsaRoundRobinSelect = document.getElementById(`rsa-rr-${userId}`) || document.getElementById(`access-rsa-rr-${userId}`);
    const saveBtn = document.getElementById(`save-rsa-rr-${userId}`);
    const originalHtml = saveBtn?.innerHTML || '';
    const skipRsaRoundRobin = String(rsaRoundRobinSelect?.value || 'include') === 'skip';

    try {
        if (saveBtn) {
            saveBtn.disabled = true;
            saveBtn.classList.add('loading');
            saveBtn.innerHTML = '<i class="fas fa-spinner fa-spin"></i> Saving...';
        }

        await updateDoc(doc(db, 'users', userId), {
            skipRsaRoundRobin,
            updatedAt: serverTimestamp(),
            updatedBy: currentUser?.email || ''
        });

        await addDoc(collection(db, 'audit'), {
            action: 'rsa_round_robin_setting_updated',
            userId,
            userEmail: user.email || '',
            userRole: user.role || '',
            skipRsaRoundRobin,
            performedBy: currentUser?.email || '',
            timestamp: serverTimestamp()
        });

        const targetIndex = allUsers.findIndex((item) => item.id === userId);
        if (targetIndex >= 0) {
            allUsers[targetIndex] = { ...allUsers[targetIndex], skipRsaRoundRobin };
        }

        renderAccessLevelManager();
        renderRsaRoundRobinManager();
        showNotification('RSA round robin setting updated successfully.', 'success');
    } catch (error) {
        showNotification('Failed to update RSA round robin setting', 'error');
    } finally {
        if (saveBtn) {
            saveBtn.disabled = false;
            saveBtn.classList.remove('loading');
            saveBtn.innerHTML = originalHtml;
        }
    }
};

window.updateAccessUserSearch = (value) => {
    currentAccessUserSearch = normalizeAccessUserSearch(value);
    updateAccessSearchInputs();
    renderAccessLevelManager();
};

async function saveUploaderRoutingRuleForMode(uploaderUserId, mode = 'normal') {
    const uploader = allUsers.find((u) => u.id === uploaderUserId);
    if (!uploader) return showNotification('Uploader not found', 'error');

    const uploaderEmail = normalizeEmail(uploader.email);
    if (!uploaderEmail) return showNotification('Uploader email is missing', 'error');

    const isSkipReviewer = mode === 'skip_reviewer';
    const saveBtn = document.getElementById(`${isSkipReviewer ? 'route-skip-save' : 'route-save'}-${uploaderUserId}`);
    const originalHtml = saveBtn?.innerHTML || '';
    const reviewerEmail = isSkipReviewer ? '' : normalizeEmail(document.getElementById(`route-reviewer-${uploaderUserId}`)?.value);
    const rsaEmail = normalizeEmail(document.getElementById(`${isSkipReviewer ? 'route-skip-rsa' : 'route-rsa'}-${uploaderUserId}`)?.value);
    const paymentEmail = normalizeEmail(document.getElementById(`${isSkipReviewer ? 'route-skip-payment' : 'route-payment'}-${uploaderUserId}`)?.value);
    if (isSkipReviewer && !rsaEmail) return showNotification('RSA is required for Skip Reviewer routing', 'error');
    const enabled = Boolean(reviewerEmail || rsaEmail || paymentEmail);
    const existingRule = findRoutingRuleForUploader(uploaderEmail);

    try {
        if (saveBtn) {
            saveBtn.disabled = true;
            saveBtn.classList.add('loading');
            saveBtn.innerHTML = '<i class="fas fa-spinner fa-spin"></i> Saving...';
        }
        const payload = {
            uploaderEmail,
            routeMode: isSkipReviewer ? 'skip_reviewer' : 'normal',
            reviewerEmail: reviewerEmail || '',
            rsaEmail: rsaEmail || '',
            paymentEmail: paymentEmail || '',
            enabled,
            updatedAt: serverTimestamp(),
            updatedBy: currentUser?.email || ''
        };
        if (!existingRule) payload.createdAt = serverTimestamp();

        await setDoc(doc(db, 'uploaderRoutingRules', encodeURIComponent(uploaderEmail)), payload, { merge: true });

        await addDoc(collection(db, 'audit'), {
            action: 'uploader_routing_rule_updated',
            uploaderEmail,
            routeMode: isSkipReviewer ? 'skip_reviewer' : 'normal',
            reviewerEmail: reviewerEmail || '',
            rsaEmail: rsaEmail || '',
            paymentEmail: paymentEmail || '',
            enabled,
            performedBy: currentUser?.email || '',
            timestamp: serverTimestamp()
        });

        showNotification(isSkipReviewer ? 'Skip Reviewer routing rule saved' : 'Uploader routing rule saved', 'success');
    } catch (error) {
        showNotification('Failed to save routing rule', 'error');
    } finally {
        if (saveBtn) {
            saveBtn.disabled = false;
            saveBtn.classList.remove('loading');
            saveBtn.innerHTML = originalHtml;
        }
    }
}

function renderSettingsSubTabState() {
    SETTINGS_DROPDOWN_TABS.forEach(({ id }) => {
        const dropdown = document.querySelector(`[data-settings-dropdown="${id}"]`);
        const button = document.querySelector(`[data-settings-dropdown-btn="${id}"]`);
        const active = activeSettingsDropdownTab === id;
        if (dropdown) dropdown.classList.toggle('active', active);
        if (button) {
            button.classList.toggle('active', active);
            button.setAttribute('aria-expanded', active ? 'true' : 'false');
        }
        const section = document.getElementById(`settings${id.charAt(0).toUpperCase()}${id.slice(1)}Section`);
        if (section) section.style.display = 'none';
    });
}

function getSettingsCardIcon(title = '') {
    return SETTINGS_CARD_ICON_MAP[title] || 'fa-sliders';
}

function initializeSettingsCardLayout() {
    if (!settingsModalStorage) return;
    SETTINGS_DROPDOWN_TABS.forEach(({ id: tabId }) => {
        const section = document.getElementById(`settings${tabId.charAt(0).toUpperCase()}${tabId.slice(1)}Section`);
        if (!section || section.dataset.cardified === 'true') return;

        const panels = Array.from(section.children).filter((child) => child instanceof HTMLElement && child.tagName === 'DIV');

        panels.forEach((panel, index) => {
            const title = String(panel.querySelector('h3')?.textContent || `${SETTINGS_SUBTAB_META[tabId]?.title || 'Settings'} ${index + 1}`).trim();
            const description = Array.from(panel.querySelectorAll('p'))
                .map((item) => String(item.textContent || '').trim())
                .find(Boolean) || 'Open this settings section in a larger workspace.';
            const panelId = `${tabId}-settings-panel-${index + 1}`;
            panel.dataset.settingsPanelId = panelId;
            panel.dataset.settingsPanelTitle = title;
            panel.dataset.settingsPanelDescription = description;
            panel.dataset.settingsPanelIcon = getSettingsCardIcon(title);
            panel.classList.add('settings-modal-panel');
            settingsModalStorage.appendChild(panel);
        });

        section.innerHTML = '';
        section.style.display = 'none';
        section.dataset.cardified = 'true';
    });
    renderSettingsDropdownNav();
    settingsCardsInitialized = true;
}

function getSettingsPanelsForTab(tabId) {
    return Array.from(document.querySelectorAll(`[data-settings-panel-id^="${tabId}-settings-panel-"]`))
        .filter((panel) => panel instanceof HTMLElement)
        .sort((a, b) => String(a.dataset.settingsPanelId || '').localeCompare(String(b.dataset.settingsPanelId || '')));
}

function renderSettingsDropdownNav() {
    if (!settingsDropdownNav) return;
    settingsDropdownNav.innerHTML = SETTINGS_DROPDOWN_TABS.map(({ id, label, icon }) => {
        const panels = getSettingsPanelsForTab(id);
        const menuItems = panels.map((panel) => {
            const panelId = String(panel.dataset.settingsPanelId || '').trim();
            const title = String(panel.dataset.settingsPanelTitle || 'Settings Section').trim();
            const description = String(panel.dataset.settingsPanelDescription || '').trim();
            const itemIcon = String(panel.dataset.settingsPanelIcon || getSettingsCardIcon(title)).trim();
            return `
                <button
                    type="button"
                    class="settings-dropdown-item"
                    data-settings-panel-open="${escapeHtml(panelId)}"
                    title="${escapeHtml(description)}"
                >
                    <span class="settings-dropdown-item-icon"><i class="fas ${itemIcon}"></i></span>
                    <span class="settings-dropdown-item-copy">
                        <strong>${escapeHtml(title)}</strong>
                        <small>${escapeHtml(description)}</small>
                    </span>
                </button>
            `;
        }).join('');

        return `
            <div class="settings-dropdown-group">
                <button
                    type="button"
                    class="admin-subtab-btn settings-dropdown-trigger ${currentSettingsSubTab === id ? 'active' : ''}"
                    data-settings-dropdown-btn="${id}"
                    aria-expanded="${activeSettingsDropdownTab === id ? 'true' : 'false'}"
                >
                    <i class="fas ${icon}"></i>
                    <span>${escapeHtml(label)}</span>
                    <i class="fas fa-chevron-down settings-dropdown-caret"></i>
                </button>
                <div class="settings-dropdown-menu ${activeSettingsDropdownTab === id ? 'active' : ''}" data-settings-dropdown="${id}">
                    ${menuItems || '<div class="settings-dropdown-empty">No subtabs available</div>'}
                </div>
            </div>
        `;
    }).join('');

    settingsDropdownNav.querySelectorAll('[data-settings-dropdown-btn]').forEach((button) => {
        button.addEventListener('click', (event) => {
            event.stopPropagation();
            const tabId = String(button.getAttribute('data-settings-dropdown-btn') || '').trim();
            window.switchSettingsSubTab(tabId);
        });
    });

    settingsDropdownNav.querySelectorAll('[data-settings-panel-open]').forEach((button) => {
        button.addEventListener('click', (event) => {
            event.stopPropagation();
            const panelId = String(button.getAttribute('data-settings-panel-open') || '').trim();
            activeSettingsDropdownTab = '';
            renderSettingsSubTabState();
            openSettingsSectionModal(panelId);
        });
    });
}

function returnActiveSettingsPanelToStorage() {
    if (!activeSettingsModalPanelId || !settingsModalStorage || !settingsSectionModalContentHost) return;
    const panel = document.querySelector(`[data-settings-panel-id="${activeSettingsModalPanelId}"]`);
    if (panel) settingsModalStorage.appendChild(panel);
    activeSettingsModalPanelId = '';
    settingsSectionModalContentHost.innerHTML = '';
}

function closeSettingsSectionModal() {
    returnActiveSettingsPanelToStorage();
    settingsSectionModal?.classList.remove('active');
}

function openSettingsSectionModal(panelId) {
    const panel = document.querySelector(`[data-settings-panel-id="${panelId}"]`);
    if (!panel || !settingsSectionModalContentHost) return;

    returnActiveSettingsPanelToStorage();
    activeSettingsModalPanelId = panelId;
    const title = String(panel.dataset.settingsPanelTitle || 'Settings Section').trim();
    const description = String(panel.dataset.settingsPanelDescription || 'Manage this settings section in a larger workspace, then save your changes.').trim();
    const icon = String(panel.dataset.settingsPanelIcon || 'fa-sliders').trim();

    if (settingsSectionModalTitle) {
        settingsSectionModalTitle.innerHTML = `<i class="fas ${icon}"></i> ${escapeHtml(title)}`;
    }
    if (settingsSectionModalHeroTitle) settingsSectionModalHeroTitle.textContent = title;
    if (settingsSectionModalDescription) settingsSectionModalDescription.textContent = description;
    if (settingsSectionModalIcon) settingsSectionModalIcon.innerHTML = `<i class="fas ${icon}"></i>`;

    settingsSectionModalContentHost.innerHTML = '';
    settingsSectionModalContentHost.appendChild(panel);
    settingsSectionModal?.classList.add('active');
}

function renderAgentBankManagement() {
    const body = document.getElementById('agentBankTableBody');
    if (!body) return;

    if (!currentAgentBankOptions.length) {
        body.innerHTML = '<tr><td colspan="3" class="no-data">No banks configured yet</td></tr>';
        return;
    }

    body.innerHTML = currentAgentBankOptions.map((bank) => {
        const encodedName = encodeURIComponent(bank.name);
        return `
            <tr>
                <td><strong>${escapeHtml(bank.name)}</strong></td>
                <td>${bank.active ? '<span class="status-badge approved">Active</span>' : '<span class="status-badge pending">Inactive</span>'}</td>
                <td style="display:flex;gap:8px;flex-wrap:wrap;">
                    <button class="action-btn" type="button" onclick="window.toggleAgentBankOption('${encodedName}')" style="background:${bank.active ? '#b45309' : '#0f766e'};color:#fff;border:none;">
                        <i class="fas ${bank.active ? 'fa-ban' : 'fa-check'}"></i> ${bank.active ? 'Deactivate' : 'Activate'}
                    </button>
                    <button class="action-btn" type="button" onclick="window.removeAgentBankOption('${encodedName}')" style="background:#b91c1c;color:#fff;border:none;">
                        <i class="fas fa-trash"></i> Remove
                    </button>
                </td>
            </tr>
        `;
    }).join('');
}

function renderPfaManagement() {
    const body = document.getElementById('pfaTableBody');
    if (!body) return;
    if (!currentPfaOptions.length) {
        body.innerHTML = '<tr><td colspan="2" class="no-data">No PFAs configured yet</td></tr>';
        return;
    }
    body.innerHTML = currentPfaOptions.map((name) => `
        <tr>
            <td><strong>${escapeHtml(name)}</strong></td>
            <td><button class="action-btn" type="button" onclick="window.removePfaOption('${encodeURIComponent(name)}')" style="background:#b91c1c;color:#fff;border:none;"><i class="fas fa-trash"></i> Remove</button></td>
        </tr>
    `).join('');
}

function renderDocumentRequirementsManager() {
    const body = document.getElementById('documentRequirementsTableBody');
    if (!body) return;
    if (!currentDocumentRequirements.length) {
        body.innerHTML = '<tr><td colspan="5" class="no-data">No document boxes configured</td></tr>';
        return;
    }
    body.innerHTML = currentDocumentRequirements.map((doc, index) => `
        <tr>
            <td><code>${escapeHtml(doc.id)}</code></td>
            <td>
                <input
                    type="text"
                    value="${escapeHtml(doc.name)}"
                    oninput="window.updateDocumentRequirementName(${index}, this.value)"
                    style="width:100%;padding:8px;border:1px solid #e2e8f0;border-radius:8px;"
                >
            </td>
            <td>
                <select onchange="window.updateDocumentRequirementField(${index},'required',this.value)" style="padding:8px;border:1px solid #e2e8f0;border-radius:8px;">
                    <option value="true" ${doc.required ? 'selected' : ''}>Required</option>
                    <option value="false" ${!doc.required ? 'selected' : ''}>Optional</option>
                </select>
            </td>
            <td>
                <select onchange="window.updateDocumentRequirementField(${index},'active',this.value)" style="padding:8px;border:1px solid #e2e8f0;border-radius:8px;">
                    <option value="true" ${doc.active !== false ? 'selected' : ''}>Active</option>
                    <option value="false" ${doc.active === false ? 'selected' : ''}>Inactive</option>
                </select>
            </td>
            <td>
                <button class="action-btn" type="button" onclick="window.removeDocumentRequirement(${index})" style="background:#b91c1c;color:#fff;border:none;">
                    <i class="fas fa-trash"></i> Remove
                </button>
            </td>
        </tr>
    `).join('');
}

function renderWorkflowLabelsManager() {
    const grid = document.getElementById('workflowLabelsGrid');
    if (!grid) return;
    const entries = Object.entries(currentWorkflowLabels || {});
    grid.innerHTML = entries.map(([key, value]) => `
        <label>${escapeHtml(formatSettingLabel(key))}
            <input type="text" value="${escapeHtml(String(value || ''))}" oninput="window.updateWorkflowLabel('${escapeHtml(key)}', this.value)" style="width:100%;margin-top:6px;padding:10px;border:1px solid #e2e8f0;border-radius:8px;">
        </label>
    `).join('');
}

function renderPropertyRulesManager() {
    const body = document.getElementById('propertyRulesTableBody');
    if (!body) return;
    if (!currentPropertyRules.length) {
        body.innerHTML = '<tr><td colspan="6" class="no-data">No property rules configured</td></tr>';
        return;
    }
    body.innerHTML = currentPropertyRules.map((rule, index) => `
        <tr>
            <td><input type="text" value="${escapeHtml(rule.name || '')}" oninput="window.updatePropertyRule(${index},'name',this.value)" style="width:100%;padding:8px;border:1px solid #e2e8f0;border-radius:8px;"></td>
            <td><input type="number" value="${Number(rule.min || 0)}" oninput="window.updatePropertyRule(${index},'min',this.value)" style="width:120px;padding:8px;border:1px solid #e2e8f0;border-radius:8px;"></td>
            <td><input type="number" value="${Number(rule.max || 0)}" oninput="window.updatePropertyRule(${index},'max',this.value)" style="width:120px;padding:8px;border:1px solid #e2e8f0;border-radius:8px;"></td>
            <td><input type="number" value="${Number(rule.value || 0)}" oninput="window.updatePropertyRule(${index},'value',this.value)" style="width:120px;padding:8px;border:1px solid #e2e8f0;border-radius:8px;"></td>
            <td><input type="number" value="${Number(rule.fee || 0)}" oninput="window.updatePropertyRule(${index},'fee',this.value)" style="width:100px;padding:8px;border:1px solid #e2e8f0;border-radius:8px;"></td>
            <td><button class="action-btn" type="button" onclick="window.removePropertyRule(${index})" style="background:#b91c1c;color:#fff;border:none;"><i class="fas fa-trash"></i> Remove</button></td>
        </tr>
    `).join('');
}

function getUserRoleLevel(user = {}) {
    const rawLevel = user?.roleLevel ?? user?.accessLevel ?? 1;
    const normalized = String(rawLevel || '').trim().toLowerCase();
    if (Number(rawLevel) === 2 || normalized === '2' || normalized === 'level 2' || normalized === 'level2') {
        return 2;
    }
    return 1;
}

function getRoleLevelCapability(user = {}) {
    const role = String(user?.role || '').trim().toLowerCase();
    const level = getUserRoleLevel(user);
    if (role === 'uploader') {
        return 'Document requirement follows Workflow settings';
    }
    if (role === 'rsa') {
        const exportScope = level === 2 ? 'Can export Excel records across all RSA users' : 'Can export only current RSA records';
        return user?.skipRsaRoundRobin ? `${exportScope}; skipped in RSA round robin` : exportScope;
    }
    return '-';
}

function renderAccessLevelManager() {
    const body = document.getElementById('accessLevelsTableBody');
    if (!body) return;
    updateAccessSearchInputs();

    const rows = allUsers
        .filter((user) => ['uploader', 'rsa'].includes(String(user.role || '').trim().toLowerCase()))
        .filter((user) => getAccessSearchMatch(user))
        .sort((a, b) => String(a.fullName || a.email || '').localeCompare(String(b.fullName || b.email || '')));

    if (!rows.length) {
        body.innerHTML = `<tr><td colspan="6" class="no-data">${currentAccessUserSearch ? 'No matching uploader or RSA users found' : 'No uploader or RSA users found'}</td></tr>`;
        renderRsaRoundRobinManager();
        return;
    }

    body.innerHTML = rows.map((user) => {
        const role = String(user.role || '').trim().toLowerCase();
        const level = getUserRoleLevel(user);
        return `
            <tr>
                <td><strong>${escapeHtml(user.fullName || user.email || 'Unknown')}</strong></td>
                <td>${escapeHtml(user.email || '-')}</td>
                <td>${escapeHtml(formatSettingLabel(role))}</td>
                <td>
                    <select id="access-level-${user.id}" style="padding:8px;border:1px solid #e2e8f0;border-radius:8px;">
                        <option value="1" ${level === 1 ? 'selected' : ''}>Level 1</option>
                        <option value="2" ${level === 2 ? 'selected' : ''}>Level 2</option>
                    </select>
                </td>
                <td>${escapeHtml(getRoleLevelCapability(user))}</td>
                <td>
                    <button id="save-access-level-${user.id}" class="action-btn" type="button" style="background:#003366;color:#fff;border:none;" onclick="window.saveUserAccessLevel('${user.id}')">
                        <i class="fas fa-save"></i> Save
                    </button>
                </td>
            </tr>
        `;
    }).join('');

    renderRsaRoundRobinManager();
}

function renderRsaRoundRobinManager() {
    const grid = document.getElementById('rsaRoundRobinAccessGrid');
    if (!grid) return;
    updateAccessSearchInputs();

    const rsaUsers = allUsers
        .filter((user) => String(user?.role || '').trim().toLowerCase() === 'rsa')
        .filter((user) => getAccessSearchMatch(user))
        .sort((a, b) => String(a.fullName || a.email || '').localeCompare(String(b.fullName || b.email || '')));

    if (!rsaUsers.length) {
        grid.innerHTML = `<div class="no-data" style="padding:16px;border:1px dashed #cbd5e1;border-radius:12px;background:#fff;">${currentAccessUserSearch ? 'No matching RSA users found' : 'No RSA users found'}</div>`;
        return;
    }

    grid.innerHTML = rsaUsers.map((user) => {
        const skipRsaRoundRobin = Boolean(user?.skipRsaRoundRobin);
        const level = getUserRoleLevel(user);
        return `
            <div style="padding:16px;border:1px solid #dbe6f2;border-radius:14px;background:#fff;box-shadow:0 8px 20px rgba(15,59,103,0.06);">
                <div style="display:flex;justify-content:space-between;align-items:flex-start;gap:12px;margin-bottom:10px;">
                    <div>
                        <div style="font-weight:700;color:#0f172a;">${escapeHtml(user.fullName || user.email || 'Unknown')}</div>
                        <div style="font-size:13px;color:#64748b;">${escapeHtml(user.email || '-')}</div>
                    </div>
                    <span style="display:inline-flex;align-items:center;padding:6px 10px;border-radius:999px;background:#e0f2fe;color:#075985;font-size:12px;font-weight:700;">Level ${level}</span>
                </div>
                <label style="display:block;color:#334155;font-size:13px;font-weight:600;margin-bottom:8px;">RSA Round Robin</label>
                <select id="rsa-rr-${user.id}" style="width:100%;padding:10px;border:1px solid #e2e8f0;border-radius:10px;">
                    <option value="include" ${!skipRsaRoundRobin ? 'selected' : ''}>Include in round robin</option>
                    <option value="skip" ${skipRsaRoundRobin ? 'selected' : ''}>Skip this RSA user</option>
                </select>
                <p style="margin:10px 0 14px;color:#64748b;font-size:13px;">${escapeHtml(getRoleLevelCapability(user))}</p>
                <button id="save-rsa-rr-${user.id}" class="action-btn" type="button" style="background:#003366;color:#fff;border:none;" onclick="window.saveRsaRoundRobinSetting('${user.id}')">
                    <i class="fas fa-save"></i> Save RSA Setting
                </button>
            </div>
        `;
    }).join('');
}

function renderRolePermissionsManager() {
    const grid = document.getElementById('rolePermissionsGrid');
    if (!grid) return;
    grid.innerHTML = ROLE_PERMISSION_FIELDS.map(({ key, label }) => `
        <label>${escapeHtml(label)}
            <select onchange="window.updateRolePermission('${key}', this.value)" style="width:100%;margin-top:6px;padding:10px;border:1px solid #e2e8f0;border-radius:8px;">
                <option value="true" ${currentRolePermissions[key] !== false ? 'selected' : ''}>Enabled</option>
                <option value="false" ${currentRolePermissions[key] === false ? 'selected' : ''}>Disabled</option>
            </select>
        </label>
    `).join('');
}

function renderDocumentRequirementRoleManager() {
    const grid = document.getElementById('documentRequirementRoleGrid');
    if (!grid) return;
    const defaults = getDefaultSystemSettings().documentRequirementRoles || {};
    grid.innerHTML = DOCUMENT_REQUIREMENT_ROLE_FIELDS.map(({ key, label }) => {
        const enforced = (key in currentDocumentRequirementRoles)
            ? currentDocumentRequirementRoles[key] !== false
            : defaults[key] !== false;
        return `
        <label>${escapeHtml(label)}
            <select onchange="window.updateDocumentRequirementRole('${key}', this.value)" style="width:100%;margin-top:6px;padding:10px;border:1px solid #e2e8f0;border-radius:8px;">
                <option value="true" ${enforced ? 'selected' : ''}>Required Before Submit</option>
                <option value="false" ${!enforced ? 'selected' : ''}>Exempt</option>
            </select>
        </label>
    `;
    }).join('');
}

function renderNotificationTemplatesManager() {
    const grid = document.getElementById('notificationTemplatesGrid');
    if (!grid) return;
    grid.innerHTML = NOTIFICATION_TEMPLATE_FIELDS.map(({ key, label }) => `
        <label>${escapeHtml(label)}
            <textarea rows="4" oninput="window.updateNotificationTemplate('${key}', this.value)" style="width:100%;margin-top:6px;padding:10px;border:1px solid #e2e8f0;border-radius:8px;resize:vertical;">${escapeHtml(String(currentNotificationTemplates[key] || ''))}</textarea>
        </label>
    `).join('');
}

function renderHouseNumberRulesManager() {
    const body = document.getElementById('houseNumberRulesTableBody');
    if (!body) return;
    if (!currentHouseNumberRules.length) {
        body.innerHTML = '<tr><td colspan="7" class="no-data">No house number rules configured</td></tr>';
        return;
    }
    body.innerHTML = currentHouseNumberRules.map((rule, index) => `
        <tr>
            <td><input type="text" value="${escapeHtml(rule.propertyType || '')}" oninput="window.updateHouseNumberRule(${index},'propertyType',this.value)" style="width:100%;padding:8px;border:1px solid #e2e8f0;border-radius:8px;"></td>
            <td>
                <select onchange="window.updateHouseNumberRule(${index},'mode',this.value)" style="padding:8px;border:1px solid #e2e8f0;border-radius:8px;">
                    <option value="alpha_suffix" ${rule.mode === 'alpha_suffix' ? 'selected' : ''}>Alpha Suffix</option>
                    <option value="block_100" ${rule.mode === 'block_100' ? 'selected' : ''}>Block 100</option>
                    <option value="house_infinite" ${rule.mode === 'house_infinite' ? 'selected' : ''}>House Infinite</option>
                </select>
            </td>
            <td><input type="text" value="${escapeHtml(rule.prefix || '')}" oninput="window.updateHouseNumberRule(${index},'prefix',this.value)" style="width:100px;padding:8px;border:1px solid #e2e8f0;border-radius:8px;"></td>
            <td><input type="number" value="${Number(rule.startNumber || 0)}" oninput="window.updateHouseNumberRule(${index},'startNumber',this.value)" style="width:120px;padding:8px;border:1px solid #e2e8f0;border-radius:8px;"></td>
            <td><input type="text" value="${escapeHtml(rule.startLetter || '')}" oninput="window.updateHouseNumberRule(${index},'startLetter',this.value)" style="width:100px;padding:8px;border:1px solid #e2e8f0;border-radius:8px;"></td>
            <td><input type="text" value="${escapeHtml(rule.startPrefix || '')}" oninput="window.updateHouseNumberRule(${index},'startPrefix',this.value)" style="width:120px;padding:8px;border:1px solid #e2e8f0;border-radius:8px;"></td>
            <td><button class="action-btn" type="button" onclick="window.removeHouseNumberRule(${index})" style="background:#b91c1c;color:#fff;border:none;"><i class="fas fa-trash"></i> Remove</button></td>
        </tr>
    `).join('');
}

async function persistAgentBankOptions(actionLabel) {
    const normalizedOptions = normalizeAgentBankOptions(currentAgentBankOptions, []);
    currentAgentBankOptions = normalizedOptions;

    await setDoc(doc(db, 'settings', 'system'), {
        agentBankOptions: normalizedOptions,
        updatedAt: serverTimestamp(),
        updatedBy: currentUser?.email || ''
    }, { merge: true });

    clearSystemSettingsCache();
    renderAgentBankManagement();

    await addDoc(collection(db, 'audit'), {
        action: actionLabel,
        bankCount: normalizedOptions.length,
        activeBankCount: normalizedOptions.filter((bank) => bank.active).length,
        performedBy: currentUser?.email || '',
        timestamp: serverTimestamp()
    });
}

function getUploaderOwnedApprovedAgents(uploaderEmail = '') {
    const normalizedUploaderEmail = normalizeEmail(uploaderEmail);
    return allAgents.filter((agent) => {
        if (String(agent.status || '').toLowerCase() !== 'approved') return false;
        const createdByUid = String(agent.createdByUid || '').trim();
        const uploader = allUsers.find((user) => normalizeEmail(user.email) === normalizedUploaderEmail);
        if (uploader && createdByUid && createdByUid === String(uploader.uid || uploader.id || '').trim()) return true;
        return normalizeEmail(agent.createdBy) === normalizedUploaderEmail;
    });
}

function renderApplicationAgentModule() {
    const body = document.getElementById('superApplicationAgentTableBody');
    if (!body) return;

    const rows = allSubmissions
        .filter((sub) => String(sub.status || '').toLowerCase() !== 'draft')
        .filter((sub) => !String(sub.agentId || '').trim())
        .slice(0, 300);

    if (!rows.length) {
        body.innerHTML = '<tr><td colspan="6" class="no-data">No applications are missing agent assignment</td></tr>';
        return;
    }

    body.innerHTML = rows.map((sub) => {
        const uploaderEmail = normalizeEmail(sub.uploadedBy);
        const options = getUploaderOwnedApprovedAgents(uploaderEmail);
        const selectOptions = options.length
            ? options.map((agent) => `<option value="${escapeHtml(agent.id)}">${escapeHtml(agent.fullName || 'Unnamed')} (${escapeHtml(agent.contactNumber || '-')})</option>`).join('')
            : '<option value="">No registered approved agents for this uploader</option>';
        return `
            <tr>
                <td><strong>${escapeHtml(sub.customerName || 'Unknown')}</strong></td>
                <td>${escapeHtml(uploaderEmail || '-')}</td>
                <td>${escapeHtml(String(sub.status || '-'))}</td>
                <td>No Agent</td>
                <td>
                    <select id="sa-app-agent-${sub.id}" style="padding:8px;border:1px solid #e2e8f0;border-radius:8px;width:100%;">
                        <option value="">Select registered agent...</option>
                        ${selectOptions}
                    </select>
                </td>
                <td>
                    <button class="action-btn" style="background:#003366;color:#fff;border:none;" onclick="window.assignApplicationAgent('${sub.id}')">
                        <i class="fas fa-save"></i> Update Agent
                    </button>
                </td>
            </tr>
        `;
    }).join('');
}

function getApplicationLinkedAgentLabel(sub = {}) {
    if (!String(sub.agentId || '').trim()) return 'No Agent';
    return String(sub.agentName || '').trim() || 'Unknown Agent';
}

function renderApplicationAgentRerouteModule() {
    const body = document.getElementById('superAgentRerouteTableBody');
    if (!body) return;

    const search = String(document.getElementById('superAgentRerouteSearch')?.value || '').trim().toLowerCase();
    const rows = allSubmissions
        .filter((sub) => String(sub.status || '').toLowerCase() !== 'draft')
        .filter((sub) => String(sub.agentId || '').trim())
        .filter((sub) => {
            const currentAgentLabel = getApplicationLinkedAgentLabel(sub);
            const searchable = [
                sub.customerName || '',
                sub.uploadedBy || '',
                currentAgentLabel,
                sub.status || ''
            ].join(' ').toLowerCase();
            return !search || searchable.includes(search);
        })
        .slice(0, 300);

    if (!rows.length) {
        body.innerHTML = '<tr><td colspan="6" class="no-data">No linked applications match this search</td></tr>';
        return;
    }

    body.innerHTML = rows.map((sub) => {
        const uploaderEmail = normalizeEmail(sub.uploadedBy);
        const options = getUploaderOwnedApprovedAgents(uploaderEmail)
            .filter((agent) => String(agent.id || '').trim() !== String(sub.agentId || '').trim());
        const selectOptions = options.length
            ? options.map((agent) => `<option value="${escapeHtml(agent.id)}">${escapeHtml(agent.fullName || 'Unnamed')} (${escapeHtml(agent.contactNumber || '-')})</option>`).join('')
            : '<option value="">No other approved agents for this uploader</option>';
        return `
            <tr>
                <td><strong>${escapeHtml(sub.customerName || 'Unknown')}</strong></td>
                <td>${escapeHtml(uploaderEmail || '-')}</td>
                <td>${escapeHtml(String(sub.status || '-'))}</td>
                <td>${escapeHtml(getApplicationLinkedAgentLabel(sub))}</td>
                <td>
                    <select id="sa-reroute-agent-${sub.id}" style="padding:8px;border:1px solid #e2e8f0;border-radius:8px;width:100%;">
                        <option value="">Select replacement agent...</option>
                        ${selectOptions}
                    </select>
                </td>
                <td>
                    <button class="action-btn" style="background:#0f766e;color:#fff;border:none;" onclick="window.rerouteApplicationAgent('${sub.id}')">
                        <i class="fas fa-right-left"></i> Re-route
                    </button>
                </td>
            </tr>
        `;
    }).join('');
}

function getAgentLinkedSubmissionCount(agentId) {
    const normalizedId = String(agentId || '').trim();
    if (!normalizedId) return 0;
    return allSubmissions.filter((sub) => String(sub.agentId || '').trim() === normalizedId).length;
}

function getAgentRegisteredByLabel(agent = {}) {
    const email = normalizeEmail(agent.createdBy);
    if (!email) return '-';
    const user = allUsers.find((item) => normalizeEmail(item.email) === email);
    if (!user) return email;
    return `${user.fullName || email} (${email})`;
}

function renderAgentRecords() {
    const body = document.getElementById('superAgentRecordsTableBody');
    if (!body) return;

    const search = String(document.getElementById('superAgentSearch')?.value || '').trim().toLowerCase();
    const statusFilter = String(document.getElementById('superAgentStatusFilter')?.value || '').trim().toLowerCase();

    const rows = allAgents
        .filter((agent) => {
            const status = String(agent.status || 'pending').toLowerCase();
            if (statusFilter && status !== statusFilter) return false;
            const searchable = [
                agent.fullName || '',
                agent.contactNumber || '',
                agent.accountBank || '',
                agent.accountNumber || '',
                agent.createdBy || '',
                getAgentRegisteredByLabel(agent)
            ].join(' ').toLowerCase();
            return !search || searchable.includes(search);
        })
        .sort((a, b) => String(a.fullName || '').localeCompare(String(b.fullName || '')));

    if (!rows.length) {
        body.innerHTML = '<tr><td colspan="8" class="no-data">No agent records found</td></tr>';
        return;
    }

    body.innerHTML = rows.map((agent) => {
        const status = String(agent.status || 'pending').toLowerCase();
        const linkedApps = getAgentLinkedSubmissionCount(agent.id);
        const badgeClass = status === 'approved' ? 'status-active' : status === 'rejected' ? 'status-deactivated' : 'status-pending';
        return `
            <tr>
                <td><strong>${escapeHtml(agent.fullName || 'Unnamed Agent')}</strong></td>
                <td>${escapeHtml(agent.contactNumber || '-')}</td>
                <td>${escapeHtml(agent.accountBank || '-')}</td>
                <td>${escapeHtml(agent.accountNumber || '-')}</td>
                <td>${escapeHtml(getAgentRegisteredByLabel(agent))}</td>
                <td><span class="status-badge ${badgeClass}">${escapeHtml(status)}</span></td>
                <td>${linkedApps}</td>
                <td>
                    <button class="action-btn edit-btn" onclick="window.editSuperAgent('${agent.id}')">
                        <i class="fas fa-edit"></i> Edit
                    </button>
                </td>
            </tr>
        `;
    }).join('');
}

window.saveUploaderRoutingRule = async (uploaderUserId) => saveUploaderRoutingRuleForMode(uploaderUserId, 'normal');
window.saveSkipReviewerRoutingRule = async (uploaderUserId) => saveUploaderRoutingRuleForMode(uploaderUserId, 'skip_reviewer');
window.switchRoutingRulesSubTab = (tabId) => {
    currentRoutingSubTab = tabId === 'skip-reviewer' ? 'skip-reviewer' : 'normal';
    if (currentTab === 'routing-rules') renderRoutingRules();
};
window.switchSettingsSubTab = (tabId) => {
    currentSettingsSubTab = SETTINGS_DROPDOWN_TABS.some((item) => item.id === String(tabId || '').trim()) ? String(tabId).trim() : 'system';
    activeSettingsDropdownTab = activeSettingsDropdownTab === currentSettingsSubTab ? '' : currentSettingsSubTab;
    closeSettingsSectionModal();
    if (currentTab === 'settings') renderSettingsSubTabState();
};

window.addAgentBankOption = async () => {
    const input = document.getElementById('agentBankNameInput');
    const bankName = String(input?.value || '').trim();
    if (!bankName) return showNotification('Enter a bank name', 'warning');

    const exists = currentAgentBankOptions.some((bank) => String(bank.name || '').trim().toLowerCase() === bankName.toLowerCase());
    if (exists) return showNotification('That bank already exists', 'warning');

    try {
        currentAgentBankOptions = normalizeAgentBankOptions([
            ...currentAgentBankOptions,
            { name: bankName, active: true }
        ], getDefaultSystemSettings().agentBankOptions);
        await persistAgentBankOptions('agent_bank_added');
        if (input) input.value = '';
        showNotification('Bank added successfully', 'success');
    } catch (_) {
        showNotification('Failed to add bank', 'error');
    }
};

window.addPfaOption = () => {
    const input = document.getElementById('pfaNameInput');
    const name = String(input?.value || '').trim();
    if (!name) return showNotification('Enter a PFA name', 'warning');
    if (currentPfaOptions.some((item) => item.toLowerCase() === name.toLowerCase())) {
        return showNotification('That PFA already exists', 'warning');
    }
    currentPfaOptions = [...currentPfaOptions, name].sort((a, b) => a.localeCompare(b));
    if (input) input.value = '';
    renderPfaManagement();
};

window.removePfaOption = (encodedName) => {
    const name = decodeURIComponent(String(encodedName || ''));
    currentPfaOptions = currentPfaOptions.filter((item) => item !== name);
    renderPfaManagement();
};

window.updateDocumentRequirementField = (index, field, value) => {
    currentDocumentRequirements = currentDocumentRequirements.map((doc, idx) => {
        if (idx !== index) return doc;
        return {
            ...doc,
            [field]: value === 'true'
        };
    });
    renderDocumentRequirementsManager();
};

window.updateDocumentRequirementName = (index, value) => {
    currentDocumentRequirements = currentDocumentRequirements.map((doc, idx) => {
        if (idx !== index) return doc;
        return {
            ...doc,
            name: String(value || '')
        };
    });
};

window.updateRolePermission = (key, value) => {
    currentRolePermissions = {
        ...currentRolePermissions,
        [key]: value === 'true'
    };
};

window.updateDocumentRequirementRole = (key, value) => {
    currentDocumentRequirementRoles = {
        ...currentDocumentRequirementRoles,
        [key]: value === 'true'
    };
};

window.updateNotificationTemplate = (key, value) => {
    currentNotificationTemplates = {
        ...currentNotificationTemplates,
        [key]: String(value || '')
    };
};

window.updateHouseNumberRule = (index, field, value) => {
    currentHouseNumberRules = currentHouseNumberRules.map((rule, idx) => {
        if (idx !== index) return rule;
        return {
            ...rule,
            [field]: field === 'startNumber' ? Number(value || 0) : String(value || '')
        };
    });
};

window.addHouseNumberRule = () => {
    currentHouseNumberRules = [
        ...currentHouseNumberRules,
        {
            propertyType: '',
            mode: 'alpha_suffix',
            prefix: '',
            startNumber: 0,
            startLetter: '',
            startPrefix: ''
        }
    ];
    renderHouseNumberRulesManager();
};

window.removeHouseNumberRule = (index) => {
    currentHouseNumberRules = currentHouseNumberRules.filter((_, idx) => idx !== index);
    renderHouseNumberRulesManager();
};

window.removeDocumentRequirement = (index) => {
    currentDocumentRequirements = currentDocumentRequirements.filter((_, idx) => idx !== index);
    renderDocumentRequirementsManager();
};

window.updateWorkflowLabel = (key, value) => {
    currentWorkflowLabels = {
        ...currentWorkflowLabels,
        [key]: String(value || '')
    };
};

window.addPropertyRule = () => {
    currentPropertyRules = [
        ...currentPropertyRules,
        { name: '', min: 0, max: 0, value: 0, fee: 0 }
    ];
    renderPropertyRulesManager();
};

window.updatePropertyRule = (index, field, value) => {
    currentPropertyRules = currentPropertyRules.map((rule, idx) => {
        if (idx !== index) return rule;
        return {
            ...rule,
            [field]: ['name'].includes(field) ? String(value || '') : Number(value || 0)
        };
    });
};

window.removePropertyRule = (index) => {
    currentPropertyRules = currentPropertyRules.filter((_, idx) => idx !== index);
    renderPropertyRulesManager();
};

window.toggleAgentBankOption = async (encodedName) => {
    const bankName = decodeURIComponent(String(encodedName || ''));
    const current = currentAgentBankOptions.find((bank) => bank.name === bankName);
    if (!current) return showNotification('Bank not found', 'error');

    try {
        currentAgentBankOptions = currentAgentBankOptions.map((bank) => (
            bank.name === bankName ? { ...bank, active: !bank.active } : bank
        ));
        await persistAgentBankOptions(current.active ? 'agent_bank_deactivated' : 'agent_bank_activated');
        showNotification(`Bank ${current.active ? 'deactivated' : 'activated'} successfully`, 'success');
    } catch (_) {
        showNotification('Failed to update bank status', 'error');
    }
};

window.removeAgentBankOption = async (encodedName) => {
    const bankName = decodeURIComponent(String(encodedName || ''));
    const confirmed = confirm(`Remove ${bankName} from the selectable bank list? Existing agent records will keep their saved bank name.`);
    if (!confirmed) return;

    try {
        currentAgentBankOptions = currentAgentBankOptions.filter((bank) => bank.name !== bankName);
        await persistAgentBankOptions('agent_bank_removed');
        showNotification('Bank removed successfully', 'success');
    } catch (_) {
        showNotification('Failed to remove bank', 'error');
    }
};

function closeSuperUserModal() {
    document.getElementById('superUserModal')?.classList.remove('active');
}

function closeSuperAgentModal() {
    document.getElementById('superAgentModal')?.classList.remove('active');
}

function closeClearCacheConfirmModal() {
    document.getElementById('clearCacheConfirmModal')?.classList.remove('active');
}

function closeDocumentRequirementModal() {
    document.getElementById('documentRequirementModal')?.classList.remove('active');
    document.getElementById('documentRequirementForm')?.reset();
    const iconInput = document.getElementById('documentRequirementIcon');
    if (iconInput) iconInput.value = 'fa-file-alt';
    const idInput = document.getElementById('documentRequirementId');
    if (idInput) idInput.dataset.touched = 'false';
}

function openDocumentRequirementModal() {
    const modal = document.getElementById('documentRequirementModal');
    const iconInput = document.getElementById('documentRequirementIcon');
    if (iconInput && !String(iconInput.value || '').trim()) {
        iconInput.value = 'fa-file-alt';
    }
    modal?.classList.add('active');
}

function openClearCacheConfirmModal() {
    return new Promise((resolve) => {
        const modal = document.getElementById('clearCacheConfirmModal');
        const confirmBtn = document.getElementById('confirmClearCacheBtn');
        const cancelBtn = document.getElementById('cancelClearCacheConfirmBtn');
        const closeBtn = document.getElementById('closeClearCacheConfirmModalBtn');
        if (!modal || !confirmBtn || !cancelBtn || !closeBtn) {
            resolve(false);
            return;
        }

        const cleanup = () => {
            modal.removeEventListener('click', onBackdrop);
            confirmBtn.removeEventListener('click', onConfirm);
            cancelBtn.removeEventListener('click', onCancel);
            closeBtn.removeEventListener('click', onCancel);
        };

        const onConfirm = () => {
            cleanup();
            closeClearCacheConfirmModal();
            resolve(true);
        };

        const onCancel = () => {
            cleanup();
            closeClearCacheConfirmModal();
            resolve(false);
        };

        const onBackdrop = (event) => {
            if (event.target === modal) onCancel();
        };

        confirmBtn.addEventListener('click', onConfirm);
        cancelBtn.addEventListener('click', onCancel);
        closeBtn.addEventListener('click', onCancel);
        modal.addEventListener('click', onBackdrop);
        modal.classList.add('active');
    });
}

window.editSuperUser = async (userId) => {
    selectedSuperUserId = userId;
    const user = allUsers.find((u) => u.id === userId);
    if (!user) return showNotification('User not found', 'error');

    const whatsapp = splitWhatsApp(user);
    document.getElementById('superModalFullName').value = user.fullName || '';
    document.getElementById('superModalLocation').value = user.location || '';
    document.getElementById('superModalEmail').value = normalizeEmail(user.email);
    document.getElementById('superModalDepartment').value = user.department || '';
    document.getElementById('superModalWhatsappCode').value = whatsapp.code || '+234';
    document.getElementById('superModalWhatsappLocalNumber').value = whatsapp.local || '';
    const modalRole = String(user.role || 'uploader').toLowerCase();
    document.getElementById('superModalRole').value = modalRole === 'viewer' ? 'reviewer' : modalRole;
    document.getElementById('superModalStatus').value = String(user.status || 'active').toLowerCase();
    document.getElementById('superUserModal')?.classList.add('active');
};

window.editSuperAgent = async (agentId) => {
    selectedSuperAgentId = agentId;
    const agent = allAgents.find((item) => item.id === agentId);
    if (!agent) return showNotification('Agent not found', 'error');

    document.getElementById('superAgentFullName').value = agent.fullName || '';
    document.getElementById('superAgentContactNumber').value = agent.contactNumber || '';
    document.getElementById('superAgentAccountNumber').value = agent.accountNumber || '';
    document.getElementById('superAgentAccountBank').value = agent.accountBank || '';
    document.getElementById('superAgentStatus').value = String(agent.status || 'pending').toLowerCase();
    document.getElementById('superAgentRegisteredBy').value = getAgentRegisteredByLabel(agent);
    document.getElementById('superAgentModal')?.classList.add('active');
};

async function saveDocumentRequirementModal(event) {
    event.preventDefault();

    const saveBtn = document.getElementById('saveDocumentRequirementModalBtn');
    const originalHtml = saveBtn?.innerHTML || '';
    const name = String(document.getElementById('documentRequirementName')?.value || '').trim();
    const requestedId = String(document.getElementById('documentRequirementId')?.value || '').trim();
    const required = String(document.getElementById('documentRequirementRequired')?.value || 'true') === 'true';
    const active = String(document.getElementById('documentRequirementActive')?.value || 'true') === 'true';
    const icon = String(document.getElementById('documentRequirementIcon')?.value || 'fa-file-alt').trim() || 'fa-file-alt';

    if (!name) {
        showNotification('Document name is required', 'warning');
        return;
    }

    let documentId = slugifyDocumentRequirementId(requestedId || name);
    if (!documentId) {
        showNotification('Enter a valid document name or document ID', 'warning');
        return;
    }

    const existingIds = new Set(currentDocumentRequirements.map((doc) => String(doc.id || '').trim().toLowerCase()));
    let uniqueId = documentId;
    let counter = 2;
    while (existingIds.has(uniqueId.toLowerCase())) {
        uniqueId = `${documentId}_${counter}`;
        counter += 1;
    }

    try {
        if (saveBtn) {
            saveBtn.disabled = true;
            saveBtn.innerHTML = '<i class="fas fa-spinner fa-spin"></i> Saving...';
        }

        currentDocumentRequirements = [
            ...currentDocumentRequirements,
            {
                id: uniqueId,
                name,
                icon,
                required,
                active
            }
        ];
        renderDocumentRequirementsManager();
        closeDocumentRequirementModal();
        showNotification('Document box added. Click Save Settings to publish it everywhere.', 'success');
    } finally {
        if (saveBtn) {
            saveBtn.disabled = false;
            saveBtn.innerHTML = originalHtml;
        }
    }
}

async function saveSuperUser(e) {
    e.preventDefault();
    if (!selectedSuperUserId) return showNotification('No user selected', 'error');

    const saveBtn = document.getElementById('saveSuperUserModalBtn');
    const originalHtml = saveBtn?.innerHTML || '';
    const whatsappCodeRaw = String(document.getElementById('superModalWhatsappCode')?.value || '').trim();
    const whatsappLocalDigits = String(document.getElementById('superModalWhatsappLocalNumber')?.value || '').replace(/\D/g, '');
    const whatsappCode = whatsappCodeRaw ? (whatsappCodeRaw.startsWith('+') ? whatsappCodeRaw : `+${whatsappCodeRaw}`) : '';
    const whatsappNumber = whatsappCode && whatsappLocalDigits ? `${whatsappCode}${whatsappLocalDigits}` : '';

    if (whatsappCode && !/^\+\d{1,4}$/.test(whatsappCode)) {
        showNotification('WhatsApp country code is invalid', 'error');
        return;
    }
    if (whatsappLocalDigits && !/^\d{10}$/.test(whatsappLocalDigits)) {
        showNotification('WhatsApp number must be exactly 10 digits', 'error');
        return;
    }

    const userData = {
        fullName: String(document.getElementById('superModalFullName')?.value || '').trim(),
        location: String(document.getElementById('superModalLocation')?.value || '').trim(),
        email: normalizeEmail(document.getElementById('superModalEmail')?.value),
        department: String(document.getElementById('superModalDepartment')?.value || '').trim(),
        whatsappCode,
        whatsappLocalNumber: whatsappLocalDigits,
        whatsappNumber,
        phone: whatsappNumber || whatsappLocalDigits || '',
        role: String(document.getElementById('superModalRole')?.value || 'uploader').trim(),
        status: String(document.getElementById('superModalStatus')?.value || 'active').trim(),
        updatedAt: serverTimestamp(),
        updatedBy: currentUser?.email || ''
    };

    if (!userData.fullName || !userData.email) {
        showNotification('Full name and email are required', 'error');
        return;
    }

    try {
        if (saveBtn) {
            saveBtn.disabled = true;
            saveBtn.innerHTML = '<i class="fas fa-spinner fa-spin"></i> Saving...';
        }

        await updateDoc(doc(db, 'users', selectedSuperUserId), userData);
        await addDoc(collection(db, 'audit'), {
            action: 'user_updated',
            userId: selectedSuperUserId,
            userEmail: userData.email,
            userFullName: userData.fullName,
            performedBy: currentUser?.email || '',
            timestamp: serverTimestamp()
        });
        closeSuperUserModal();
        showNotification('User details updated successfully', 'success');
    } catch (error) {
        showNotification('Failed to update user details', 'error');
    } finally {
        if (saveBtn) {
            saveBtn.disabled = false;
            saveBtn.innerHTML = originalHtml;
        }
    }
}

async function saveSuperAgent(e) {
    e.preventDefault();
    if (!selectedSuperAgentId) return showNotification('No agent selected', 'error');

    const saveBtn = document.getElementById('saveSuperAgentModalBtn');
    const originalHtml = saveBtn?.innerHTML || '';
    const fullName = String(document.getElementById('superAgentFullName')?.value || '').trim();
    const contactNumber = String(document.getElementById('superAgentContactNumber')?.value || '').trim();
    const accountNumber = String(document.getElementById('superAgentAccountNumber')?.value || '').trim();
    const accountBank = String(document.getElementById('superAgentAccountBank')?.value || '').trim();
    const status = String(document.getElementById('superAgentStatus')?.value || 'pending').trim().toLowerCase();

    if (!fullName || !contactNumber || !accountNumber || !accountBank) {
        return showNotification('All agent fields are required', 'warning');
    }

    const existingAgent = allAgents.find((item) => item.id === selectedSuperAgentId);
    if (!existingAgent) return showNotification('Agent not found', 'error');

    try {
        if (saveBtn) {
            saveBtn.disabled = true;
            saveBtn.innerHTML = '<i class="fas fa-spinner fa-spin"></i> Saving...';
        }

        const agentPayload = {
            fullName,
            contactNumber,
            accountNumber,
            accountBank,
            status,
            updatedAt: serverTimestamp(),
            updatedBy: currentUser?.email || ''
        };

        await updateDoc(doc(db, 'agents', selectedSuperAgentId), agentPayload);

        const systemSettings = await getSystemSettings(db, { force: true });
        const linkedSubmissions = allSubmissions.filter((sub) => String(sub.agentId || '').trim() === selectedSuperAgentId);
        if (linkedSubmissions.length && systemSettings.agentEditSyncEnabled) {
            await Promise.all(linkedSubmissions.map((sub) => updateDoc(doc(db, 'submissions', sub.id), {
                agentName: fullName,
                agentContactNumber: contactNumber,
                agentAccountNumber: accountNumber,
                agentAccountBank: accountBank,
                updatedAt: serverTimestamp(),
                updatedBy: currentUser?.email || ''
            })));
        }

        await addDoc(collection(db, 'audit'), {
            action: 'agent_record_updated_by_super_admin',
            agentId: selectedSuperAgentId,
            agentName: fullName,
            previousAgentName: existingAgent.fullName || '',
            previousContactNumber: existingAgent.contactNumber || '',
            previousAccountNumber: existingAgent.accountNumber || '',
            previousAccountBank: existingAgent.accountBank || '',
            newContactNumber: contactNumber,
            newAccountNumber: accountNumber,
            newAccountBank: accountBank,
            newStatus: status,
            linkedSubmissionCount: linkedSubmissions.length,
            agentEditSyncEnabled: systemSettings.agentEditSyncEnabled,
            performedBy: currentUser?.email || '',
            timestamp: serverTimestamp()
        });

        closeSuperAgentModal();
        showNotification('Agent record updated successfully', 'success');
    } catch (error) {
        showNotification('Failed to update agent record', 'error');
    } finally {
        if (saveBtn) {
            saveBtn.disabled = false;
            saveBtn.innerHTML = originalHtml;
        }
    }
}

window.saveAdminManager = async (userId) => {
    const roleEl = document.getElementById(`sa-role-${userId}`);
    const statusEl = document.getElementById(`sa-status-${userId}`);
    const role = String(roleEl?.value || '').trim();
    const status = String(statusEl?.value || '').trim();
    if (!role) return showNotification('Role is required', 'error');

    try {
        try {
            await updateDoc(doc(db, 'users', userId), {
                role,
                status,
                updatedAt: serverTimestamp(),
                updatedBy: currentUser?.email || ''
            });
        } catch (error) {
            if (error?.code !== 'permission-denied') throw error;
            await ensureCurrentSuperAdminWritableProfile();
            await updateDoc(doc(db, 'users', userId), {
                role,
                status,
                updatedAt: serverTimestamp(),
                updatedBy: currentUser?.email || ''
            });
        }
        await addDoc(collection(db, 'audit'), {
            action: 'user_management_updated',
            userId,
            newRole: role,
            newStatus: status,
            performedBy: currentUser?.email || '',
            timestamp: serverTimestamp()
        });
        showNotification('User updated successfully', 'success');
    } catch (error) {
        showNotification('Failed to update user', 'error');
    }
};

window.redirectSubmission = async (submissionId) => {
    const reviewer = String(document.getElementById(`rd-reviewer-${submissionId}`)?.value || '').trim().toLowerCase();
    const rsa = String(document.getElementById(`rd-rsa-${submissionId}`)?.value || '').trim().toLowerCase();
    const payment = String(document.getElementById(`rd-payment-${submissionId}`)?.value || '').trim().toLowerCase();
    const reason = String(document.getElementById(`rd-reason-${submissionId}`)?.value || '').trim();
    if (!reason) return showNotification('Reason is required for redirect', 'warning');

    const prev = allSubmissions.find((s) => s.id === submissionId) || {};
    try {
        await updateDoc(doc(db, 'submissions', submissionId), {
            assignedTo: reviewer || '',
            assignedToRSA: rsa || '',
            assignedToPayment: payment || '',
            reassignedAt: serverTimestamp(),
            reassignedBy: currentUser?.email || '',
            reassignmentReason: reason
        });
        await addDoc(collection(db, 'audit'), {
            action: 'application_redirected',
            submissionId,
            customerName: prev.customerName || '',
            oldReviewer: prev.assignedTo || '',
            oldRSA: prev.assignedToRSA || '',
            oldPayment: prev.assignedToPayment || '',
            newReviewer: reviewer,
            newRSA: rsa,
            newPayment: payment,
            reason,
            performedBy: currentUser?.email || '',
            timestamp: serverTimestamp()
        });
        showNotification('Application redirected successfully', 'success');
    } catch (error) {
        showNotification('Failed to redirect application', 'error');
    }
};

window.resetRoundRobinCounter = async (counterDocId) => {
    const confirmed = confirm(`Reset ${counterDocId} counter?`);
    if (!confirmed) return;
    try {
        await setDoc(doc(db, 'counters', counterDocId), {
            lastIndex: -1,
            lastDate: '',
            updatedAt: serverTimestamp(),
            updatedBy: currentUser?.email || ''
        }, { merge: true });
        await addDoc(collection(db, 'audit'), {
            action: 'round_robin_counter_reset',
            counterDocId,
            performedBy: currentUser?.email || '',
            timestamp: serverTimestamp()
        });
        showNotification(`${counterDocId} reset successful`, 'success');
        renderRoundRobin();
    } catch (error) {
        showNotification('Failed to reset counter', 'error');
    }
};

window.saveSuperSettings = async (triggerButton = null) => {
    const saveBtn = triggerButton || document.getElementById('saveSettingsBtn');
    const originalHtml = saveBtn?.innerHTML || '';
    const maintenanceMode = String(document.getElementById('settingMaintenanceMode')?.value || 'false') === 'true';
    const maintenanceMessage = String(document.getElementById('settingMaintenanceMessage')?.value || '').trim() || getDefaultSystemSettings().maintenanceMessage;
    const commissionRatePercent = Number(document.getElementById('settingCommissionRate')?.value || 0);
    const commissionEffectiveFromValue = String(document.getElementById('settingCommissionEffectiveFrom')?.value || '').trim() || '2026-05-07';
    const maxImageUploadMb = Number(document.getElementById('settingMaxImageUploadMb')?.value || getDefaultSystemSettings().maxImageUploadMb);
    const maxPdfUploadMb = Number(document.getElementById('settingMaxPdfUploadMb')?.value || getDefaultSystemSettings().maxPdfUploadMb);
    const reviewerRoundRobinEnabled = String(document.getElementById('settingReviewerRoundRobinEnabled')?.value || 'true') === 'true';
    const rsaRoundRobinEnabled = String(document.getElementById('settingRsaRoundRobinEnabled')?.value || 'true') === 'true';
    const paymentRoundRobinEnabled = String(document.getElementById('settingPaymentRoundRobinEnabled')?.value || 'true') === 'true';
    const agentEditSyncEnabled = String(document.getElementById('settingAgentEditSyncEnabled')?.value || 'true') === 'true';
    const notificationsEmailEnabled = String(document.getElementById('settingNotificationsEmailEnabled')?.value || 'true') === 'true';
    const notificationsPushEnabled = String(document.getElementById('settingNotificationsPushEnabled')?.value || 'true') === 'true';
    const announcementEnabled = String(document.getElementById('settingAnnouncementEnabled')?.value || 'false') === 'true';
    const announcementTone = String(document.getElementById('settingAnnouncementTone')?.value || 'info').trim() || 'info';
    const announcementMessage = String(document.getElementById('settingAnnouncementMessage')?.value || '').trim();
    const defaultRouteMode = String(document.getElementById('settingDefaultRouteMode')?.value || 'normal').trim() || 'normal';
    const fallbackAssignmentMode = String(document.getElementById('settingFallbackAssignmentMode')?.value || 'round_robin').trim() || 'round_robin';
    const agentRegistrationApprovalRequired = String(document.getElementById('settingAgentRegistrationApprovalRequired')?.value || 'true') === 'true';
    const rejectionMinLength = Number(document.getElementById('settingRejectionMinLength')?.value || 0);
    const reviewerRejectRequired = String(document.getElementById('settingReviewerRejectRequired')?.value || 'true') === 'true';
    const rsaRejectRequired = String(document.getElementById('settingRsaRejectRequired')?.value || 'true') === 'true';
    const sessionTimeoutMinutes = Number(document.getElementById('settingSessionTimeoutMinutes')?.value || getDefaultSystemSettings().securityControls.sessionTimeoutMinutes);
    const auditRetentionDays = Number(document.getElementById('settingAuditRetentionDays')?.value || getDefaultSystemSettings().auditControls.retentionDays);
    if (!Number.isFinite(commissionRatePercent) || commissionRatePercent < 0) {
        showNotification('Enter a valid commission rate percentage', 'warning');
        return false;
    }
    if (!Number.isFinite(maxImageUploadMb) || maxImageUploadMb <= 0 || !Number.isFinite(maxPdfUploadMb) || maxPdfUploadMb <= 0) {
        showNotification('Upload size limits must be greater than 0', 'warning');
        return false;
    }
    if (!Number.isFinite(sessionTimeoutMinutes) || sessionTimeoutMinutes <= 0) {
        showNotification('Session timeout must be greater than 0', 'warning');
        return false;
    }
    if (!Number.isFinite(auditRetentionDays) || auditRetentionDays <= 0) {
        showNotification('Audit retention days must be greater than 0', 'warning');
        return false;
    }
    const commissionRate = commissionRatePercent / 100;
    const commissionRateEffectiveFrom = `${commissionEffectiveFromValue}T00:00:00+01:00`;
    let workflowLabels;
    let documentRequirementRoles = { ...getDefaultSystemSettings().documentRequirementRoles, ...currentDocumentRequirementRoles };
    let rolePermissions = { ...getDefaultSystemSettings().rolePermissions, ...currentRolePermissions };
    let routingPolicies = { ...getDefaultSystemSettings().routingPolicies, ...currentRoutingPolicies };
    let notificationTemplates = { ...currentNotificationTemplates };
    let houseNumberRules = {};
    workflowLabels = { ...currentWorkflowLabels };
    if (!Object.keys(workflowLabels).length) workflowLabels = { ...getDefaultSystemSettings().workflowLabels };
    const pfaOptions = [...currentPfaOptions];
    const seenDocumentIds = new Set();
    const documentRequirements = currentDocumentRequirements
        .map((doc, index) => {
            const name = String(doc.name || '').trim();
            const generatedId = slugifyDocumentRequirementId(doc.id || name || `document_${index + 1}`);
            let uniqueId = generatedId || `document_${index + 1}`;
            let counter = 2;
            while (seenDocumentIds.has(uniqueId)) {
                uniqueId = `${generatedId || `document_${index + 1}`}_${counter}`;
                counter += 1;
            }
            seenDocumentIds.add(uniqueId);
            return {
                ...doc,
                id: uniqueId,
                name
            };
        })
        .filter((doc) => doc.name);
    const bulkImportRequiredColumns = parseLinesTextarea('settingBulkImportRequiredColumns');
    const propertyRules = currentPropertyRules
        .map((rule) => ({
            name: String(rule.name || '').trim(),
            min: Number(rule.min || 0),
            max: Number(rule.max || 0),
            value: Number(rule.value || 0),
            fee: Number(rule.fee || 0)
        }))
        .filter((rule) => rule.name && rule.max >= rule.min);
    const normalizedNotificationTemplates = Object.fromEntries(
        Object.entries(notificationTemplates)
            .map(([key, value]) => [key, String(value || '').trim()])
            .filter(([, value]) => value)
    );
    const normalizedHouseNumberRules = {};
    currentHouseNumberRules
        .map((rule) => ({
            propertyType: String(rule.propertyType || '').trim(),
            mode: String(rule.mode || 'alpha_suffix').trim() || 'alpha_suffix',
            prefix: String(rule.prefix || '').trim(),
            startNumber: Number(rule.startNumber || 0),
            startLetter: String(rule.startLetter || '').trim(),
            startPrefix: String(rule.startPrefix || '').trim()
        }))
        .filter((rule) => rule.propertyType)
        .forEach((rule) => {
            normalizedHouseNumberRules[rule.propertyType] = {
                mode: rule.mode,
                prefix: rule.prefix,
                startNumber: rule.startNumber,
                startLetter: rule.startLetter,
                startPrefix: rule.startPrefix
            };
        });
    rolePermissions = Object.fromEntries(
        ROLE_PERMISSION_FIELDS.map(({ key }) => [key, rolePermissions[key] !== false])
    );
    documentRequirementRoles = Object.fromEntries(
        DOCUMENT_REQUIREMENT_ROLE_FIELDS.map(({ key }) => [key, documentRequirementRoles[key] !== false])
    );
    routingPolicies = {
        ...routingPolicies,
        defaultRouteMode,
        fallbackAssignmentMode
    };
    houseNumberRules = normalizedHouseNumberRules;
    if (!documentRequirements.length) {
        showNotification('Add at least one document box before saving settings.', 'warning');
        return false;
    }
    try {
        if (saveBtn) {
            saveBtn.disabled = true;
            saveBtn.classList.add('loading');
            saveBtn.innerHTML = '<i class="fas fa-spinner fa-spin"></i> Saving Settings...';
        }
        const existingSnap = await getDoc(doc(db, 'settings', 'system'));
        const existing = existingSnap.exists() ? (existingSnap.data() || {}) : {};
        await setDoc(doc(db, 'settings', 'system'), {
            maintenanceMode,
            maintenanceMessage,
            dashboardAnnouncement: {
                enabled: announcementEnabled,
                tone: announcementTone,
                message: announcementMessage
            },
            commissionRate,
            commissionRateEffectiveFrom,
            maxImageUploadMb,
            maxPdfUploadMb,
            reviewerRoundRobinEnabled,
            rsaRoundRobinEnabled,
            paymentRoundRobinEnabled,
            agentEditSyncEnabled,
            notificationsEmailEnabled,
            notificationsPushEnabled,
            pfaOptions,
            documentRequirements,
            documentRequirementRoles,
            rolePermissions,
            routingPolicies,
            workflowLabels,
            rejectionRules: {
                reviewerRequired: reviewerRejectRequired,
                rsaRequired: rsaRejectRequired,
                minLength: Math.max(0, rejectionMinLength)
            },
            agentRegistrationRules: {
                approvalRequired: agentRegistrationApprovalRequired
            },
            bulkImportRules: {
                requiredColumns: bulkImportRequiredColumns
            },
            securityControls: {
                sessionTimeoutMinutes,
                forceLogoutToken: String(existing?.securityControls?.forceLogoutToken || '').trim()
            },
            notificationTemplates: normalizedNotificationTemplates,
            auditControls: {
                retentionDays: auditRetentionDays
            },
            propertyRules,
            houseNumberRules: normalizedHouseNumberRules,
            updatedAt: serverTimestamp(),
            updatedBy: currentUser?.email || ''
        }, { merge: true });
        clearCommissionSettingsCache();
        clearSystemSettingsCache();
        const effectiveFromMs = new Date(commissionRateEffectiveFrom).getTime();
        const submissionsToBackfill = (allSubmissions?.length ? allSubmissions : (await getDocs(collection(db, 'submissions'))).docs.map((snap) => ({ id: snap.id, ...(snap.data() || {}) })))
            .filter((sub) => String(sub?.status || '').toLowerCase() !== 'draft')
            .filter((sub) => !Number.isFinite(Number(sub?.commissionRate)));
        if (submissionsToBackfill.length) {
            const settingsForBackfill = {
                rate: commissionRate,
                effectiveFromIso: commissionRateEffectiveFrom,
                effectiveFromMs
            };
            await Promise.all(submissionsToBackfill.map((sub) => {
                const resolvedRate = resolveSubmissionCommissionRate(sub, settingsForBackfill);
                return updateDoc(doc(db, 'submissions', sub.id), {
                    commissionRate: resolvedRate,
                    commissionRatePercent: Number((resolvedRate * 100).toFixed(4)),
                    commissionRateLabel: formatCommissionRateLabel(resolvedRate),
                    commissionRateEffectiveFrom,
                    commissionRateAssignedAtIso: new Date().toISOString(),
                    updatedAt: serverTimestamp(),
                    updatedBy: currentUser?.email || ''
                });
            }));
        }
        await addDoc(collection(db, 'audit'), {
            action: 'super_admin_settings_updated',
            previousCommissionRate: Number(existing.commissionRate ?? getDefaultCommissionSettings().rate),
            newCommissionRate: commissionRate,
            previousCommissionRateEffectiveFrom: String(existing.commissionRateEffectiveFrom || getDefaultCommissionSettings().effectiveFromIso),
            newCommissionRateEffectiveFrom: commissionRateEffectiveFrom,
            commissionBackfillCount: submissionsToBackfill.length,
            maintenanceMode,
            maintenanceMessage,
            maxImageUploadMb,
            maxPdfUploadMb,
            reviewerRoundRobinEnabled,
            rsaRoundRobinEnabled,
            paymentRoundRobinEnabled,
            agentEditSyncEnabled,
            notificationsEmailEnabled,
            notificationsPushEnabled,
            announcementEnabled,
            defaultRouteMode,
            agentRegistrationApprovalRequired,
            rejectionMinLength,
            pfaCount: pfaOptions.length,
            documentRequirementCount: Array.isArray(documentRequirements) ? documentRequirements.length : 0,
            documentRequirementRoles,
            bulkImportColumnCount: bulkImportRequiredColumns.length,
            sessionTimeoutMinutes,
            auditRetentionDays,
            performedBy: currentUser?.email || '',
            timestamp: serverTimestamp()
        });
        showNotification('System settings saved successfully.', 'success');
        return true;
    } catch (error) {
        showNotification('Failed to save settings', 'error');
        return false;
    } finally {
        if (saveBtn) {
            saveBtn.disabled = false;
            saveBtn.classList.remove('loading');
            saveBtn.innerHTML = originalHtml;
        }
    }
};

window.assignApplicationAgent = async (submissionId) => {
    const submission = allSubmissions.find((item) => item.id === submissionId);
    if (!submission) return showNotification('Submission not found', 'error');

    const select = document.getElementById(`sa-app-agent-${submissionId}`);
    const agentId = String(select?.value || '').trim();
    if (!agentId) return showNotification('Select an agent first', 'warning');

    const allowedAgents = getUploaderOwnedApprovedAgents(submission.uploadedBy);
    const selectedAgent = allowedAgents.find((agent) => agent.id === agentId);
    if (!selectedAgent) {
        showNotification('Selected agent is not registered by this uploader', 'error');
        return;
    }

    try {
        await updateDoc(doc(db, 'submissions', submissionId), {
            agentId: selectedAgent.id,
            agentName: selectedAgent.fullName || '',
            agentContactNumber: selectedAgent.contactNumber || '',
            agentAccountNumber: selectedAgent.accountNumber || '',
            agentAccountBank: selectedAgent.accountBank || '',
            updatedAt: serverTimestamp(),
            updatedBy: currentUser?.email || ''
        });
        await addDoc(collection(db, 'audit'), {
            action: 'submission_agent_attached_by_super_admin',
            submissionId,
            customerName: submission.customerName || '',
            uploaderEmail: submission.uploadedBy || '',
            agentId: selectedAgent.id,
            agentName: selectedAgent.fullName || '',
            performedBy: currentUser?.email || '',
            timestamp: serverTimestamp()
        });
        showNotification('Application agent updated successfully', 'success');
    } catch (error) {
        showNotification('Failed to update application agent', 'error');
    }
};

window.rerouteApplicationAgent = async (submissionId) => {
    const submission = allSubmissions.find((item) => item.id === submissionId);
    if (!submission) return showNotification('Submission not found', 'error');

    const currentAgentId = String(submission.agentId || '').trim();
    if (!currentAgentId) return showNotification('This application has no linked agent to re-route', 'warning');

    const select = document.getElementById(`sa-reroute-agent-${submissionId}`);
    const agentId = String(select?.value || '').trim();
    if (!agentId) return showNotification('Select a replacement agent first', 'warning');
    if (agentId === currentAgentId) return showNotification('Select a different agent to re-route this application', 'warning');

    const allowedAgents = getUploaderOwnedApprovedAgents(submission.uploadedBy);
    const selectedAgent = allowedAgents.find((agent) => String(agent.id || '').trim() === agentId);
    if (!selectedAgent) {
        showNotification('Selected agent is not registered by this uploader', 'error');
        return;
    }

    const currentAgentName = getApplicationLinkedAgentLabel(submission);
    const nextAgentName = String(selectedAgent.fullName || '').trim() || 'Selected Agent';
    const confirmed = window.confirm(`Re-route ${submission.customerName || 'this application'} from ${currentAgentName} to ${nextAgentName}?`);
    if (!confirmed) return;

    try {
        await updateDoc(doc(db, 'submissions', submissionId), {
            agentId: selectedAgent.id,
            agentName: selectedAgent.fullName || '',
            agentContactNumber: selectedAgent.contactNumber || '',
            agentAccountNumber: selectedAgent.accountNumber || '',
            agentAccountBank: selectedAgent.accountBank || '',
            updatedAt: serverTimestamp(),
            updatedBy: currentUser?.email || ''
        });
        await addDoc(collection(db, 'audit'), {
            action: 'submission_agent_rerouted_by_super_admin',
            submissionId,
            customerName: submission.customerName || '',
            uploaderEmail: submission.uploadedBy || '',
            previousAgentId: currentAgentId,
            previousAgentName: currentAgentName,
            nextAgentId: selectedAgent.id,
            nextAgentName: selectedAgent.fullName || '',
            performedBy: currentUser?.email || '',
            timestamp: serverTimestamp()
        });
        showNotification('Application re-routed successfully', 'success');
    } catch (error) {
        showNotification('Failed to re-route application agent', 'error');
    }
};

window.clearAppCacheForAllUsers = async () => {
    const confirmed = await openClearCacheConfirmModal();
    if (!confirmed) return;

    const cacheClearToken = new Date().toISOString();
    const clearBtn = document.getElementById('clearAppCacheBtn');
    const originalHtml = clearBtn?.innerHTML || '';
    try {
        if (clearBtn) {
            clearBtn.disabled = true;
            clearBtn.classList.add('loading');
            clearBtn.innerHTML = '<i class="fas fa-spinner fa-spin"></i> Clearing Cache...';
        }

        await setDoc(doc(db, 'settings', 'system'), {
            cacheClearToken,
            cacheClearRequestedAt: serverTimestamp(),
            updatedAt: serverTimestamp(),
            updatedBy: currentUser?.email || ''
        }, { merge: true });
        clearSystemSettingsCache();

        await addDoc(collection(db, 'audit'), {
            action: 'app_cache_clear_requested',
            cacheClearToken,
            performedBy: currentUser?.email || '',
            timestamp: serverTimestamp()
        });

        showNotification('Cache clear sent. Open dashboards will refresh shortly.', 'success');
    } catch (error) {
        showNotification('Failed to trigger cache clear', 'error');
    } finally {
        if (clearBtn) {
            clearBtn.disabled = false;
            clearBtn.classList.remove('loading');
            clearBtn.innerHTML = originalHtml;
        }
    }
};

window.forceLogoutAllUsers = async () => {
    const confirmed = confirm('Force logout all signed-in users? They will need to sign in again.');
    if (!confirmed) return;

    try {
        const securityControls = {
            ...(await getSystemSettings(db, { force: true })).securityControls,
            forceLogoutToken: new Date().toISOString()
        };
        await setDoc(doc(db, 'settings', 'system'), {
            securityControls,
            updatedAt: serverTimestamp(),
            updatedBy: currentUser?.email || ''
        }, { merge: true });
        clearSystemSettingsCache();
        await addDoc(collection(db, 'audit'), {
            action: 'force_logout_all_users',
            performedBy: currentUser?.email || '',
            timestamp: serverTimestamp()
        });
        showNotification('Force logout signal sent successfully', 'success');
    } catch (_) {
        showNotification('Failed to send force logout signal', 'error');
    }
};

window.exportAuditCsv = () => {
    try {
        const rows = [
            ['Timestamp', 'Action', 'Performed By', 'Details'],
            ...allAudits.map((entry) => [
                formatDate(entry.timestamp),
                String(entry.action || ''),
                String(entry.performedBy || entry.userEmail || ''),
                JSON.stringify(entry)
            ])
        ];
        const csv = rows.map((row) => row.map((cell) => `"${String(cell ?? '').replace(/"/g, '""')}"`).join(',')).join('\n');
        const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
        const url = URL.createObjectURL(blob);
        const link = document.createElement('a');
        link.href = url;
        link.download = `cmbank_audit_export_${Date.now()}.csv`;
        link.click();
        URL.revokeObjectURL(url);
        showNotification('Audit CSV exported', 'success');
    } catch (_) {
        showNotification('Failed to export audit CSV', 'error');
    }
};

window.signOutUser = async () => {
    try { await signOut(auth); } catch (_) {}
    window.location.href = 'index.html';
};

function setupRealtimeData() {
    onSnapshot(query(collection(db, 'users')), (snap) => {
        allUsers = snap.docs.map((d) => ({ id: d.id, ...d.data() }));
        updateNavigationCounts();
        renderCurrentTab();
        if (currentTab !== 'global') renderGlobalView();
    });

    onSnapshot(query(collection(db, 'submissions'), orderBy('uploadedAt', 'desc')), (snap) => {
        allSubmissions = snap.docs.map((d) => ({ id: d.id, ...d.data() }));
        updateNavigationCounts();
        renderCurrentTab();
        if (currentTab !== 'global') renderGlobalView();
    });

    onSnapshot(query(collection(db, 'audit'), orderBy('timestamp', 'desc'), limit(200)), (snap) => {
        allAudits = snap.docs.map((d) => ({ id: d.id, ...d.data() }));
        updateNavigationCounts();
        if (currentTab === 'audit') renderAudit();
    });

    onSnapshot(query(collection(db, 'uploaderRoutingRules')), (snap) => {
        allRoutingRules = snap.docs.map((d) => ({ id: d.id, ...d.data() }));
        updateNavigationCounts();
        if (currentTab === 'routing-rules') renderRoutingRules();
    });

    onSnapshot(query(collection(db, 'agents')), (snap) => {
        allAgents = snap.docs.map((d) => ({ id: d.id, ...d.data() }));
        if (currentTab === 'agents') renderCurrentTab();
    });
}

document.querySelectorAll('.nav-item[data-tab]').forEach((item) => {
    item.addEventListener('click', (e) => {
        e.preventDefault();
        switchTab(item.dataset.tab);
    });
});

document.getElementById('superUserSearch')?.addEventListener('input', renderAdminManagement);
document.getElementById('superUserRoleFilter')?.addEventListener('change', renderAdminManagement);
document.getElementById('superUserStatusFilter')?.addEventListener('change', renderAdminManagement);
document.getElementById('routingRulesSearch')?.addEventListener('input', () => {
    if (currentTab === 'routing-rules') renderRoutingRules();
});
document.getElementById('clearAppCacheBtn')?.addEventListener('click', () => {
    window.clearAppCacheForAllUsers?.();
});
document.getElementById('forceLogoutAllUsersBtn')?.addEventListener('click', () => {
    window.forceLogoutAllUsers?.();
});
document.getElementById('exportAuditCsvBtn')?.addEventListener('click', () => {
    window.exportAuditCsv?.();
});
document.getElementById('saveSettingsBtn')?.addEventListener('click', () => {
    window.saveSuperSettings?.();
});
document.getElementById('saveSettingsSectionModalBtn')?.addEventListener('click', async (event) => {
    const ok = await window.saveSuperSettings?.(event.currentTarget);
    if (ok) closeSettingsSectionModal();
});
document.getElementById('closeSettingsSectionModalBtn')?.addEventListener('click', closeSettingsSectionModal);
document.getElementById('cancelSettingsSectionModalBtn')?.addEventListener('click', closeSettingsSectionModal);
document.getElementById('settingsSectionModal')?.addEventListener('click', (event) => {
    if (event.target?.id === 'settingsSectionModal') {
        closeSettingsSectionModal();
    }
});
document.addEventListener('click', (event) => {
    if (!settingsDropdownNav) return;
    if (!settingsDropdownNav.contains(event.target)) {
        if (activeSettingsDropdownTab) {
            activeSettingsDropdownTab = '';
            renderSettingsSubTabState();
        }
    }
});
document.addEventListener('keydown', (event) => {
    if (event.key === 'Escape' && settingsSectionModal?.classList.contains('active')) {
        closeSettingsSectionModal();
    }
    if (event.key === 'Escape' && activeSettingsDropdownTab) {
        activeSettingsDropdownTab = '';
        renderSettingsSubTabState();
    }
});
document.getElementById('addAgentBankBtn')?.addEventListener('click', () => {
    window.addAgentBankOption?.();
});
document.getElementById('addPfaBtn')?.addEventListener('click', () => {
    window.addPfaOption?.();
});
document.getElementById('addDocumentRequirementBtn')?.addEventListener('click', () => {
    openDocumentRequirementModal();
});
document.getElementById('pfaNameInput')?.addEventListener('keydown', (event) => {
    if (event.key === 'Enter') {
        event.preventDefault();
        window.addPfaOption?.();
    }
});
document.getElementById('addPropertyRuleBtn')?.addEventListener('click', () => {
    window.addPropertyRule?.();
});
document.getElementById('addHouseNumberRuleBtn')?.addEventListener('click', () => {
    window.addHouseNumberRule?.();
});
document.getElementById('agentBankNameInput')?.addEventListener('keydown', (event) => {
    if (event.key === 'Enter') {
        event.preventDefault();
        window.addAgentBankOption?.();
    }
});
document.getElementById('superAgentSearch')?.addEventListener('input', () => {
    if (currentTab === 'agents' && currentAgentSubTab === 'agent-records') renderAgentRecords();
});
document.getElementById('superAgentRerouteSearch')?.addEventListener('input', () => {
    if (currentTab === 'agents' && currentAgentSubTab === 'agent-reroute') renderApplicationAgentRerouteModule();
});
document.getElementById('superAgentStatusFilter')?.addEventListener('change', () => {
    if (currentTab === 'agents' && currentAgentSubTab === 'agent-records') renderAgentRecords();
});
document.getElementById('superUserForm')?.addEventListener('submit', saveSuperUser);
document.getElementById('closeSuperUserModalBtn')?.addEventListener('click', closeSuperUserModal);
document.getElementById('cancelSuperUserModalBtn')?.addEventListener('click', closeSuperUserModal);
document.getElementById('documentRequirementForm')?.addEventListener('submit', saveDocumentRequirementModal);
document.getElementById('closeDocumentRequirementModalBtn')?.addEventListener('click', closeDocumentRequirementModal);
document.getElementById('cancelDocumentRequirementModalBtn')?.addEventListener('click', closeDocumentRequirementModal);
document.getElementById('documentRequirementModal')?.addEventListener('click', (event) => {
    if (event.target?.id === 'documentRequirementModal') {
        closeDocumentRequirementModal();
    }
});
document.getElementById('documentRequirementName')?.addEventListener('input', (event) => {
    const idInput = document.getElementById('documentRequirementId');
    if (!idInput) return;
    if (String(idInput.dataset.touched || '') === 'true') return;
    idInput.value = slugifyDocumentRequirementId(event.target?.value || '');
});
document.getElementById('documentRequirementId')?.addEventListener('input', (event) => {
    const target = event.target;
    if (!target) return;
    target.dataset.touched = String(Boolean(String(target.value || '').trim()));
});
document.getElementById('superAgentForm')?.addEventListener('submit', saveSuperAgent);
document.getElementById('closeSuperAgentModalBtn')?.addEventListener('click', closeSuperAgentModal);
document.getElementById('cancelSuperAgentModalBtn')?.addEventListener('click', closeSuperAgentModal);

document.getElementById('forceRefreshBtn')?.addEventListener('click', (e) => {
    e.preventDefault();
    const url = new URL(window.location.href);
    url.searchParams.set('_', Date.now().toString());
    window.location.replace(url.toString());
});

auth.onAuthStateChanged(async (user) => {
    if (!user) return void (window.location.href = 'index.html');
    currentUser = user;

    try {
        const sharedProfile = await getCurrentUserProfileShared(db, user);
        if (!sharedProfile) return void (window.location.href = 'index.html');

        let currentUserDoc = doc(db, 'users', String(sharedProfile.__docId || user.uid || '').trim());
        let currentUserSnap = await getDoc(currentUserDoc);
        if (!currentUserSnap.exists()) {
            const fallbackRef = doc(db, 'users', String(user.uid || '').trim());
            currentUserSnap = await getDoc(fallbackRef);
            if (currentUserSnap.exists()) currentUserDoc = fallbackRef;
        }
        if (!currentUserSnap.exists()) return void (window.location.href = 'index.html');

        currentUserDoc = await ensureCurrentUserProfileAtUid(user, currentUserSnap);
        currentUserData = currentUserDoc.data() || {};
        const role = String(currentUserData.role || '').toLowerCase();
        if (role !== 'super_admin') {
            window.location.href = roleHome(role);
            return;
        }

        if (superAdminName) {
            superAdminName.textContent = currentUserData.fullName || user.email || 'Super Admin';
        }

        setupRealtimeData();
        switchTab('global');
    } catch (error) {
        showNotification('Failed to validate session', 'error');
        window.location.href = 'index.html';
    }
});
